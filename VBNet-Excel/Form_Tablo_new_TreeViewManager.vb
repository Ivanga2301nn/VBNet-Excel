Imports System.Drawing
Imports System.Windows.Forms

Public Class Form_Tablo_new_TreeViewManager
    ' ========================================================================
    ' 📌 ОСНОВНИ ПОЛЕТА И СЪБИТИЯ
    ' ========================================================================
    ' TreeView контролът, който се управлява от класа
    Private ReadOnly _tv As TreeView
    ' Основният списък с данни за табла и токови кръгове
    Private ReadOnly _listTokow As List(Of Form_Tablo_new.strTokow)
    ''' <summary>
    ''' Събитие при избор на обект от TreeView.
    ''' Изпраща избрания запис към основната форма.
    ''' </summary>
    Public Event ObjectSelected(ByVal selectedItem As Form_Tablo_new.strTokow)
    ''' <summary>
    ''' Събитие при заявка за преместване чрез Drag & Drop.
    ''' Изпраща източника и целевия обект към бизнес логиката.
    ''' </summary>
    Public Event RequestMoveObject(ByVal source As Form_Tablo_new.strTokow,
                               ByVal target As Form_Tablo_new.strTokow)
    ' ========================================================================
    ' 🎨 UI КОНСТАНТИ И ВИЗУАЛНИ ШАБЛОНИ
    ' ========================================================================
    Private Const ICON_BUILDING As String = "🏢"     ' Иконка за сграда
    Private Const ICON_PANEL As String = "🗄️"        ' Иконка за табло
    Private Const ICON_CIRCUITS As String = "🔵"     ' Иконка за токов кръг
    Private Const LABEL_CIRCUITS As String = "ТК"    ' Кратък етикет за токов кръг
    Private Const POWER_UNIT As String = "kW"        ' Единица за мощност
    Private Const DECIMAL_PLACES As Integer = 2     ' Брой знаци след десетичната запетая при визуализация.
    ''' <summary>
    ''' Форматира текста на възел за табло.
    ''' Добавя иконка и обща мощност.
    ''' </summary>
    Private Function FormatPanelText(item As Form_Tablo_new.strTokow) As String
        ' Създава формат според зададения брой десетични знаци
        Dim formatSpecifier As String = "F" & DECIMAL_PLACES
        ' Форматира мощността
        Dim formattedPower As String = item.Мощност.ToString(formatSpecifier)
        ' Връща готов текст за визуализация
        Return $"{ICON_PANEL} {item.Tablo} ({formattedPower} {POWER_UNIT})"
    End Function
    ''' <summary>
    ''' Форматира текста на възел за токов кръг.
    ''' </summary>
    Private Function FormatCircuitText(item As Form_Tablo_new.strTokow) As String
        ' Връща готов текст за визуализация
        Return $"{ICON_CIRCUITS} {LABEL_CIRCUITS} {item.ТоковКръг} - {item.Device}"
    End Function
    ''' <summary>
    ''' Конструктор на TreeViewManager.
    ''' 
    ''' Инициализира:
    ''' - референция към TreeView контрола
    ''' - референция към основния списък с токови кръгове
    ''' - обработчици за избор на възел
    ''' - Drag & Drop логиката за преместване на табла
    ''' 
    ''' Логика:
    ''' 1. Запазва подадените референции
    ''' 2. Разрешава Drag & Drop върху TreeView
    ''' 3. Закача всички необходими събития:
    '''    - AfterSelect
    '''    - ItemDrag
    '''    - DragEnter
    '''    - DragOver
    '''    - DragDrop
    ''' </summary>
    Public Sub New(ByVal targetTreeView As TreeView, ByRef data As List(Of Form_Tablo_new.strTokow))
        _tv = targetTreeView
        _listTokow = data
        ' Събитие при избор на възел
        AddHandler _tv.AfterSelect, AddressOf HandleAfterSelect
        ' Разрешаваме Drag & Drop върху TreeView
        _tv.AllowDrop = True
        ' Закачаме събитията за Drag & Drop
        AddHandler _tv.ItemDrag, AddressOf HandleItemDrag
        AddHandler _tv.DragEnter, AddressOf HandleDragEnter
        AddHandler _tv.DragOver, AddressOf HandleDragOver
        AddHandler _tv.DragDrop, AddressOf HandleDragDrop
    End Sub
    ' =============================================================
    ' Процедура: RefreshTree
    ' =============================================================
    ' <summary>
    ' Основна процедура за обновяване (refresh) на TreeView контролата (_tv),
    ' която визуализира йерархична структура от данни, съдържащи:
    ' - Сгради (Buildings)
    ' - Табла (Panels / Boards)
    ' - Консуматори / токови кръгове (Circuits)
    '
    ' Данните се вземат от колекцията _listTokow и се организират в дървовидна структура.
    '
    ' Целта на метода е:
    ' - да изгради TreeView от нулата при всяко извикване
    ' - да гарантира уникалност на възлите
    ' - да поддържа йерархия: Сграда → Табло → Табло (вложено) → Консуматори
    ' </summary>
    Public Sub RefreshTree()
        ' <summary>
        ' BeginUpdate спира визуалното обновяване на TreeView-а,
        ' за да се избегне трептене и да се подобри производителността
        ' при масово добавяне/изтриване на възли.
        ' </summary>
        _tv.BeginUpdate()
        ' <summary>
        ' Изчистване на текущата дървовидна структура.
        ' Това гарантира, че всяко извикване започва от "чисто състояние".
        ' </summary>
        _tv.Nodes.Clear()
        Try
            ' =============================================================
            ' Речници за бърз достъп до вече създадени възли
            ' =============================================================
            ' <summary>
            ' buildingNodes:
            ' Ключ: име на сграда (String)
            ' Стойност: TreeNode, представляващ съответната сграда в TreeView
            '
            ' Използва се за:
            ' - предотвратяване на дублиране на сгради
            ' - бързо намиране на root node за дадена сграда
            ' </summary>
            Dim buildingNodes As New Dictionary(Of String, TreeNode)
            ' <summary>
            ' allTabloNodes:
            ' Ключ: уникален идентификатор "Сграда_Табло"
            ' Стойност: TreeNode за конкретно табло
            '
            ' Използва се за:
            ' - гарантиране на уникалност на таблата
            ' - позволяване на вложени табла (табло в табло)
            ' - бърз достъп при добавяне на консуматори
            ' </summary>
            Dim allTabloNodes As New Dictionary(Of String, TreeNode)
            ' =============================================================
            ' 1. ПЪРВИ ПАС: Създаване на възлите за СГРАДИ
            ' =============================================================
            ' <summary>
            ' Обхождаме всички елементи в _listTokow, за да създадем
            ' уникален набор от сгради (root nodes).
            ' </summary>
            For Each item In _listTokow
                ' <summary>
                ' bName:
                ' Нормализирано име на сграда.
                ' Ако BuildingName е празно или whitespace → използва "Обект"
                '
                ' Това гарантира:
                ' - че няма "празни" корени в TreeView
                ' - че всеки елемент винаги принадлежи към някакъв root
                ' </summary>
                Dim bName = If(String.IsNullOrWhiteSpace(item.BuildingName), "Обект", item.BuildingName)
                ' <summary>
                ' Проверка за дублиране:
                ' Ако сградата вече съществува, не я създаваме повторно.
                ' </summary>
                If Not buildingNodes.ContainsKey(bName) Then
                    ' <summary>
                    ' Създаване на TreeNode за сграда.
                    ' ICON_BUILDING е визуален индикатор (икона/символ).
                    ' </summary>
                    Dim bNode As New TreeNode($"{ICON_BUILDING} {bName}")
                    ' Добавяне към TreeView root ниво
                    _tv.Nodes.Add(bNode)
                    ' Съхраняване в речника за бъдеща референция
                    buildingNodes.Add(bName, bNode)
                End If
            Next
            ' =============================================================
            ' 2. ВТОРИ ПАС: Създаване на уникални ТАБЛА (без йерархия)
            ' =============================================================
            ' <summary>
            ' В този пас се създават всички възли от тип "Табло",
            ' но без да се подреждат в дървото.
            '
            ' Причина:
            ' - първо се гарантира уникалност
            ' - после се прави йерархично свързване (в 3-ти пас)
            ' </summary>
            For Each item In _listTokow
                If item.Device = "Табло" Then
                    ' <summary>
                    ' tabloKey:
                    ' Уникален ключ за табло в рамките на сграда.
                    '
                    ' Формат:
                    ' "Сграда_Табло"
                    '
                    ' Това предотвратява конфликт при:
                    ' - еднакви имена на табла в различни сгради
                    ' </summary>
                    Dim tabloKey = item.BuildingName & "_" & item.Tablo
                    If Not allTabloNodes.ContainsKey(tabloKey) Then
                        ' <summary>
                        ' Създаване на TreeNode за табло.
                        '
                        ' FormatPanelText(item):
                        ' централен шаблон за визуализация:
                        ' - икона
                        ' - име
                        ' - формат/десетични знаци
                        ' </summary>
                        Dim tNode As New TreeNode(FormatPanelText(item))
                        ' <summary>
                        ' Tag:
                        ' Съхранява оригиналния обект item,
                        ' за да може по-късно да се извлича логика/данни
                        ' при селекция в TreeView.
                        ' </summary>
                        tNode.Tag = item
                        allTabloNodes(tabloKey) = tNode
                    End If
                End If
            Next
            ' =============================================================
            ' 3. ТРЕТИ ПАС: ЙЕРАРХИЧНО ПОДРЕЖДАНЕ НА ТАБЛАТА
            ' =============================================================
            ' <summary>
            ' Тук вече се определя къде точно отива всяко табло:
            ' - в друго табло (вложена структура)
            ' - или директно под сграда
            ' </summary>
            For Each item In _listTokow
                If item.Device = "Табло" Then
                    Dim tabloKey = item.BuildingName & "_" & item.Tablo
                    Dim currentNode = allTabloNodes(tabloKey)
                    ' <summary>
                    ' Защита срещу дублирано закачане:
                    ' Ако node вече има родител → пропускаме
                    ' (предотвратява многократно добавяне)
                    ' </summary>
                    If currentNode.Parent IsNot Nothing Then Continue For
                    ' <summary>
                    ' Проверка за родителско табло:
                    ' item.Табло_Родител съдържа име на "родителско" табло
                    ' </summary>
                    If Not String.IsNullOrEmpty(item.Табло_Родител) Then
                        Dim parentKey = item.BuildingName & "_" & item.Табло_Родител
                        If allTabloNodes.ContainsKey(parentKey) Then
                            ' Вложено табло → добавяне под родителя
                            allTabloNodes(parentKey).Nodes.Add(currentNode)
                        End If
                    Else
                        ' <summary>
                        ' Ако няма родител:
                        ' таблото се добавя директно под съответната сграда
                        ' </summary>
                        Dim bName = If(String.IsNullOrWhiteSpace(item.BuildingName), "Обект", item.BuildingName)
                        buildingNodes(bName).Nodes.Add(currentNode)
                    End If
                End If
            Next
            ' =============================================================
            ' 4. ЧЕТВЪРТИ ПАС: ДОБАВЯНЕ НА КОНСУМАТОРИ (КРЪГОВЕ)
            ' =============================================================
            ' <summary>
            ' В този етап се добавят всички елементи,
            ' които НЕ са "Табло" → т.е. консуматори / токови кръгове.
            ' </summary>
            For Each item In _listTokow
                If item.Device <> "Табло" Then
                    Dim tabloKey = item.BuildingName & "_" & item.Tablo
                    If allTabloNodes.ContainsKey(tabloKey) Then
                        ' <summary>
                        ' Създаване на възел за токов кръг (консуматор).
                        '
                        ' FormatCircuitText(item):
                        ' шаблон за визуализация на електрически кръг
                        ' </summary>
                        Dim cNode As New TreeNode(FormatCircuitText(item))
                        cNode.Tag = item
                        allTabloNodes(tabloKey).Nodes.Add(cNode)
                    End If
                End If
            Next
        Finally
            ' =============================================================
            ' ФИНАЛНА ВИЗУАЛНА ОРГАНИЗАЦИЯ НА ДЪРВОТО
            ' =============================================================
            ' <summary>
            ' CollapseAll:
            ' Свива всички възли в TreeView.
            ' </summary>
            _tv.CollapseAll()
            ' <summary>
            ' Разгъване само на root нивото (сградите),
            ' за да се покаже първо структурното ниво на проекта.
            ' </summary>
            For Each rootNode As TreeNode In _tv.Nodes
                rootNode.Expand()
            Next
            ' <summary>
            ' EndUpdate възстановява визуалното обновяване
            ' и показва финалната структура наведнъж.
            ' </summary>
            _tv.EndUpdate()
        End Try
    End Sub
    ''' <summary>
    ''' Обработва избора на възел в TreeView.
    ''' 
    ''' Ако избраният възел съдържа обект от тип strTokow
    ''' в свойството Tag, събитието ObjectSelected се
    ''' извиква и предава избрания обект към външния код.
    ''' 
    ''' Използва се за синхронизация между TreeView,
    ''' DataGridView и останалата логика на формата.
    ''' </summary>
    ''' <param name="sender">
    ''' TreeView контролът, който е генерирал събитието.
    ''' </param>
    ''' <param name="e">
    ''' Данни за избрания възел.
    ''' </param>
    Private Sub HandleAfterSelect(sender As Object, e As TreeViewEventArgs)
        If e.Node.Tag IsNot Nothing AndAlso TypeOf e.Node.Tag Is Form_Tablo_new.strTokow Then
            RaiseEvent ObjectSelected(DirectCast(e.Node.Tag, Form_Tablo_new.strTokow))
        End If
    End Sub
    ''' <summary>
    ''' Стартира Drag & Drop операция при влачене на възел от TreeView.
    ''' 
    ''' Методът:
    ''' 1. Маркира текущо влачения възел като SelectedNode
    ''' 2. Стартира операция по преместване (Move)
    ''' 
    ''' Използва се като начална точка за прехвърляне
    ''' на табла и токови кръгове в йерархията.
    ''' </summary>
    ''' <param name="sender">
    ''' TreeView контролът, който генерира събитието.
    ''' </param>
    ''' <param name="e">
    ''' Данни за влачения обект.
    ''' </param>
    Private Sub HandleItemDrag(sender As Object, e As ItemDragEventArgs)
        _tv.SelectedNode = DirectCast(e.Item, TreeNode)
        _tv.DoDragDrop(e.Item, DragDropEffects.Move)
    End Sub
    ''' <summary>
    ''' Активира Drag & Drop операция при навлизане на мишката
    ''' в зоната на TreeView.
    ''' 
    ''' Задава ефект "Move", което указва,
    ''' че възелът може да бъде преместен.
    ''' </summary>
    ''' <param name="sender">
    ''' TreeView контролът, който приема Drag & Drop операцията.
    ''' </param>
    ''' <param name="e">
    ''' Данни за текущата Drag & Drop операция.
    ''' </param>
    Private Sub HandleDragEnter(sender As Object, e As DragEventArgs)
        e.Effect = DragDropEffects.Move
    End Sub
    ''' <summary>
    ''' Обработва движението на влачения елемент над TreeView.
    ''' 
    ''' Методът:
    ''' 1. Преобразува координатите на мишката към TreeView
    ''' 2. Намира възела под курсора
    ''' 3. Маркира възела визуално като текуща цел
    ''' 
    ''' Използва се за по-удобна визуална навигация
    ''' по време на Drag & Drop операция.
    ''' </summary>
    ''' <param name="sender">
    ''' TreeView контролът, върху който се извършва влаченето.
    ''' </param>
    ''' <param name="e">
    ''' Данни за текущата Drag & Drop операция.
    ''' </param>
    Private Sub HandleDragOver(sender As Object, e As DragEventArgs)
        Dim targetPoint As Point = _tv.PointToClient(New Point(e.X, e.Y))
        _tv.SelectedNode = _tv.GetNodeAt(targetPoint) ' Визуално маркираме целта
    End Sub
    ''' <summary>
    ''' Финализира Drag & Drop операцията при пускане на възел.
    ''' 
    ''' Методът:
    ''' 1. Определя върху кой възел е пуснат елементът
    ''' 2. Извлича влачения и целевия обект от Tag
    ''' 3. Проверява дали операцията е валидна
    ''' 4. Изпраща заявка към формата за преместване
    ''' 
    ''' Позволява:
    ''' • местене на токови кръгове между табла
    ''' • преместване на табла в други табла
    ''' • изграждане на йерархична структура
    ''' 
    ''' Забранява:
    ''' • пускане върху самия себе си
    ''' • невалидни цели
    ''' </summary>
    ''' <param name="sender">
    ''' TreeView контролът, който приема операцията.
    ''' </param>
    ''' <param name="e">
    ''' Данни за Drag & Drop операцията.
    ''' </param>
    Private Sub HandleDragDrop(sender As Object, e As DragEventArgs)
        ' Вземаме точката, в която е пуснат бутонът
        Dim targetPoint As Point = _tv.PointToClient(New Point(e.X, e.Y))
        ' Намираме възела под курсора
        Dim targetNode As TreeNode = _tv.GetNodeAt(targetPoint)
        ' Вземаме влачения възел
        Dim draggedNode As TreeNode =
        DirectCast(e.Data.GetData(GetType(TreeNode)), TreeNode)
        ' Проверка:
        ' • дали имаме валидни възли
        ' • дали не местим възела върху самия него
        If draggedNode IsNot Nothing AndAlso
       targetNode IsNot Nothing AndAlso
       Not draggedNode.Equals(targetNode) Then
            ' Вземаме обектите от Tag-а
            Dim sourceObj As Form_Tablo_new.strTokow =
            DirectCast(draggedNode.Tag, Form_Tablo_new.strTokow)
            Dim targetObj As Form_Tablo_new.strTokow =
            DirectCast(targetNode.Tag, Form_Tablo_new.strTokow)
            ' Разрешаваме местене само:
            ' • в друго табло
            ' • или в корен/сграда
            If targetObj.Device = "Табло" OrElse
           String.IsNullOrEmpty(targetObj.Tablo) Then
                ' Изпращаме заявка към формата,
                ' която реално ще обнови данните
                RaiseEvent RequestMoveObject(sourceObj, targetObj)
            End If
        End If
    End Sub
End Class