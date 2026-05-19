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
    Private Const ICON_FOLDER As String = "📂"        ' New: Иконка за папката с токови кръгове
    Private Const LABEL_CIRCUITS As String = "ТК"    ' Кратък етикет за токов кръг
    Private Const POWER_UNIT As String = "kW"        ' Единица за мощност
    Private Const DECIMAL_PLACES As Integer = 2      ' Брой знаци след десетичната запетая при визуализация.
    ' ========================================================================
    ' 🖱️ КОНТЕКСТНО МЕНЮ (ДЕ СЕН БУТОН)
    ' ========================================================================
    Private WithEvents _contextMenu As ContextMenuStrip
    ' Дефинираме максималното ниво на разгъване. 
    ' За момента е 0 (само сградите), но лесно можеш да го промениш на 1, 2 или 3.
    Private MaxExpandLevel As Integer = 0
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
        ' Инициализиране на менюто за десен бутон
        InitializeContextMenu()
    End Sub
    ' ========================================================================
    ' Метод: InitializeContextMenu
    ' ========================================================================
    ''' <summary>
    ''' Конструира елементите на контекстното меню и го обвързва с TreeView контролата.
    ''' </summary>
    Private Sub InitializeContextMenu()
        _contextMenu = New ContextMenuStrip()

        ' Създаване на бутони с иконки и шаблони за управление
        Dim menuAddPanel As New ToolStripMenuItem("➕ Добави под-табло", Nothing, AddressOf MenuAddPanel_Click)
        Dim menuDelete As New ToolStripMenuItem("❌ Изтрий избран елемент", Nothing, AddressOf MenuDelete_Click)
        Dim menuSeparator As New ToolStripSeparator()
        Dim menuExpandAll As New ToolStripMenuItem("📂 Разгъни всичко", Nothing, AddressOf MenuExpandAll_Click)
        Dim menuCollapseAll As New ToolStripMenuItem("📁 Свий всичко", Nothing, AddressOf MenuCollapseAll_Click)

        ' Набиване на елементите в менюто
        _contextMenu.Items.Add(menuAddPanel)
        _contextMenu.Items.Add(menuDelete)
        _contextMenu.Items.Add(menuSeparator)
        _contextMenu.Items.Add(menuExpandAll)
        _contextMenu.Items.Add(menuCollapseAll)

        ' Закачане на менюто към твоя TreeView (_tv)
        _tv.ContextMenuStrip = _contextMenu
        AddHandler _tv.NodeMouseClick, AddressOf _tv_NodeMouseClick
    End Sub
    ' ========================================================================
    ' Контекстни команди (Действия при клик)
    ' ========================================================================
    Private Sub MenuAddPanel_Click(sender As Object, e As EventArgs)
        Dim selectedNode = _tv.SelectedNode
        If selectedNode IsNot Nothing Then
            Dim currentItem As Form_Tablo_new.strTokow = TryCast(selectedNode.Tag, Form_Tablo_new.strTokow)
            If currentItem IsNot Nothing Then
                MessageBox.Show($"Тук ще добавим ново табло, чийто родител ще бъде: {currentItem.Tablo}")
                ' След добавяне в _listTokow ще викаме твоя RefreshTree()
            End If
        End If
    End Sub
    Private Sub MenuDelete_Click(sender As Object, e As EventArgs)
        Dim selectedNode = _tv.SelectedNode
        If selectedNode IsNot Nothing Then
            ' Ако е папка "Токови кръгове", тя няма Tag, но таблата и кръговете имат
            Dim labelText As String = selectedNode.Text
            Dim result = MessageBox.Show($"Сигурни ли сте, че искате да изтриете {labelText}?", "Потвърждение", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

            If result = DialogResult.Yes Then
                ' Логика за триене от _listTokow и обновяване на дървото
                RefreshTree()
            End If
        End If
    End Sub
    Private Sub MenuExpandAll_Click(sender As Object, e As EventArgs)
        _tv.BeginUpdate()
        _tv.ExpandAll()
        _tv.EndUpdate()
    End Sub
    Private Sub MenuCollapseAll_Click(sender As Object, e As EventArgs)
        _tv.BeginUpdate()
        _tv.CollapseAll()
        ' Прилагаме динамичното свиване/разгъване, което направихме, 
        ' за да се отворят само сградите (Ниво 0)
        For Each rootNode As TreeNode In _tv.Nodes
            ExpandNodesToLevel(rootNode, MaxExpandLevel)
        Next
        _tv.EndUpdate()
    End Sub
    ' ========================================================================
    ' Събитие: _tv_NodeMouseClick
    ' ========================================================================
    ''' <summary>
    ''' Гарантира, че десният клик маркира възела под мишката и управлява видимостта на бутоните.
    ''' </summary>
    Private Sub _tv_NodeMouseClick(sender As Object, e As TreeNodeMouseClickEventArgs)

        If e.Button = MouseButtons.Right Then
            ' Автоматично селектираме възела, върху който е кликнато
            _tv.SelectedNode = e.Node

            ' Проверка дали маркираният възел е реално Табло (чрез Tag структурата ти)
            Dim currentItem As Form_Tablo_new.strTokow = TryCast(e.Node.Tag, Form_Tablo_new.strTokow)
            Dim isPanel As Boolean = (currentItem IsNot Nothing AndAlso currentItem.Device = "Табло")

            ' Защита: Бутонът "Добави под-табло" е активен САМО ако сме кликнали върху Табло
            _contextMenu.Items(0).Enabled = isPanel
        End If
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
            ' =============================================================
            ' 1. ПЪРВИ ПАС: Създаване на възлите за СГРАДИ
            ' =============================================================
            For Each item In _listTokow
                ' Заменяме "Обект" с динамичната променлива от формата
                Dim bName = If(String.IsNullOrWhiteSpace(item.BuildingName), Form_Tablo_new.ROOT_NODE_TEXT, item.BuildingName)

                If Not buildingNodes.ContainsKey(bName) Then
                    Dim bNode As New TreeNode($"{ICON_BUILDING} {bName}")
                    _tv.Nodes.Add(bNode)
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
            For Each item In _listTokow
                If item.Device = "Табло" Then ' (Забележка: провери дали е "Табло" или "Taбло" с латинско 'a')
                    Dim tabloKey = item.BuildingName & "_" & item.Tablo
                    Dim currentNode = allTabloNodes(tabloKey)

                    If currentNode.Parent IsNot Nothing Then Continue For

                    If Not String.IsNullOrEmpty(item.Табло_Родител) Then
                        Dim parentKey = item.BuildingName & "_" & item.Табло_Родител
                        If allTabloNodes.ContainsKey(parentKey) Then
                            allTabloNodes(parentKey).Nodes.Add(currentNode)
                        End If
                    Else
                        ' Заменяме "Обект" с динамичната променлива от формата и тук
                        Dim bName = If(String.IsNullOrWhiteSpace(item.BuildingName), Form_Tablo_new.ROOT_NODE_TEXT, item.BuildingName)
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
            ' 1. Свиваме АБСОЛЮТНО ВСИЧКО
            _tv.CollapseAll()
            ' 2. Разгъваме САМО сградите (root ниво)
            For Each rootNode As TreeNode In _tv.Nodes
                rootNode.Expand()
                ' 3. Експлицитно гарантираме, че всички табла (и тяхното съдържание) остават свити
                For Each panelNode As TreeNode In rootNode.Nodes
                    panelNode.Collapse()
                Next
            Next
            ' 4. Връщаме визуалното обновяване
            _tv.EndUpdate()
        End Try
    End Sub
    ''' <summary>
    ''' Контролирано разгъва TreeView структурата до определено ниво в дълбочина.
    ''' </summary>
    Private Sub ExpandNodesToLevel(node As TreeNode, maxLevel As Integer)
        ' Ако текущото ниво е в разрешения диапазон, разгъваме възела
        If node.Level <= maxLevel Then
            node.Expand()
            ' Рекурсивно проверяваме и разгъваме децата му
            For Each childNode As TreeNode In node.Nodes
                ExpandNodesToLevel(childNode, maxLevel)
            Next
        Else
            ' Ако сме надвишили максималното ниво, за всеки случай свиваме нагоре
            node.Collapse()
        End If
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