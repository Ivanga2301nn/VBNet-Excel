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
    Private Const ICON_FOLDER As String = "📂"       ' New: Иконка за папката с токови кръгове
    Private Const LABEL_CIRCUITS As String = "ТК"    ' Кратък етикет за токов кръг
    Private Const CONSUMERS_NODE_TEXT As String = "Консуматори"
    Private Const POWER_UNIT As String = "kW"        ' Единица за мощност
    Private Const DECIMAL_PLACES As Integer = 2      ' Брой знаци след десетичната запетая при визуализация.
    ' ========================================================================
    ' 🖱️ КОНТЕКСТНО МЕНЮ (ДЕСЕН БУТОН)
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

        ' Създаване на бутони
        Dim menuAddPanel As New ToolStripMenuItem("➕ Добави под-табло", Nothing, AddressOf MenuAddPanel_Click)
        Dim menuDelete As New ToolStripMenuItem("❌ Изтрий избран елемент", Nothing, AddressOf MenuDelete_Click)

        Dim menuExpandNode As New ToolStripMenuItem("➕ Разгъни този възел", Nothing, AddressOf MenuExpandNode_Click)
        Dim menuCollapseNode As New ToolStripMenuItem("➖ Свий този възел", Nothing, AddressOf MenuCollapseNode_Click)

        Dim menuExpandAll As New ToolStripMenuItem("📂 Разгъни всичко", Nothing, AddressOf MenuExpandAll_Click)
        Dim menuCollapseAll As New ToolStripMenuItem("📁 Свий всичко", Nothing, AddressOf MenuCollapseAll_Click)

        ' Добавяне в менюто (със сепаратори, създадени на място)
        _contextMenu.Items.Add(menuAddPanel)
        _contextMenu.Items.Add(menuDelete)
        _contextMenu.Items.Add(New ToolStripSeparator()) ' Нов обект всеки път!
        _contextMenu.Items.Add(menuExpandNode)
        _contextMenu.Items.Add(menuCollapseNode)
        _contextMenu.Items.Add(New ToolStripSeparator()) ' Нов обект всеки път!
        _contextMenu.Items.Add(menuExpandAll)
        _contextMenu.Items.Add(menuCollapseAll)

        _tv.ContextMenuStrip = _contextMenu
        AddHandler _tv.NodeMouseClick, AddressOf _tv_NodeMouseClick
    End Sub
    ' ========================================================================
    ' Контекстни команди (БЕЗ ДУБЛИРАНИ ИМЕНА)
    ' ========================================================================
    Private Sub MenuAddPanel_Click(sender As Object, e As EventArgs)
        ' Твоята логика тук
    End Sub
    Private Sub MenuDelete_Click(sender As Object, e As EventArgs)
        ' Твоята логика тук
    End Sub
    Private Sub MenuExpandNode_Click(sender As Object, e As EventArgs)
        If _tv.SelectedNode IsNot Nothing Then _tv.SelectedNode.ExpandAll()
    End Sub
    Private Sub MenuCollapseNode_Click(sender As Object, e As EventArgs)
        If _tv.SelectedNode IsNot Nothing Then _tv.SelectedNode.Collapse()
    End Sub
    Private Sub MenuExpandAll_Click(sender As Object, e As EventArgs)
        _tv.BeginUpdate()
        _tv.ExpandAll()
        _tv.EndUpdate()
    End Sub
    Private Sub MenuCollapseAll_Click(sender As Object, e As EventArgs)
        _tv.BeginUpdate()
        _tv.CollapseAll()
        ' Ако ползваш ExpandNodesToLevel, се увери, че е дефиниран някъде
        For Each rootNode As TreeNode In _tv.Nodes
            ' Тук сложи твоята логика, ако е необходимо
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
    ''' <summary>
    ''' Основна процедура за пълно изграждане и обновяване на TreeView (_tv),
    ''' съдържащ йерархична структура на:
    '''
    ''' - Сгради
    ''' - Табла
    ''' - Подтабла
    ''' - Консуматори / токови кръгове
    '''
    ''' Процедурата:
    ''' 1. Изчиства текущото дърво
    ''' 2. Създава всички сгради
    ''' 3. Създава всички табла
    ''' 4. Свързва таблата йерархично
    ''' 5. Добавя консуматорите в специални групи
    ''' 6. Изчислява сумарна мощност на консуматорите
    ''' 7. Форматира финалния изглед
    '''
    ''' Използва се многопроходна (multi-pass) обработка,
    ''' за да се избегнат проблеми със зависимости между възлите.
    ''' </summary>
    Public Sub RefreshTree()
        _tv.BeginUpdate()
        _tv.Nodes.Clear()
        Try
            ' РЕЧНИЦИ ЗА УПРАВЛЕНИЕ НА ВЪЗЛИТЕ
            ' buildingNodes:
            ' Съдържа всички root възли за сградите.
            ' Ключ:
            ' - име на сграда
            ' Стойност:
            ' - TreeNode на сградата
            ' StringComparer.OrdinalIgnoreCase:
            ' игнорира разлики между малки/главни букви.
            Dim buildingNodes As New Dictionary(Of String, TreeNode)(StringComparer.OrdinalIgnoreCase)
            ' allTabloNodes:
            ' Централен речник за всички табла.
            ' Ключ:
            ' - "Сграда_Табло"
            ' Стойност:
            ' - TreeNode на таблото
            ' Използва се за:
            ' - уникалност
            ' - бърз достъп
            ' - йерархично свързване
            Dim allTabloNodes As New Dictionary(Of String, TreeNode)(StringComparer.OrdinalIgnoreCase)
            ' tabloToBuilding:
            ' Свързва всяко табло със съответната му сграда.
            ' Използва се по-късно при:
            ' - автоматично закачане
            ' - fallback логика
            Dim tabloToBuilding As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            ' tabloToParent:
            ' Съдържа информация за родителското табло.
            ' Ключ:
            ' - текущо табло
            ' Стойност:
            ' - parent tablo key
            Dim tabloToParent As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            ' 1. ПЪРВИ ПАС: СЪЗДАВАНЕ НА СГРАДИТЕ
            For Each item In _listTokow
                ' Нормализиране на името на сградата.
                ' Ако липсва име:
                ' използва се ROOT_NODE_TEXT като fallback root.
                Dim bName = If(String.IsNullOrWhiteSpace(item.BuildingName),
                           Form_Tablo_new.ROOT_NODE_TEXT,
                           item.BuildingName.Trim())
                ' Проверка дали сградата вече е създадена.
                If Not buildingNodes.ContainsKey(bName) Then
                    ' ICON_BUILDING:
                    ' визуален символ/икона.
                    Dim bNode As New TreeNode($"{ICON_BUILDING} {bName}")
                    ' Добавяне като root node
                    _tv.Nodes.Add(bNode)
                    ' Записване в речника
                    buildingNodes.Add(bName, bNode)
                End If
            Next
            ' 2. ВТОРИ ПАС: СЪЗДАВАНЕ НА ТАБЛАТА
            For Each item In _listTokow
                If item.Device IsNot Nothing AndAlso
                        item.Device.Trim().Equals("Табло", StringComparison.OrdinalIgnoreCase) Then
                    Dim bName = If(String.IsNullOrWhiteSpace(item.BuildingName),
                               Form_Tablo_new.ROOT_NODE_TEXT,
                               item.BuildingName.Trim())
                    ' нормализирано име на табло.
                    Dim tName = If(item.Tablo Is Nothing, "", item.Tablo.Trim())
                    ' tabloKey:
                    ' уникален идентификатор за табло.
                    Dim tabloKey = bName & "_" & tName
                    ' СЪЗДАВАНЕ НА TREE NODE ЗА ТАБЛО
                    If Not allTabloNodes.ContainsKey(tabloKey) Then
                        ' FormatPanelText:
                        ' централизирана визуализация на табло:
                        ' - име
                        ' - икони
                        ' - мощности
                        ' - формат
                        Dim tNode As New TreeNode(FormatPanelText(item))
                        ' пази оригиналния бизнес обект.
                        tNode.Tag = item
                        allTabloNodes(tabloKey) = tNode
                        ' Свързване към сграда
                        tabloToBuilding(tabloKey) = bName
                        ' По подразбиране няма родител
                        tabloToParent(tabloKey) = ""
                    End If
                    ' ЗАПИС НА РОДИТЕЛСКО ТАБЛО
                    ' Ако е зададен родител:
                    ' записва се ключ към родителското табло.
                    If Not String.IsNullOrWhiteSpace(item.Табло_Родител) AndAlso
                   item.Табло_Родител.Trim() <> Form_Tablo_new.ROOT_NODE_TEXT Then
                        tabloToParent(tabloKey) =
                        bName & "_" & item.Табло_Родител.Trim()
                    End If
                End If
            Next
            ' 3. ТРЕТИ ПАС: ЙЕРАРХИЧНО СВЪРЗВАНЕ НА ТАБЛАТА
            For Each tabloKey In allTabloNodes.Keys
                Dim currentNode = allTabloNodes(tabloKey)
                Dim parentKey = tabloToParent(tabloKey)
                Dim bName = tabloToBuilding(tabloKey)
                ' А) ТАБЛО С РОДИТЕЛ
                ' Ако има валиден parent:
                ' текущото табло се закача към него.
                ' защита срещу самореференция.
                If Not String.IsNullOrEmpty(parentKey) AndAlso
               parentKey <> tabloKey AndAlso
               allTabloNodes.ContainsKey(parentKey) Then
                    allTabloNodes(parentKey).Nodes.Add(currentNode)
                    ' Б) ГЛАВНО РАЗПРЕДЕЛИТЕЛНО ТАБЛО (Гл.Р.Т.)
                ElseIf tabloKey.EndsWith("_Гл.Р.Т.", StringComparison.OrdinalIgnoreCase) Then
                    ' Главното табло се поставя директно под сградата.
                    If buildingNodes.ContainsKey(bName) Then
                        buildingNodes(bName).Nodes.Add(currentNode)
                    End If
                    ' В) ОБИКНОВЕНО ТАБЛО БЕЗ РОДИТЕЛ
                Else
                    ' Ако таблото няма родител:
                    ' опитва се автоматично да се закачи към Гл.Р.Т.
                    Dim glrtKey = bName & "_Гл.Р.Т."
                    If allTabloNodes.ContainsKey(glrtKey) Then
                        ' Закачане към главното табло
                        allTabloNodes(glrtKey).Nodes.Add(currentNode)
                    ElseIf buildingNodes.ContainsKey(bName) Then
                        ' Fallback:
                        ' ако липсва Гл.Р.Т. → директно под сградата
                        buildingNodes(bName).Nodes.Add(currentNode)
                    End If
                End If
            Next
            ' 4. ЧЕТВЪРТИ ПАС: ДОБАВЯНЕ НА КОНСУМАТОРИ
            ' <summary>
            ' consumerSums:
            ' пази сумарна мощност за всяка група консуматори.
            ' Ключ:
            ' - group TreeNode
            ' Стойност:
            ' - total power
            Dim consumerSums As New Dictionary(Of TreeNode, Double)
            For Each item In _listTokow
                ' ПРОВЕРКА ДАЛИ Е КОНСУМАТОР
                ' Всички елементи, които НЕ са "Табло",
                ' се третират като консуматори/кръгове.
                If item.Device Is Nothing OrElse
               Not item.Device.Trim().Equals("Табло", StringComparison.OrdinalIgnoreCase) Then
                    Dim bName = If(String.IsNullOrWhiteSpace(item.BuildingName),
                               Form_Tablo_new.ROOT_NODE_TEXT,
                               item.BuildingName.Trim())
                    Dim tName = If(item.Tablo Is Nothing, "", item.Tablo.Trim())
                    Dim tabloKey = bName & "_" & tName
                    If allTabloNodes.ContainsKey(tabloKey) Then
                        ' НАМИРАНЕ НА РОДИТЕЛСКОТО ТАБЛО
                        Dim parentPanelNode = allTabloNodes(tabloKey)
                        ' специален възел-група за всички консуматори.
                        Dim consumerGroupNode As TreeNode = Nothing
                        ' ТЪРСЕНЕ НА ГРУПА "КОНСУМАТОРИ"
                        For Each node In parentPanelNode.Nodes
                            ' Проверява дали вече има група за консуматори.
                            If node.Text.Contains(CONSUMERS_NODE_TEXT) Then
                                consumerGroupNode = node
                                Exit For
                            End If
                        Next
                        ' СЪЗДАВАНЕ НА ГРУПА ПРИ ЛИПСА
                        If consumerGroupNode Is Nothing Then
                            consumerGroupNode =
                            New TreeNode($"{ICON_FOLDER} {CONSUMERS_NODE_TEXT}")
                            parentPanelNode.Nodes.Add(consumerGroupNode)
                            ' Начална стойност за сумата
                            consumerSums(consumerGroupNode) = 0
                        End If
                        ' ДОБАВЯНЕ НА КОНСУМАТОР
                        Dim cNode As New TreeNode(FormatCircuitText(item))
                        cNode.Tag = item
                        consumerGroupNode.Nodes.Add(cNode)
                        ' СУМИРАНЕ НА МОЩНОСТТА
                        Dim power As Double = 0
                        ' безопасно преобразуване към Double.
                        Double.TryParse(item.Мощност.ToString(), power)
                        consumerSums(consumerGroupNode) += power
                    End If
                End If
            Next
            ' =============================================================
            ' ОБНОВЯВАНЕ НА ТЕКСТА НА ГРУПИТЕ
            ' =============================================================
            ' Формат за десетични знаци.
            Dim formatSpecifier As String = "F" & DECIMAL_PLACES
            For Each groupNode In consumerSums.Keys
                Dim sumValue As Double = consumerSums(groupNode)
                ' Обновяване текста на групата:
                groupNode.Text =
                $"{ICON_FOLDER} {CONSUMERS_NODE_TEXT} ({sumValue.ToString(formatSpecifier)} {POWER_UNIT})"
            Next
        Finally
            ' Свива всички възли
            _tv.CollapseAll()
            For Each rootNode As TreeNode In _tv.Nodes
                ' Разгъване само на root нивото.
                rootNode.Expand()
                ' Свиване на таблата за по-прегледна начална визуализация.
                For Each panelNode As TreeNode In rootNode.Nodes
                    panelNode.Collapse()
                Next
            Next
            ' Възстановяване на визуалното обновяване.
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