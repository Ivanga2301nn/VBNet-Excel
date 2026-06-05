Imports System.Drawing
Imports System.Windows.Forms

Public Class TreeViewManager

    ' ========================================================================
    ' 📌 СЪБИТИЯ (МОСТ КЪМ ФОРМАТА)
    ' ========================================================================
    ''' <summary>
    ''' Задейства се при ляв клик върху табло или консуматор. Подава обекта на формата.
    ''' </summary>
    Public Event NodeLeftClick(ByVal selectedObject As clsTokow)

    ''' <summary>
    ''' Задейства се при успешно пускане на елемент чрез Drag & Drop.
    ''' </summary>
    Public Event RequestMoveObject(ByVal source As clsTokow, ByVal target As clsTokow)


    ' ========================================================================
    ' 📌 ПОЛЕТА И КОНСТАНТИ
    ' ========================================================================
    Private WithEvents _tv As TreeView
    Private WithEvents _contextMenu As ContextMenuStrip

    Friend Shared ROOT_NODE_TEXT As String = "Гл.Р.Т."

    Private Const ICON_BUILDING As String = "🏢"     ' Иконка за сграда
    Private Const ICON_PANEL As String = "🗄️"        ' Иконка за табло
    Private Const ICON_CIRCUITS As String = "🔵"     ' Иконка за токов кръг
    Private Const ICON_FOLDER As String = "📂"       ' Иконка за папката с токови кръгове
    Private Const LABEL_CIRCUITS As String = "ТК"    ' Кратък етикет за токов кръг
    Private Const CONSUMERS_NODE_TEXT As String = "Консуматори"
    Private Const POWER_UNIT As String = "kW"        ' Единица за мощност
    Private Const DECIMAL_PLACES As Integer = 2      ' Брой знаци след десетичната запетая

    ' TRUE: Свива дървото до Сгради/Табла при пускане на формата за спретнат начален изглед.
    ' FALSE: След Drag & Drop позволява на дървото да се разгърне автоматично, за да не се затваря пред очите на потребителя.
    Private _isFirstLoad As Boolean = True


    ' ========================================================================
    ' 📌 КОНСТРУКТОР
    ' ========================================================================
    Public Sub New(ByVal targetTreeView As TreeView)
        _tv = targetTreeView

        ' Визуални настройки
        _tv.ForeColor = Color.FromArgb(45, 45, 45)
        _tv.LineColor = Color.FromArgb(180, 180, 180)
        _tv.AllowDrop = True

        ' Ръчно закачаме само Drag събитията, които нямат WithEvents/Handles
        AddHandler _tv.ItemDrag, AddressOf HandleItemDrag
        AddHandler _tv.DragEnter, AddressOf HandleDragEnter
        AddHandler _tv.DragOver, AddressOf HandleDragOver
        AddHandler _tv.DragDrop, AddressOf HandleDragDrop

        ' Инициализиране на контекстното меню
        InitializeContextMenu()
    End Sub


    ' ========================================================================
    ' 🎨 ПОМОЩНИ ФОРМАТИРАЩИ ФУНКЦИИ
    ' ========================================================================
    Private Function FormatPanelText(item As clsTokow) As String
        Dim formatSpecifier As String = "F" & DECIMAL_PLACES
        Dim formattedPower As String = item.Мощност.ToString(formatSpecifier)
        Return $"{ICON_PANEL} {item.Tablo} ({formattedPower} {POWER_UNIT})"
    End Function

    Private Function FormatCircuitText(item As clsTokow) As String
        Return $"{ICON_CIRCUITS} {LABEL_CIRCUITS} {item.ТоковКръг} - {item.Device}"
    End Function


    ' ========================================================================
    ' 🖱️ УПРАВЛЕНИЕ НА КЛИКОВЕ И СЪБИТИЯ С МИШКАТА
    ' ========================================================================
    ''' <summary>
    ''' Реагира при клик с мишката върху възел в дървото.
    ''' </summary>
    Private Sub _tv_NodeMouseClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles _tv.NodeMouseClick
        Dim currentItem As clsTokow = TryCast(e.Node.Tag, clsTokow)

        Select Case e.Button
            Case MouseButtons.Left
                ' --- ЛЯВ КЛИК: Избор на обект за таблицата ---
                ' Използваме HitTest, за да разберем точно къде е кликнал потребителят
                Dim hitInfo As TreeViewHitTestInfo = _tv.HitTest(e.Location)
                ' Проверяваме дали кликът е върху самия текст или иконката на възела (а не върху "+" или празното място)
                If hitInfo.Location = TreeViewHitTestLocations.Label OrElse
               hitInfo.Location = TreeViewHitTestLocations.Image Then
                    ' Само тогава изстрелваме събитието към формата за таблицата!
                    If currentItem IsNot Nothing Then
                        RaiseEvent NodeLeftClick(currentItem)
                    End If
                End If
            Case MouseButtons.Right
                ' --- ДЕСЕН КЛИК: Маркиране и Контекстно меню ---
                _tv.SelectedNode = e.Node
                If currentItem IsNot Nothing Then
                    ' Бутонът "Добави под-табло" е активен САМО ако устройството е "Табло"
                    Dim isPanel As Boolean = (currentItem.Device IsNot Nothing AndAlso
                                              currentItem.Device.Trim().Equals("Табло", StringComparison.OrdinalIgnoreCase))
                    _contextMenu.Items(0).Enabled = isPanel
                Else
                    _contextMenu.Items(0).Enabled = False
                End If
        End Select
    End Sub

    ''' <summary>
    ''' Спира разгръщането на папките "Консуматори".
    ''' </summary>
    Private Sub _tv_BeforeExpand(sender As Object, e As TreeViewCancelEventArgs) Handles _tv.BeforeExpand
        If e.Node.Text.Contains(ICON_FOLDER) Then
            e.Cancel = True
        End If
    End Sub


    ' ========================================================================
    ' 🖱️ КОНТЕКСТНО МЕНЮ (МЕНЮ ПРИ ДЕСЕН БУТОН)
    ' ========================================================================
    Private Sub InitializeContextMenu()
        _contextMenu = New ContextMenuStrip()

        Dim menuAddPanel As New ToolStripMenuItem("➕ Добави под-табло", Nothing, AddressOf MenuAddPanel_Click)
        Dim menuDelete As New ToolStripMenuItem("❌ Изтрий избран елемент", Nothing, AddressOf MenuDelete_Click)
        Dim menuExpandNode As New ToolStripMenuItem("➕ Разгъни този възел", Nothing, AddressOf MenuExpandNode_Click)
        Dim menuCollapseNode As New ToolStripMenuItem("➖ Свий този възел", Nothing, AddressOf MenuCollapseNode_Click)
        Dim menuExpandAll As New ToolStripMenuItem("📂 Разгъни всичко", Nothing, AddressOf MenuExpandAll_Click)
        Dim menuCollapseAll As New ToolStripMenuItem("📁 Свий всичко", Nothing, AddressOf MenuCollapseAll_Click)

        _contextMenu.Items.Add(menuAddPanel)
        _contextMenu.Items.Add(menuDelete)
        _contextMenu.Items.Add(New ToolStripSeparator())
        _contextMenu.Items.Add(menuExpandNode)
        _contextMenu.Items.Add(menuCollapseNode)
        _contextMenu.Items.Add(New ToolStripSeparator())
        _contextMenu.Items.Add(menuExpandAll)
        _contextMenu.Items.Add(menuCollapseAll)

        _tv.ContextMenuStrip = _contextMenu
    End Sub

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
        _tv.EndUpdate()
    End Sub
    ' ========================================================================
    ' 🔄 ГЕНЕРИРАНЕ И ОБНОВЯВАНЕ НА ДЪРВОТО
    ' ========================================================================
    Public Sub RefreshTree()
        _tv.BeginUpdate()
        _tv.Nodes.Clear()
        Try
            Dim buildingNodes As New Dictionary(Of String, TreeNode)(StringComparer.OrdinalIgnoreCase)
            Dim allTabloNodes As New Dictionary(Of String, TreeNode)(StringComparer.OrdinalIgnoreCase)
            Dim tabloToBuilding As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Dim tabloToParent As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            ' 1. ПЪРВИ ПАС: СЪЗДАВАНЕ НА СГРАДИТЕ
            For Each item In AppSettings.ListTokow
                Dim bName = If(String.IsNullOrWhiteSpace(item.BuildingName), ROOT_NODE_TEXT, item.BuildingName.Trim())
                If Not buildingNodes.ContainsKey(bName) Then
                    Dim bNode As New TreeNode($"{ICON_BUILDING} {bName}")
                    bNode.ForeColor = _tv.ForeColor
                    _tv.Nodes.Add(bNode)
                    buildingNodes.Add(bName, bNode)
                End If
            Next
            ' 2. ВТОРИ ПАС: СЪЗДАВАНЕ НА ТАБЛАТА
            For Each item In AppSettings.ListTokow
                If item.Device IsNot Nothing AndAlso item.Device.Trim().Equals("Табло", StringComparison.OrdinalIgnoreCase) Then
                    Dim bName = If(String.IsNullOrWhiteSpace(item.BuildingName), ROOT_NODE_TEXT, item.BuildingName.Trim())
                    Dim tName = If(item.Tablo Is Nothing, "", item.Tablo.Trim())
                    Dim tabloKey = bName & "_" & tName
                    If Not allTabloNodes.ContainsKey(tabloKey) Then
                        Dim tNode As New TreeNode(FormatPanelText(item))
                        tNode.ForeColor = _tv.ForeColor
                        tNode.Tag = item
                        allTabloNodes(tabloKey) = tNode
                        tabloToBuilding(tabloKey) = bName
                        tabloToParent(tabloKey) = ""
                    End If
                    If Not String.IsNullOrWhiteSpace(item.Табло_Родител) AndAlso item.Табло_Родител.Trim() <> ROOT_NODE_TEXT Then
                        tabloToParent(tabloKey) = bName & "_" & item.Табло_Родител.Trim()
                    End If
                End If
            Next
            ' 3. ТРЕТИ ПАС: ЙЕРАРХИЧНО СВЪРЗВАНЕ НА ТАБЛАТА
            For Each tabloKey In allTabloNodes.Keys
                Dim currentNode = allTabloNodes(tabloKey)
                Dim parentKey = tabloToParent(tabloKey)
                Dim bName = tabloToBuilding(tabloKey)

                If Not String.IsNullOrEmpty(parentKey) AndAlso parentKey <> tabloKey AndAlso allTabloNodes.ContainsKey(parentKey) Then
                    allTabloNodes(parentKey).Nodes.Add(currentNode)
                ElseIf tabloKey.EndsWith("_Гл.Р.Т.", StringComparison.OrdinalIgnoreCase) Then
                    If buildingNodes.ContainsKey(bName) Then
                        buildingNodes(bName).Nodes.Add(currentNode)
                    End If
                Else
                    Dim glrtKey = bName & "_Гл.Р.Т."
                    If allTabloNodes.ContainsKey(glrtKey) Then
                        allTabloNodes(glrtKey).Nodes.Add(currentNode)
                    ElseIf buildingNodes.ContainsKey(bName) Then
                        buildingNodes(bName).Nodes.Add(currentNode)
                    End If
                End If
            Next
            ' 4. ЧЕТВЪРТИ ПАС: ДОБАВЯНЕ НА КОНСУМАТОРИ
            Dim consumerSums As New Dictionary(Of TreeNode, Double)
            For Each item In AppSettings.ListTokow
                If item.Device Is Nothing OrElse Not item.Device.Trim().Equals("Табло", StringComparison.OrdinalIgnoreCase) Then
                    Dim bName = If(String.IsNullOrWhiteSpace(item.BuildingName), ROOT_NODE_TEXT, item.BuildingName.Trim())
                    Dim tName = If(item.Tablo Is Nothing, "", item.Tablo.Trim())
                    Dim tabloKey = bName & "_" & tName
                    If allTabloNodes.ContainsKey(tabloKey) Then
                        Dim parentPanelNode = allTabloNodes(tabloKey)
                        Dim consumerGroupNode As TreeNode = Nothing
                        For Each node In parentPanelNode.Nodes
                            If node.Text.Contains(CONSUMERS_NODE_TEXT) Then
                                consumerGroupNode = node
                                Exit For
                            End If
                        Next
                        If consumerGroupNode Is Nothing Then
                            consumerGroupNode = New TreeNode($"{ICON_FOLDER} {CONSUMERS_NODE_TEXT}")
                            consumerGroupNode.ForeColor = _tv.ForeColor
                            parentPanelNode.Nodes.Add(consumerGroupNode)
                            consumerSums(consumerGroupNode) = 0
                        End If
                        Dim cNode As New TreeNode(FormatCircuitText(item))
                        cNode.ForeColor = _tv.ForeColor
                        cNode.Tag = item
                        consumerGroupNode.Nodes.Add(cNode)
                        Dim power As Double = 0
                        Double.TryParse(item.Мощност.ToString(), power)
                        consumerSums(consumerGroupNode) += power
                    End If
                End If
            Next
            ' ОБНОВЯВАНЕ НА МОЩНОСТИТЕ НА ПАПКИТЕ
            Dim formatSpecifier As String = "F" & DECIMAL_PLACES
            For Each groupNode In consumerSums.Keys
                Dim sumValue As Double = consumerSums(groupNode)
                groupNode.Text = $"{ICON_FOLDER} {CONSUMERS_NODE_TEXT} ({sumValue.ToString(formatSpecifier)} {POWER_UNIT})"
            Next
        Finally
            ' ИНТЕНЛИГЕНТНО РАЗГЪВАНЕ СПОРЕД СЪСТОЯНИЕТО
            If _isFirstLoad Then
                _tv.CollapseAll()
                For Each rootNode As TreeNode In _tv.Nodes
                    rootNode.Expand()
                    For Each panelNode As TreeNode In rootNode.Nodes
                        panelNode.Collapse()
                    Next
                Next
                _isFirstLoad = False
            Else
                For Each rootNode As TreeNode In _tv.Nodes
                    rootNode.ExpandAll()
                Next
            End If
            _tv.EndUpdate()
        End Try
    End Sub
    ' ========================================================================
    ' 🚚 DRAG & DROP ЛОГИКА
    ' ========================================================================
    Private Sub HandleItemDrag(sender As Object, e As ItemDragEventArgs)
        _tv.SelectedNode = DirectCast(e.Item, TreeNode)
        _tv.DoDragDrop(e.Item, DragDropEffects.Move)
    End Sub
    Private Sub HandleDragEnter(sender As Object, e As DragEventArgs)
        e.Effect = DragDropEffects.Move
    End Sub
    Private Sub HandleDragOver(sender As Object, e As DragEventArgs)
        Dim targetPoint As Point = _tv.PointToClient(New Point(e.X, e.Y))
        _tv.SelectedNode = _tv.GetNodeAt(targetPoint)
    End Sub
    Private Sub HandleDragDrop(sender As Object, e As DragEventArgs)
        Dim targetPoint As Point = _tv.PointToClient(New Point(e.X, e.Y))
        Dim targetNode As TreeNode = _tv.GetNodeAt(targetPoint)
        Dim draggedNode As TreeNode = DirectCast(e.Data.GetData(GetType(TreeNode)), TreeNode)

        ' Защита: невалидни възли или пускане върху себе си
        If draggedNode Is Nothing OrElse targetNode Is Nothing OrElse draggedNode.Equals(targetNode) Then Return

        Dim sourceObj As clsTokow = TryCast(draggedNode.Tag, clsTokow)
        Dim targetObj As clsTokow = TryCast(targetNode.Tag, clsTokow)

        ' Влаченият обект задължително трябва да има данни
        If sourceObj Is Nothing Then Return

        ' СЛУЧАЙ А: Пускане директно върху Сграда (Корен)
        If targetNode.Level = 0 Then
            Dim buildingTarget As New clsTokow()
            buildingTarget.BuildingName = targetNode.Text.Replace(ICON_BUILDING, "").Trim()
            buildingTarget.Device = "Сграда"
            buildingTarget.Tablo = ""

            RaiseEvent RequestMoveObject(sourceObj, buildingTarget)
            Return
        End If

        ' СЛУЧАЙ Б: Пускане върху друго валидно Табло
        If targetObj IsNot Nothing AndAlso targetObj.Device IsNot Nothing AndAlso
           targetObj.Device.Trim().Equals("Табло", StringComparison.OrdinalIgnoreCase) Then

            RaiseEvent RequestMoveObject(sourceObj, targetObj)
            Return
        End If

        ' СЛУЧАЙ В: Токови кръгове и папки се подминават тихо и безопасно
    End Sub
End Class