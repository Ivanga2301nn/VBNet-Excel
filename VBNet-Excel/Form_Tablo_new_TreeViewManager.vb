Imports System.Windows.Forms
Imports System.Linq
Imports System.Drawing

Public Class Form_Tablo_new_TreeViewManager
    ' Референции към контролата и данните
    Private tv As TreeView
    Private dataList As List(Of Form_Tablo_new.strTokow)

    ''' <summary>
    ''' Конструктор: Приема контролата и списъка с данни
    ''' </summary>
    Public Sub New(treeViewControl As TreeView, data As List(Of Form_Tablo_new.strTokow))
        tv = treeViewControl
        dataList = data
        ' Активиране на Drag & Drop
        tv.AllowDrop = True
        AddHandler tv.ItemDrag, AddressOf Tv_ItemDrag
        AddHandler tv.DragEnter, AddressOf Tv_DragEnter
        AddHandler tv.DragOver, AddressOf Tv_DragOver
        AddHandler tv.DragDrop, AddressOf Tv_DragDrop
    End Sub
    ''' <summary>
    ''' Главен входен метод за инициализация и изграждане на дървовидната структура.
    ''' Логиката:
    ''' 1. Проверява дали има налични данни
    ''' 2. Подготвя връзките между таблата (feeder структура)
    ''' 3. Преизчислява агрегатните стойности по йерархията
    ''' 4. Изгражда визуалното TreeView дърво
    ''' 
    ''' Този метод заменя старите отделни извиквания и гарантира,
    ''' че данните и визуализацията винаги са синхронизирани.
    ''' </summary>
    Public Sub InitializeAndBuild(rootNodeText As String)
        ' Ако няма данни → прекратяваме
        If dataList Is Nothing OrElse dataList.Count = 0 Then Return
        ' Подготвяме логическите връзки между таблата
        PreparePanelFeeders(rootNodeText)
        ' Преизчисляваме всички родителски суми по йерархията
        RecalculateParentSummary(rootNodeText)
        ' Изграждаме визуалното дърво
        BuildTree(rootNodeText)
    End Sub
    ' ========================================================================
    ' 📊 ЛОГИКА ЗА ДАННИ (ПРЕДИ ВИЗУАЛИЗАЦИЯ)
    ' ========================================================================
    ''' <summary>
    ''' Подготвя логическите връзки (feeder-и) между главното табло и подтаблата.
    ''' Логиката:
    ''' 1. Извлича всички уникални табла от данните
    ''' 2. Пропуска кореновото табло
    ''' 3. Проверява дали вече съществува връзка "Дете"
    ''' 4. Ако няма → създава нова връзка към подтабло
    ''' 5. Използва копие на оригиналния запис, за да не се губят данни
    ''' 6. Добавя връзката в основния списък
    ''' 
    ''' Цел: изграждане на йерархия между таблата за последващо изчисление и визуализация
    ''' </summary>
    Private Sub PreparePanelFeeders(rootName As String)
        ' Взимаме всички уникални табла от списъка
        Dim uniquePanels = dataList.
                        Where(Function(x) Not String.IsNullOrWhiteSpace(x.Tablo)).
                        Select(Function(x) x.Tablo.Trim()).
                        Distinct().
                        ToList()
        For Each pName In uniquePanels
            ' Пропускаме кореновото табло
            If String.Equals(pName, rootName, StringComparison.OrdinalIgnoreCase) Then Continue For
            ' Проверка дали вече съществува връзка към това табло
            Dim linkExists = dataList.Any(
                            Function(x)
                                Return x.Tablo = rootName AndAlso
                                       x.Device = "Дете" AndAlso
                                       String.Equals(x.ТоковКръг, pName, StringComparison.OrdinalIgnoreCase)
                            End Function)
            ' Ако връзката не съществува → създаваме я
            If Not linkExists Then
                ' Намираме оригиналния запис на подтаблото
                Dim childMaster = dataList.FirstOrDefault(
                                    Function(x)
                                        Return x.Tablo = pName AndAlso
                                               x.Device = "Табло"
                                    End Function)
                If childMaster IsNot Nothing Then
                    ' Създаваме независимо копие на обекта
                    Dim feeder As Form_Tablo_new.strTokow = childMaster.Clone()
                    ' Превръщаме го във връзка "Дете"
                    feeder.Tablo = rootName
                    feeder.Device = "Дете"
                    feeder.ТоковКръг = pName
                    feeder.Табло_Родител = ""
                    feeder.Консуматор = "Табло"
                    feeder.предназначение = pName
                    ' Добавяме връзката към основния списък
                    dataList.Add(feeder)
                End If
            End If
        Next
    End Sub
    ''' <summary>
    ''' Преизчислява мощност и ток на родителското табло.
    ''' (Преместена логика от BuildPanelSummaryRecord)
    ''' </summary>
    Private Sub RecalculateParentSummary(rootName As String)
        ' Намираме записа на родителското табло
        Dim parentRecord = dataList.FirstOrDefault(Function(x) x.Tablo = rootName AndAlso x.Device = "Табло")
        If parentRecord Is Nothing Then Return
        ' Сумираме мощността и тока от всички деца, които са "Дете" (фийдъри)
        Dim children = dataList.Where(Function(x) x.Tablo = rootName AndAlso x.Device = "Дете")
        parentRecord.Мощност = children.Sum(Function(c) c.Мощност)
        parentRecord.Ток = children.Sum(Function(c) c.Ток)
    End Sub
    ' ========================================================================
    ' 🌲 ЛОГИКА ЗА ВИЗУАЛИЗАЦИЯ (TREEVIEW) - РЕКУРСИВНА
    ' ========================================================================
    ''' <summary>
    ''' Рисува TreeView с автоматична йерархия (линейна или сложна).
    ''' </summary>
    Private Sub BuildTree(rootNodeText As String)
        tv.Nodes.Clear()
        tv.BeginUpdate() ' За по-гладко обновяване
        Try
            ' ✅ ПРОВЕРКА: Колко сгради има?
            Dim distinctBuildings = dataList.
                Where(Function(x) Not String.IsNullOrEmpty(x.BuildingName)).
                Select(Function(x) x.BuildingName).
                Distinct().Count()
            If distinctBuildings <= 1 Then
                ' 🏢 Една сграда → започваме от корена
                Dim rootNode As New TreeNode(rootNodeText)
                tv.Nodes.Add(rootNode)
                ' Намираме всички табла, които са директно под корена
                Dim rootPanels = FindChildPanels(rootNodeText)
                For Each panelName In rootPanels
                    AddPanelNodeRecursive(panelName, rootNode.Nodes)
                Next
            Else
                ' 🏘️ Няколко сгради → първо групираме по сграда
                Dim buildingGroups = dataList.GroupBy(Function(x) x.BuildingName)
                For Each bGrp In buildingGroups
                    Dim bName = If(String.IsNullOrEmpty(bGrp.Key), "Неизвестна сграда", bGrp.Key)
                    Dim bNode As New TreeNode(bName)
                    tv.Nodes.Add(bNode)

                    ' За всяка сграда намираме кореновите табла (тези, които не са деца на други)
                    Dim rootPanels = FindRootPanelsForBuilding(bName)
                    For Each panelName In rootPanels
                        AddPanelNodeRecursive(panelName, bNode.Nodes)
                    Next
                Next
            End If
        Finally
            If tv.Nodes.Count > 0 Then
                tv.Nodes(0).Expand() ' Разгъваме първото ниво за по-добър изглед
            End If
            tv.EndUpdate()
        End Try
    End Sub
    ''' <summary>
    ''' Намира всички табла, които са директно под даден родител
    ''' </summary>
    Private Function FindChildPanels(parentName As String) As List(Of String)
        ' Търсим записи с Device="Дете", където Tablo=parentName
        ' ТоковКръг съдържа името на детското табло
        Return dataList.
            Where(Function(x) x.Device = "Дете" AndAlso
                  String.Equals(x.Tablo, parentName, StringComparison.OrdinalIgnoreCase)).
            Select(Function(x) x.ТоковКръг).
            Distinct().
            OrderBy(Function(x) x).
            ToList()
    End Function
    ''' <summary>
    ''' Намира кореновите табла за дадена сграда (тези, които нямат родител)
    ''' </summary>
    Private Function FindRootPanelsForBuilding(buildingName As String) As List(Of String)
        ' Всички табла в тази сграда
        Dim allPanelsInBuilding = dataList.
            Where(Function(x) x.BuildingName = buildingName AndAlso
                  x.Device = "Табло").
            Select(Function(x) x.Tablo).
            Distinct()
        ' Табла, които са деца на други (имат родител)
        Dim childPanels = dataList.
            Where(Function(x) x.Device = "Дете").
            Select(Function(x) x.ТоковКръг).
            Distinct()
        ' Коренови са тези, които са в сградата, но НЕ са деца
        Return allPanelsInBuilding.Except(childPanels).OrderBy(Function(x) x).ToList()
    End Function
    ''' <summary>
    ''' РЕКУРСИВНО добавя табло и всичките му подтабла.
    ''' Токовите кръгове са групирани под един общ нод (сгънат).
    ''' </summary>
    Private Sub AddPanelNodeRecursive(panelName As String, parentNodes As TreeNodeCollection)
        ' Намираме записа за това табло
        Dim panelRecord = dataList.FirstOrDefault(Function(x) x.Tablo = panelName AndAlso x.Device = "Табло")
        If panelRecord Is Nothing Then Return
        ' Изчисляваме общата мощност (сумираме децата + собствените кръгове)
        Dim totalPower = CalculatePanelTotalPower(panelName)
        ' Създаваме възела за таблото
        Dim pNode As New TreeNode($"{panelName} ({totalPower:F2} kW)")
        parentNodes.Add(pNode)
        ' ✅ РЕКУРСИЯ
        Dim childPanels = FindChildPanels(panelName)
        For Each childName In childPanels
            AddPanelNodeRecursive(childName, pNode.Nodes)
        Next
        ' Добавяме токовите кръгове
        Dim circuits = dataList.Where(Function(x) x.Tablo = panelName AndAlso
                                     x.Device <> "Табло" AndAlso
                                     x.Device <> "Дете" AndAlso
                                     Not String.IsNullOrEmpty(x.ТоковКръг) AndAlso
                                     x.ТоковКръг <> "ОБЩО").ToList()
        If circuits.Count > 0 Then
            Dim totalCircuitPower = circuits.Sum(Function(c) c.Мощност)
            Dim circuitsNode As New TreeNode($"🔵 Токови кръгове ({totalCircuitPower:F2} kW)")
            ' --- ШАРЕНИЯ ---
            circuitsNode.ForeColor = Color.DarkBlue     ' Тъмно син цвят на текста
            circuitsNode.BackColor = Color.Ivory        ' Нежен фонов цвят
            circuitsNode.NodeFont = New Font(tv.Font, FontStyle.Bold) ' Удебелен шрифт
            ' ----------------
            pNode.Nodes.Add(circuitsNode)
            For Each tok In circuits
                Dim tkNode As New TreeNode($"{tok.ТоковКръг} ({tok.Мощност:F2} kW)")
                circuitsNode.Nodes.Add(tkNode)
            Next
            ' Сгъваме възела с токовите кръгове
            circuitsNode.Collapse()
        End If
        ' 🔥 КЛЮЧЪТ: Сгъваме самото табло (Т-маг), след като всичко в него е заредено
        pNode.Collapse()
    End Sub
    ''' <summary>
    ''' Изчислява общата мощност на табло (деца + собствени кръгове)
    ''' </summary>
    Private Function CalculatePanelTotalPower(panelName As String) As Double
        ' Мощност от подтабла (чрез фийдърите)
        Dim feederPower = dataList.
            Where(Function(x) x.Tablo = panelName AndAlso x.Device = "Дете").
            Sum(Function(x) x.Мощност)
        ' Мощност от собствени кръгове
        Dim circuitPower = dataList.
            Where(Function(x) x.Tablo = panelName AndAlso
                  x.Device <> "Табло" AndAlso
                  x.Device <> "Дете").
            Sum(Function(x) x.Мощност)
        Return feederPower + circuitPower
    End Function
End Class