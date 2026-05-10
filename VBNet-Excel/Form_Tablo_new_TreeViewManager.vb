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
    ''' Преизчислява общите стойности (мощност и ток) за родителското табло.
    ''' Логиката:
    ''' 1. Намира главния запис на таблото
    ''' 2. Взима всички директни подчинени елементи ("Дете")
    ''' 3. Сумира мощност и ток от всички деца
    ''' 4. Записва резултата в родителския запис
    ''' 
    ''' Използва се като част от йерархичното преизчисляване на таблата.
    ''' </summary>
    Private Sub RecalculateParentSummary(rootName As String)
        ' Намираме основния запис на родителското табло
        Dim parentRecord =
        dataList.FirstOrDefault(Function(x)
                                    Return x.Tablo = rootName AndAlso
                                    x.Device = "Табло"
                                End Function)
        ' Ако няма такова табло → прекратяваме
        If parentRecord Is Nothing Then Return
        ' Взимаме всички директни деца на таблото
        Dim children = dataList.Where(
                            Function(x)
                                Return x.Tablo = rootName AndAlso
                                       x.Device = "Дете"
                            End Function)
        ' Сумираме мощността от всички деца
        parentRecord.Мощност = children.Sum(Function(c) c.Мощност)
        ' Сумираме тока от всички деца
        parentRecord.Ток = children.Sum(Function(c) c.Ток)
    End Sub
    ''' <summary>
    ''' КЛАС ЗА УПРАВЛЕНИЕ НА ЕЛЕКТРИЧЕСКИ ТАБЛА И ЙЕРАРХИЯ ОТ ТОКОВИ КРЪГОВЕ.
    ''' 
    ''' Основни отговорности:
    ''' - Зареждане и съхранение на проектни данни (dataList)
    ''' - Генериране и поддръжка на йерархична структура от табла и подтабла
    ''' - Автоматично създаване на "фийдъри" (връзки между табла)
    ''' - Рекурсивно преизчисляване на мощност и ток по йерархията
    ''' - Изграждане на визуално дърво (TreeView)
    ''' - Управление на JSON сериализация/десериализация на проектите
    ''' - Извличане и нормализиране на име на сграда от DWG файл
    ''' 
    ''' Класът служи като централен слой между:
    ''' UI (TreeView / Forms)
    ''' и
    ''' бизнес логиката за електрическите табла.
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
    ''' Намира всички директно свързани подтабла (деца) към дадено родителско табло.
    ''' 
    ''' Логика:
    ''' - Търси записи в dataList с Device = "Дете"
    ''' - Филтрира по Tablo = parentName
    ''' - Връща уникалните стойности от полето ТоковКръг (имената на подтаблата)
    ''' - Подрежда резултата по азбучен ред
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
    ''' Намира кореновите табла за дадена сграда.
    ''' 
    ''' Кореново табло е такова, което:
    ''' - принадлежи към дадената сграда
    ''' - не се среща като "дете" (няма родителска връзка в йерархията)
    ''' 
    ''' Логика:
    ''' 1. Взимат се всички табла (Device = "Табло") за конкретната сграда
    ''' 2. Взимат се всички табла, които са маркирани като "Дете"
    ''' 3. От първия списък се изключват всички деца
    ''' 4. Резултатът се подрежда по азбучен ред
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
    ''' Рекурсивно изгражда възел в TreeView за дадено табло и всички негови подтабла.
    ''' 
    ''' Структура:
    ''' - Създава възел за текущото табло
    ''' - Добавя подтабла чрез рекурсия (йерархично)
    ''' - Групира всички токови кръгове в отделен под-възел
    ''' - Създава листа за всеки токов кръг
    ''' - Сгъва възлите за по-компактен визуален изглед
    ''' 
    ''' Логика:
    ''' 1. Намира таблото в dataList
    ''' 2. Изчислява общата мощност (включително деца)
    ''' 3. Добавя рекурсивно всички подтабла
    ''' 4. Добавя токовите кръгове като отделна група
    ''' 5. Сгъва структурата за по-добра четимост
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
        ' РЕКУРСИЯ за подтаблата
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
            circuitsNode.ForeColor = Color.DarkBlue
            circuitsNode.BackColor = Color.Ivory
            circuitsNode.NodeFont = New Font(tv.Font, FontStyle.Bold)
            pNode.Nodes.Add(circuitsNode)
            For Each tok In circuits
                Dim tkNode As New TreeNode($"{tok.ТоковКръг} ({tok.Мощност:F2} kW)")
                circuitsNode.Nodes.Add(tkNode)
            Next
            circuitsNode.Collapse()
        End If
        pNode.Collapse()
    End Sub
    ''' <summary>
    ''' Изчислява общата мощност на дадено табло.
    ''' 
    ''' Включва:
    ''' - Мощност от подтабла (чрез "Дете" връзки / фийдъри)
    ''' - Мощност от директно свързани токови кръгове
    ''' 
    ''' Резултатът представлява сумарното натоварване на таблото.
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



    Private Sub Tv_ItemDrag(sender As Object, e As ItemDragEventArgs)
        ' Започваме Drag операция с избрания възел
        ' AllowDrop ефект: Copy (не местим, а копираме референцията)
        tv.DoDragDrop(e.Item, DragDropEffects.Move)
    End Sub
    Private Sub Tv_DragEnter(sender As Object, e As DragEventArgs)
        ' Проверяваме дали влаченият обект е TreeNode
        If e.Data.GetDataPresent(GetType(TreeNode)) Then
            ' Позволяваме Move операция
            e.Effect = DragDropEffects.Move
        Else
            ' Не позволяваме drop на други обекти
            e.Effect = DragDropEffects.None
        End If
    End Sub
    Private Sub Tv_DragOver(sender As Object, e As DragEventArgs)
        ' Получаваме точката, където е мишката (в координати на TreeView)
        Dim targetPoint As Point = tv.PointToClient(New Point(e.X, e.Y))
        ' Намираме възела под мишката
        Dim targetNode As TreeNode = tv.GetNodeAt(targetPoint)
        If targetNode IsNot Nothing Then
            ' Получаваме влачения възел
            Dim draggedNode As TreeNode = CType(e.Data.GetData(GetType(TreeNode)), TreeNode)
            ' ВАЛИДАЦИЯ: Не позволяваме да drop-нем възел върху себе си или свои деца
            If targetNode Is draggedNode OrElse IsChildOf(draggedNode, targetNode) Then
                e.Effect = DragDropEffects.None
                Return
            End If
            ' ВАЛИДАЦИЯ: Проверяваме дали и двата възела са табла (не "Токови кръгове")
            If Not IsPanelNode(draggedNode) OrElse Not IsPanelNode(targetNode) Then
                e.Effect = DragDropEffects.None
                Return
            End If
            ' Всичко е наред - позволяваме drop
            e.Effect = DragDropEffects.Move
            ' Визуална обратна връзка: маркираме целевия възел
            tv.SelectedNode = targetNode
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub
    Private Function IsChildOf(node As TreeNode, parentNode As TreeNode) As Boolean
        Dim parent As TreeNode = node.Parent
        While parent IsNot Nothing
            If parent Is parentNode Then Return True
            parent = parent.Parent
        End While
        Return False
    End Function
    Private Function IsPanelNode(node As TreeNode) As Boolean
        Return node.Text.Contains("kW)") AndAlso Not node.Text.Contains("Токови кръгове")
    End Function
    Private Function ExtractPanelName(nodeText As String) As String
        If String.IsNullOrEmpty(nodeText) Then Return String.Empty
        ' Намираме позицията на " ("
        Dim idx = nodeText.IndexOf(" (")
        If idx > 0 Then
            Return nodeText.Substring(0, idx).Trim()
        End If
        Return nodeText.Trim()
    End Function
    Private Sub Tv_DragDrop(sender As Object, e As DragEventArgs)
        ' 1. Взимаме възлите
        Dim draggedNode As TreeNode = CType(e.Data.GetData(GetType(TreeNode)), TreeNode)
        Dim targetPoint As Point = tv.PointToClient(New Point(e.X, e.Y))
        Dim targetNode As TreeNode = tv.GetNodeAt(targetPoint)
        If draggedNode Is Nothing OrElse targetNode Is Nothing OrElse draggedNode Is targetNode Then Return
        ' 2. Валидация: само табла могат да се местят
        If Not IsPanelNode(draggedNode) OrElse Not IsPanelNode(targetNode) Then Return
        If IsChildOf(draggedNode, targetNode) Then Return ' Не позволяваме циклична зависимост
        ' 3. Извличаме имената на таблата от текста (формат: "T-1 (15.54 kW)")
        Dim draggedPanelName As String = ExtractPanelName(draggedNode.Text)
        Dim targetPanelName As String = ExtractPanelName(targetNode.Text)
        If String.IsNullOrEmpty(draggedPanelName) OrElse String.IsNullOrEmpty(targetPanelName) Then Return
        ' 4. 🔥 АКТУАЛИЗИРАМЕ ДАННИТЕ (ListTokow)
        ' Намираме всички записи, които принадлежат на местеното табло
        Dim panelRecords = dataList.Where(Function(x) x.Tablo = draggedPanelName AndAlso x.Device = "Табло").ToList()
        If panelRecords.Count = 0 Then Return
        For Each rec In panelRecords
            ' Променяме родителската връзка
            rec.Табло_Родител = targetPanelName
        Next
        ' 5. АКТУАЛИЗИРАМЕ ВРЪЗКИТЕ (фийдърите)
        ' Премахваме старата връзка (ако съществува)
        Dim oldFeeders = dataList.Where(Function(x) x.Device = "Дете" AndAlso
                                        String.Equals(x.ТоковКръг, draggedPanelName, StringComparison.OrdinalIgnoreCase)).ToList()
        For Each f In oldFeeders
            dataList.Remove(f)
        Next
        ' Създаваме нова връзка под новия родител
        Dim masterRecord = panelRecords.FirstOrDefault()
        If masterRecord IsNot Nothing Then
            Dim newFeeder As Form_Tablo_new.strTokow = masterRecord.Clone()
            newFeeder.Tablo = targetPanelName
            newFeeder.Device = "Дете"
            newFeeder.ТоковКръг = draggedPanelName
            newFeeder.Табло_Родител = ""
            newFeeder.Консуматор = "Табло"
            newFeeder.предназначение = draggedPanelName
            dataList.Add(newFeeder)
        End If
        ' 6. ПРЕИЗЧИСЛЯВАНЕ НА МОЩНОСТИТЕ
        ' Преизчисляваме стария родител (ако го намерим)
        Dim oldParentName = panelRecords.FirstOrDefault()?.Табло_Родител
        If Not String.IsNullOrEmpty(oldParentName) Then
            RecalculateParentSummary(oldParentName)
        End If
        ' Преизчисляваме новия родител
        RecalculateParentSummary(targetPanelName)
        ' 7. 🔄 ОБНОВЯВАНЕ НА TREEVIEW
        ' Премахваме възела от старото място
        If draggedNode.Parent IsNot Nothing Then
            draggedNode.Parent.Nodes.Remove(draggedNode)
        End If
        ' Добавяме го под новия родител
        targetNode.Nodes.Add(draggedNode)
        ' Маркираме и разгъваме новия родител за визуална обратна връзка
        tv.SelectedNode = targetNode
        targetNode.Expand()
        ' 8. (Опционално) Информиране на формата, че данните са променени
        ' Ако формата има нужда да знае, че ListTokow е променен, тук може да се вдигне събитие
    End Sub
End Class