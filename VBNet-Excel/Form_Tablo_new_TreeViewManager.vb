Imports System.Windows.Forms
Imports System.Linq
Imports System.Drawing

Public Class Form_Tablo_new_TreeViewManager
    ' Референции към контролата и данните
    Private tv As TreeView
    Private dataList As List(Of Form_Tablo_new.strTokow)
    Private rootText As String
    ' ========================================================================
    ' 🎨 UI КОНСТАНТИ (лесни за промяна на едно място)
    ' ========================================================================
    ' ========================================================================
    ' 🎨 UI КОНСТАНТИ & ШАБЛОНИ (едно място за всички визуални промени)
    ' ========================================================================
    Private Const ICON_BUILDING As String = "🏢"        ' Сграда / Комплекс
    Private Const ICON_PANEL As String = "🗄️"           ' Разпределително табло
    Private Const ICON_CIRCUITS As String = "🔵"        ' Група токови кръгове
    Private Const LABEL_CIRCUITS As String = "Т.К."     ' Име на групата токови кръгове
    Private Const POWER_UNIT As String = "kW"

    '  ШАБЛОНИ ЗА ФОРМАТИРАНЕ (използват String.Format)
    Private ReadOnly Property BuildingTemplate As String = ICON_BUILDING & " {0} ({1:F2} " & POWER_UNIT & ")"
    Private ReadOnly Property PanelTemplate As String = ICON_PANEL & " {0} ({1:F2} " & POWER_UNIT & ")"
    Private ReadOnly Property CircuitsTemplate As String = ICON_CIRCUITS & " " & LABEL_CIRCUITS & " ({0:F2} " & POWER_UNIT & ")"
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
        ' ✅ 1. ПЪРВО ЗАПАЗВАМЕ СЪСТОЯНИЕТО (за да е достъпно навсякъде в класа)
        rootText = rootNodeText
        ' 2. Подготвяме логическите връзки между таблата
        PreparePanelFeeders(rootText)
        ' ✅ НОВО: Гарантираме съществуването на кореновия запис
        EnsureRootPanelExists(rootText)
        ' 3. Преизчисляваме всички родителски суми по йерархията
        UpdatePanelSummary(rootNodeText)
        ' 4. Изграждаме визуалното дърво
        BuildTree(rootText)
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
        tv.BeginUpdate()
        Try
            ' ✅ ПРОВЕРКА: Колко сгради има?
            Dim distinctBuildings = dataList.
            Where(Function(x) Not String.IsNullOrEmpty(x.BuildingName)).
            Select(Function(x) x.BuildingName).
            Distinct().Count()

            If distinctBuildings <= 1 Then
                ' 🏢 Една сграда → Синхронизираме данните и взимаме общата мощност
                UpdatePanelSummary(rootNodeText)
                ' Взимаме актуализираната мощност от списъка (ако няма запис → 0)
                Dim rootRecord = dataList.FirstOrDefault(Function(x) x.Tablo = rootNodeText AndAlso x.Device = "Табло")
                Dim totalPower As Double = If(rootRecord?.Мощност, 0)
                ' Създаваме името на корена: "Име (0.00 kW)"
                Dim nodeName As String = $"{rootNodeText} ({totalPower:F2} kW)"
                Dim rootNode As New TreeNode(nodeName)
                tv.Nodes.Add(rootNode)
                ' Намираме всички табла под корена
                Dim rootPanels = FindChildPanels(rootNodeText)
                For Each panelName In rootPanels
                    AddPanelNodeRecursive(panelName, rootNode.Nodes)
                Next
            Else
                ' 🏘️ Няколко сгради
                ' Тук можеш да направиш същото за всяка сграда в нейния собствен цикъл
                Dim buildingGroups = dataList.GroupBy(Function(x) x.BuildingName)
                For Each bGrp In buildingGroups
                    Dim bName = If(String.IsNullOrEmpty(bGrp.Key), "Неизвестна сграда", bGrp.Key)

                    ' Мощност за конкретната сграда (само кореновите табла за нея)
                    Dim bPower = bGrp.Where(Function(x) x.Device = "Табло" AndAlso
                                        Not dataList.Any(Function(d) d.Device = "Дете" AndAlso d.ТоковКръг = x.Tablo)).
                                  Sum(Function(x) x.Мощност)
                    Dim bNode As New TreeNode(String.Format(BuildingTemplate, bName, bPower))
                    'Dim bNode As New TreeNode($"{bName} ({bPower:F2} kW)")
                    tv.Nodes.Add(bNode)

                    Dim rootPanels = FindRootPanelsForBuilding(bName)
                    For Each panelName In rootPanels
                        AddPanelNodeRecursive(panelName, bNode.Nodes)
                    Next
                Next
            End If
        Finally
            If tv.Nodes.Count > 0 Then
                tv.Nodes(0).Expand()
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
        ' Dim pNode As New TreeNode($"{panelName} ({totalPower:F2} kW)")
        Dim pNode As New TreeNode(String.Format(PanelTemplate, panelName, totalPower))
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
            'Dim circuitsNode As New TreeNode($"🔵 Т.К.({totalCircuitPower:F2} kW)")
            Dim circuitsNode As New TreeNode(String.Format(CircuitsTemplate, totalCircuitPower))
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
    ''' <summary>
    ''' Преизчислява и актуализира обобщените стойности за дадено табло.
    ''' Сумира мощност, брой контакти и брой лампи от всички записи под него,
    ''' като игнорира самите обобщения (Device = "Табло").
    ''' Поддържа филтър по сграда за бъдеща универсалност.
    ''' </summary>
    Private Sub UpdatePanelSummary(panelName As String, Optional buildingName As String = Nothing)
        ' 1. Намираме целевия запис (Device = "Табло")
        Dim targetRecord = dataList.FirstOrDefault(Function(x) x.Tablo = panelName AndAlso x.Device = "Табло")
        If targetRecord Is Nothing Then Return

        ' 2. Филтрираме източника по сграда, ако е подадена (за бъдеща поддръжка на много сгради)
        Dim sourceData = If(String.IsNullOrEmpty(buildingName),
                            dataList,
                            dataList.Where(Function(x) x.BuildingName = buildingName))

        ' 3. Взимаме всички записи под това табло, БЕЗ да броим самото обобщение
        Dim itemsToSum = sourceData.Where(Function(x) x.Tablo = panelName AndAlso x.Device <> "Табло")

        ' 4. Изчисляваме сумите (LINQ автоматично връща 0 за празни колекции)
        Dim totalPower As Double = itemsToSum.Sum(Function(x) x.Мощност)
        Dim totalKontakt As Integer = itemsToSum.Sum(Function(x) x.brKontakt)
        Dim totalLamp As Integer = itemsToSum.Sum(Function(x) x.brLamp)

        ' 5. Записваме резултатите обратно в обекта
        targetRecord.Мощност = totalPower
        targetRecord.brKontakt = totalKontakt
        targetRecord.brLamp = totalLamp
    End Sub
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
        Dim draggedNode As TreeNode = CType(e.Data.GetData(GetType(TreeNode)), TreeNode)
        Dim targetPoint As Point = tv.PointToClient(New Point(e.X, e.Y))
        Dim targetNode As TreeNode = tv.GetNodeAt(targetPoint)
        If draggedNode Is Nothing OrElse targetNode Is Nothing OrElse draggedNode Is targetNode Then Return
        If Not IsPanelNode(draggedNode) OrElse Not IsPanelNode(targetNode) Then Return
        If IsChildOf(draggedNode, targetNode) Then Return
        Dim draggedPanelName As String = ExtractPanelName(draggedNode.Text)
        Dim targetPanelName As String = ExtractPanelName(targetNode.Text)
        If String.IsNullOrEmpty(draggedPanelName) OrElse String.IsNullOrEmpty(targetPanelName) Then Return
        ' 1. Намираме мастер записа на местеното табло
        Dim draggedMaster = dataList.FirstOrDefault(Function(x) x.Tablo = draggedPanelName AndAlso x.Device = "Табло")
        If draggedMaster Is Nothing Then Return
        ' 2. Запомняме стария родител преди промяна
        Dim oldParentName As String = draggedMaster.Табло_Родител
        ' 3. Актуализираме йерархията в данните
        draggedMaster.Табло_Родител = targetPanelName
        ' 4. Премахваме всички стари фийдъри, сочещи към това табло
        Dim oldFeeders = dataList.Where(Function(x) x.Device = "Дете" AndAlso
                                        String.Equals(x.ТоковКръг, draggedPanelName, StringComparison.OrdinalIgnoreCase)).ToList()
        For Each f In oldFeeders
            dataList.Remove(f)
        Next
        ' 5. Създаваме нов фийдър (връзка) под новия родител
        Dim newFeeder As Form_Tablo_new.strTokow = draggedMaster.Clone()
        newFeeder.Device = "Дете"
        newFeeder.Tablo = targetPanelName
        newFeeder.ТоковКръг = draggedPanelName
        newFeeder.Табло_Родител = ""
        newFeeder.Консуматор = "Табло"
        newFeeder.предназначение = draggedPanelName
        ' 🔥 КЛЮЧОВА ПОПРАВКА: MemberwiseClone копира референцията на списъка.
        ' Трябва да създадем нов празен списък, за да не "споделя" консуматори с оригинала.
        newFeeder.Konsumator = New List(Of Form_Tablo_new.strKonsumator)
        dataList.Add(newFeeder)
        ' 6. Преизчисляване на мощностите
        If Not String.IsNullOrEmpty(oldParentName) AndAlso oldParentName <> targetPanelName Then
            UpdatePanelSummary(oldParentName)
        End If
        UpdatePanelSummary(targetPanelName)
        ' 7. Пълно обновяване на дървото
        BuildTree(rootText)
    End Sub
    ''' <summary>
    ''' Гарантира, че в ListTokow съществува агрегиращ запис за корена.
    ''' Не се дублира при повторно извикване.
    ''' BuildingName се взима от първия наличен запис (консистентно с CreateTokowList).
    ''' </summary>
    Private Sub EnsureRootPanelExists(rootName As String)
        ' Ако вече има запис → излизаме
        If dataList.Any(Function(x) x.Tablo = rootName AndAlso x.Device = "Табло") Then Return

        ' Взимаме BuildingName от първия наличен запис (ако има)
        ' Това гарантира, че коренът е в същата "сграда" като останалите данни
        Dim buildingName = If(dataList.FirstOrDefault()?.BuildingName, String.Empty)

        ' Създаваме синтетичен запис само за сумиране
        Dim rootPanel As New Form_Tablo_new.strTokow()
        rootPanel.Tablo = rootName
        rootPanel.Device = "Табло"
        rootPanel.ТоковКръг = "ОБЩО"
        rootPanel.Мощност = 0
        rootPanel.Ток = 0
        rootPanel.brKontakt = 0
        rootPanel.brLamp = 0
        rootPanel.BuildingName = buildingName
        rootPanel.Табло_Родител = "" ' Коренът няма родител

        dataList.Add(rootPanel)
    End Sub
End Class