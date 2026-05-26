Imports System.Windows.Forms

Public Class PanelBalanceManager
    ' 1. Пазим локални референции на ниво клас
    Private _rcdCatalog As RCDCatalog
    Private _listTokow As List(Of strTokow)
    Private _disconnectorCatalog As DisconnectorCatalog
    Private _cableCatalog As CableCatalog
    Private _electricalCalculationEngine As ElectricalCalculationEngine
    ''' <summary>
    ''' КОНСТРУКТОР: Приема създадените каталози и списъка с токови кръгове от формата
    ''' </summary>
    Public Sub New(rcdCat As RCDCatalog,
                   tokowList As List(Of strTokow),
                   disconnectorCat As DisconnectorCatalog,
                   cableCat As CableCatalog,
                   electricalCalcEngine As ElectricalCalculationEngine)
        Me._rcdCatalog = rcdCat                                 ' 👈 Запаметяваме каталога с RCD в класа при създаването му
        Me._listTokow = tokowList                               ' 👈 Запаметяваме списъка в класа при създаването му
        Me._disconnectorCatalog = disconnectorCat               ' 👈 Запаметяваме каталога с прекъсвачи в класа при създаването му
        Me._cableCatalog = cableCat                             ' 👈 Запаметяваме каталога с кабели в класа при създаването му   
        Me._electricalCalculationEngine = electricalCalcEngine  ' 👈 Запаметяваме електрическия калкулатор в класа при създаването му     
    End Sub
    ''' <summary>
    ''' Клас за групиране на токови кръгове за балансиране на фазите
    ''' </summary>
    Public Class BalanceGroup
        Public Circuits As List(Of strTokow) ' Списък с токови кръгове в групата
        Public GroupType As String ' Тип на групата: "ThreePhase", "RCD", "SmallBus", "LargeBus", "Normal"
        Public GroupKey As String ' Ключ на групата: RCD_Нула (N1, N2...), "Bus" или Nothing
        Public TotalCurrent As Double ' Сумарен ток на групата (сума от токовете на всички ТК)
        Public AssignedPhase As String ' Зададена фаза след балансиране (L1, L2, L3 или "L1,L2,L3")
        Public Sub New()
            Circuits = New List(Of strTokow) ' Конструктор - инициализира списъка с ТК
        End Sub
        Public ReadOnly Property CircuitCount As Integer ' Брой токови кръгове в групата
            Get
                Return Circuits.Count
            End Get
        End Property
        Public ReadOnly Property TotalPower As Double ' Сумарна мощност на групата (сума от мощностите на всички ТК)
            Get
                Return Circuits.Sum(Function(t) t.Мощност)
            End Get
        End Property
    End Class
#Region "📂 EnsureTotalRecordsExists"
    ' Това е "Архитектът на сградите" - неговата единствена задача е да намери сградите
    Public Sub AddFeederRecords()
        Dim buildings As List(Of String) = _listTokow.Select(Function(t) t.BuildingName).Distinct().ToList()
        For Each bName As String In buildings
            ' Делегираме цялата логика за топологията на сградата
            ProcessBuildingTopology(bName)
        Next
    End Sub
    Private Sub ProcessBuildingTopology(buildingName As String)
        ' 1. Създаваме речник за нивата (кеш)
        Dim levels As New Dictionary(Of String, Integer)
        Dim allPanels = _listTokow.Where(Function(t) t.BuildingName = buildingName AndAlso t.Device = "Табло").ToList()
        ' 2. Попълваме речника (пресмятаме само веднъж за всяко табло)
        For Each p In allPanels
            levels(p.Tablo) = CalculateLevel(buildingName, p.Tablo)
        Next
        ' 3. Сортираме, използвайки речника (много бързо!)
        Dim sortedPanels = allPanels.OrderByDescending(Function(t) levels(t.Tablo)) _
                                .Select(Function(t) t.Tablo).Distinct().ToList()
        ' 4. Оркестрация
        For Each tName As String In sortedPanels
            BuildPanelSummaryRecord(buildingName, tName)
        Next
    End Sub
    ' Помощна функция само за изчислението, която ползваме за кеширането
    Private Function CalculateLevel(buildingName As String, tabloName As String) As Integer
        ' Тук остава твоята логика с While, тя е стабилна
        Dim level As Integer = 0
        Dim currentName As String = tabloName
        While True
            'Dim t = ListTokow.FirstOrDefault(Function(x) _
            '                                     x.BuildingName = buildingName AndAlso
            '                                     x.Tablo = currentName)
            ' Смени този ред вътре в CalculateLevel:
            Dim t = _listTokow.FirstOrDefault(Function(x) x.BuildingName = buildingName AndAlso
                                             x.Tablo = currentName AndAlso
                                             x.ТоковКръг = "ОБЩО")
            If t Is Nothing OrElse String.IsNullOrEmpty(t.Табло_Родител) Then Exit While
            level += 1
            currentName = t.Табло_Родител
        End While
        Return level
    End Function
    ''' <summary>
    ''' Изгражда обобщен запис "ОБЩО" за дадено табло.
    ''' Логиката включва:
    ''' - събиране на всички кръгове
    ''' - изчисляване на мощности и консуматори
    ''' - определяне на фази и полюси
    ''' - намиране или създаване на запис "ОБЩО"
    ''' - изчисляване на ток и избор на апаратура
    ''' </summary>
    Private Sub BuildPanelSummaryRecord(buildingName As String, tabloName As String)
        ' ВЗИМАМЕ САМО ДИРЕКТНИТЕ ДЕЦА
        ' 1. Преки консуматори в това табло
        ' 2. Преки "ОБЩО" записи на табла, чийто Табло_Родител Е ТОВА табло
        Dim panelCircuits As List(Of strTokow) = _listTokow.Where(Function(t)
                                                                      Dim isOwn = (t.BuildingName = buildingName AndAlso t.Tablo = tabloName AndAlso t.ТоковКръг <> "ОБЩО")
                                                                      Dim isDirectChild = (t.BuildingName = buildingName AndAlso t.Табло_Родител = tabloName AndAlso t.ТоковКръг = "ОБЩО")
                                                                      Return isOwn OrElse isDirectChild
                                                                  End Function).ToList()
        ' Ако няма нищо за обработка, излизаме
        If panelCircuits.Count = 0 Then Exit Sub
        ' 2. ИЗЧИСЛЕНИЯ
        Dim totalPower As Double = panelCircuits.Sum(Function(c) c.Мощност)
        Dim totalLamps As Integer = panelCircuits.Sum(Function(c) c.brLamp)
        Dim totalContacts As Integer = panelCircuits.Sum(Function(c) c.brKontakt)
        Dim hasThreePhase As Boolean = panelCircuits.Any(Function(c) c.Брой_Полюси = 3)
        Dim mostCommonPoles As Integer = If(hasThreePhase, 3, 1)
        Dim totalPhase As String = If(hasThreePhase, "L1,L2,L3", "L")
        ' 3. НАМИРАНЕ НА ЗАПИСА "ОБЩО" ЗА ТОВА ТАБЛО В ТАЗИ СГРАДА
        Dim totalTokow = _listTokow.FirstOrDefault(Function(t)
                                                       Return t.BuildingName = buildingName AndAlso
                                                      t.Tablo = tabloName AndAlso t.ТоковКръг = "ОБЩО"
                                                   End Function)
        ' 4. ПОПЪЛВАНЕ НА ДАННИТЕ
        With totalTokow
            .Табло_Родител = GetParentForTablo(buildingName, tabloName) ' Помощна функция за намиране на родителя
            .Device = "Табло"
            .Брой_Полюси = mostCommonPoles
            .Мощност = totalPower
            .Фаза = totalPhase
            .brLamp = totalLamps
            .brKontakt = totalContacts
            ' Специфична логика за главното табло на сградата
            If .Tablo.Contains("Гл.Р.Т.") Then
                .Консуматор = "Ке="
                .предназначение = "Рпр.=15кW"
            End If
        End With
        ' 5. ЕЛЕКТРИЧЕСКИ ИЗЧИСЛЕНИЯ
        If hasThreePhase Then
            BalancePhases(buildingName, tabloName) ' Тук също трябва да подаваме (bName, tName)
            ' Извличане на стойности от текстови полета (формат "X>стойност")
            Dim valL1 As Double = CDbl(totalTokow.RCD_Клас.Split(">"c)(1))
            Dim valL2 As Double = CDbl(totalTokow.RCD_Ток.Split(">"c)(1))
            Dim valL3 As Double = CDbl(totalTokow.RCD_Чувствителност.Split(">"c)(1))
            ' Изчисляване на максимален ток
            totalTokow.Ток = Math.Max(valL1, Math.Max(valL2, valL3))
        Else
            totalTokow.Ток = _electricalCalculationEngine.calc_Inom(totalTokow.Мощност, totalTokow.Брой_Полюси)
        End If
        _disconnectorCatalog.CalculateDisconnector(totalTokow)
        _cableCatalog.CalculateCable(totalTokow)
    End Sub
    ' =============================================================
    ' Функция: GetParentForTablo
    ' =============================================================
    ' <summary>
    ' Връща родителското табло
    ' за подадено табло и сграда.
    '
    ' Използва се за:
    ' - изграждане на йерархия
    ' - TreeView структура
    ' - проследяване на вложени табла
    ' </summary>
    Private Function GetParentForTablo(buildingName As String, tabloName As String) As String
        ' Търси записа за текущото табло
        ' само измежду записите от тип "Табло"
        Dim currentTabloRecord = _listTokow.FirstOrDefault(Function(t)
                                                               Return t.BuildingName = buildingName AndAlso
                                                                  t.Tablo = tabloName AndAlso
                                                                  t.Device = "Табло"
                                                           End Function)

        ' Ако е намерен запис:
        ' връща стойността на Табло_Родител
        ' Ако Табло_Родител е празно:
        ' връща празен низ
        If currentTabloRecord IsNot Nothing Then
            Return If(String.IsNullOrEmpty(currentTabloRecord.Табло_Родител),
                  "",
                  currentTabloRecord.Табло_Родител)
        End If
        ' Ако няма намерен запис:
        ' връща празен низ
        Return ""
    End Function
    ''' <summary>
    ''' Балансира фазите (L1, L2, L3) за дадено табло и изчислява резултатните токове.
    ''' </summary>
    ''' <param name="selectedTablo">Име на таблото, за което ще се извърши балансиране.</param>
    ''' <remarks>
    ''' Процедурата извършва пълно балансиране на токовите кръгове и изчислява
    ''' крайното натоварване по фази, което записва в реда "ОБЩО".
    '''
    ''' Основна логика:
    ''' 1. Извлича всички кръгове за таблото (без ред "ОБЩО")
    ''' 2. Проверява за наличие на трифазни консуматори
    '''    - ако няма → пита потребителя дали да продължи
    ''' 3. Намира реда "ОБЩО" и го маркира като трифазен
    ''' 4. Създава групи за балансиране (Bus, RCD, Normal)
    ''' 5. Инициализира токовете по фази
    ''' 6. Добавя трифазните товари към всички фази
    ''' 7. Разпределя групите към най-слабо натоварената фаза (greedy алгоритъм)
    ''' 8. Преизчислява реалните токове по фази след разпределението
    ''' 9. Записва резултатите в ред "ОБЩО"
    ''' 10. Определя максималния фазов ток
    '''
    ''' Важни особености:
    ''' - Редът "ОБЩО" се използва като обобщение на резултатите
    ''' - Трифазните консуматори се разпределят равномерно към всички фази
    ''' - Еднофазните се разпределят чрез групиране
    '''
    ''' Потенциални рискове:
    ''' - Ако няма ред "ОБЩО" → ще възникне грешка (NullReference)
    ''' - Използва CDbl → зависи от регионалните настройки (десетичен разделител)
    ''' - Split(">") предполага винаги валиден формат на текста
    '''
    ''' Възможни подобрения:
    ''' - Проверка за Nothing при totalRow
    ''' - Използване на числови стойности вместо парсване на текст
    ''' - Сортиране на групите по ток преди балансиране
    ''' </remarks>
    Private Sub BalancePhases(buildingName As String, tabloName As String)
        ' 1. ВЗЕМИ КРЪГОВЕТЕ (БЕЗ "ОБЩО")
        ' Вече филтрираме по сграда AND табло
        Dim panelCircuits As List(Of strTokow) = _listTokow.Where(Function(t)
                                                                      Return (t.BuildingName = buildingName AndAlso t.Tablo = tabloName AndAlso t.ТоковКръг <> "ОБЩО") OrElse
               (t.BuildingName = buildingName AndAlso t.Табло_Родител = tabloName AndAlso t.ТоковКръг = "ОБЩО")
                                                                  End Function).ToList()
        ' Ако няма кръгове → прекратяване
        If panelCircuits.Count = 0 Then Return
        ' =====================================================
        ' 2. ПРОВЕРКА ЗА ТРИФАЗНИ КОНСУМАТОРИ
        ' =====================================================
        Dim hasThreePhase As Boolean = panelCircuits.Any(
                                       Function(t) t.Брой_Полюси = 3 OrElse t.Фаза = "L1,L2,L3"
                                       )
        ' Ако няма → пита потребителя
        If Not hasThreePhase Then
            Dim result As DialogResult = MessageBox.Show(
                "Няма трифазни консуматори в това табло." &
                vbCrLf & vbCrLf & "Искате ли да балансирате таблото?",
                "Балансиране на фазите",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
                )
            If result = MsgBoxResult.No Then Return
        End If
        ' 3. НАМИРАНЕ НА РЕД "ОБЩО"
        Dim totalRow = _listTokow.FirstOrDefault(Function(t)
                                                     Return t.BuildingName = buildingName AndAlso
                                                       t.Tablo = tabloName AndAlso
                                                       t.ТоковКръг = "ОБЩО"
                                                 End Function)
        ' Маркиране като трифазен
        totalRow.Брой_Полюси = 3
        totalRow.Фаза = "L1,L2,L3"
        ' =====================================================
        ' 4. СЪЗДАВАНЕ НА ГРУПИ
        ' =====================================================
        Dim balanceGroups As List(Of BalanceGroup) = CreateBalanceGroups(panelCircuits)
        ' =====================================================
        ' 5. ИНИЦИАЛИЗАЦИЯ НА ФАЗИТЕ
        ' =====================================================
        Dim phaseCurrents As New Dictionary(Of String, Double) From {
                                {"L1", 0},
                                {"L2", 0},
                                {"L3", 0}
        }
        ' =====================================================
        ' 6. ДОБАВЯНЕ НА ТРИФАЗНИ ТОВАРИ
        ' =====================================================
        Dim threePhaseCircuits = panelCircuits.Where(
                                 Function(t) t.Брой_Полюси = 3 OrElse t.Фаза = "L1,L2,L3"
                                 ).ToList()
        For Each circuit In threePhaseCircuits
            phaseCurrents("L1") += circuit.Ток
            phaseCurrents("L2") += circuit.Ток
            phaseCurrents("L3") += circuit.Ток
        Next
        ' =====================================================
        ' 7. БАЛАНСИРАНЕ (GREEDY)
        ' =====================================================
        For Each group In balanceGroups
            ' Най-слабо натоварена фаза
            Dim minPhase As String = phaseCurrents.Keys.
                                     OrderBy(Function(p) phaseCurrents(p)).
                                     First()
            ' Присвояване
            group.AssignedPhase = minPhase
            ' Запис в кръговете
            For Each circuit In group.Circuits
                circuit.Фаза = group.AssignedPhase
            Next
            ' Добавяне на товар
            phaseCurrents(minPhase) += group.TotalCurrent
        Next
        ' =====================================================
        ' 8. ПРЕИЗЧИСЛЕНИЕ НА ФАЗИТЕ
        ' =====================================================
        phaseCurrents("L1") = 0
        phaseCurrents("L2") = 0
        phaseCurrents("L3") = 0
        For Each circuit In panelCircuits
            ' Трифазен → към всички
            If circuit.Брой_Полюси = 3 OrElse circuit.Фаза = "L1,L2,L3" Then
                phaseCurrents("L1") += circuit.Ток
                phaseCurrents("L2") += circuit.Ток
                phaseCurrents("L3") += circuit.Ток

            Else
                ' Еднофазен → към конкретната фаза
                Dim p As String = circuit.Фаза.Trim().ToUpper()
                If phaseCurrents.ContainsKey(p) Then
                    phaseCurrents(p) += circuit.Ток
                End If
            End If
        Next
        ' =====================================================
        ' 9. ЗАПИС В "ОБЩО"
        ' =====================================================
        totalRow.RCD_Тип = "Ток фази"
        totalRow.RCD_Клас = "Фаза L1->" & phaseCurrents("L1").ToString("N2")
        totalRow.RCD_Ток = "Фаза L2->" & phaseCurrents("L2").ToString("N2")
        totalRow.RCD_Чувствителност = "Фаза L3->" & phaseCurrents("L3").ToString("N2")
        ' =====================================================
        ' 10. ОПРЕДЕЛЯНЕ НА МАКСИМАЛЕН ТОК
        ' =====================================================
        Dim valL1 As Double = CDbl(totalRow.RCD_Клас.Split(">"c)(1))
        Dim valL2 As Double = CDbl(totalRow.RCD_Ток.Split(">"c)(1))
        Dim valL3 As Double = CDbl(totalRow.RCD_Чувствителност.Split(">"c)(1))
        totalRow.Ток = Math.Max(valL1, Math.Max(valL2, valL3))
    End Sub
    ''' <summary>
    ''' Създава групи от токови кръгове за целите на балансиране на фазите.
    ''' </summary>
    ''' <param name="panelCircuits">Списък от токови кръгове (strTokow), принадлежащи към едно табло.</param>
    ''' <returns>Списък от групи (BalanceGroup), използвани за по-нататъшно разпределение по фази.</returns>
    ''' <remarks>
    ''' Функцията разделя всички токови кръгове в три основни типа групи:
    '''
    ''' 1. Шинни групи (Bus):
    '''    - Включва кръгове, които са маркирани с Шина = True и са еднофазни.
    '''    - Изчислява се процентното участие на шинните консуматори спрямо общата мощност.
    '''    - В зависимост от процента:
    '''          под 10% → "SmallBus"
    '''          над 10% → "LargeBus"
    '''
    ''' 2. Групи по ДТЗ (RCD):
    '''    - Включва еднофазни кръгове, които имат зададена RCD_Нула (N1, N2, ...)
    '''    - Изключва кръговете, които вече са част от шинна група.
    '''    - Групира се по стойността на RCD_Нула.
    '''
    ''' 3. Нормални групи (Normal):
    '''    - Включва всички останали еднофазни кръгове:
    '''         - без ДТЗ
    '''         - не са част от шинна група
    '''
    ''' За всяка група се изчислява:
    ''' - списък с кръгове
    ''' - общ ток (TotalCurrent)
    ''' - тип на групата (GroupType)
    ''' - ключ (GroupKey), използван за идентификация
    '''
    ''' Функцията връща списък от BalanceGroup, които могат да се използват за:
    ''' - балансиране на фазите
    ''' - оптимално разпределение на товарите
    ''' - анализ на натоварването
    '''
    ''' Потенциални особености:
    ''' - Само еднофазни кръгове (Брой_Полюси = 1) се включват в логиката за балансиране.
    ''' - Трифазните консуматори не се обработват тук (вероятно се третират отделно).
    ''' - Ако общата мощност е 0, се избягва деление на нула при изчисляване на процента.
    ''' - Използването на IIf може да доведе до изпълнение и на двата клона (VB особеност),
    '''   но в този контекст няма странични ефекти.
    ''' - Debug.Print се използва за диагностика и проследяване на създадените групи.
    ''' </remarks>
    Private Function CreateBalanceGroups(panelCircuits As List(Of strTokow)) As List(Of BalanceGroup)
        ' Списък с резултатните групи
        Dim groups As New List(Of BalanceGroup)
        ' ----------------------------------------------------
        ' 1. ШИННИ ГРУПИ (Bus)
        ' ----------------------------------------------------
        Dim busCircuits = panelCircuits.Where(
                                Function(t) t.Шина = True AndAlso t.Брой_Полюси = 1
                                ).ToList()
        If busCircuits.Count > 0 Then
            ' Обща мощност на таблото
            Dim totalPower As Double = panelCircuits.Sum(Function(t) t.Мощност)
            ' Мощност на шинните консуматори
            Dim busPower As Double = busCircuits.Sum(Function(t) t.Мощност)
            ' Процентно участие на шината
            Dim busPowerPercent As Double = 0
            If totalPower > 0 Then
                busPowerPercent = (busPower / totalPower) * 100
            End If
            ' Създаване на група за шината
            Dim busGroup As New BalanceGroup With {
                                .GroupType = IIf(busPowerPercent < 10, "SmallBus", "LargeBus"),
                                .GroupKey = "Bus",
                                .Circuits = busCircuits,
                                .TotalCurrent = busCircuits.Sum(Function(t) t.Ток)
            }

            groups.Add(busGroup)
        End If
        ' ----------------------------------------------------
        ' 2. ГРУПИ ПО ДТЗ (RCD)
        ' ----------------------------------------------------
        Dim rcdGroups = panelCircuits.Where(
        Function(t) t.Брой_Полюси = 1 AndAlso
                            Not String.IsNullOrEmpty(t.RCD_Нула) AndAlso
                            t.Шина = False   ' Изключва вече включените в шинна група
                            ).GroupBy(Function(t) t.RCD_Нула)
        For Each rcdGroup In rcdGroups
            Dim balanceGroup As New BalanceGroup With {
                            .GroupType = "RCD",
                            .GroupKey = rcdGroup.Key,  ' Например: N1, N2, N3
                            .Circuits = rcdGroup.ToList(),
                            .TotalCurrent = rcdGroup.Sum(Function(t) t.Ток)
                        }
            groups.Add(balanceGroup)
        Next
        ' ----------------------------------------------------
        ' 3. НОРМАЛНИ ГРУПИ (без ДТЗ и без шина)
        ' ----------------------------------------------------
        Dim normalCircuits = panelCircuits.Where(
                                Function(t) t.Брой_Полюси = 1 AndAlso
                                String.IsNullOrEmpty(t.RCD_Нула) AndAlso
                                t.Шина = False
                                ).ToList()

        For Each circuit In normalCircuits
            Dim normalGroup As New BalanceGroup With {
                        .GroupType = "Normal",
                        .GroupKey = Nothing,
                        .Circuits = New List(Of strTokow) From {circuit},
                        .TotalCurrent = circuit.Ток
                    }
            groups.Add(normalGroup)
        Next
        groups = groups.OrderByDescending(Function(g) g.TotalCurrent).ToList()
        Return groups
    End Function
#End Region
End Class
