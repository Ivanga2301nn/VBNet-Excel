Imports Microsoft.Office.Interop.Excel

Public Class BoardStructureManager
    ' 1. Пазим локални референции на ниво клас
    Private _rcdCatalog As RCDCatalog
    ''' <summary>
    ''' КОНСТРУКТОР: Приема създадените каталози и списъка с токови кръгове от формата
    ''' </summary>
    Public Sub New(rcdCat As RCDCatalog)
        Me._rcdCatalog = rcdCat
    End Sub

    ''' <summary>
    ''' Клас за групиране на токови кръгове за балансиране на фазите
    ''' </summary>
    Public Class BalanceGroup
        Public Circuits As List(Of clsTokow) ' Списък с токови кръгове в групата
        Public GroupType As String ' Тип на групата: "ThreePhase", "RCD", "SmallBus", "LargeBus", "Normal"
        Public GroupKey As String ' Ключ на групата: RCD_Нула (N1, N2...), "Bus" или Nothing
        Public TotalCurrent As Double ' Сумарен ток на групата (сума от токовете на всички ТК)
        Public AssignedPhase As String ' Зададена фаза след балансиране (L1, L2, L3 или "L1,L2,L3")
        ''' <summary>
        ''' Конструктор - инициализира списъка с ТК
        ''' </summary>
        Public Sub New()
            Circuits = New List(Of clsTokow)
        End Sub
        ''' <summary>
        ''' Брой токови кръгове в групата
        ''' </summary>
        Public ReadOnly Property CircuitCount As Integer
            Get
                Return Circuits.Count
            End Get
        End Property
        ''' <summary>
        ''' Сумарна мощност на групата
        ''' </summary>
        Public ReadOnly Property TotalPower As Double
            Get
                Return Circuits.Sum(Function(t) t.Мощност)
            End Get
        End Property
    End Class

#Region "Групиране на контактни кръгове с ДТЗ (RCD)"
    ''' <summary>
    ''' Групира контактните токови кръгове в ДЗТ (RCD) групи, 
    ''' като разделя таблата с еднакви имена в различните сгради.
    ''' </summary>
    Public Sub GroupContactsForRCD()
        ' Проверка за празен списък (защита от грешки)
        If AppSettings.ListTokow Is Nothing OrElse AppSettings.ListTokow.Count = 0 Then Exit Sub
        ' =========================================================================================
        ' КРИТИЧНА ПРОМЯНА: Групираме по анонимен тип (двоен ключ - Сграда и Табло едновременно).
        ' Това гарантира, че "Табло 1" в "Сграда А" и "Табло 1" в "Сграда Б" ще бъдат две отделни групи.
        ' =========================================================================================
        Dim panels = AppSettings.ListTokow.GroupBy(Function(t) New With {Key t.BuildingName, Key t.Tablo})
        For Each panelGroup In panels
            ' Избор само на кръговете, които съдържат контакти и не са самото главно табло
            ' panelGroup вече съдържа само кръгове от конкретното табло в конкретната сграда
            Dim contactCircuits = panelGroup.Where(
                                  Function(t) t.brKontakt > 0 AndAlso t.Device <> "Табло"
                                  ).ToList()
            ' Брой на контактните кръгове в това конкретно табло
            Dim n As Integer = contactCircuits.Count
            ' Ако в това табло няма контакти – преминава към следващото табло/сграда
            If n = 0 Then Continue For
            ' Брояч за номера на ДТЗ в рамките на ТОВА ТАБЛО
            Dim rcdCounter As Integer = 0
            ' Динамично разпределяне според броя на контактите
            Select Case n
                ' Обединяваме Case 1 и Case 2, тъй като логиката им е идентична (отиват под една ДЗТ)
                Case 1, 2
                    rcdCounter += 1
                    CreateRCDGroup(contactCircuits, rcdCounter)
                ' Три или повече контактни кръга → задейства се разпределителния алгоритъм
                Case Is >= 3
                    GroupByThrees(contactCircuits, n, rcdCounter)
            End Select
            ProcessPanelRCDLogic(panelGroup.ToList())
        Next
    End Sub
    ''' <summary>
    ''' Разделя списък от токови кръгове на групи по 3 за защита с ДТЗ.
    ''' </summary>
    ''' <param name="circuits">Списък от токови кръгове.</param>
    ''' <param name="n">Общият брой кръгове.</param>
    ''' <param name="rcdCounter">Брояч на ДТЗ, предаван по референция.</param>
    ''' <remarks>
    ''' Основната цел е да се разпределят контактните кръгове в групи,
    ''' които да бъдат защитени с една ДТЗ.
    '''
    ''' Алгоритъм:
    ''' - Определя се броят на пълните групи по 3 (fullGroups).
    ''' - Определя се остатъкът (remainder).
    '''
    ''' Възможни случаи:
    ''' - remainder = 0 → всички групи са по 3 кръга.
    ''' - remainder = 1 → последните 4 кръга се групират заедно.
    ''' - remainder = 2 → последната група съдържа 2 кръга.
    '''
    ''' След създаване на групите:
    ''' - за всяка група се увеличава броячът на ДТЗ
    ''' - извиква се CreateRCDGroup() за създаване на защитата.
    '''
    ''' Потенциална особеност:
    ''' - При малък брой групи (например 4 кръга) алгоритъмът създава една група от 4,
    '''   вместо 3+1, което е по-практично при реални електрически табла.
    ''' </remarks>
    Private Sub GroupByThrees(circuits As List(Of clsTokow), n As Integer, ByRef rcdCounter As Integer)
        ' Брой пълни групи по 3
        Dim fullGroups = n \ 3
        ' Остатък след групиране
        Dim remainder As Integer = n Mod 3
        ' Списък със създадените групи
        Dim groups As New List(Of List(Of clsTokow))
        Select Case remainder
        ' Всички групи са по 3
            Case 0
                For i As Integer = 0 To fullGroups - 1
                    groups.Add(circuits.Skip(i * 3).Take(3).ToList())
                Next
        ' Последната група става 4
            Case 1
                For i As Integer = 0 To fullGroups - 2
                    groups.Add(circuits.Skip(i * 3).Take(3).ToList())
                Next
                groups.Add(circuits.Skip((fullGroups - 1) * 3).Take(4).ToList())
        ' Последната група е 2
            Case 2
                For i As Integer = 0 To fullGroups - 1
                    groups.Add(circuits.Skip(i * 3).Take(3).ToList())
                Next
                groups.Add(circuits.Skip(fullGroups * 3).Take(2).ToList())
        End Select
        ' Създаване на ДТЗ за всяка група
        For Each group In groups
            rcdCounter += 1
            CreateRCDGroup(group, rcdCounter)
        Next
    End Sub
    ''' <summary>
    ''' Създава група от токови кръгове, защитени от една ДТЗ.
    ''' </summary>
    ''' <param name="circuits">Списък от кръгове, които ще бъдат защитени от една ДТЗ.</param>
    ''' <param name="rcdNumber">Номер на ДТЗ в рамките на таблото.</param>
    ''' <remarks>
    ''' Процедурата извършва следните действия:
    '''
    ''' 1. Изчислява сумарния ток на всички кръгове в групата.
    ''' 2. Избира последния кръг в списъка като представителен за изчисленията.
    ''' 3. Проверява дали групата съдържа трифазен консуматор.
    ''' 4. Ако има трифазен консуматор:
    '''    - броят на полюсите се принудително задава на 3.
    ''' 5. Временно се задава сумарният ток на избрания кръг.
    ''' 6. Извиква се SetRCD(), която избира подходяща ДТЗ от каталога.
    ''' 7. На всички кръгове в групата се задава обща нула:
    '''    - "N1", "N2", "N3" и т.н.
    ''' 8. След това се възстановяват оригиналните стойности
    '''    на ток и брой полюси на последния кръг.
    '''
    ''' Потенциални особености:
    ''' - Методът използва последния кръг като временен носител на сумарния ток.
    ''' - Това е практично решение, но изисква внимателно възстановяване
    '''   на оригиналните стойности след изчислението.
    '''
    ''' Важна забележка:
    ''' - Ако структурата strTokow е Value Type (Structure),
    '''   промените върху елементите може да не се отразят в оригиналния списък,
    '''   ако не се използват по референция.
    ''' </remarks>
    Private Sub CreateRCDGroup(circuits As List(Of clsTokow), rcdNumber As Integer)
        ' Сумарен ток на групата
        Dim totalCurrent As Double = circuits.Sum(Function(t) t.Ток)
        ' Последният кръг се използва като представителен за изчисленията
        Dim lastCircuit As clsTokow = circuits.Last()
        ' Запазване на оригиналните параметри
        Dim originalTok As Double = lastCircuit.Ток
        Dim originalPoles As Integer = lastCircuit.Брой_Полюси
        ' Проверка дали има трифазен консуматор в групата
        Dim hasThreePhase As Boolean = circuits.Any(Function(t) t.Брой_Полюси = 3)
        ' Ако има трифазен консуматор → използва се 3-полюсна конфигурация
        If hasThreePhase Then lastCircuit.Брой_Полюси = 3
        ' Временно задаване на сумарния ток
        lastCircuit.Ток = totalCurrent
        ' Избор на подходяща ДТЗ
        Dim matchingRCD = _rcdCatalog.SelectRcd(totalCurrent, hasThreePhase, False)
        ' Задаване на обща нула за всички кръгове в групата
        For Each circuit In circuits
            circuit.RCD_Нула = "N" & rcdNumber.ToString()
        Next
        ' Възстановяване на оригиналните стойности
        lastCircuit.Ток = originalTok
        lastCircuit.Брой_Полюси = originalPoles
    End Sub
    ''' <summary>
    ''' Анализира ДТЗ групите в таблото, изчислява параметрите им 
    ''' и записва избраната ДТЗ в последния кръг от всяка "N" група.
    ''' </summary>
    ''' <param name="panelCircuits">Списък с всички токови кръгове за текущото табло.</param>
    Public Sub ProcessPanelRCDLogic(panelCircuits As List(Of clsTokow))
        ' 1. Филтрираме само кръговете, които имат назначена "N" група, 
        ' и ги групираме по техния N-номер (напр. "N1", "N2"...)
        Dim rcdGroups = panelCircuits.
                        Where(Function(t) Not String.IsNullOrEmpty(t.RCD_Нула) AndAlso
                        t.RCD_Нула.StartsWith("N")).
                        GroupBy(Function(t) t.RCD_Нула)
        ' =========================================================================
        ' ТУК Е ПЕРФЕКТНОТО МЯСТО: Създаваме енджина локално за процедурата
        ' =========================================================================
        ' 2. Обхождаме всяка намерена "N" група в таблото
        For Each group In rcdGroups
            ' Вземаме списъка с кръгове, които участват в тази конкретна "N" група
            Dim groupCircuits As List(Of clsTokow) = group.ToList()
            ' 3. СУМИРАНЕ: Изчисляваме сумарния ток на групата
            Dim totalCurrent As Double = groupCircuits.Sum(Function(t) t.Ток)
            ' 4. ПРОВЕРКА ЗА ПОЛЮСИ: Проверяваме дали поне един от кръговете е 3-полюсен
            Dim hasThreePhase As Boolean = groupCircuits.Any(Function(t) t.Брой_Полюси = 3)
            ' 5. ВЗЕМАМЕ ПОСЛЕДНИЯ: Намираме последния токов кръг от групата
            Dim lastCircuit As clsTokow = groupCircuits.Last()
            ' 6. ЗАПАЗВАНЕ НА ОРИГИНАЛНИТЕ ДАННИ
            Dim originalTok As Double = lastCircuit.Ток
            Dim originalPoles As Integer = lastCircuit.Брой_Полюси
            ' 7. ВРЕМЕННО НАГНАЖДАНЕ НА ПАРАМЕТРИТЕ ЗА СУМАРНАТА ГРУПА
            lastCircuit.Ток = totalCurrent
            If hasThreePhase Then lastCircuit.Брой_Полюси = 3
            _rcdCatalog.SetRCD(lastCircuit)
            ' 8. ВЪЗСТАНОВЯВАНЕ НА ОРИГИНАЛНИТЕ СТОЙНОСТИ НА КРЪГА
            ' След като SetRCD е записвала вътре в обекта параметрите на ДТЗ, 
            ' връщаме оригиналния ток и полюси на самия токов кръг.
            lastCircuit.Ток = originalTok
            lastCircuit.Брой_Полюси = originalPoles
        Next
    End Sub
#End Region

    ''' <summary>
    ''' Структурира данните в списъка, като гарантира, че съществуват 
    ''' коренните записи за всяка сграда и сумарните ("ОБЩО") записи за всяко табло.
    ''' </summary>
    Public Sub EnsureAllStructureRecords()
        ' Бърза защита: ако няма прочетени данни от AutoCAD, няма какво да структурираме
        If AppSettings.ListTokow Is Nothing OrElse AppSettings.ListTokow.Count = 0 Then Exit Sub
        ' 1. Взимаме всички уникални сгради, които съществуват в списъка на един ход
        Dim allBuildings As List(Of String) = AppSettings.ListTokow.Select(Function(x) x.BuildingName).Distinct().ToList()
        ' 2. Започваме обхождането на всяка сграда
        For Each bName As String In allBuildings
            ' ==========================================
            ' ЧАСТ 1: ГАРАНТИРАНЕ НА КОРЕНЕН ЗАПИС (Root Node) ЗА СГРАДАТА
            ' ==========================================
            Dim rootExists As Boolean = AppSettings.ListTokow.Any(Function(x) x.Tablo = ROOT_NODE_TEXT AndAlso
                                                              x.BuildingName = bName)
            If Not rootExists Then
                Dim rootPanel As New clsTokow With {
                .BuildingName = bName,
                .Tablo = ROOT_NODE_TEXT,
                .Device = "Табло",
                .Табло_Родител = "",
                .ТоковКръг = "ОБЩО"
            }
                AppSettings.ListTokow.Add(rootPanel)
            End If
            ' ==========================================
            ' ЧАСТ 2: ГАРАНТИРАНЕ НА ЗАПИС "ОБЩО" ЗА ВСЯКО ТАБЛО В СГРАДАТА
            ' ==========================================
            ' Намираме уникалните имена на табла в текущата сграда (като изключваме корена)
            Dim panelsInCurrentBuilding = AppSettings.ListTokow.Where(Function(t) t.BuildingName = bName AndAlso
                                                                  t.Tablo <> ROOT_NODE_TEXT) _
                                               .Select(Function(t) t.Tablo) _
                                               .Distinct() _
                                               .ToList()
            ' Обхождаме реалните табла, за да им подсигурим сумарен ред "ОБЩО"
            For Each tName As String In panelsInCurrentBuilding
                Dim totalExists As Boolean = AppSettings.ListTokow.Any(Function(x) x.BuildingName = bName AndAlso
                                                                   x.Tablo = tName AndAlso
                                                                   x.ТоковКръг = "ОБЩО")
                If Not totalExists Then
                    Dim totalRecord As New clsTokow With {
                    .BuildingName = bName,
                    .Tablo = tName,
                    .ТоковКръг = "ОБЩО",
                    .Device = "Табло",
                    .Табло_Родител = ROOT_NODE_TEXT ' Всяко табло се закача за корена на сградата
                }
                    AppSettings.ListTokow.Add(totalRecord)
                End If
            Next
        Next
    End Sub
End Class