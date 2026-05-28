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
        Public Circuits As List(Of strTokow) ' Списък с токови кръгове в групата
        Public GroupType As String ' Тип на групата: "ThreePhase", "RCD", "SmallBus", "LargeBus", "Normal"
        Public GroupKey As String ' Ключ на групата: RCD_Нула (N1, N2...), "Bus" или Nothing
        Public TotalCurrent As Double ' Сумарен ток на групата (сума от токовете на всички ТК)
        Public AssignedPhase As String ' Зададена фаза след балансиране (L1, L2, L3 или "L1,L2,L3")
        ''' <summary>
        ''' Конструктор - инициализира списъка с ТК
        ''' </summary>
        Public Sub New()
            Circuits = New List(Of strTokow)
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
#Region "Сортиране на списък с токови кръгове (List(Of strTokow))"
    ''' <summary>
    ''' Извършва йерархично сортиране на списъка с токови кръгове (Natural Sort).
    ''' </summary>
    ''' <param name="tokowList">
    ''' Списъкът с токови кръгове (List(Of strTokow)), който се подава по референция (ByRef).
    ''' Тъй като се пренаписва оригиналната референция, промените се отразяват директно в извикващата форма.
    ''' </param>
    ''' <remarks>
    ''' Сортирането следва строга географско-електрическа йерархия в 3 стъпки:
    ''' 1. Сграда (BuildingName) - защитава от празни стойности в AutoCAD.
    ''' 2. Електрическо табло (Tablo) - групира кръговете към съответното табло.
    ''' 3. Токов кръг (ТоковКръг) - естествено сортиране по инженерни приоритети чрез GetCircuitSortKey.
    ''' -----------------------------------------------------------------------------------------
    ''' ПЪРВО НИВО: ГЕОГРАФСКО ГРУПИРАНЕ (ПО СГРАДА)
    ''' Използваме вградена проверка If(функция), за да се подсигурим срещу пропуски в AutoCAD.
    ''' Ако атрибутът за сграда е празен, обектът се насочва към служебна група "БЕЗ СГРАДА".
    ''' -----------------------------------------------------------------------------------------
    ''' ВТОРО НИВО: ЛОКАЛНО ГРУПИРАНЕ (ПО ЕЛЕКТРИЧЕСКО ТАБЛО)
    ''' Изпълнява се строго в границите на всяка отделна сграда, дефинирана на първото ниво.
    ''' Подрежда таблата азбучно по име (напр. ГРТ, Т1, ТО).
    ''' -----------------------------------------------------------------------------------------
    ''' ТРЕТО НИВО: ИНЖЕНЕРНО СОРТИРАНЕ (ПО ТОКОВ КРЪГ)
    ''' Извиква спомагателната функция GetCircuitSortKey, която анализира текстовия низ на кръга
    ''' и генерира изкуствен алфа-нумеричен ключ за "естествено подреждане" по приоритети
    ''' (напр. Съществуващи -> Аварийни -> Допълнителни -> Числа -> Резерви -> Общо).
    ''' -----------------------------------------------------------------------------------------
    ''' ФИНАЛИЗИРАНЕ НА СОРТИРАНЕТО
    ''' Материализира LINQ заявката и я превръща обратно в нов, чист и подреден List(Of strTokow).
    ''' -----------------------------------------------------------------------------------------
    ''' </remarks>
    Public Sub SortListTokow()
        ' Изпълняваме тристепенна LINQ щафета, за да подредим оригиналния списък.
        ' Тъй като променливата е предадена с ByRef, промените ще се отразят веднага.
        AppSettings.ListTokow = AppSettings.ListTokow.
                    OrderBy(Function(t) If(String.IsNullOrEmpty(t.BuildingName), "БЕЗ СГРАДА", t.BuildingName)). ' ПЪРВО НИВО: Групиране по Сграда.
                    ThenBy(Function(t) t.Tablo). ' ВТОРО НИВО: Подреждане по име на Електрическото Табло.
                    ThenBy(Function(t) GetCircuitSortKey(t.ТоковКръг)). 'ТРЕТО НИВО: Естествено сортиране на самите токови кръгове.
                    ToList() ' ФИНАЛИЗИРАНЕ: Превръщаме подредената колекция обратно в чист списък (List).
    End Sub
    ''' <summary>
    ''' Връща ключ за сортиране на токов кръг със специален приоритет
    ''' Порядок: 1.ав. → 2.до. → 3.други букви → 4.числа → 5.само букви
    ''' </summary>
    Private Function GetCircuitSortKey(circuitName As String) As String
        If String.IsNullOrEmpty(circuitName) Then Return "ZZZZZZZZZZ"
        Dim name As String = circuitName.Trim().ToUpper()
        Dim priority As String = "9"  ' По подразбиране най-нисък приоритет
        Dim numberPart As String = ""
        Dim letterPart As String = ""
        ' ============================================================
        ' 1. ОПРЕДЕЛИ КАТЕГОРИЯТА (ПРИОРИТЕТ)
        ' ============================================================
        Select Case True
            ' 1. СЪЩЕСТВУВАЩИ
            Case name = "СЪЩ."
                priority = "0"
                numberPart = ExtractNumber(name)
                letterPart = "СЪЩ"
            ' 2. АВАРИЙНИ
            Case name.Contains("АВ")
                priority = "1"
                numberPart = ExtractNumber(name)
                letterPart = "АВ"
            ' 3. ДОПЪЛНИТЕЛНИ
            Case name.Contains("ДО")
                priority = "2"
                numberPart = ExtractNumber(name)
                letterPart = "ДО"
            ' 4. ЧИСТИ ЧИСЛА
            Case IsNumeric(name)
                priority = "3"
                numberPart = name
                letterPart = ""
            ' 5. ЧИСЛО + БУКВА (напр. 1а, 2б)
            Case HasNumberAndLetters(name) AndAlso Char.IsDigit(name(0))
                priority = "4"
                numberPart = ExtractNumber(name)
                letterPart = ExtractLetters(name)
            ' 6. ОБЩО (Провери го ПРЕДИ общия случай за букви)
            Case name = "ОБЩО"
                priority = "9"
                numberPart = ""
                letterPart = "ZZZZZ"
            ' 7. РЕЗЕРВА
            Case name = "РЕЗ."
                priority = "8"
                numberPart = ""
                letterPart = "РЕЗ"
            ' 8. ВСИЧКО ЗАПОЧВАЩО С БУКВА (Основни кръгове като А1, Б1 и т.н.)
            Case Not String.IsNullOrEmpty(name) AndAlso Char.IsLetter(name(0))
                priority = "5"
                numberPart = ExtractNumber(name)
                letterPart = name
                ' 9. ВСИЧКО ОСТАНАЛО
            Case Else
                priority = "8"
                numberPart = ""
                letterPart = name
        End Select
        ' ============================================================
        ' 2. СЪЗДАЙ КЛЮЧ ЗА СОРТИРАНЕ
        ' ============================================================
        ' Формат: Приоритет + Номер (с водещи нули) + Букви
        ' Пример: "10000000001АВ" за "1ав."
        If numberPart.Length > 0 Then
            ' Подравняване на числото с водещи нули (до 10 цифри)
            numberPart = numberPart.PadLeft(10, "0"c)
            Return priority & numberPart & letterPart
        Else
            ' Само букви - сортирай азбучно
            Return priority & "0000000000" & letterPart
        End If
    End Function
    ''' <summary>
    ''' Извлича числото от низ (напр. "1АВ" → "1", "А2Б" → "2")
    ''' </summary>
    Private Function ExtractNumber(text As String) As String
        Dim result As String = ""
        For Each c As Char In text
            If Char.IsDigit(c) Then
                result &= c
            End If
        Next
        Return result
    End Function
    ''' <summary>
    ''' Извлича буквите от низ (напр. "1АВ" → "АВ", "А2Б" → "АБ")
    ''' </summary>
    Private Function ExtractLetters(text As String) As String
        Dim result As String = ""
        For Each c As Char In text
            If Char.IsLetter(c) Then
                result &= c
            End If
        Next
        Return result
    End Function
    ''' <summary>
    ''' Проверява дали низът съдържа и букви и числа
    ''' </summary>
    Private Function HasNumberAndLetters(text As String) As Boolean
        Dim hasNumber As Boolean = False
        Dim hasLetter As Boolean = False
        For Each c As Char In text
            If Char.IsDigit(c) Then hasNumber = True
            If Char.IsLetter(c) Then hasLetter = True
        Next
        Return hasNumber AndAlso hasLetter
    End Function
    ''' <summary>
    ''' Проверява дали низът е само число
    ''' </summary>
    Private Function IsNumeric(text As String) As Boolean
        For Each c As Char In text
            If Not Char.IsDigit(c) Then Return False
        Next
        Return text.Length > 0
    End Function
#End Region

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
    Private Sub GroupByThrees(circuits As List(Of strTokow), n As Integer, ByRef rcdCounter As Integer)
        ' Брой пълни групи по 3
        Dim fullGroups = n \ 3
        ' Остатък след групиране
        Dim remainder As Integer = n Mod 3
        ' Списък със създадените групи
        Dim groups As New List(Of List(Of strTokow))
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
    Private Sub CreateRCDGroup(circuits As List(Of strTokow), rcdNumber As Integer)
        ' Сумарен ток на групата
        Dim totalCurrent As Double = circuits.Sum(Function(t) t.Ток)
        ' Последният кръг се използва като представителен за изчисленията
        Dim lastCircuit As strTokow = circuits.Last()
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
                Dim rootPanel As New strTokow With {
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
                    Dim totalRecord As New strTokow With {
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