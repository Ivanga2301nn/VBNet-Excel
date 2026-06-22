#Region "КЛАС: CableCatalog (Кабели)"
Public Class CableCatalog
    Public Class CableInfo
        Public CableType As String         ' "СВТ", "САВТ", "Al/R"
        Public InsulationType As String    ' ("ПВЦ", "XLPE", "GUM")
        Public Material As String          ' "Cu", "Al"
        Public PhaseSize As String         ' "2,5", "4", и т.н.
        Public NeutralSize As String       ' "0", "1,5", "2,5", и т.н.
        Public MaxCurrent_Air As Double    ' Допустим ток във въздух
        Public MaxCurrent_Ground As Double ' Допустим ток в земя
        Public MaxWorkingTemp As Double    ' (65, 70, 90°C)
    End Class
    ''' <summary>
    ''' КОНСТРУКТОР: Извиква се автоматично при New CableCatalog()
    ''' </summary>
    Public Sub New()
        ' Веднага напълва каталога с данни в паметта, за да е готов за DataGridView
        LoadCatalog()
    End Sub
    ' Стандартните сечения, подредени по големина
    Public ReadOnly StandardPhaseSizes As String() = {
                           "1,5", "2,5", "4", "6", "10", "16", "25", "35", "50", "70", "95", "120", "150", "185", "240"
                            }
    ' Стандартните сечения на неутралното жило, подредени по големина
    Public ReadOnly StandardNeutralSizes As String() = {
                          "1,5", "2,5", "4", "6", "10", "16", "16", "16", "25", "35", "50", "70", "70", "95", "120"
}
    ' Складът за всички кабели (замества DataList)
    Public Property CableList As New List(Of CableInfo)()
    ' Списъкът с методи за монтаж
    Public Property LiMountMethod As New List(Of strMountMethod)()
    Public Structure strMountMethod
        Public Simbol As String
        Public Text As String
    End Structure
    ' Списък с уникалните типове кабели за запълване на ComboBox (замества Cable_For_combo)
    Public Property CableTypesForCombo As New List(Of String)()
    ''' <summary>
    ''' Зарежда каталожните данни за кабелите.
    ''' </summary>
    Public Sub LoadCatalog()
        CableList.Clear()
        ' 1. СВТ (Cu, 70°C, PVC)
        AddCableSeries("СВТ", "Cu", 70, "PVC",
                   {19, 25, 34, 43, 59, 79, 105, 126, 157, 199, 246, 285, 326, 374, 445},
                   {25, 34, 45, 55, 76, 96, 126, 151, 178, 225, 270, 306, 346, 390, 458})
        ' 2. САВТ (Al, 70°C, PVC)
        AddCableSeries("САВТ", "Al", 70, "PVC",
                   {0, 20, 26, 34, 43, 64, 82, 100, 119, 152, 185, 215, 245, 285, 338},
                   {0, 25, 32, 42, 53, 75, 92, 110, 134, 170, 210, 245, 274, 310, 360})
        ' 3. NYY (Cu, 70°C, PVC)
        AddCableSeries("NYY", "Cu", 70, "PVC",
                   {19.5, 25, 34, 43, 59, 79, 106, 129, 157, 199, 246, 285, 326, 374, 445},
                   {27, 36, 47, 59, 79, 102, 133, 159, 188, 232, 280, 318, 359, 406, 473})
        ' 4. NAYY (Al, 70°C, PVC)
        AddCableSeries("NAYY", "Al", 70, "PVC",
                   {0, 0, 0, 0, 0, 0, 82, 100, 119, 152, 186, 216, 246, 285, 338},
                   {0, 0, 0, 0, 0, 0, 102, 123, 144, 179, 215, 245, 275, 313, 364})
        ' 5. N2XY (Cu, 90°C, XLPE)
        AddCableSeries("N2XY", "Cu", 90, "XLPE",
                   {24, 32, 42, 53, 74, 98, 133, 162, 197, 250, 308, 359, 412, 475, 564},
                   {31, 40, 52, 64, 86, 112, 145, 174, 206, 254, 305, 348, 392, 444, 517})
        ' 6. NA2XY (Al, 90°C, XLPE)
        AddCableSeries("NA2XY", "Al", 90, "XLPE",
                   {0, 0, 0, 0, 0, 0, 102, 126, 149, 191, 234, 273, 311, 360, 427},
                   {0, 0, 0, 0, 0, 0, 112, 135, 158, 196, 234, 268, 300, 342, 398})
        ' --- Специални кабели (Остават ръчно или с единично добавяне) ---
        ' Al/R (Al, 90°C, XLPE) - Поради специфичните му неутрални жила (35+54, 50+54, 70+54...)
        CableList.Add(New CableInfo With {.CableType = "Al/R", .Material = "Al", .PhaseSize = "16", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 83, .MaxCurrent_Ground = 0, .NeutralSize = "16"})
        CableList.Add(New CableInfo With {.CableType = "Al/R", .Material = "Al", .PhaseSize = "25", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 111, .MaxCurrent_Ground = 0, .NeutralSize = "25"})
        CableList.Add(New CableInfo With {.CableType = "Al/R", .Material = "Al", .PhaseSize = "35", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 138, .MaxCurrent_Ground = 0, .NeutralSize = "35"})
        CableList.Add(New CableInfo With {.CableType = "Al/R", .Material = "Al", .PhaseSize = "35", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 138, .MaxCurrent_Ground = 0, .NeutralSize = "54"})
        CableList.Add(New CableInfo With {.CableType = "Al/R", .Material = "Al", .PhaseSize = "50", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 164, .MaxCurrent_Ground = 0, .NeutralSize = "54"})
        CableList.Add(New CableInfo With {.CableType = "Al/R", .Material = "Al", .PhaseSize = "70", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 213, .MaxCurrent_Ground = 0, .NeutralSize = "54"})
        CableList.Add(New CableInfo With {.CableType = "Al/R", .Material = "Al", .PhaseSize = "95", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 258, .MaxCurrent_Ground = 0, .NeutralSize = "70"})
        CableList.Add(New CableInfo With {.CableType = "Al/R", .Material = "Al", .PhaseSize = "150", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 344, .MaxCurrent_Ground = 0, .NeutralSize = "70"})

        ' ПВ-А1 (Cu, 70°C, PVC) - Тъй като всички нули са "0"
        Dim pvCurrentAir() As Double = {20, 27, 36, 45, 63, 82, 113, 138, 168, 210, 262, 307, 352, 405, 482}
        Dim pvCurrentGround() As Double = {29, 38, 49, 62, 83, 104, 136, 162, 192, 236, 285, 322, 363, 410, 475}
        For i As Integer = 0 To StandardPhaseSizes.Length - 1
            CableList.Add(New CableInfo With {
                         .CableType = "ПВ-А1",
                         .Material = "Cu",
                         .PhaseSize = StandardPhaseSizes(i),
                         .MaxWorkingTemp = 70,
                         .InsulationType = "PVC",
                         .MaxCurrent_Air = pvCurrentAir(i),
                         .MaxCurrent_Ground = pvCurrentGround(i),
                         .NeutralSize = "0"
                         })
        Next
    End Sub
    ''' <summary>
    ''' Връща уникалните типове кабели от наличния списък.
    ''' </summary>
    Public Function GetUniqueCableTypes() As List(Of String)
        Return CableList.Select(Function(b) b.CableType) _
                    .Distinct() _
                    .ToList()
    End Function
    ''' <summary>
    ''' Автоматично генерира и добавя пълна серия кабели на базата на синхронизираните масиви за сечения.
    ''' </summary>
    Private Sub AddCableSeries(cableType As String, material As String, maxTemp As Integer, insType As String, currentAir() As Double, currentGround() As Double)
        For i As Integer = 0 To StandardPhaseSizes.Length - 1
            CableList.Add(New CableInfo With {
            .CableType = cableType,
            .Material = material,
            .PhaseSize = StandardPhaseSizes(i),
            .MaxWorkingTemp = maxTemp,
            .InsulationType = insType,
            .MaxCurrent_Air = currentAir(i),
            .MaxCurrent_Ground = currentGround(i),
            .NeutralSize = StandardNeutralSizes(i)
        })
        Next
    End Sub
    ''' <summary>
    ''' Изчислява необходимото сечение на кабел според тока и условията на полагане
    ''' Оптимизиран за сградни инсталации
    ''' </summary>
    Public Sub CalculateCable(ByRef tokow As clsTokow,
                                 Optional Type As String = "СВТ",        ' Тип кабел (СВТ, САВТ, NYY...)
                                Optional layMethod As Integer = 0,      ' 0=въздух (35°C), 1=земя (15°C)
                                Optional mountMethod As String = "B1",  ' "A1"=гипсокартон, "B2"=под мазилка, "C"=над таван
                                Optional Broj_Cable As Integer = 1,     ' Брой паралелни кабели
                                Optional Tipe_Cable As Integer = 0,     ' 0=кабел (3-жилен), 1=проводник (1-жилен)
                                Optional matType As Integer = 0,        ' 0=мед (Cu), 1=алуминий (Al)
                                Optional RetType As Integer = 1         ' 0=само сечение, 1=пълно означение
                                )

        If tokow.Device = "Разединител" OrElse
           tokow.Device = "Съществуващ" OrElse
           tokow.Device = "Резерва" Then Exit Sub

        Dim Ibreaker As String = tokow.Breaker_Номинален_Ток
        Dim NumberPoles As String = tokow.Брой_Полюси

        ' 1. МАТЕРИАЛ И ФИЛТРИРАНЕ НА КАТАЛОГА (използваме DataList вместо Catalog_Cables)
        Dim material As String = If(matType = 1, "Al", "Cu")
        Dim filteredCables = CableList.Where(
                                Function(c) c.CableType = Type AndAlso c.Material = material
                             ).OrderBy(
                                Function(c) CDbl(c.PhaseSize.Replace(",", "."))
                             ).ToList()

        If filteredCables.Count = 0 Then Exit Sub ' Защита, ако каталогът е празен

        ' 2. КОРЕКЦИОННИ КОЕФИЦИЕНТИ
        ' K1 - брой кабели на скара
        Dim K1_Table As New Dictionary(Of Integer, Double) From {
            {1, 1.0}, {2, 0.88}, {3, 0.82}, {4, 0.77}, {5, 0.73}, {6, 0.7}
        }
        Dim K1 As Double = If(K1_Table.ContainsKey(Broj_Cable), K1_Table(Broj_Cable), 0.7)

        ' K2 - температура
        Dim Qok As Double = If(layMethod = 1, 15, 35) ' 15°C земя, 35°C въздух
        Const Qokdef As Double = 25
        Dim Q As Double = filteredCables(0).MaxWorkingTemp
        Dim K2 As Double = 1.0
        Dim ratio As Double = (Q - Qok) / (Q - Qokdef)
        If ratio > 0 Then K2 = Math.Sqrt(ratio)

        ' ТАБЛИЦА С КОЕФИЦИЕНТИ ЗА МОНТАЖ
        Dim MountCoefficients As New Dictionary(Of String, Double) From {
            {"A1", 1.0}, {"B1", 1.0}, {"C", 1.0}, {"D1", 1.0}, {"D2", 1.0}, {"E", 1.0}, {"F", 1.0}
        }
        Dim K3 As Double = If(MountCoefficients.ContainsKey(mountMethod), MountCoefficients(mountMethod), 1.0)

        ' 3. ИЗБОР НА СЕЧЕНИЕ
        Dim calc As String = "######"
        Dim Inom As Double = Val(Ibreaker)
        Dim Idop As Double = Inom / (K1 * K2 * K3)

        ' ТЪРСИМ ПЪРВОТО СЕЧЕНИЕ КОЕТО ИЗДЪРЖА Idop
        For i As Integer = 0 To filteredCables.Count - 1
            Dim cable As CableInfo = filteredCables(i)
            Dim Imax As Double = If(layMethod = 1, cable.MaxCurrent_Ground, cable.MaxCurrent_Air)

            If Imax >= Idop Then
                calc = cable.PhaseSize
                Exit For
            End If
        Next

        ' 4. ИЗВЛИЧАНЕ НА ТОКОВЕ ЗА ГОЛЕМИ СЕЧЕНИЯ (за паралелни кабели)
        Dim bestSection As String = ""
        Dim bestNum As Integer = 0
        Dim bestNeutral As String = ""

        If calc = "######" Then
            Dim Current_120 As Double = 0 : Dim Current_150 As Double = 0
            Dim Current_185 As Double = 0 : Dim Current_240 As Double = 0
            Dim Neutral_120 As String = "" : Dim Neutral_150 As String = ""
            Dim Neutral_185 As String = "" : Dim Neutral_240 As String = ""
            For Each cable As CableInfo In filteredCables
                Select Case cable.PhaseSize
                    Case "120"
                        Current_120 = If(layMethod = 1, cable.MaxCurrent_Ground, cable.MaxCurrent_Air)
                        Neutral_120 = cable.NeutralSize
                    Case "150"
                        Current_150 = If(layMethod = 1, cable.MaxCurrent_Ground, cable.MaxCurrent_Air)
                        Neutral_150 = cable.NeutralSize
                    Case "185"
                        Current_185 = If(layMethod = 1, cable.MaxCurrent_Ground, cable.MaxCurrent_Air)
                        Neutral_185 = cable.NeutralSize
                    Case "240"
                        Current_240 = If(layMethod = 1, cable.MaxCurrent_Ground, cable.MaxCurrent_Air)
                        Neutral_240 = cable.NeutralSize
                End Select
            Next
            Dim Idop_Adjusted As Double = Idop * 0.95
            Dim cables_120 As Integer = If(Current_120 > 0, CInt(Math.Ceiling(Idop_Adjusted / Current_120)), 0)
            Dim cables_150 As Integer = If(Current_150 > 0, CInt(Math.Ceiling(Idop_Adjusted / Current_150)), 0)
            Dim cables_185 As Integer = If(Current_185 > 0, CInt(Math.Ceiling(Idop_Adjusted / Current_185)), 0)
            Dim cables_240 As Integer = If(Current_240 > 0, CInt(Math.Ceiling(Idop_Adjusted / Current_240)), 0)

            Dim data = {
                New With {.Size = "120", .Price = 21.17 * cables_120, .Nom = cables_120, .Neutral = Neutral_120},
                New With {.Size = "150", .Price = 24.62 * cables_150, .Nom = cables_150, .Neutral = Neutral_150},
                New With {.Size = "185", .Price = 30.7 * cables_185, .Nom = cables_185, .Neutral = Neutral_185},
                New With {.Size = "240", .Price = 39.86 * cables_240, .Nom = cables_240, .Neutral = Neutral_240}
            }
            Dim bestMatch = data.Where(Function(x) x.Price > 0).OrderBy(Function(x) x.Price).FirstOrDefault()
            If bestMatch IsNot Nothing Then
                calc = bestMatch.Size
                bestSection = bestMatch.Size
                bestNum = bestMatch.Nom
                bestNeutral = bestMatch.Neutral
            End If
        End If
        ' 5. ФОРМАТИРАНЕ НА РЕЗУЛТАТА
        Dim Text As String = ""
        If RetType = 0 Then
            Text = calc
        Else
            Dim Poles As String = If(NumberPoles = "1", "3x", "5x")
            Dim calc_N As String = ""
            If Val(calc.Replace(",", ".")) > 16 Then
                Poles = "4х"
                Dim index = filteredCables.FindIndex(Function(c) c.PhaseSize = calc)
                If index >= 0 Then calc_N = filteredCables(index).NeutralSize
            End If

            Text = If(bestNum > 1, bestNum & "x", "")
            Text += Type & " "
            If Poles = "4х" AndAlso Not String.IsNullOrEmpty(calc_N) Then
                Text += "3х" & calc & "+" & calc_N
            Else
                Text += Poles & calc
            End If
            Text += "mm²"
        End If
        tokow.Кабел_Брой_Фаза = bestNum
        tokow.Кабел_Брой_Група = Broj_Cable
        tokow.Кабел_Сечение = Text
        tokow.Кабел_Тип = Type
        tokow.Кабел_Полагане = If(layMethod = 0, "Във въздух", "В земя")
        tokow.Кабел_Монтаж = GetMountMethodInfo(mountMethod)
    End Sub
    ''' <summary>
    ''' Напълва списъка с дефинираните начини на монтаж
    ''' </summary>
    Public Sub LoadMountMethods()
        LiMountMethod = New List(Of strMountMethod) From {
        New strMountMethod With {.Simbol = "A1", .Text = "В изолация"},
        New strMountMethod With {.Simbol = "B1", .Text = "Тръба (стена)"},
        New strMountMethod With {.Simbol = "C", .Text = "Върху стена"},
        New strMountMethod With {.Simbol = "D1", .Text = "Тръба (земя)"},
        New strMountMethod With {.Simbol = "D2", .Text = "Кабел (земя)"},
        New strMountMethod With {.Simbol = "E", .Text = "Кабелна скара"},
        New strMountMethod With {.Simbol = "F", .Text = "Многож. скара"},
        New strMountMethod With {.Simbol = "G", .Text = "Свободен въздух"}
    }
    End Sub
    ''' <summary>
    ''' Помощен метод за правилно конвертиране на сечение от стринг към double (справя се с запетаи)
    ''' </summary>
    Private Function Method_StringSizeToDouble(sizeStr As String) As Double
        If String.IsNullOrEmpty(sizeStr) Then Return 0
        Return Val(sizeStr.Replace(",", "."))
    End Function
    ''' <summary>
    ''' Връща информация за начин на монтаж на база подадена стойност (символ или текст).
    ''' </summary>
    Public Function GetMountMethodInfo(inputValue As String) As String
        If String.IsNullOrEmpty(inputValue) Then Return "Не е намерено"
        ' Търсене в локалния списък на класа
        Dim result = LiMountMethod.FirstOrDefault(Function(m) m.Simbol.Equals(inputValue, StringComparison.OrdinalIgnoreCase) OrElse m.Text = inputValue)
        If Not String.IsNullOrEmpty(result.Simbol) Then
            Return If(result.Simbol.Equals(inputValue, StringComparison.OrdinalIgnoreCase), result.Text, result.Simbol)
        End If
        Return "Не е намерено"
    End Function
    ''' <summary>
    ''' Връща типа на материала на кабела: 0 за Мед (Cu), 1 за Алуминий (Al).
    ''' </summary>
    Public Function GetCableTypeResult(cableName As String) As Integer
        If String.IsNullOrEmpty(cableName) Then Return 0
        ' Списък с алуминиеви кабели
        Dim targetCables As String() = {"САВТ", "NA2XY", "Al/R", "NAYY"}
        ' Сравняваме, като превръщаме входа в главни букви (за по-сигурно)
        ' Забележка: Тъй като Al/R съдържа латински букви, inputValue.ToUpper() ще го направи AL/R
        Dim upperName As String = cableName.ToUpper()
        ' Правим и масива с главни букви, за да съвпаднат перфектно
        Dim targetCablesUpper As String() = {"САВТ", "NA2XY", "AL/R", "NAYY"}
        If targetCablesUpper.Contains(upperName) Then
            Return 1 ' Алуминий
        Else
            Return 0 ' Мед
        End If
    End Function
End Class
#End Region

#Region "КЛАС: BreakerCatalog (Прекъсвачи)"
Public Class BreakerCatalog
    ''' <summary>
    ''' КАТАЛОГ автоматичен прекъсвач – MCB, MCCB или ACB.
    ''' Може да се използва за избор на прекъсвач за генераторни табла,
    ''' както и за по-сложни сценарии с селективност и късо съединение.
    ''' </summary>
    Public Class BreakerInfo
        Public Brand As String                      ' Производител на прекъсвача (например "Schneider").
        Public Series As String                     ' Серия или модел на прекъсвача (например "EZ9", "C120").
        Public Category As String                   ' "MCB", "MCCB" или "ACB"
        Public NominalCurrent As Integer            ' Номинален ток в ампери.
        Public Poles As Integer                     ' Брой полюси (1, 2, 3, 4).
        Public Ics_kA As Decimal                    ' Прекъсвателна способност.
        Public Curve As String                      ' Крива (B, C, D...).
        Public TripUnit As String                   ' Защитен блок (TM-D, Micrologic...).        Public Sub New(brand As String, type As String, Inom As Double, curve As String, kA As Double, poles As Integer, trip As String)
    End Class
    ''' <summary>
    ''' Списъкът, който държи филтрирания каталог за ИЗБРАНИЯ производител.
    ''' </summary>
    Public Property Breakers As New List(Of BreakerInfo)()
    ' ✅ Списъци, подготвени специално за ComboBox-овете във формата
    Public Property Brand_For_combo As New List(Of String)()
    Public Property Breakers_For_combo As New List(Of String)()
    Public Property TripUnit_For_combo As New List(Of String)()
    Public Property Curve_For_combo As New List(Of String)()
    Public Sub New()
        ' ✅ Твърдо дефинираме поддържаните марки, за да заредим първия ComboBox на формата
        Brand_For_combo = New List(Of String) From {"Schneider", "Schrack", "Siemens", "ABB", "Бат Генчо"}
        LoadCatalog()
    End Sub
    ''' <summary>
    ''' Генерира всички възможни комбинации на параметрите 
    ''' и ги добавя в списъка 'Breakers'.
    ''' </summary>
    Public Sub LoadCatalog()
        Breakers.Clear()
        ' MCB
        AddBreakerSeries("Schneider", "EZ9 MCB", "MCB",
                         {6, 10, 16, 20, 25, 32, 40, 50, 63},
                         {"C"}, {1, 2, 3, 4}, {6}, Nothing)
        AddBreakerSeries("Schneider", "iC60N", "MCB",
                         {2, 3, 4, 6, 10, 16, 20, 25, 32, 40, 50, 63},
                         {"C", "B", "D"}, {1, 2, 3, 4}, {6}, Nothing)
        AddBreakerSeries("Schneider", "C120", "MCB",
                         {63, 80, 100, 125}, {"C", "B", "D"},
                         {1, 2, 3, 4}, {10}, Nothing)
        ' MCCB
        AddBreakerSeries("Schneider", "NSXm", "MCCB",
                         {16, 25, 32, 40, 50, 63, 80, 100},
                         {"E", "B", "F", "N", "H"}, {3, 4}, {25}, {"TM-D", "TM-DC"})
        AddBreakerSeries("Schneider", "NSX100", "MCCB",
                         {16, 25, 32, 40, 63, 80, 100},
                         {"B", "F", "N", "H", "S", "L"}, {3}, {25}, {"TM-D", "TM-DC"})
        AddBreakerSeries("Schneider", "NSX160", "MCCB",
                         {80, 100, 125, 160},
                         {"B", "F", "N", "H", "S", "L"}, {3}, {36}, {"TM-D"})
        AddBreakerSeries("Schneider", "NSX250", "MCCB",
                         {125, 160, 200, 250},
                         {"B", "F", "N", "H", "S", "L"}, {3}, {50},
                         {"TM-D", "Micrologic 2.0", "Micrologic 5.0"})
        AddBreakerSeries("Schneider", "NSX400", "MCCB",
                         {250, 320, 400},
                         {"F", "N", "H", "S", "L"}, {3}, {70}, {"Micrologic 2.3"})
        AddBreakerSeries("Schneider", "NSX630", "MCCB",
                         {400, 500, 630},
                         {"F", "N", "H", "S", "L"}, {3}, {100}, {"Micrologic 2.3"})
        ' ACB
        AddBreakerSeries("Schneider", "MTZ1", "ACB",
                         {630, 800, 1000, 1250, 1600},
                         {"H1", "H2"},
                         {3, 4}, {42, 65, 100}, {"Micrologic 6.0"})
        AddBreakerSeries("Schneider", "MTZ2", "ACB",
                         {800, 1000, 1250, 1600, 2000, 2500, 3200, 4000, 5000, 6300},
                         {"N1", "H1", "H2"},
                         {3, 4}, {42, 65, 100}, {"Micrologic 6.0"})
        ' TODO: Тук утре по същия начин се добавят редове за Siemens без промяна на друга логика
        ' AddBreakerSeries("Siemens", "5SY", "MCB", {6, 10, 16...}, ...)
        ' TODO: Тук се добавят редове за ABB утре
        ' Amparo 6kA (Битови серии) - обикновено B и C крива, 1-полюсни и 3-полюсни
        AddBreakerSeries("Schrack", "Amparo 6kA", "MCB",
                                 {6, 10, 13, 16, 20, 25, 32, 40, 50, 63},
                                 {"B", "C"}, {1, 3}, {6}, Nothing)
        ' BMS-6 (Индустриална / стандартна серия 6kA) - B, C и D крива, включва и 2p / 4p
        AddBreakerSeries("Schrack", "BMS-6", "MCB",
                         {1, 2, 4, 6, 10, 13, 16, 20, 25, 32, 40, 50, 63},
                         {"B", "C", "D"}, {1, 2, 3, 4}, {6}, Nothing)
        ' BMS-10 (Усилена серия 10kA)
        AddBreakerSeries("Schrack", "BMS-10", "MCB",
                         {0.5, 1, 2, 4, 6, 10, 16, 20, 25, 32, 40, 50, 63},
                         {"B", "C", "D"}, {1, 2, 3, 4}, {10}, Nothing)
        ' MC1 - по-малките ляти корпуси (до 160A)
        AddBreakerSeries("Schrack", "MC1", "MCCB",
                         {20, 25, 32, 40, 50, 63, 80, 100, 125, 160},
                         Nothing, {3, 4}, {25, 50}, {"TM (Thermal-Magnetic)"})
        ' MC2 - среден размер (до 300A)
        AddBreakerSeries("Schrack", "MC2", "MCCB",
                         {160, 200, 250, 300},
                         Nothing, {3, 4}, {50}, {"TM", "Electronic"})
    End Sub
    ''' <summary>
    ''' МЕГА ВАЖНО: Филтрира помощните списъци САМО за избраната в момента марка!
    ''' </summary>
    Public Sub FilterComboLists(brandName As String)
        ' Вземаме само прекъсвачите на избрания производител
        Dim filtered = Breakers.Where(Function(b) b.Brand = brandName).ToList()
        ' Сега вече списъците за комбо кутиите съдържат само вярната апаратура!
        Breakers_For_combo = filtered.Select(Function(b) b.Series).Distinct().ToList()
        TripUnit_For_combo = filtered.Select(Function(b) b.TripUnit).Distinct().ToList()
        Curve_For_combo = filtered.Select(Function(b) b.Curve).Distinct().ToList()
    End Sub

    ''' <summary>
    ''' Универсален метод за вътрешно генериране на комбинациите.
    ''' </summary>
    Private Sub AddBreakerSeries(brand As String, series As String, category As String,
                                 currents As Integer(), curves As String(), polesList As Integer(),
                                 icsValues As Decimal(), tripUnits As String())

        Dim localCurves As String() = If(curves, New String() {"-"})
        Dim localPoles As Integer() = If(polesList, New Integer() {0})
        Dim localIcs As Decimal() = If(icsValues, New Decimal() {0})
        Dim localTrips As String() = If(tripUnits, New String() {Nothing})
        For Each Inom In currents
            For Each curve In localCurves
                For Each poles In localPoles
                    For Each ics In localIcs
                        For Each trip In localTrips
                            Breakers.Add(New BreakerInfo With {
                                .Brand = brand,
                                .Series = series,
                                .Category = category,
                                .NominalCurrent = Inom,
                                .Poles = poles,
                                .Curve = curve,
                                .Ics_kA = ics,
                                .TripUnit = trip
                            })
                        Next
                    Next
                Next
            Next
        Next
    End Sub
    ''' <summary>
    ''' Определя и задава подходящ прекъсвач за даден токов кръг.
    ''' </summary>
    Public Sub CalculateBreaker(ByRef tokow As clsTokow)
        ' Деклариране на променлива за намерения прекъсвач от новия клас
        Dim breaker As BreakerInfo = Nothing
        ' ------------------------------------------------------------
        ' Избор на серия прекъсвач според изчисления ток и тип устройство
        ' ------------------------------------------------------------
        Select Case tokow.Device
            Case "Разединител"
                ' Пропускаме или добавяш специфична логика, ако е нужно
            Case "Бойлер"
                ' За бойлери използваме по-строги критерии (крива C) и минимум 17A
                Dim searchCurrent As Double = If(tokow.Ток > 17, tokow.Ток, 17)
                breaker = SelectBreaker(searchCurrent, tokow.Брой_Полюси, "C")
            Case "Контакт"
                ' За контакти също използваме крива C и минимум 17A
                Dim searchCurrent As Double = If(tokow.Ток > 17, tokow.Ток, 17)
                breaker = SelectBreaker(searchCurrent, tokow.Брой_Полюси, "C")
            Case "Лампа"
                ' За лампи – минимум 8.5А и крива C
                breaker = SelectBreaker(8.5, tokow.Брой_Полюси, "C")
            Case Else
                ' За други устройства – селекция по диапазон на тока според твоята скала
                Select Case tokow.Ток
                    Case Is <= 63
                        ' Модулни прекъсвачи (EZ9, iC60N) -> Крива С
                        breaker = SelectBreaker(tokow.Ток, tokow.Брой_Полюси, "C")
                    Case Is <= 125
                        ' Модулни прекъсвачи с по-висок номинал (C120) -> Крива С
                        breaker = SelectBreaker(tokow.Ток, tokow.Брой_Полюси, "C")
                    Case Is <= 160
                        ' Компактни MCCB прекъсвачи (NSXm) -> Крива N
                        breaker = SelectBreaker(tokow.Ток, tokow.Брой_Полюси, "N")
                    Case Is <= 630
                        ' MCCB прекъсвачи за по-големи товари (NSX100-630) -> Крива N
                        breaker = SelectBreaker(tokow.Ток, tokow.Брой_Полюси, "N")
                    Case Else
                        ' Въздушни прекъсвачи (ACB) -> Търсим по серия "MTZ"
                        breaker = SelectBreaker(tokow.Ток, tokow.Брой_Полюси, "MTZ")
                End Select
        End Select
        ' ------------------------------------------------------------
        ' Проверка дали е намерен подходящ прекъсвач и запис чрез референцията
        ' ------------------------------------------------------------
        If breaker Is Nothing Then
            Dim info As String =
                    $"Внимание: Не е намерен прекъсвач в {tokow.Tablo}!" & vbCrLf &
                    "Детайли:" & vbCrLf &
                    $"- Табло: {tokow.Tablo}" & vbCrLf &
                    $"- Кръг: {tokow.ТоковКръг}" & vbCrLf &
                    $"- Мощност: {tokow.Мощност} kW" & vbCrLf &
                    $"- Ток: {tokow.Ток} A"
            MsgBox(info, MsgBoxStyle.Exclamation, "Инфо за LayerPair")
        Else
            ' Обновяваме директно стойностите в оригиналния обект в паметта
            tokow.Breaker_Номинален_Ток = breaker.NominalCurrent.ToString()
            tokow.Breaker_Тип_Апарат = breaker.Series
            tokow.Breaker_Крива = breaker.Curve
            tokow.Breaker_Изкл_Възможност = breaker.Ics_kA & "kA"
            tokow.Брой_Полюси = breaker.Poles
            tokow.Breaker_Защитен_блок = If(breaker.TripUnit, "-")
        End If
    End Sub
    ' =============================================================
    ' Функция: SelectBreaker
    ' =============================================================
    ''' <summary>
    ''' Избира подходящ прекъсвач (BreakerInfo) от колекцията Breakers
    ''' според изчислен ток, брой полюси и крива/серия. Марката се взема автоматично от текущото състояние.
    ''' </summary>
    ''' <param name="calculatedCurrent">Изчислен работен ток.</param>
    ''' <param name="poles">Необходим брой полюси.</param>
    ''' <param name="curveOrSeries">Крива или серия на прекъсвача. По подразбиране: "C"</param>
    ''' <returns>BreakerInfo: избраният прекъсвач, или Nothing при липса на подходящ.</returns>
    Public Function SelectBreaker(calculatedCurrent As Double,
                                  poles As Integer,
                                  Optional curveOrSeries As String = "C") As BreakerInfo
        Const SAFETY_FACTOR As Double = 1.15
        Dim minimumRequiredCurrent As Double = calculatedCurrent * SAFETY_FACTOR
        ' ✅ ПРОМЕНЕНО: Прочиташ го директно чрез Глобалния модул AppSettings
        Dim _activeBrand As String = AppSettings.CurrentManufacturer
        ' ====================================================================================
        ' ПОДРОБНО ОБЯСНЕНИЕ НА ТЪРСЕНЕТО (LINQ):
        ' ====================================================================================
        ' 1. b.Poles = poles -> Филтрира по точния брой полюси.
        ' 2. b.Brand.Equals(manufacturer...) -> Филтрира по марката, която каталогът ВЕЧЕ Е ЗАПОМНИЛ!
        ' 3. Търси по крива ИЛИ по серия.
        ' 4. Сортира по номинален ток във възходящ ред.
        ' 5. Взема първия, който покрива минималния ток с коефициента за сигурност.
        ' ====================================================================================
        Dim selectedBreaker = Breakers _
            .Where(Function(b) b.Poles = poles) _
            .Where(Function(b) b.Brand.Equals(_activeBrand, StringComparison.OrdinalIgnoreCase)) _
            .Where(Function(b) b.Curve.Equals(curveOrSeries, StringComparison.OrdinalIgnoreCase) OrElse
                               b.Series.Equals(curveOrSeries, StringComparison.OrdinalIgnoreCase)) _
            .OrderBy(Function(b) b.NominalCurrent) _
            .FirstOrDefault(Function(b) b.NominalCurrent >= minimumRequiredCurrent)

        ' Fallback избор (ако няма намерен апарат за конкретната крива, търсим по-друг параметър)
        If selectedBreaker Is Nothing Then
            selectedBreaker = Breakers _
                .Where(Function(b) b.Poles = poles) _
                .Where(Function(b) b.Brand.Equals(_activeBrand, StringComparison.OrdinalIgnoreCase)) _
                .OrderBy(Function(b) b.NominalCurrent) _
                .FirstOrDefault(Function(b) b.NominalCurrent >= minimumRequiredCurrent)
        End If
        Return selectedBreaker
    End Function
    ' Изчиства данните за прекъсвач (MCB)
    Public Sub ClearBreaker(ByRef tokow As clsTokow)
        tokow.Breaker_Тип_Апарат = ""           ' Серия апарат (EZ9, C120, NSX, MTZ)
        tokow.Breaker_Крива = ""                ' Характеристика (B, C, D)
        tokow.Breaker_Номинален_Ток = ""        ' Номинален ток (пример: "16A")
        tokow.Breaker_Изкл_Възможност = ""      ' Изключвателна способност ("6000A", "10000A")
        tokow.Breaker_Защитен_блок = ""         ' Изключвателна способност ("6000A", "10000A")
    End Sub
    ''' <summary>
    ''' Извлича цифрата от стрингове като "1p", "2p", "3p", "4p". Ако не успее, връща 0.
    ''' </summary>
    Private Function ParsePoles(polesStr As String) As Integer
        If String.IsNullOrWhiteSpace(polesStr) Then Return 0
        ' Вземаме само първия символ и пробваме да го превърнем в число
        Dim firstChar As String = polesStr.Trim().Substring(0, 1)
        Dim poles As Integer
        If Integer.TryParse(firstChar, poles) Then
            Return poles
        End If
        Return 0
    End Function

    ''' <summary>
    ''' Връща уникалните серии прекъсвачи за текущата марка (от AppSettings), 
    ''' които поддържат подадения номинален ток и брой полюси (напр. "3p").
    ''' </summary>
    Public Function GetUniqueBreakerTypes(targetCurrent As String, polesStr As String) As List(Of String)
        Dim result As New List(Of String)()
        If String.IsNullOrWhiteSpace(targetCurrent) Then
            result.Add("---")
            Return result
        End If
        ' 1. Изчистваме и парсваме тока
        Dim cleanCurrentStr As String = targetCurrent.ToUpper().Replace("A", "").Replace("А", "").Replace(" ", "")
        Dim currentDecimal As Decimal
        If Not Decimal.TryParse(cleanCurrentStr, currentDecimal) Then
            result.Add("---")
            Return result
        End If
        ' 2. Извличаме броя полюси като число
        Dim targetPoles As Integer = ParsePoles(polesStr)
        ' 3. Вземаме активната марка
        Dim activeBrand As String = AppSettings.CurrentManufacturer
        ' 4. Филтрираме каталога
        result = Breakers.Where(Function(b) b.Brand.Equals(activeBrand, StringComparison.OrdinalIgnoreCase) AndAlso
                                          b.NominalCurrent = currentDecimal AndAlso
                                          b.Poles = targetPoles) _
                         .Select(Function(b) b.Series) _
                         .Distinct() _
                         .ToList()
        If result.Count = 0 Then result.Add("---")
        Return result
    End Function
    ''' <summary>
    ''' Връща уникалните амперажи за избраната серия и брой полюси. Ако няма намерени, връща ["---"].
    ''' </summary>
    Public Function GetUniqueBreakerCurrents(seriesName As String, polesStr As String) As List(Of String)
        If String.IsNullOrWhiteSpace(seriesName) Then Return New List(Of String) From {"---"}
        Dim targetPoles As Integer = ParsePoles(polesStr)
        Dim activeBrand As String = AppSettings.CurrentManufacturer
        Dim result = Breakers.Where(Function(b) b.Brand.Equals(activeBrand, StringComparison.OrdinalIgnoreCase) AndAlso
                                              b.Series.Equals(seriesName, StringComparison.OrdinalIgnoreCase) AndAlso
                                              b.Poles = targetPoles) _
                           .Select(Function(b) b.NominalCurrent.ToString()) _
                           .Distinct() _
                           .ToList()
        If result.Count = 0 Then result.Add("---")
        Return result
    End Function
    ''' <summary>
    ''' Връща уникалните криви за избраната серия и брой полюси. Ако няма, връща ["---"].
    ''' </summary>
    Public Function GetUniqueBreakerCurves(seriesName As String, polesStr As String) As List(Of String)
        If String.IsNullOrWhiteSpace(seriesName) Then Return New List(Of String) From {"---"}
        Dim targetPoles As Integer = ParsePoles(polesStr)
        Dim activeBrand As String = AppSettings.CurrentManufacturer
        Dim result = Breakers.Where(Function(b) b.Brand.Equals(activeBrand, StringComparison.OrdinalIgnoreCase) AndAlso
                                              b.Series.Equals(seriesName, StringComparison.OrdinalIgnoreCase) AndAlso
                                              b.Poles = targetPoles) _
                           .Select(Function(b) b.Curve) _
                           .Where(Function(c) Not String.IsNullOrEmpty(c) AndAlso c <> "-") _
                           .Distinct() _
                           .ToList()
        If result.Count = 0 Then result.Add("---")
        Return result
    End Function
    ''' <summary>
    ''' Връща уникалните защитни блокове за избраната серия и брой полюси. Ако няма, връща ["---"].
    ''' </summary>
    Public Function GetUniqueBreakerUnits(seriesName As String, polesStr As String) As List(Of String)
        If String.IsNullOrWhiteSpace(seriesName) Then Return New List(Of String) From {"---"}
        Dim targetPoles As Integer = ParsePoles(polesStr)
        Dim activeBrand As String = AppSettings.CurrentManufacturer
        Dim result = Breakers.Where(Function(b) b.Brand.Equals(activeBrand, StringComparison.OrdinalIgnoreCase) AndAlso
                                              b.Series.Equals(seriesName, StringComparison.OrdinalIgnoreCase) AndAlso
                                              b.Poles = targetPoles) _
                           .Select(Function(b) b.TripUnit) _
                           .Where(Function(t) Not String.IsNullOrEmpty(t) AndAlso t <> "-") _
                           .Distinct() _
                           .ToList()
        If result.Count = 0 Then result.Add("---")
        Return result
    End Function
End Class
#End Region

#Region "КЛАС: MotorProtectionCatalog (Моторни защити - GV)"
''' <summary>
''' Централен клас за управление на моторни защити.
''' Съдържа продуктовата база данни и алгоритъма за автоматичен избор.
''' Подготвен за централизирано зареждане от външни процедури.
''' </summary>
Public Class MotorProtectionCatalog
    Public Class MotorProtect
        Public Brand As String = ""
        Public MinCurrent As Double = 0.0
        Public MaxCurrent As Double = 0.0
        Public Type As String = ""
        Public MotorPower As String = ""
        Public SettingRange As String = ""
    End Class
    ''' <summary>
    ''' КОНСТРУКТОР: Зарежда каталога веднага при създаване, гледайки AppSettings
    ''' </summary>
    Public Sub New()
        LoadCatalog()
    End Sub
    Public Property ProtectMotor As New List(Of MotorProtect)()
    Public Sub LoadCatalog()
        ProtectMotor.Clear()
        ' Case "Schneider Electric", "Schneider"
        ' Коригирани и напълно подравнени масиви (точно 14 елемента всеки)
        AddProtectMotorSeries("Schneider",
                     {"GV2ME01", "GV2ME02", "GV2ME03", "GV2ME04", "GV2ME05", "GV2ME06", "GV2ME07", "GV2ME08", "GV2ME10", "GV2ME14", "GV2ME16", "GV2ME20", "GV2ME22", "GV2ME32"},
                     {0.1, 0.16, 0.25, 0.4, 0.63, 1.0, 1.6, 2.5, 4.0, 6.0, 9.0, 13.0, 17.0, 24.0},
                     {0.16, 0.25, 0.4, 0.63, 1.0, 1.6, 2.5, 4.0, 6.3, 10.0, 14.0, 18.0, 25.0, 32.0},
                     {"<0.06", "0.06", "0.12", "0.18", "0.25", "0.55", "0.75", "1.5", "2.2", "4.0", "5.5", "7.5", "11.0", "15.0"},
                     {"0.1-0.16A", "0.16-0.25A", "0.25-0.40A", "0.40-0.63A", "0.63-1.0A", "1.0-1.6A", "1.6-2.5A", "2.5-4.0A", "4.0-6.3A", "6.0-10A", "9.0-14A", "13-18A", "17-25A", "24-32A"}
                )
        ' Case "Siemens"
        ' Case "ABB"
        ' Case "Schrack Technik", "Schrack"
        ''
    End Sub
    ''' <summary>
    ''' Универсален метод за вътрешно генериране на комбинациите.
    ''' </summary>
    Private Sub AddProtectMotorSeries(brand As String, series As String(),
                                      minCurrent As Double(), maxCurrent As Double(),
                                      motorPower As String(), settingRange As String())

        Dim expectedLength As Integer = series.Length
        ' Защитна проверка срещу разминаване в масивите
        If minCurrent.Length <> expectedLength OrElse
           maxCurrent.Length <> expectedLength OrElse
           motorPower.Length <> expectedLength OrElse
           settingRange.Length <> expectedLength Then
            Throw New ArgumentException("Грешка в базата данни! Масивите за марка '" & brand & "' имат различен брой елементи.")
        End If
        For i As Integer = 0 To expectedLength - 1
            Dim item As New MotorProtect()
            item.Brand = brand
            item.Type = series(i)
            item.MinCurrent = minCurrent(i)
            item.MaxCurrent = maxCurrent(i)
            item.MotorPower = motorPower(i)
            item.SettingRange = settingRange(i)
            ProtectMotor.Add(item)
        Next
    End Sub
    ''' <summary>
    ''' Избира подходящ моторен прекъсвач на база подаден ток.
    ''' </summary>
    Public Function Calculate_GV2(Ток As String, Връща As Integer) As String
        ' 1. Преобразуване на входния ток - сигурен метод без значение от регионалните настройки
        Dim cleanInput As String = Ток.Replace(",", ".")
        Dim I_double As Double
        ' Използваме InvariantCulture, за да сме сигурни, че винаги чете "10.5" правилно
        If Not Double.TryParse(cleanInput, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, I_double) Then
            Return "N/A"
        End If
        If I_double <= 0 Then Return "N/A"
        ' 2. Търсене в базата данни
        ' Използваме AndAlso и четем директно от AppSettings, за да сме в крак с промените!
        Dim match = ProtectMotor.FirstOrDefault(Function(x) _
                   x.Brand.Equals(AppSettings.CurrentManufacturer,
                                  StringComparison.OrdinalIgnoreCase) AndAlso
                   I_double >= x.MinCurrent AndAlso
                   I_double <= x.MaxCurrent)
        If match Is Nothing Then
            Return "Out of range (" & I_double.ToString("F2", System.Globalization.CultureInfo.InvariantCulture) & "A)"
        End If
        ' 3. Връщане на резултат
        Select Case Връща
            Case 1 : Return match.Type               ' Напр. GV2ME08
            Case 2 : Return match.MotorPower         ' Напр. 1.5
            Case 3 : Return match.SettingRange       ' Напр. 2.5-4.0A
            Case Else : Return "Грешен параметър"
        End Select
    End Function
End Class
#End Region

#Region "КЛАС: DisconnectorCatalog (Товариви прекъсвачи)"
Public Class DisconnectorCatalog
    ' Тази структура съдържа информация за прекъсвач (изключвател/разединител),
    Public Class DisconnectorInfo
        Public Brand As String           ' Производител на прекъсвача (например "Schneider").
        Public NominalCurrent As Integer ' Номинален ток на прекъсвача в ампери.
        Public Type As String             ' Тип на прекъсвача.
        Public Poles As Integer           ' Брой полюси на прекъсвача.
    End Class
    ''' <summary>
    ''' Списъкът, който държи каталога за разединителите.
    ''' </summary>
    Public Property Disconnectors As New List(Of DisconnectorInfo)()
    Public Property Disconnectors_For_combo As New List(Of String)()
    Public Property Discon_Tok_For_combo As New List(Of String)()
    Public Sub New()
        LoadCatalog()
        ' Генериране на списъците за ComboBox
        Disconnectors_For_combo = Disconnectors.Select(Function(b) b.Type).Distinct().ToList()
        Discon_Tok_For_combo = Disconnectors.
                               Select(Function(d) d.NominalCurrent).
                               Distinct().
                               OrderBy(Function(n) n).
                               Select(Function(n) n.ToString()).
                               ToList()
    End Sub
    ''' <summary>
    ''' Процедура за автоматично запълване на каталога чрез серии
    ''' </summary>
    Private Sub LoadCatalog()
        Disconnectors.Clear()
        ' Schneider iSW (1, 2, 3, 4 полюса, от 20 до 125А)
        AddDisconnectorSeries("Schneider", "iSW", {20, 32, 40, 63, 100, 125}, {1, 2, 3, 4})
        ' Schneider INS (3 и 4 полюса, от 100 до 1600А)
        AddDisconnectorSeries("Schneider", "INS", {40, 63, 80, 100, 125, 160, 200, 250, 320, 400, 500, 630, 800, 1000, 1250, 1600}, {3, 4})
        ' Schneider IN (3 и 4 полюса, от 1600 до 2500А)
        AddDisconnectorSeries("Schneider", "IN", {1600, 1600, 2500}, {3, 4})
    End Sub
    ''' <summary>
    ''' Помощен метод за добавяне на цяла серия разединители по твоя модел.
    ''' </summary>
    Private Sub AddDisconnectorSeries(brand As String, type As String, currents As Integer(), poles As Integer())
        For Each p As Integer In poles
            For Each current As Integer In currents
                Dim discon As New DisconnectorInfo With {
                    .Brand = brand,
                    .Type = type,
                    .NominalCurrent = current,
                    .Poles = p
                }
                Disconnectors.Add(discon)
            Next
        Next
    End Sub
    ''' <summary>
    ''' Избира подходящ разединител (прекъсвач) според тока на токовия кръг и подадената марка.
    ''' </summary>
    ''' <param name="tokow">Токов кръг</param>
    ''' <param name="brand">Марка на апарата (по подразбиране "Schneider")</param>
    Public Sub CalculateDisconnector(tokow As clsTokow,
                                     Optional brand As String = "Schneider")
        ' 1️⃣ КОНСТАНТИ (КОЕФИЦИЕНТИ)
        Const MIN_FACTOR As Double = 1.15
        Const MAX_FACTOR As Double = 1.25
        ' 2️⃣ ИЗЧИСЛЯВАНЕ НА ДИАПАЗОН
        Dim minRange As Double = tokow.Ток * MIN_FACTOR
        Dim maxRange As Double = tokow.Ток * MAX_FACTOR
        ' ====================================================================================
        ' 3️⃣ ПОДРОБНО ОБЯСНЕНИЕ НА ТЪРСЕНЕТО (LINQ):
        ' ====================================================================================
        ' Използваме LINQ филтриране върху списъка "Disconnectors". Тъй като "DisconnectorInfo" 
        ' е Клас (Референтен тип), променливата "suitable" ще съдържа конкретния намерен апарат 
        ' или ще бъде "Nothing", ако няма съвпадение.
        '
        ' ЛОГИКА НА ФИЛТРИРАНЕТО (СТЪПКА ПО СТЪПКА):
        ' 1. .Where() -> Филтрира списъка по 3 критерия едновременно:
        '    - СЪВПАДЕНИЕ НА ПОЛЮСИ: d.Poles трябва да е равно на tokow.Брой_Полюси.
        '    - СЪВПАДЕНИЕ НА МАРКА: d.Brand трябва да отговаря на подадения аргумент 'brand'.
        '      Използваме "OrdinalIgnoreCase", за да игнорираме главни/малки букви при сравнение.
        '    - ДОПУСТИМ ТОК: Номиналният ток d.NominalCurrent трябва да е >= изчисления minRange.
        '
        ' 2. .OrderBy() -> Сортира останалите след филтъра апарати във възходящ ред по техния ток.
        '    Така най-малкият възможен номинал (който все пак ни върши работа) застава първи.
        '
        ' 3. .FirstOrDefault() -> Взема първия елемент от сортирания списък (най-оптималния ток).
        '    Ако никой апарат не е отговорил на филтъра, методът автоматично ни връща "Nothing".
        ' ====================================================================================
        Dim suitable As DisconnectorInfo = Disconnectors.
                        Where(Function(d) d.Poles = tokow.Брой_Полюси AndAlso
                        d.Brand.Equals(AppSettings.CurrentManufacturer,
                                       StringComparison.OrdinalIgnoreCase) AndAlso
                        d.NominalCurrent >= minRange).
                        OrderBy(Function(d) d.NominalCurrent).
                        FirstOrDefault()
        ' 4️⃣ ПРОВЕРКА И ЗАПИС
        If suitable IsNot Nothing Then
            tokow.Breaker_Номинален_Ток = suitable.NominalCurrent
            tokow.Breaker_Тип_Апарат = suitable.Type
            tokow.Breaker_Крива = "-"
        Else
            ' Съобщение за грешка с конкретно търсената марка
            MsgBox(String.Format("Грешка: Не е намерен разединител от марка '{0}' за {1}А с {2} полюса.", brand, tokow.Ток, tokow.Брой_Полюси))
        End If
    End Sub
    ''' <summary>
    ''' Помощна функция: Извлича цифрата от стрингове като "1p", "2p", "3p", "4p". Ако не успее, връща 0.
    ''' </summary>
    Private Function ParsePoles(polesStr As String) As Integer
        If String.IsNullOrWhiteSpace(polesStr) Then Return 0
        Dim firstChar As String = polesStr.Trim().Substring(0, 1)
        Dim poles As Integer
        If Integer.TryParse(firstChar, poles) Then
            Return poles
        End If
        Return 0
    End Function
    ''' <summary>
    ''' Връща уникалните типове разединители за текущата марка (от AppSettings),
    ''' които поддържат подадения ток и брой полюси. Ако няма, връща ["---"].
    ''' </summary>
    Public Function GetUniqueDisconnectorTypes(targetCurrent As String, polesStr As String) As List(Of String)
        Dim result As New List(Of String)()
        If String.IsNullOrWhiteSpace(targetCurrent) Then
            result.Add("---")
            Return result
        End If
        ' 1. Изчистваме тока от "А" (латиница/кирилица) и интервали
        Dim cleanCurrentStr As String = targetCurrent.ToUpper().Replace("A", "").Replace("А", "").Replace(" ", "")
        ' 2. Превръщаме в Integer, тъй като в DisconnectorInfo NominalCurrent е Integer
        Dim currentInt As Integer
        If Not Integer.TryParse(cleanCurrentStr, currentInt) Then
            result.Add("---")
            Return result
        End If
        ' 3. Парсваме полюсите
        Dim targetPoles As Integer = ParsePoles(polesStr)
        ' 4. Вземаме активната марка от AppSettings
        Dim activeBrand As String = AppSettings.CurrentManufacturer
        ' 5. Филтрираме списъка Disconnectors по твоите точни полета: .Brand, .NominalCurrent, .Poles и селектираме .Type
        result = Disconnectors.Where(Function(d) d.Brand.Equals(activeBrand, StringComparison.OrdinalIgnoreCase) AndAlso
                                                 d.NominalCurrent = currentInt AndAlso
                                                 d.Poles = targetPoles) _
                              .Select(Function(d) d.Type) _
                              .Distinct() _
                              .ToList()
        If result.Count = 0 Then result.Add("---")
        Return result
    End Function
    ''' <summary>
    ''' Връща уникалните амперажи за избран тип разединител и брой полюси. Ако няма, връща ["---"].
    ''' </summary>
    Public Function GetUniqueDisconnectorCurrents(typeName As String, polesStr As String) As List(Of String)
        If String.IsNullOrWhiteSpace(typeName) Then Return New List(Of String) From {"---"}
        Dim targetPoles As Integer = ParsePoles(polesStr)
        Dim activeBrand As String = AppSettings.CurrentManufacturer
        ' Филтрираме по .Type вместо по .Series
        Dim result = Disconnectors.Where(Function(d) d.Brand.Equals(activeBrand, StringComparison.OrdinalIgnoreCase) AndAlso
                                                     d.Type.Equals(typeName, StringComparison.OrdinalIgnoreCase) AndAlso
                                                     d.Poles = targetPoles) _
                                  .Select(Function(d) d.NominalCurrent.ToString()) _
                                  .Distinct() _
                                  .ToList()
        If result.Count = 0 Then result.Add("---")
        Return result
    End Function
End Class
#End Region

#Region "КЛАС: RCDCatalog (Дефектнотокови защити - RCD)"
Public Class RCDCatalog
    ''' <summary>
    ''' Клас: RCDInfo
    ''' </summary>
    ''' <remarks>
    ''' Този клас съхранява информация за защитно устройство от тип диференциална токова защита (ДТЗ/RCD).
    ''' Променен на Клас за надеждна проверка за "Nothing", ако апаратът не бъде намерен в каталога.
    ''' </remarks>
    Public Class RCDInfo
        Public Brand As String              ' Производител на RCD устройството
        Public NominalCurrent As Integer    ' Номинален ток на RCD в ампери
        Public Type As String               ' Тип на чувствителността на RCD спрямо диференциален ток.
        ' - "AC" – реагира на синусоидален променлив ток
        ' - "A" – реагира на променлив и пулсиращ постоянен ток
        ' - "F" – висока чувствителност, бърза реакция на различни видове ток
        Public Poles As String              ' Брой полюси на устройството ("2p", "4p")
        Public Sensitivity As Integer       ' Чувствителност на RCD в милиампери
        Public DeviceType As String         ' Вид на устройството ("RCCB", "RCBO", "iID")
        Public Breaker As Boolean
        ' True – устройството е RCBO (комбиниран прекъсвач + ДТЗ);
        ' False – устройството е RCCB (само диференциална защита).
    End Class
    ''' <summary>
    ''' Списъкът, който държи каталога за дефектнотоковите защити.
    ''' </summary>
    Public Rcds As New List(Of RCDInfo)()
    ' Списъци за ComboBox-овете във формата
    Public Rcds_For_combo As New List(Of String)()
    Public Rcd_Tok_For_combo As New List(Of String)()
    Public Sub New()
        ' ✅ ИЗВИКВАНЕ НА ПРОЦЕДУРАТА ЗА ЗАПЪЛВАНЕ
        LoadCatalog()
        ' Автоматично генериране на списъците за ComboBox от заредената база
        Rcds_For_combo = Rcds.Select(Function(r) r.DeviceType).Distinct().ToList()
        Rcd_Tok_For_combo = Rcds.
                            Select(Function(r) r.NominalCurrent).
                            Distinct().
                            OrderBy(Function(n) n).
                            Select(Function(n) n.ToString()).
                            ToList()
    End Sub
    ''' <summary>
    ''' Процедура за автоматично запълване на каталога чрез серии
    ''' </summary>
    Private Sub LoadCatalog()
        ' 1. Серия EZ9 RCCB (Чисто RCCB, Тип AC, 30mA, 2p и 4p -> Breaker = False)
        AddRcdSeries("Schneider", {25, 40, 63}, {"AC"}, {"2p", "4p"}, {30, 300}, "EZ9 RCCB", False)
        ' 2. Серия EZ9 RCBO (Комбиниран прекъсвач, Тип AC, 30mA, само 2p -> Breaker = True)
        AddRcdSeries("Schneider", {6, 10, 16, 20, 25, 32, 40}, {"AC"}, {"2p"}, {30}, "EZ9 RCBO", True)
        ' 3. Серия iID - тип "si" (Специална защита, 300mA, 2p и 4p -> Breaker = False)
        AddRcdSeries("Schneider", {25, 40, 63}, {"AC", "si"}, {"2p", "4p"}, {300}, "iID", False)
        ' 4. Серия iID - тип "AC" (Стандартна защита, 30mA, само 4p -> Breaker = False)
        ' Тъй като токовете са специфични (добавен е 80А и 100А), ги подаваме отделно
        AddRcdSeries("Schneider", {25, 40, 80, 100}, {"AC"}, {"4p"}, {30}, "iID", False)
        ' 5. Серия Vigi iC60 (Блок/Модул с прекъсвач, Тип AC, 30mA, 2p и 4p -> Breaker = True)
        AddRcdSeries("Schneider", {25, 40, 63}, {"AC"}, {"2p", "4p"}, {30}, "Vigi iC60", True)
        ' Пример 3: ABB Серия (Комбиниран прекъсвач + ДТЗ -> Breaker = True)
        AddRcdSeries("ABB", {10, 16, 25, 32, 40}, {"A"}, {"2p"}, {30}, "RCBO", True)
    End Sub
    ''' <summary>
    ''' Помощен метод за автоматизирано добавяне на серии ДТЗ по класа RCDInfo.
    ''' </summary>
    Private Sub AddRcdSeries(brand As String, currents As Integer(), types As String(), poles As String(), sensitivities As Integer(), deviceType As String, hasBreaker As Boolean)
        For Each p As String In poles
            For Each sensitivity As Integer In sensitivities
                For Each current As Integer In currents
                    For Each type As String In types
                        Dim rcd As New RCDInfo With {
                            .Brand = brand,
                            .NominalCurrent = current,
                            .Type = type,
                            .Poles = p.ToString(),
                            .Sensitivity = sensitivity,
                            .DeviceType = deviceType,
                            .Breaker = hasBreaker
                        }
                        Rcds.Add(rcd)
                    Next
                Next
            Next
        Next
    End Sub
    ''' <summary>
    ''' Избира подходяща дефектнотокова защита (RCD) от каталога.
    ''' </summary>
    ''' <param name="calculatedCurrent">Изчисления или изискуем минимален ток на кръга</param>
    ''' <param name="poles">Брой полюси, подадени като String (напр. "2p" или "4p")</param>
    ''' <param name="sensitivity">Търсена чувствителност в mA (по подразбиране 30 mA)</param>
    Public Function SelectRcd(calculatedCurrent As Double,
                              poles As String,
                              Breaker As Boolean,
                              Optional sensitivity As Integer = 30
                              ) As RCDInfo

        Dim minimumRequiredCurrent As Double = calculatedCurrent
        ' ====================================================================================
        ' ПОДРОБНО ОБЯСНЕНИЕ НА ТЪРСЕНЕТО (LINQ):
        ' ====================================================================================
        ' Тъй като RCDInfo вече е КЛАС (Референтен тип), променливата "selectedRcd" ще съдържа
        ' конкретната намерена инстанция или ще бъде равен на "Nothing", ако няма съвпадение.
        '
        ' ЛОГИКА НА ФИЛТРИРАНЕТО (СТЪПКА ПО СТЪПКА):
        ' 1. Първи .Where -> Сравнява полюсите като текст (r.Poles съвпада с poles, напр. "2p").
        '                    Използва се StringComparison за пълна сигурност.
        ' 2. Втори .Where -> Филтрира по марка (r.Brand), като изрично игнорира главни/малки букви.
        ' 3. Трети .Where -> Филтрира по точно съвпадение на чувствителността на утечката (mA).
        ' 4. .OrderBy -> Сортира останалите ДТЗ във възходящ ред по техния номинален ток.
        ' 5. .FirstOrDefault(...) -> Търси първия апарат от сортирания списък, чийто номинален
        '    ток е по-голям или равен на изчисления "minimumRequiredCurrent".
        ' ====================================================================================
        Dim _activeBrand As String = AppSettings.CurrentManufacturer
        Dim selectedRcd = Rcds _
            .Where(Function(r) r.Poles.Equals(poles, StringComparison.OrdinalIgnoreCase)) _
            .Where(Function(r) r.Brand.Equals(_activeBrand, StringComparison.OrdinalIgnoreCase)) _
            .Where(Function(r) r.Sensitivity = sensitivity) _
            .Where(Function(r) r.Breaker = Breaker) _
            .OrderBy(Function(r) r.NominalCurrent) _
            .FirstOrDefault(Function(r) r.NominalCurrent >= minimumRequiredCurrent)
        Return selectedRcd
    End Function
    ''' <summary>
    ''' Определя подходяща диференциална токова защита (RCD/ДЗТ) за даден токов кръг (strTokow).
    ''' </summary>
    ''' <param name="tokow">Обект от тип strTokow, представляващ токов кръг или консуматор.</param>
    ''' <remarks>
    ''' Функцията избира RCD от каталога RCD_Catalog според следните критерии:
    ''' 1. Номинален ток >= 1.2 * ток на токовия кръг (минимум 20 A)
    ''' 2. Брой полюси (2p или 4p) спрямо фазовостта на кръга
    ''' 3. Дали устройството трябва да бъде RCBO (комбиниран с прекъсвач) или само RCCB
    '''
    ''' Стъпки на логиката:
    ''' - Определя се броят на полюсите според tokow.Брой_Полюси
    ''' - Изчислява се минималният необходим номинален ток (1.2 пъти токът на кръга или минимум 20 A)
    ''' - Филтрира се каталога RCD_Catalog по номинален ток, брой полюси и тип устройство (RCBO/RCCB)
    ''' - Ако няма съвпадение:
    '''   - Показва се предупреждение с всички търсени параметри и местоположението на токовия кръг (табло, токов кръг)
    ''' - Ако има съвпадение:
    '''   - Избира се първият подходящ RCD
    '''   - Актуализират се параметрите на tokow, включително:
    '''     Brand, DeviceType, Type, Sensitivity, NominalCurrent, Poles, Нула (N) и RCD_Автомат (Breaker)
    '''
    ''' Потенциални забележки:
    ''' - Ако RCD_Catalog е празен или няма подходящ RCD, се показва съобщение, но функцията не връща грешка програмно.
    ''' - Използването на First() предполага, че списъкът matchingRCDs е сортиран или е достатъчно добър избор първият елемент.
    ''' - Полето tokow се модифицира по стойност; ако strTokow е структура (Value Type), може да се наложи връщане на обновения обект или използване на ByRef.
    ''' - Изчислението на requiredCurrent включва коефициент 1.2; 
    ''' това е запас за безопасност според стандарти.
    ''' </remarks>
    Public Sub SetRCD(tokow As clsTokow)
        If tokow.ТоковКръг = "ОБЩО" Then Return
        If tokow.ТоковКръг = "Разединител" Then Return
        ' Определяне на броя полюси на RCD: 4p за трифазен, 2p за еднофазен
        Dim poles As String = If(tokow.Брой_Полюси = 3, "4p", "2p")
        ' Минимален номинален ток: 1.2 * ток на кръга, но не по-малко от 20 A
        Dim requiredCurrent As Double = If(tokow.Ток * 1.2 < 20, 20, tokow.Ток * 1.2)
        ' Проверка дали е необходим RCBO (RCD с прекъсвач)
        Dim needRCBO As Boolean = tokow.RCD_Автомат
        Dim matchingRCD = SelectRcd(requiredCurrent, poles, needRCBO)
        ' ----------------------------------------------------
        ' Ако не е намерена подходяща ДЗТ
        ' ----------------------------------------------------
        If matchingRCD Is Nothing Then
            Dim info As String = $"ВНИМАНИЕ: Не е намерена подходяща ДЗТ!{vbCrLf}{vbCrLf}" &
                                 $"Търсени параметри:{vbCrLf}" &
                                 $"- Мин. номинален ток: {requiredCurrent} A{vbCrLf}" &
                                 $"- Комбинирана (RCBO): {If(needRCBO, "Да", "Не")}{vbCrLf}" &
                                 $"- Брой полюси: {poles}{vbCrLf}{vbCrLf}" &
                                 $"Местоположение:{vbCrLf}" &
                                 $"- Табло: {tokow.Tablo}{vbCrLf}" &
                                 $"- Токов кръг: {tokow.ТоковКръг}"
            MsgBox(info, MsgBoxStyle.Exclamation, "Липсваща апаратура в каталога")
        Else
            ' ------------------------------------------------
            ' Актуализиране на параметрите на токовия кръг
            ' според избраната ДЗТ
            ' ------------------------------------------------
            tokow.RCD_Бранд = matchingRCD.Brand
            tokow.RCD_Тип = matchingRCD.DeviceType
            tokow.RCD_Клас = matchingRCD.Type
            tokow.RCD_Чувствителност = matchingRCD.Sensitivity
            tokow.RCD_Ток = matchingRCD.NominalCurrent
            tokow.RCD_Полюси = matchingRCD.Poles
            If String.IsNullOrEmpty(tokow.RCD_Нула) Then tokow.RCD_Нула = "N"
            tokow.RCD_Автомат = matchingRCD.Breaker
        End If
    End Sub
    Public Sub ClearRCD(ByRef tokow As clsTokow)
        tokow.RCD_Бранд = ""
        tokow.RCD_Тип = ""
        tokow.RCD_Клас = ""
        tokow.RCD_Чувствителност = ""
        tokow.RCD_Ток = 0
        tokow.RCD_Полюси = 0
    End Sub
End Class
#End Region