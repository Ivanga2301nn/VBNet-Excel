#Region "КЛАС: CableCatalog (Кабели)"
Public Class CableCatalog
    Public Class CableInfo
        Public PhaseSize As String         ' "2,5", "4", и т.н.
        Public NeutralSize As String       ' "0", "1,5", "2,5", и т.н.
        Public MaxCurrent_Air As Double    ' Допустим ток във въздух
        Public MaxCurrent_Ground As Double ' Допустим ток в земя
        Public Material As String          ' "Cu", "Al"
        Public CableType As String         ' "СВТ", "САВТ", "Al/R"
        Public MaxWorkingTemp As Double    ' (65, 70, 90°C)
        Public InsulationType As String    ' ("ПВЦ", "XLPE", "GUM")
    End Class
    ' Складът за всички кабели (замества DataList)
    Public Property DataList As New List(Of CableInfo)()
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
        DataList.Clear()
        DataList.Add(New CableInfo With {.CableType = "СВТ", .Material = "Cu", .PhaseSize = "1,5", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 19, .MaxCurrent_Ground = 25, .NeutralSize = "1,5"})
        DataList.Add(New CableInfo With {.CableType = "СВТ", .Material = "Cu", .PhaseSize = "2,5", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 25, .MaxCurrent_Ground = 34, .NeutralSize = "2,5"})
        DataList.Add(New CableInfo With {.CableType = "СВТ", .Material = "Cu", .PhaseSize = "4", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 34, .MaxCurrent_Ground = 45, .NeutralSize = "4"})
        DataList.Add(New CableInfo With {.CableType = "СВТ", .Material = "Cu", .PhaseSize = "6", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 43, .MaxCurrent_Ground = 55, .NeutralSize = "6"})
        DataList.Add(New CableInfo With {.CableType = "СВТ", .Material = "Cu", .PhaseSize = "10", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 59, .MaxCurrent_Ground = 76, .NeutralSize = "10"})
        DataList.Add(New CableInfo With {.CableType = "СВТ", .Material = "Cu", .PhaseSize = "16", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 79, .MaxCurrent_Ground = 96, .NeutralSize = "16"})
        DataList.Add(New CableInfo With {.CableType = "СВТ", .Material = "Cu", .PhaseSize = "25", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 105, .MaxCurrent_Ground = 126, .NeutralSize = "16"})
        DataList.Add(New CableInfo With {.CableType = "СВТ", .Material = "Cu", .PhaseSize = "35", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 126, .MaxCurrent_Ground = 151, .NeutralSize = "16"})
        DataList.Add(New CableInfo With {.CableType = "СВТ", .Material = "Cu", .PhaseSize = "50", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 157, .MaxCurrent_Ground = 178, .NeutralSize = "25"})
        DataList.Add(New CableInfo With {.CableType = "СВТ", .Material = "Cu", .PhaseSize = "70", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 199, .MaxCurrent_Ground = 225, .NeutralSize = "35"})
        DataList.Add(New CableInfo With {.CableType = "СВТ", .Material = "Cu", .PhaseSize = "95", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 246, .MaxCurrent_Ground = 270, .NeutralSize = "50"})
        DataList.Add(New CableInfo With {.CableType = "СВТ", .Material = "Cu", .PhaseSize = "120", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 285, .MaxCurrent_Ground = 306, .NeutralSize = "70"})
        DataList.Add(New CableInfo With {.CableType = "СВТ", .Material = "Cu", .PhaseSize = "150", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 326, .MaxCurrent_Ground = 346, .NeutralSize = "70"})
        DataList.Add(New CableInfo With {.CableType = "СВТ", .Material = "Cu", .PhaseSize = "185", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 374, .MaxCurrent_Ground = 390, .NeutralSize = "95"})
        DataList.Add(New CableInfo With {.CableType = "СВТ", .Material = "Cu", .PhaseSize = "240", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 445, .MaxCurrent_Ground = 458, .NeutralSize = "120"})

        DataList.Add(New CableInfo With {.CableType = "САВТ", .Material = "Al", .PhaseSize = "1,5", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 0, .MaxCurrent_Ground = 0, .NeutralSize = "1,5"})
        DataList.Add(New CableInfo With {.CableType = "САВТ", .Material = "Al", .PhaseSize = "2,5", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 20, .MaxCurrent_Ground = 25, .NeutralSize = "2,5"})
        DataList.Add(New CableInfo With {.CableType = "САВТ", .Material = "Al", .PhaseSize = "4", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 26, .MaxCurrent_Ground = 32, .NeutralSize = "4"})
        DataList.Add(New CableInfo With {.CableType = "САВТ", .Material = "Al", .PhaseSize = "6", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 34, .MaxCurrent_Ground = 42, .NeutralSize = "6"})
        DataList.Add(New CableInfo With {.CableType = "САВТ", .Material = "Al", .PhaseSize = "10", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 43, .MaxCurrent_Ground = 53, .NeutralSize = "10"})
        DataList.Add(New CableInfo With {.CableType = "САВТ", .Material = "Al", .PhaseSize = "16", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 64, .MaxCurrent_Ground = 75, .NeutralSize = "16"})
        DataList.Add(New CableInfo With {.CableType = "САВТ", .Material = "Al", .PhaseSize = "25", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 82, .MaxCurrent_Ground = 92, .NeutralSize = "16"})
        DataList.Add(New CableInfo With {.CableType = "САВТ", .Material = "Al", .PhaseSize = "35", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 100, .MaxCurrent_Ground = 110, .NeutralSize = "16"})
        DataList.Add(New CableInfo With {.CableType = "САВТ", .Material = "Al", .PhaseSize = "50", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 119, .MaxCurrent_Ground = 134, .NeutralSize = "25"})
        DataList.Add(New CableInfo With {.CableType = "САВТ", .Material = "Al", .PhaseSize = "70", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 152, .MaxCurrent_Ground = 170, .NeutralSize = "35"})
        DataList.Add(New CableInfo With {.CableType = "САВТ", .Material = "Al", .PhaseSize = "95", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 185, .MaxCurrent_Ground = 210, .NeutralSize = "50"})
        DataList.Add(New CableInfo With {.CableType = "САВТ", .Material = "Al", .PhaseSize = "120", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 215, .MaxCurrent_Ground = 245, .NeutralSize = "70"})
        DataList.Add(New CableInfo With {.CableType = "САВТ", .Material = "Al", .PhaseSize = "150", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 245, .MaxCurrent_Ground = 274, .NeutralSize = "70"})
        DataList.Add(New CableInfo With {.CableType = "САВТ", .Material = "Al", .PhaseSize = "185", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 285, .MaxCurrent_Ground = 310, .NeutralSize = "95"})
        DataList.Add(New CableInfo With {.CableType = "САВТ", .Material = "Al", .PhaseSize = "240", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 338, .MaxCurrent_Ground = 360, .NeutralSize = "120"})

        DataList.Add(New CableInfo With {.CableType = "Al/R", .Material = "Al", .PhaseSize = "16", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 83, .MaxCurrent_Ground = 0, .NeutralSize = "16"})
        DataList.Add(New CableInfo With {.CableType = "Al/R", .Material = "Al", .PhaseSize = "25", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 111, .MaxCurrent_Ground = 0, .NeutralSize = "25"})
        DataList.Add(New CableInfo With {.CableType = "Al/R", .Material = "Al", .PhaseSize = "35", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 138, .MaxCurrent_Ground = 0, .NeutralSize = "35"})
        DataList.Add(New CableInfo With {.CableType = "Al/R", .Material = "Al", .PhaseSize = "35", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 138, .MaxCurrent_Ground = 0, .NeutralSize = "54"})
        DataList.Add(New CableInfo With {.CableType = "Al/R", .Material = "Al", .PhaseSize = "50", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 164, .MaxCurrent_Ground = 0, .NeutralSize = "54"})
        DataList.Add(New CableInfo With {.CableType = "Al/R", .Material = "Al", .PhaseSize = "70", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 213, .MaxCurrent_Ground = 0, .NeutralSize = "54"})
        DataList.Add(New CableInfo With {.CableType = "Al/R", .Material = "Al", .PhaseSize = "95", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 258, .MaxCurrent_Ground = 0, .NeutralSize = "70"})
        DataList.Add(New CableInfo With {.CableType = "Al/R", .Material = "Al", .PhaseSize = "150", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 344, .MaxCurrent_Ground = 0, .NeutralSize = "70"})

        DataList.Add(New CableInfo With {.CableType = "ПВ-А1", .Material = "Cu", .PhaseSize = "1,5", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 20, .MaxCurrent_Ground = 29, .NeutralSize = "0"})
        DataList.Add(New CableInfo With {.CableType = "ПВ-А1", .Material = "Cu", .PhaseSize = "2,5", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 27, .MaxCurrent_Ground = 38, .NeutralSize = "0"})
        DataList.Add(New CableInfo With {.CableType = "ПВ-А1", .Material = "Cu", .PhaseSize = "4", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 36, .MaxCurrent_Ground = 49, .NeutralSize = "0"})
        DataList.Add(New CableInfo With {.CableType = "ПВ-А1", .Material = "Cu", .PhaseSize = "6", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 45, .MaxCurrent_Ground = 62, .NeutralSize = "0"})
        DataList.Add(New CableInfo With {.CableType = "ПВ-А1", .Material = "Cu", .PhaseSize = "10", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 63, .MaxCurrent_Ground = 83, .NeutralSize = "0"})
        DataList.Add(New CableInfo With {.CableType = "ПВ-А1", .Material = "Cu", .PhaseSize = "16", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 82, .MaxCurrent_Ground = 104, .NeutralSize = "0"})
        DataList.Add(New CableInfo With {.CableType = "ПВ-А1", .Material = "Cu", .PhaseSize = "25", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 113, .MaxCurrent_Ground = 136, .NeutralSize = "0"})
        DataList.Add(New CableInfo With {.CableType = "ПВ-А1", .Material = "Cu", .PhaseSize = "35", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 138, .MaxCurrent_Ground = 162, .NeutralSize = "0"})
        DataList.Add(New CableInfo With {.CableType = "ПВ-А1", .Material = "Cu", .PhaseSize = "50", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 168, .MaxCurrent_Ground = 192, .NeutralSize = "0"})
        DataList.Add(New CableInfo With {.CableType = "ПВ-А1", .Material = "Cu", .PhaseSize = "70", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 210, .MaxCurrent_Ground = 236, .NeutralSize = "0"})
        DataList.Add(New CableInfo With {.CableType = "ПВ-А1", .Material = "Cu", .PhaseSize = "95", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 262, .MaxCurrent_Ground = 285, .NeutralSize = "0"})
        DataList.Add(New CableInfo With {.CableType = "ПВ-А1", .Material = "Cu", .PhaseSize = "120", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 307, .MaxCurrent_Ground = 322, .NeutralSize = "0"})
        DataList.Add(New CableInfo With {.CableType = "ПВ-А1", .Material = "Cu", .PhaseSize = "150", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 352, .MaxCurrent_Ground = 363, .NeutralSize = "0"})
        DataList.Add(New CableInfo With {.CableType = "ПВ-А1", .Material = "Cu", .PhaseSize = "185", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 405, .MaxCurrent_Ground = 410, .NeutralSize = "0"})
        DataList.Add(New CableInfo With {.CableType = "ПВ-А1", .Material = "Cu", .PhaseSize = "240", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 482, .MaxCurrent_Ground = 475, .NeutralSize = "0"})

        DataList.Add(New CableInfo With {.CableType = "NYY", .Material = "Cu", .PhaseSize = "1,5", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 19.5, .MaxCurrent_Ground = 27, .NeutralSize = "1,5"})
        DataList.Add(New CableInfo With {.CableType = "NYY", .Material = "Cu", .PhaseSize = "2,5", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 25, .MaxCurrent_Ground = 36, .NeutralSize = "2,5"})
        DataList.Add(New CableInfo With {.CableType = "NYY", .Material = "Cu", .PhaseSize = "4", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 34, .MaxCurrent_Ground = 47, .NeutralSize = "4"})
        DataList.Add(New CableInfo With {.CableType = "NYY", .Material = "Cu", .PhaseSize = "6", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 43, .MaxCurrent_Ground = 59, .NeutralSize = "6"})
        DataList.Add(New CableInfo With {.CableType = "NYY", .Material = "Cu", .PhaseSize = "10", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 59, .MaxCurrent_Ground = 79, .NeutralSize = "10"})
        DataList.Add(New CableInfo With {.CableType = "NYY", .Material = "Cu", .PhaseSize = "16", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 79, .MaxCurrent_Ground = 102, .NeutralSize = "16"})
        DataList.Add(New CableInfo With {.CableType = "NYY", .Material = "Cu", .PhaseSize = "25", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 106, .MaxCurrent_Ground = 133, .NeutralSize = "16"})
        DataList.Add(New CableInfo With {.CableType = "NYY", .Material = "Cu", .PhaseSize = "35", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 129, .MaxCurrent_Ground = 159, .NeutralSize = "16"})
        DataList.Add(New CableInfo With {.CableType = "NYY", .Material = "Cu", .PhaseSize = "50", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 157, .MaxCurrent_Ground = 188, .NeutralSize = "25"})
        DataList.Add(New CableInfo With {.CableType = "NYY", .Material = "Cu", .PhaseSize = "70", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 199, .MaxCurrent_Ground = 232, .NeutralSize = "35"})
        DataList.Add(New CableInfo With {.CableType = "NYY", .Material = "Cu", .PhaseSize = "95", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 246, .MaxCurrent_Ground = 280, .NeutralSize = "50"})
        DataList.Add(New CableInfo With {.CableType = "NYY", .Material = "Cu", .PhaseSize = "120", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 285, .MaxCurrent_Ground = 318, .NeutralSize = "70"})
        DataList.Add(New CableInfo With {.CableType = "NYY", .Material = "Cu", .PhaseSize = "150", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 326, .MaxCurrent_Ground = 359, .NeutralSize = "70"})
        DataList.Add(New CableInfo With {.CableType = "NYY", .Material = "Cu", .PhaseSize = "185", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 374, .MaxCurrent_Ground = 406, .NeutralSize = "95"})
        DataList.Add(New CableInfo With {.CableType = "NYY", .Material = "Cu", .PhaseSize = "240", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 445, .MaxCurrent_Ground = 473, .NeutralSize = "120"})

        DataList.Add(New CableInfo With {.CableType = "NAYY", .Material = "Al", .PhaseSize = "1,5", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 0, .MaxCurrent_Ground = 0, .NeutralSize = "1,5"})
        DataList.Add(New CableInfo With {.CableType = "NAYY", .Material = "Al", .PhaseSize = "2,5", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 0, .MaxCurrent_Ground = 0, .NeutralSize = "2,5"})
        DataList.Add(New CableInfo With {.CableType = "NAYY", .Material = "Al", .PhaseSize = "4", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 0, .MaxCurrent_Ground = 0, .NeutralSize = "4"})
        DataList.Add(New CableInfo With {.CableType = "NAYY", .Material = "Al", .PhaseSize = "6", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 0, .MaxCurrent_Ground = 0, .NeutralSize = "6"})
        DataList.Add(New CableInfo With {.CableType = "NAYY", .Material = "Al", .PhaseSize = "10", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 0, .MaxCurrent_Ground = 0, .NeutralSize = "10"})
        DataList.Add(New CableInfo With {.CableType = "NAYY", .Material = "Al", .PhaseSize = "16", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 0, .MaxCurrent_Ground = 0, .NeutralSize = "16"})
        DataList.Add(New CableInfo With {.CableType = "NAYY", .Material = "Al", .PhaseSize = "25", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 82, .MaxCurrent_Ground = 102, .NeutralSize = "16"})
        DataList.Add(New CableInfo With {.CableType = "NAYY", .Material = "Al", .PhaseSize = "35", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 100, .MaxCurrent_Ground = 123, .NeutralSize = "16"})
        DataList.Add(New CableInfo With {.CableType = "NAYY", .Material = "Al", .PhaseSize = "50", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 119, .MaxCurrent_Ground = 144, .NeutralSize = "25"})
        DataList.Add(New CableInfo With {.CableType = "NAYY", .Material = "Al", .PhaseSize = "70", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 152, .MaxCurrent_Ground = 179, .NeutralSize = "35"})
        DataList.Add(New CableInfo With {.CableType = "NAYY", .Material = "Al", .PhaseSize = "95", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 186, .MaxCurrent_Ground = 215, .NeutralSize = "50"})
        DataList.Add(New CableInfo With {.CableType = "NAYY", .Material = "Al", .PhaseSize = "120", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 216, .MaxCurrent_Ground = 245, .NeutralSize = "70"})
        DataList.Add(New CableInfo With {.CableType = "NAYY", .Material = "Al", .PhaseSize = "150", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 246, .MaxCurrent_Ground = 275, .NeutralSize = "70"})
        DataList.Add(New CableInfo With {.CableType = "NAYY", .Material = "Al", .PhaseSize = "185", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 285, .MaxCurrent_Ground = 313, .NeutralSize = "95"})
        DataList.Add(New CableInfo With {.CableType = "NAYY", .Material = "Al", .PhaseSize = "240", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 338, .MaxCurrent_Ground = 364, .NeutralSize = "120"})

        DataList.Add(New CableInfo With {.CableType = "N2XY", .Material = "Cu", .PhaseSize = "1,5", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 24, .MaxCurrent_Ground = 31, .NeutralSize = "1,5"})
        DataList.Add(New CableInfo With {.CableType = "N2XY", .Material = "Cu", .PhaseSize = "2,5", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 32, .MaxCurrent_Ground = 40, .NeutralSize = "2,5"})
        DataList.Add(New CableInfo With {.CableType = "N2XY", .Material = "Cu", .PhaseSize = "4", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 42, .MaxCurrent_Ground = 52, .NeutralSize = "4"})
        DataList.Add(New CableInfo With {.CableType = "N2XY", .Material = "Cu", .PhaseSize = "6", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 53, .MaxCurrent_Ground = 64, .NeutralSize = "6"})
        DataList.Add(New CableInfo With {.CableType = "N2XY", .Material = "Cu", .PhaseSize = "10", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 74, .MaxCurrent_Ground = 86, .NeutralSize = "10"})
        DataList.Add(New CableInfo With {.CableType = "N2XY", .Material = "Cu", .PhaseSize = "16", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 98, .MaxCurrent_Ground = 112, .NeutralSize = "16"})
        DataList.Add(New CableInfo With {.CableType = "N2XY", .Material = "Cu", .PhaseSize = "25", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 133, .MaxCurrent_Ground = 145, .NeutralSize = "16"})
        DataList.Add(New CableInfo With {.CableType = "N2XY", .Material = "Cu", .PhaseSize = "35", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 162, .MaxCurrent_Ground = 174, .NeutralSize = "16"})
        DataList.Add(New CableInfo With {.CableType = "N2XY", .Material = "Cu", .PhaseSize = "50", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 197, .MaxCurrent_Ground = 206, .NeutralSize = "25"})
        DataList.Add(New CableInfo With {.CableType = "N2XY", .Material = "Cu", .PhaseSize = "70", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 250, .MaxCurrent_Ground = 254, .NeutralSize = "35"})
        DataList.Add(New CableInfo With {.CableType = "N2XY", .Material = "Cu", .PhaseSize = "95", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 308, .MaxCurrent_Ground = 305, .NeutralSize = "50"})
        DataList.Add(New CableInfo With {.CableType = "N2XY", .Material = "Cu", .PhaseSize = "120", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 359, .MaxCurrent_Ground = 348, .NeutralSize = "70"})
        DataList.Add(New CableInfo With {.CableType = "N2XY", .Material = "Cu", .PhaseSize = "150", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 412, .MaxCurrent_Ground = 392, .NeutralSize = "70"})
        DataList.Add(New CableInfo With {.CableType = "N2XY", .Material = "Cu", .PhaseSize = "185", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 475, .MaxCurrent_Ground = 444, .NeutralSize = "95"})
        DataList.Add(New CableInfo With {.CableType = "N2XY", .Material = "Cu", .PhaseSize = "240", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 564, .MaxCurrent_Ground = 517, .NeutralSize = "120"})

        DataList.Add(New CableInfo With {.CableType = "NA2XY", .Material = "Al", .PhaseSize = "1,5", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 0, .MaxCurrent_Ground = 0, .NeutralSize = "1,5"})
        DataList.Add(New CableInfo With {.CableType = "NA2XY", .Material = "Al", .PhaseSize = "2,5", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 0, .MaxCurrent_Ground = 0, .NeutralSize = "2,5"})
        DataList.Add(New CableInfo With {.CableType = "NA2XY", .Material = "Al", .PhaseSize = "4", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 0, .MaxCurrent_Ground = 0, .NeutralSize = "4"})
        DataList.Add(New CableInfo With {.CableType = "NA2XY", .Material = "Al", .PhaseSize = "6", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 0, .MaxCurrent_Ground = 0, .NeutralSize = "6"})
        DataList.Add(New CableInfo With {.CableType = "NA2XY", .Material = "Al", .PhaseSize = "10", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 0, .MaxCurrent_Ground = 0, .NeutralSize = "10"})
        DataList.Add(New CableInfo With {.CableType = "NA2XY", .Material = "Al", .PhaseSize = "16", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 0, .MaxCurrent_Ground = 0, .NeutralSize = "16"})
        DataList.Add(New CableInfo With {.CableType = "NA2XY", .Material = "Al", .PhaseSize = "25", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 102, .MaxCurrent_Ground = 112, .NeutralSize = "16"})
        DataList.Add(New CableInfo With {.CableType = "NA2XY", .Material = "Al", .PhaseSize = "35", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 126, .MaxCurrent_Ground = 135, .NeutralSize = "16"})
        DataList.Add(New CableInfo With {.CableType = "NA2XY", .Material = "Al", .PhaseSize = "50", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 149, .MaxCurrent_Ground = 158, .NeutralSize = "25"})
        DataList.Add(New CableInfo With {.CableType = "NA2XY", .Material = "Al", .PhaseSize = "70", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 191, .MaxCurrent_Ground = 196, .NeutralSize = "35"})
        DataList.Add(New CableInfo With {.CableType = "NA2XY", .Material = "Al", .PhaseSize = "95", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 234, .MaxCurrent_Ground = 234, .NeutralSize = "50"})
        DataList.Add(New CableInfo With {.CableType = "NA2XY", .Material = "Al", .PhaseSize = "120", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 273, .MaxCurrent_Ground = 268, .NeutralSize = "70"})
        DataList.Add(New CableInfo With {.CableType = "NA2XY", .Material = "Al", .PhaseSize = "150", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 311, .MaxCurrent_Ground = 300, .NeutralSize = "70"})
        DataList.Add(New CableInfo With {.CableType = "NA2XY", .Material = "Al", .PhaseSize = "185", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 360, .MaxCurrent_Ground = 342, .NeutralSize = "95"})
        DataList.Add(New CableInfo With {.CableType = "NA2XY", .Material = "Al", .PhaseSize = "240", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 427, .MaxCurrent_Ground = 398, .NeutralSize = "120"})

        CableTypesForCombo = DataList.Select(Function(b) b.CableType).Distinct().ToList()
    End Sub
    ''' <summary>
    ''' Изчислява необходимото сечение на кабел според тока и условията на полагане
    ''' Оптимизиран за сградни инсталации
    ''' </summary>
    Public Sub CalculateCable(ByRef tokow As strTokow,
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
        Dim filteredCables = DataList.Where(
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
        Public TripUnit As String                   ' Защитен блок (TM-D, Micrologic...).
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
        Brand_For_combo = New List(Of String) From {"Schneider Electric", "Schrack Technik", "Siemens", "ABB"}
    End Sub
    ''' <summary>
    ''' Генерира декартово произведение от всички възможни комбинации на параметрите 
    ''' и ги добавя в списъка 'Breakers'.
    ''' </summary>
    ''' <param name="brand">Марка на апаратурата (напр. "Schneider", "Schrack")</param>
    ''' <param name="series">Серия на модела (напр. "EZ9 MCB", "Amparo 6kA")</param>
    ''' <param name="category">Категория на уреда (напр. "MCB", "RCCB", "RCBO", "MCCB")</param>
    ''' <param name="currents">Номинални токове Inom в Ампери - Integer() {6, 10, 16...}</param>
    ''' <param name="curves">Криви на изключване (B, C, D) - String(). Ако е излишно (напр. за ДТЗ), подай: Nothing (автоматично става "-")</param>
    ''' <param name="polesList">Брой полюси (1, 2, 3, 4) - Integer(). Ако е Nothing, автоматично става: 0</param>
    ''' <param name="icsValues">Изключвателна способност в kA - Decimal(). Ако е Nothing, автоматично става: 0</param>
    ''' <param name="tripUnits">Тип защита/утечка (напр. "30mA", "TM", "Electronic") - String(). Ако е Nothing, остава: Nothing</param>
    ''' <remarks>
    ''' ПОДРЕДБА НА ПАРАМЕТРИТЕ ПРИ ИЗВИКВАНЕ:
    ''' 1. Brand (String)      -> "Schrack"
    ''' 2. Series (String)     -> "Amparo"
    ''' 3. Category (String)   -> "MCB"
    ''' 4. Currents (Integer()) -> {6, 10, 16, 20}
    ''' 5. Curves (String())   -> {"B", "C"}       <-- (Подай Nothing за уреди без крива)
    ''' 6. Poles (Integer())   -> {1, 3}
    ''' 7. Ics kA (Decimal())  -> {6}
    ''' 8. TripUnit (String()) -> Nothing          <-- (За ДТЗ/електронни защити, напр. {"30mA"})
    ''' 
    ''' КАК РАБОТИ ВГРАДЕНАТА ЗАЩИТА С If(..., New ...):
    ''' Ако подадеш 'Nothing' за масив, циклите няма да гръмнат, а ще се завъртят точно веднъж със следните служебни стойности:
    ''' - Curves    -> "-"
    ''' - Poles     -> 0
    ''' - Ics_kA    -> 0
    ''' - TripUnit  -> Nothing
    ''' </remarks>
    Public Sub LoadCatalog(selectedBrand As String)
        Breakers.Clear()
        ' В момента генерираме кода софтуерно в зависимост от избора
        Select Case selectedBrand
            Case "Schneider Electric", "Schneider"
                ' MCB
                AddBreakerSeries("Schneider", "EZ9 MCB", "MCB",
                                 {6, 10, 16, 20, 25, 32, 40, 50, 63},
                                 {"C", "B", "D"}, {1, 3}, {6}, Nothing)
                AddBreakerSeries("Schneider", "iC60N", "MCB",
                                 {2, 3, 4, 6, 10, 16, 20, 25, 32, 40, 50, 63},
                                 {"C", "B", "D"}, {1, 3}, {6}, Nothing)
                AddBreakerSeries("Schneider", "C120", "MCB",
                                 {80, 100, 125}, {"C", "D"},
                                 {1, 3}, {10}, Nothing)
                ' MCCB
                AddBreakerSeries("Schneider", "NSXm", "MCCB",
                                 {16, 25, 32, 40, 50, 63, 80, 100, 125, 160},
                                 {"E", "B", "F", "N", "H"}, {3}, {25}, {"TM-D", "TM-DC"})
                AddBreakerSeries("Schneider", "NSX100", "MCCB", {16, 25, 32, 40, 63, 80, 100}, {"B", "F", "N", "H", "S", "L"}, {3}, {25}, {"TM-D", "TM-DC"})
                AddBreakerSeries("Schneider", "NSX160", "MCCB", {80, 100, 125, 160}, {"B", "F", "N", "H", "S", "L"}, {3}, {36}, {"TM-D"})
                AddBreakerSeries("Schneider", "NSX250", "MCCB", {125, 160, 200, 250}, {"B", "F", "N", "H", "S", "L"}, {3}, {50}, {"TM-D", "Micrologic 2.0", "Micrologic 5.0"})
                AddBreakerSeries("Schneider", "NSX400", "MCCB", {250, 320, 400}, {"F", "N", "H", "S", "L"}, {3}, {70}, {"Micrologic 2.3"})
                AddBreakerSeries("Schneider", "NSX630", "MCCB", {400, 500, 630}, {"F", "N", "H", "S", "L"}, {3}, {100}, {"Micrologic 2.3"})
                ' ACB
                AddBreakerSeries("Schneider", "MTZ", "ACB", {800, 1000, 1250, 1600, 2000, 2500, 3200, 4000, 5000, 6300}, {"MTZ"}, {3, 4}, {42, 65, 100}, {"Micrologic 6.0"})
            Case "Siemens"
                ' TODO: Тук утре по същия начин се добавят редове за Siemens без промяна на друга логика
                ' AddBreakerSeries("Siemens", "5SY", "MCB", {6, 10, 16...}, ...)
            Case "ABB"
                ' TODO: Тук се добавят редове за ABB утре
            Case "Schrack Technik", "Schrack"
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
        End Select
        ' ✅ Автоматично пълним масивите за останалите ComboBox-ове на база избраната марка
        Breakers_For_combo = Breakers.Select(Function(b) b.Series).Distinct().ToList()
        TripUnit_For_combo = Breakers.Select(Function(b) b.TripUnit).Distinct().ToList()
        Curve_For_combo = Breakers.Select(Function(b) b.Curve).Distinct().ToList()
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
    Public Sub CalculateBreaker(ByRef tokow As strTokow)
        ' Деклариране на променлива за намерения прекъсвач от новия клас
        Dim breaker As BreakerCatalog.BreakerInfo = Nothing
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
    ' <summary>
    ' Избира подходящ прекъсвач (BreakerInfo) от колекцията Breakers
    ' според:
    '
    ' - изчислен ток
    ' - брой полюси
    ' - крива или серия
    '
    ' Логиката работи на два етапа:
    '
    ' 1. Основен избор:
    '    търси най-близкия стандартен прекъсвач,
    '    който покрива тока с резерв.
    '
    ' 2. Fallback избор:
    '    ако няма намерен прекъсвач за зададената
    '    крива/серия → търси произволен подходящ прекъсвач.
    '
    ' Цел:
    ' - автоматичен подбор на защитен апарат
    ' - избягване на undersized прекъсвач
    ' - използване на стандартни номинали
    ' </summary>
    '
    ' <param name="calculatedCurrent">
    ' Изчислен работен ток.
    ' </param>
    '
    ' <param name="poles">
    ' Необходим брой полюси.
    ' </param>
    '
    ' <param name="curveOrSeries">
    ' Крива или серия на прекъсвача.
    '
    ' По подразбиране:
    ' "C"
    ' </param>
    '
    ' <returns>
    ' BreakerInfo:
    ' избраният прекъсвач,
    ' или Nothing при липса на подходящ.
    ' </returns>
    Public Function SelectBreaker(calculatedCurrent As Double,
                              poles As Integer,
                              Optional curveOrSeries As String = "C") As BreakerInfo
        ' <summary>
        ' резерв при подбора на прекъсвача.        '
        ' Прекъсвачът трябва да бъде
        ' по-голям от изчисления ток.
        ' </summary>
        Const SAFETY_FACTOR As Double = 1.15
        ' <summary>
        ' minimumRequiredCurrent:
        ' минимален необходим номинален ток
        ' след прилагане на резерва.
        ' </summary>
        Dim minimumRequiredCurrent As Double = calculatedCurrent * SAFETY_FACTOR
        ' СТЪПКА 1: Опит за намиране на прекъсвач по точни критерии (Полюси + Крива/Серия + Ток)
        Dim selectedBreaker = Breakers _
            .Where(Function(b) b.Poles = poles) _
            .Where(Function(b) b.Curve.Equals(curveOrSeries, StringComparison.OrdinalIgnoreCase) OrElse
            b.Series.Equals(curveOrSeries, StringComparison.OrdinalIgnoreCase)) _
            .OrderBy(Function(b) b.NominalCurrent) _
            .FirstOrDefault(Function(b) b.NominalCurrent >= minimumRequiredCurrent)        ' 1. Филтър по полюси
        ' 2. Филтър по Крива или Серия (игнорирайки регистъра на буквите)
        ' 3. Сортиране по номинален ток (от най-малък към най-голям)
        ' 4. Вземане на първия прекъсвач, чийто ток е по-голям или равен на минимално изисквания
        ' СТЪПКА 2: Резервен вариант (Fallback), ако първото търсене не е върнало резултат
        If selectedBreaker Is Nothing Then
            ' Ако не е намерен съвпадащ прекъсвач, търсим алтернативен само по полюси и изчислен ток
            selectedBreaker = Breakers _
                .Where(Function(b) b.Poles = poles) _
                .Where(Function(b) b.NominalCurrent >= calculatedCurrent) _
                .OrderBy(Function(b) b.NominalCurrent) _
                .FirstOrDefault()
            ' 1. Филтрираме по същия брой полюси
            ' 2. Филтрираме прекъсвачите с ток, по-голям или равен на изчисления (calculatedCurrent)
            ' 3. Сортираме ги по ток възходящо
            ' 4. Вземаме първия (най-икономичния/близкия по стойност) прекъсвач
        End If
        ' =============================================================
        ' ВРЪЩАНЕ НА РЕЗУЛТАТ
        ' =============================================================
        Return selectedBreaker
    End Function
    ''' <summary>
    ''' Изчиства данните за прекъсвач (MCB)
    ''' </summary>
    Public Sub ClearBreaker(ByRef tokow As strTokow)
        tokow.Breaker_Тип_Апарат = ""           ' Серия апарат (EZ9, C120, NSX, MTZ)
        tokow.Breaker_Крива = ""                ' Характеристика (B, C, D)
        tokow.Breaker_Номинален_Ток = ""        ' Номинален ток (пример: "16A")
        tokow.Breaker_Изкл_Възможност = ""      ' Изключвателна способност ("6000A", "10000A")
        tokow.Breaker_Защитен_блок = ""         ' Изключвателна способност ("6000A", "10000A")
    End Sub
End Class
#End Region