' -------------------------------------------------------------------------
#Region "КЛАС: CableCatalog (Кабели)"
' -------------------------------------------------------------------------
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
    Public Sub CalculateCable(ByRef tokow As Form_Tablo_new.strTokow,
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
    ''' Капсулирана помощна функция, пренесена от формата за разчитане на метода за монтаж.
    ''' </summary>
    Private Function GetMountMethodInfo(mountMethod As String) As String
        Select Case mountMethod
            Case "A1" : Return "В тръба в изолация"
            Case "B1" : Return "В тръба на стена"
            Case "B2" : Return "В тръба в мазилка"
            Case "C" : Return "Директно на стена"
            Case "D1" : Return "В тръба в земя"
            Case "D2" : Return "Директно в земя"
            Case "E" : Return "На въздух/скара"
            Case "F" : Return "В пакет"
            Case Else : Return "Под мазилка"
        End Select
    End Function
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
End Class
#End Region