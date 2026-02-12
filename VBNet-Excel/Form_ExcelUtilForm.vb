Imports System.IO
Imports System.Net
Imports System.Security.Policy
Imports System.Text.RegularExpressions
Imports System.Web
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Microsoft.Office.Interop
Imports Microsoft.VisualBasic.Devices
Imports excel = Microsoft.Office.Interop.Excel
Imports word = Microsoft.Office.Interop.Word

Public Class Form_ExcelUtilForm
    Dim PI As Double = 3.1415926535897931
    Dim cu As CommonUtil = New CommonUtil()

    Dim wordApp As New word.Application
    Dim excel_Workbook As excel.Workbook

    Dim wsKS As excel.Worksheet
    Dim wsLines As excel.Worksheet
    Dim wsKontakti As excel.Worksheet
    Dim wsLight As excel.Worksheet
    Dim wsElboard As excel.Worksheet
    Dim wsCableTrays As excel.Worksheet
    Dim wsKol_Smetka As excel.Worksheet
    Dim wsSpecefikaciq As excel.Worksheet
    Dim wsInternet As excel.Worksheet
    Dim wsKoef As excel.Worksheet
    Dim wsPIC As excel.Worksheet
    Dim wsDOMOF As excel.Worksheet
    Dim wsBOLNICA As excel.Worksheet
    Dim wwsKONTROL As excel.Worksheet
    Dim wsOPOWES As excel.Worksheet
    Dim wsSOT As excel.Worksheet
    Dim wsVIDEO As excel.Worksheet
    Dim wsKa4vane As excel.Worksheet

    Dim fullName As String
    Dim filePath As String
    Dim fileName As String
    Dim nameExcel As String

    Dim Kz As Double
    Dim Бр_Етажи As Double
    Dim H_Етаж As Double
    Dim Кабел_Розетка As Double
    Dim Кабел_кутия As Double
    Dim H_Контакт As Double
    Dim H_Ключ As Double
    Dim koefYes As Boolean = vbFalse

    Dim ProgressBar_Maximum As Integer = 700

    Private NoMyExcelProcesses() As Process

    Const red_Фотоволтаици As Integer = 40 ' Ред в wsKoef от който започват да се записват фотоволтайчните панели и инвертори
    Const red_Външно As Integer = 17 ' Ред в wsKoef от който започват да се записват фотоволтайчните панели и инвертори

    Dim File_PV As Boolean = False
    Public Structure strZazeml
        Dim blVisibility As String
        Dim blТАБЛО As String
        Dim blНадпис As String
        Dim blName As String
        Dim count As Integer
    End Structure
    Public Structure strKontakt
        Dim count As Integer
        Dim blName As String
        Dim blVisibility As String
        Dim blText As String
        Dim blInsert As Boolean
        Dim blWis As String
        Dim blMO6T_TRAFO As String
        Dim blLamp_Power As String
        Dim blLED_Lenta_Length As String
        Dim blLED_Lamp_Montav As String
        Dim blLED_Lamp_SW_Potok As String
        Dim blLED_Lamp_TIP As String
        Dim blRACH_NAIMENO As String
        Dim blRACH_UNIT As String
        Dim blRACH_Wiso4ina As String
    End Structure
    Public Structure strKabel
        Dim blType As String
        Dim blPol As String
        Dim blLength As Double
        Dim blCount As Double
    End Structure
    Public Structure strКачване
        Dim KOTA_1 As String
        Dim KOTA_2 As String
        Dim ТРЪБА_1 As String
        Dim ТРЪБА_2 As String
        Dim Kabel_d_0 As String
        Dim Kabel_d_1 As String
        Dim Kabel_d_2 As String
        Dim Kabel_d_3 As String
        Dim Kabel_d_4 As String
        Dim Kabel_d_5 As String
        Dim Kabel_d_6 As String
        Dim Kabel_d_7 As String
        Dim Kabel_d_8 As String
        Dim Kabel_d_9 As String
        Dim Kabel_d_10 As String
        Dim Kabel_g_0 As String
        Dim Kabel_g_1 As String
        Dim Kabel_g_2 As String
        Dim Kabel_g_3 As String
        Dim Kabel_g_4 As String
        Dim Kabel_g_5 As String
        Dim Kabel_g_6 As String
        Dim Kabel_g_7 As String
        Dim Kabel_g_8 As String
        Dim Kabel_g_9 As String
        Dim Kabel_g_10 As String
    End Structure
    Public Structure strТабло
        Dim bl_Табло As String
        Dim bl_Брой As Integer
        Dim bl_ИмеБлок As String
        Dim bl_1 As String
        Dim bl_2 As String
        Dim bl_3 As String
        Dim bl_4 As String
        Dim bl_5 As String
        Dim bl_6 As String
        Dim bl_7 As String
        Dim bl_8 As String
        Dim bl_9 As String
        Dim bl_10 As String
        Dim bl_DESIGNATION As String
        Dim bl_LONGNAME As String
        Dim bl_REFNB As String
        Dim bl_SHORTNAME As String
        Dim bl_RABATY As String
        Dim bl_RABATY2 As String
    End Structure
    Public Structure strСкара
        Dim bl_Ширина As Integer
        Dim bl_Височина As Integer
        Dim bl_Дължина As Double
        Dim bl_ИмеБлок As String
        Dim bl_Брой As Integer
        Dim bl_Visible As String
    End Structure
    Private Sub NewToolStripButton_Click(sender As Object, e As EventArgs) Handles NewToolStripButton.Click
        excel_Workbook = GetExcelWorksheet()

        Call Open_WS()

        NewToolStripButton.Enabled = False
        OpenToolStripButton.Enabled = False
        SaveToolStripButton.Enabled = True
        SplitContainer1.Enabled = vbTrue

        koefYes = vbTrue
        Call Брой_Етажи_ValueChanged(sender, e)

        Call InsertObekt()

    End Sub
    Private Function InsertObekt() As String
        Dim zapis As New Dictionary(Of String, String)
        ' Получаване на активния документ
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        ' Получаване на базата данни на активния документ
        Dim acCurDb As Database = acDoc.Database
        ' Започване на транзакция
        Using actrans As Transaction = acDoc.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = actrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            If Not acBlkTbl.Has("Insert_Signature") Then
                Return ""
            End If
            ' Получаване на ID на записа на блока "Insert_Signature" в таблицата на блоковете
            Dim blkRecId As ObjectId = acBlkTbl("Insert_Signature")
            ' Получаване на записа на блока
            Dim acBlkTblRec As BlockTableRecord = actrans.GetObject(blkRecId, OpenMode.ForRead)

            ' Обхождане на всички блокови референции за блока "Insert_Signature"
            For Each blkRefId As ObjectId In acBlkTblRec.GetBlockReferenceIds(True, True)
                ' Получаване на блоковата референция
                Dim acBlkRef As BlockReference = actrans.GetObject(blkRefId, OpenMode.ForRead)
                ' Получаване на колекцията от атрибути на блоковата референция
                Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                ' Обхождане на всички атрибути
                For Each objID As ObjectId In attCol
                    ' Получаване на атрибута
                    Dim dbObj As DBObject = actrans.GetObject(objID, OpenMode.ForRead)
                    Dim acAttRef As AttributeReference = dbObj
                    ' Проверка тагът на атрибута и промяна на текста на атрибута
                    zapis.Add(acAttRef.Tag, acAttRef.TextString)
                Next
            Next
        End Using

        Dim Obekt As String = ""
        Dim Място As String = ""
        Dim Proektant As String = ""

        zapis.TryGetValue("ОБЕКТ", Obekt) ' ако ключът липсва, Obekt остава ""
        zapis.TryGetValue("МЕСТОПОЛОЖЕНИЕ", Място) ' ако ключът липсва, Obekt остава ""
        zapis.TryGetValue("ПРОЕКТАНТ", Proektant) ' ако ключът липсва, Obekt остава ""

        ' Обработка само ако и двете са непразни
        If Obekt <> "" AndAlso Място <> "" Then
            If Not Obekt.Contains(Място) Then
                Obekt = Obekt & ", " & Място
            End If
        ElseIf Място <> "" Then
            ' Ако Обект е празен, а Място не е – използваме само Място
            Obekt = Място
        End If
    End Function
    Public Function CorrectText(Text As String) As String
        Do
            Dim originalText As String = Text

            Text = Text.Replace("..", ".")
            Text = Text.Replace("  ", " ")
            Text = Text.Replace(",,", ",")
            Text = Text.Replace(" ,", ",")
            Text = Text.Replace(",.", ".")
            Text = Text.Replace(vbCrLf, " ")
            ' Прекъсване на цикъла, ако няма повече замени
            If Text = originalText Then Exit Do
        Loop
        ' Дефиниране на регулярен израз за търсене на точки след главни букви
        Dim pattern As String = "\.\p{L}"
        Dim regex As New Regex(pattern)
        ' Намиране на съвпадения
        Dim matches As MatchCollection = regex.Matches(Text)
        For Each match As Match In matches
            Text = Text.Insert(match.Index + 1, " ") ' Добавя интервал след точката
        Next
        Return Text
    End Function
    Private Sub Open_WS()
        wsLines = excel_Workbook.Worksheets("Кабели")
        wsKontakti = excel_Workbook.Worksheets("Контакти")
        wsLight = excel_Workbook.Worksheets("Осветителни тела")
        wsElboard = excel_Workbook.Worksheets("Табла")
        wsCableTrays = excel_Workbook.Worksheets("Скари и канали")
        wsKol_Smetka = excel_Workbook.Worksheets("Количествена сметка")
        wsSpecefikaciq = excel_Workbook.Worksheets("Спецификация")
        wsInternet = excel_Workbook.Worksheets("Интернет")
        wsKoef = excel_Workbook.Worksheets("Коефициенти")
        wsPIC = excel_Workbook.Worksheets("Пожароизвестяване")
        wsDOMOF = excel_Workbook.Worksheets("Домофонна")

        wsBOLNICA = excel_Workbook.Worksheets("Болнична система")
        wwsKONTROL = excel_Workbook.Worksheets("Контрол достъп")
        wsOPOWES = excel_Workbook.Worksheets("Оповестяване")
        wsSOT = excel_Workbook.Worksheets("СОТ")
        wsVIDEO = excel_Workbook.Worksheets("Видеонаблюдение")
        wsKa4vane = excel_Workbook.Worksheets("Силова качване")

    End Sub
    Private Function GetExcelWorksheet() As excel.Workbook
        fullName = Application.DocumentManager.MdiActiveDocument.Name
        filePath = Mid(fullName, 1, InStrRev(fullName, "\"))
        fileName = Mid(fullName, InStrRev(fullName, "\") + 1, Len(fullName) - 6)

        'Създава обект на приложение на Excel и го прави видим
        Dim objExcel As excel.Application
        Dim wb As excel.Workbook
        objExcel = CreateObject("Excel.Application")
        objExcel.Visible = True
        MsgBox("Оправи EXCEL!!!")

        nameExcel = filePath & "KS__" & ".xlsx"
        'Създава нова работна книга
        wb = objExcel.Workbooks.Add
        'Възадва нови работени листове
        Dim ws1 As excel.Worksheet = CType(wb.Sheets(1), excel.Worksheet)
        Dim ws10 As excel.Worksheet = CType(wb.Sheets.Add(), excel.Worksheet)
        Dim ws11 As excel.Worksheet = CType(wb.Sheets.Add(), excel.Worksheet)
        Dim ws12 As excel.Worksheet = CType(wb.Sheets.Add(), excel.Worksheet)
        Dim ws13 As excel.Worksheet = CType(wb.Sheets.Add(), excel.Worksheet)
        Dim ws14 As excel.Worksheet = CType(wb.Sheets.Add(), excel.Worksheet)
        Dim ws15 As excel.Worksheet = CType(wb.Sheets.Add(), excel.Worksheet)
        Dim ws16 As excel.Worksheet = CType(wb.Sheets.Add(), excel.Worksheet)
        Dim ws17 As excel.Worksheet = CType(wb.Sheets.Add(), excel.Worksheet)
        Dim ws18 As excel.Worksheet = CType(wb.Sheets.Add(), excel.Worksheet)
        Dim ws9 As excel.Worksheet = CType(wb.Sheets.Add(), excel.Worksheet)
        Dim ws2 As excel.Worksheet = CType(wb.Sheets.Add(), excel.Worksheet)
        Dim ws3 As excel.Worksheet = CType(wb.Sheets.Add(), excel.Worksheet)
        Dim ws4 As excel.Worksheet = CType(wb.Sheets.Add(), excel.Worksheet)
        Dim ws5 As excel.Worksheet = CType(wb.Sheets.Add(), excel.Worksheet)
        Dim ws6 As excel.Worksheet = CType(wb.Sheets.Add(), excel.Worksheet)
        Dim ws7 As excel.Worksheet = CType(wb.Sheets.Add(), excel.Worksheet)
        Dim ws8 As excel.Worksheet = CType(wb.Sheets.Add(), excel.Worksheet)

        ws17.Name = "Силова качване"
        ws16.Name = "Болнична система"
        ws15.Name = "Контрол достъп"
        ws14.Name = "Оповестяване"
        ws13.Name = "СОТ"
        ws12.Name = "Видеонаблюдение"
        ws11.Name = "Домофонна"
        ws10.Name = "Коефициенти"
        ws9.Name = "Интернет"
        ws7.Name = "Спецификация"
        ws8.Name = "Количествена сметка"
        ws6.Name = "Кабели"
        ws5.Name = "Контакти"
        ws4.Name = "Осветителни тела"
        ws3.Name = "Табла"
        ws2.Name = "Скари и канали"
        ws1.Name = "Пожароизвестяване"
        '                    
        ' Настройва ws2 - Скари и канали
        '
        clearКбелнаскара(ws2)

        ws2.Range("D:H").Group()
        ws2.Range("J:K").Group()
        ws2.Range("M:Q").Group()

        '                    
        ' Настройва ws11 - ДОМОФОННА
        ' Болнична система
        ' Контрол достъп
        ' Оповестяване
        ' СОТ
        ' Видеонаблюдение
        '
        clearDomofon(ws11)
        clearDomofon(ws12)
        clearDomofon(ws13)
        clearDomofon(ws14)
        clearDomofon(ws15)
        clearDomofon(ws16)
        clearKa4vane(ws17)
        '                    
        ' Настройва ws10 - Коефициенти
        '
        With ws10
            With .Range("A:Z")
                .Font.Name = "Cambria"
                .Font.Size = 12
                .WrapText = vbTrue
            End With
            .Columns("A").ColumnWidth = 45
            .Columns("B").ColumnWidth = 8
            .Columns("C").ColumnWidth = 4
            .Columns("D").ColumnWidth = 33
            .Columns("E").ColumnWidth = 10
            .Columns("F").ColumnWidth = 15
            .Columns("G").ColumnWidth = 15
            .Columns("H").ColumnWidth = 15
            .Columns("I").ColumnWidth = 15

            .Columns("J:Z").ColumnWidth = 12

            .Cells(1, 1).Value = "Брой етажи"
            .Cells(2, 1).Value = "Височина на етаж"
            .Cells(3, 1).Value = "Коефициент на запаса"
            .Cells(4, 1).Value = "Кабел в розетка"
            .Cells(5, 1).Value = "Кабел в кутия"
            .Cells(6, 1).Value = "Височина на контактите"
            .Cells(7, 1).Value = "Височина на ключовете"
            .Cells(10, 1).Value = "Брой разкл. кутии"

            .Cells(12, 1).Value = "Мълниезащита Тип"
            .Cells(13, 1).Value = "Мълниезащита Категория"
            .Cells(14, 1).Value = "Мълниеприемник"
            .Cells(15, 1).Value = "Радиус"

            .Cells(1, 2).Value = 3
            .Cells(2, 2).Value = 2.6
            .Cells(3, 2).Value = 1.3
            .Cells(4, 2).Value = 0.3
            .Cells(5, 2).Value = 0.3
            .Cells(6, 2).Value = 0.5
            .Cells(7, 2).Value = 1

            .Cells(1, 4).Value = "Брой защитни"
            .Cells(2, 4).Value = "Брой мълния"
            .Cells(3, 4).Value = "Брой ДЗТ"
            .Cells(4, 4).Value = "Брой контакти"
            .Cells(5, 4).Value = "Брой кабели"
            .Cells(6, 4).Value = "Брой съоръжение до 2,5мм²"
            .Cells(7, 4).Value = "Брой съоръжение до 16мм²"
            .Cells(1, 6).Value = "Tабла"
            .Cells(1, 7).Value = "В ниша"
            .Cells(1, 8).Value = "Окачено"
            .Cells(1, 9).Value = "Фундамент"


            .Cells(red_Фотоволтаици + 0, 1).Value = "Брой панели"
            .Cells(red_Фотоволтаици + 1, 1).Value = "Тип панели"
            .Cells(red_Фотоволтаици + 2, 1).Value = "Брой инвертори"
            .Cells(red_Фотоволтаици + 3, 1).Value = "Тип инвертори"
            .Cells(red_Фотоволтаици + 4, 1).Value = "Брой табло DC"
            .Cells(red_Фотоволтаици + 5, 1).Value = "Брой външна единична"
            .Cells(red_Фотоволтаици + 6, 1).Value = "Брой вътрешна двойна"
            .Cells(red_Фотоволтаици + 7, 1).Value = "Брой конектор МС"

        End With
        '                    
        ' Настройва ws9 - Интернет
        '
        clearInternet(ws9)
        '                    
        ' Настройва ws7 - Спецификация
        '
        With ws7
            With .PageSetup
                .LeftMargin = objExcel.InchesToPoints(0.984251968503937)
                .RightMargin = objExcel.InchesToPoints(0.196850393700787)
                .TopMargin = objExcel.InchesToPoints(0.590551181102362)
                .BottomMargin = objExcel.InchesToPoints(0.393700787401575)
                .HeaderMargin = objExcel.InchesToPoints(0.31496062992126)
                .FooterMargin = objExcel.InchesToPoints(0.31496062992126)
                .PrintTitleRows = "$4:$4"
                .RightFooter = "Стр.&P от &N"
            End With
            With .Range("A:D")
                .Font.Name = "Cambria"
                .Font.Size = 12
                .WrapText = vbTrue
                .RowHeight = 15.75
                .VerticalAlignment = excel.XlVAlign.xlVAlignCenter
                .HorizontalAlignment = excel.XlHAlign.xlHAlignCenter
                .Font.Bold = vbFalse
                .Font.Size = 12
                .WrapText = vbTrue
                .Font.ColorIndex = excel.XlColorIndex.xlColorIndexAutomatic
                .Borders(excel.XlBordersIndex.xlDiagonalDown).LineStyle = excel.XlLineStyle.xlLineStyleNone
                .Borders(excel.XlBordersIndex.xlDiagonalUp).LineStyle = excel.XlLineStyle.xlLineStyleNone
            End With
            .Columns("A:A").ColumnWidth = 6
            .Columns("B:B").ColumnWidth = 64
            .Columns("C:C").ColumnWidth = 8
            .Columns("D:D").ColumnWidth = 9
            With .Cells(1, 1)
                .Value = "СПЕЦИФИКАЦИЯ НА МАТЕРИАЛИТЕ"
                .RowHeight = 31
                .VerticalAlignment = excel.XlVAlign.xlVAlignCenter
                .HorizontalAlignment = excel.XlHAlign.xlHAlignCenter
                .Font.Bold = vbTrue
                .Font.Size = 16
            End With

            With .Cells(2, 1)
                .Value = "ОБЕКТ: " + InsertObekt()
                .RowHeight = 42
                .VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Font.Bold = vbTrue
                .Font.Size = 14
            End With

            .Cells(4, 1).Value = "№ по ред"
            .Cells(4, 2).Value = "Описание на материала"
            .Cells(4, 3).Value = "Ед. мярка"
            .Cells(4, 4).Value = "Кол-во"
            .Cells(5, 2).Value = "ЧАСТ ЕЛЕКТРО"

            With .Range("A4:D4")
                .RowHeight = 31
                .VerticalAlignment = excel.XlVAlign.xlVAlignCenter
                .HorizontalAlignment = excel.XlHAlign.xlHAlignCenter
                With .Interior
                    .Pattern = excel.XlPattern.xlPatternSolid
                    .PatternColorIndex = 24
                    .ThemeColor = excel.XlThemeColor.xlThemeColorDark1
                    .TintAndShade = -0.2
                    .PatternTintAndShade = 0
                End With
                .Font.Bold = vbTrue
                .Font.Size = 12
                .WrapText = vbTrue
                .Font.ColorIndex = excel.XlColorIndex.xlColorIndexAutomatic
                .Borders(excel.XlBordersIndex.xlDiagonalDown).LineStyle = excel.XlLineStyle.xlLineStyleNone
                .Borders(excel.XlBordersIndex.xlDiagonalUp).LineStyle = excel.XlLineStyle.xlLineStyleNone
                .Borders(excel.XlBordersIndex.xlInsideHorizontal).LineStyle = excel.XlLineStyle.xlLineStyleNone
                With .Borders(excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = excel.XlLineStyle.xlContinuous
                    .ColorIndex = excel.XlColorIndex.xlColorIndexAutomatic
                    .TintAndShade = 0
                    .Weight = excel.XlBorderWeight.xlMedium
                End With
                With .Borders(excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = excel.XlLineStyle.xlContinuous
                    .ColorIndex = excel.XlColorIndex.xlColorIndexAutomatic
                    .TintAndShade = 0
                    .Weight = excel.XlBorderWeight.xlMedium
                End With
                With .Borders(excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = excel.XlLineStyle.xlContinuous
                    .ColorIndex = excel.XlColorIndex.xlColorIndexAutomatic
                    .TintAndShade = 0
                    .Weight = excel.XlBorderWeight.xlMedium
                End With
                With .Borders(excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = excel.XlLineStyle.xlContinuous
                    .ColorIndex = excel.XlColorIndex.xlColorIndexAutomatic
                    .TintAndShade = 0
                    .Weight = excel.XlBorderWeight.xlMedium
                End With
                With .Borders(excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = excel.XlLineStyle.xlContinuous
                    .ColorIndex = excel.XlColorIndex.xlColorIndexAutomatic
                    .TintAndShade = 0
                    .Weight = excel.XlBorderWeight.xlThin
                End With
            End With
            With .Range("A5:D5")
                .RowHeight = 18.75
                .VerticalAlignment = excel.XlVAlign.xlVAlignCenter
                .HorizontalAlignment = excel.XlHAlign.xlHAlignCenter
                '
                ' Статия за цветовете в Ексел
                ' https://renenyffenegger.ch/notes/Microsoft/Office/Excel/Object-Model/_colors/index
                '
                With .Interior
                    .Pattern = excel.XlPattern.xlPatternSolid
                    .PatternColorIndex = 24
                    .ThemeColor = excel.XlThemeColor.xlThemeColorAccent4
                    .TintAndShade = 0.4
                    .PatternTintAndShade = 0
                End With
                .Font.Bold = vbTrue
                .Font.Size = 14
                .WrapText = vbTrue
                .Font.ColorIndex = excel.XlColorIndex.xlColorIndexAutomatic
                .Borders(excel.XlBordersIndex.xlDiagonalDown).LineStyle = excel.XlLineStyle.xlLineStyleNone
                .Borders(excel.XlBordersIndex.xlDiagonalUp).LineStyle = excel.XlLineStyle.xlLineStyleNone
                .Borders(excel.XlBordersIndex.xlInsideHorizontal).LineStyle = excel.XlLineStyle.xlLineStyleNone
                .Borders(excel.XlBordersIndex.xlInsideVertical).LineStyle = excel.XlLineStyle.xlLineStyleNone
                With .Borders(excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = excel.XlLineStyle.xlContinuous
                    .ColorIndex = excel.XlColorIndex.xlColorIndexAutomatic
                    .TintAndShade = 0
                    .Weight = excel.XlBorderWeight.xlMedium
                End With
                With .Borders(excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = excel.XlLineStyle.xlContinuous
                    .ColorIndex = excel.XlColorIndex.xlColorIndexAutomatic
                    .TintAndShade = 0
                    .Weight = excel.XlBorderWeight.xlMedium
                End With
                With .Borders(excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = excel.XlLineStyle.xlContinuous
                    .ColorIndex = excel.XlColorIndex.xlColorIndexAutomatic
                    .TintAndShade = 0
                    .Weight = excel.XlBorderWeight.xlMedium
                End With
                With .Borders(excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = excel.XlLineStyle.xlContinuous
                    .ColorIndex = excel.XlColorIndex.xlColorIndexAutomatic
                    .TintAndShade = 0
                    .Weight = excel.XlBorderWeight.xlMedium
                End With
            End With
            '
            ' Клеки за разрел
            '
            Call Excel_Kol_smetka_Razdel(ws7, "Раздел", "A6", "D6")
            .Range("A1:D1").HorizontalAlignment = excel.XlHAlign.xlHAlignCenterAcrossSelection
            .Range("A2:D2").HorizontalAlignment = excel.XlHAlign.xlHAlignCenterAcrossSelection
            .Range("A5:D5").HorizontalAlignment = excel.XlHAlign.xlHAlignCenterAcrossSelection
        End With
        '                    
        ' Настройва ws8 - Количествена сметка
        '
        With ws8
            With .PageSetup
                .LeftMargin = objExcel.InchesToPoints(0.984251968503937)
                .RightMargin = objExcel.InchesToPoints(0.196850393700787)
                .TopMargin = objExcel.InchesToPoints(0.590551181102362)
                .BottomMargin = objExcel.InchesToPoints(0.590551181102362)
                .HeaderMargin = objExcel.InchesToPoints(0.31496062992126)
                .FooterMargin = objExcel.InchesToPoints(0.31496062992126)
                .PrintTitleRows = "$4:$4"
                .RightFooter = "Стр.&P от &N"
            End With

            With .Range("A:D")
                .Font.Name = "Cambria"
                .Font.Size = 12
                .WrapText = vbTrue
                .RowHeight = 15.75
                .VerticalAlignment = excel.XlVAlign.xlVAlignCenter
                .HorizontalAlignment = excel.XlHAlign.xlHAlignCenter
                .Font.Bold = vbFalse
                .Font.Size = 12
                .WrapText = vbTrue
                .Font.ColorIndex = excel.XlColorIndex.xlColorIndexAutomatic
                .Borders(excel.XlBordersIndex.xlDiagonalDown).LineStyle = excel.XlLineStyle.xlLineStyleNone
                .Borders(excel.XlBordersIndex.xlDiagonalUp).LineStyle = excel.XlLineStyle.xlLineStyleNone
            End With
            .Columns("A:A").ColumnWidth = 6
            .Columns("B:B").ColumnWidth = 67
            .Columns("C:C").ColumnWidth = 7
            .Columns("D:D").ColumnWidth = 7
            With .Cells(1, 1)
                .Value = "КОЛИЧЕСТВЕНА СМЕТКА"
                .RowHeight = 31
                .VerticalAlignment = excel.XlVAlign.xlVAlignCenter
                .HorizontalAlignment = excel.XlHAlign.xlHAlignCenter
                .Font.Bold = vbTrue
                .Font.Size = 16
            End With

            With .Cells(2, 1)
                .Value = "ОБЕКТ:" + InsertObekt()
                .RowHeight = 42
                .VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Font.Bold = vbTrue
                .Font.Size = 14
            End With

            .Cells(4, 1).Value = "№ по ред"
            .Cells(4, 2).Value = "Описание на строително-монтажни и демонтажни работи"
            .Cells(4, 3).Value = "Ед. мярка"
            .Cells(4, 4).Value = "Кол-во"
            .Cells(5, 2).Value = "ЧАСТ ЕЛЕКТРО"

            With .Range("A4:D4")
                .RowHeight = 31
                .VerticalAlignment = excel.XlVAlign.xlVAlignCenter
                .HorizontalAlignment = excel.XlHAlign.xlHAlignCenter
                With .Interior
                    .Pattern = excel.XlPattern.xlPatternSolid
                    .PatternColorIndex = 24
                    .ThemeColor = excel.XlThemeColor.xlThemeColorDark1
                    .TintAndShade = -0.2
                    .PatternTintAndShade = 0
                End With
                .Font.Bold = vbTrue
                .Font.Size = 12
                .WrapText = vbTrue
                .Font.ColorIndex = excel.XlColorIndex.xlColorIndexAutomatic
                .Borders(excel.XlBordersIndex.xlDiagonalDown).LineStyle = excel.XlLineStyle.xlLineStyleNone
                .Borders(excel.XlBordersIndex.xlDiagonalUp).LineStyle = excel.XlLineStyle.xlLineStyleNone
                .Borders(excel.XlBordersIndex.xlInsideHorizontal).LineStyle = excel.XlLineStyle.xlLineStyleNone
                With .Borders(excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = excel.XlLineStyle.xlContinuous
                    .ColorIndex = excel.XlColorIndex.xlColorIndexAutomatic
                    .TintAndShade = 0
                    .Weight = excel.XlBorderWeight.xlMedium
                End With
                With .Borders(excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = excel.XlLineStyle.xlContinuous
                    .ColorIndex = excel.XlColorIndex.xlColorIndexAutomatic
                    .TintAndShade = 0
                    .Weight = excel.XlBorderWeight.xlMedium
                End With
                With .Borders(excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = excel.XlLineStyle.xlContinuous
                    .ColorIndex = excel.XlColorIndex.xlColorIndexAutomatic
                    .TintAndShade = 0
                    .Weight = excel.XlBorderWeight.xlMedium
                End With
                With .Borders(excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = excel.XlLineStyle.xlContinuous
                    .ColorIndex = excel.XlColorIndex.xlColorIndexAutomatic
                    .TintAndShade = 0
                    .Weight = excel.XlBorderWeight.xlMedium
                End With
                With .Borders(excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = excel.XlLineStyle.xlContinuous
                    .ColorIndex = excel.XlColorIndex.xlColorIndexAutomatic
                    .TintAndShade = 0
                    .Weight = excel.XlBorderWeight.xlThin
                End With
            End With
            Call Excel_Kol_smetka_Razdel(ws8, "ЧАСТ ЕЛЕКТРО", "A5", "D5")
            With .Range("A5:D5").Interior
                .Pattern = excel.XlPattern.xlPatternSolid
                .PatternColorIndex = 24
                .ThemeColor = excel.XlThemeColor.xlThemeColorAccent4
                .TintAndShade = 0.4
                .PatternTintAndShade = 0
            End With
            '
            ' Клеки за разрел
            '
            Call Excel_Kol_smetka_Razdel(ws8, "Раздел", "A6", "D6")
            .Range("A1:D1").HorizontalAlignment = excel.XlHAlign.xlHAlignCenterAcrossSelection
            .Range("A2:D2").HorizontalAlignment = excel.XlHAlign.xlHAlignCenterAcrossSelection
            .Range("A5:D5").HorizontalAlignment = excel.XlHAlign.xlHAlignCenterAcrossSelection
        End With
        '                    
        'Настройва ws6 - Кабели
        '  
        clearKabeli(ws6, 3, 1000)
        '                    
        ' Настройва ws5 - Контакти
        '  
        clearKонтакти(ws5, "GetExcel")
        '                    
        ' Настройва ws4 - Осветителни тела
        '  
        With ws4
            With .Range("A:Z")
                .Font.Name = "Cambria"
                .Font.Size = 12
                .WrapText = vbTrue
            End With
            .Columns("A").ColumnWidth = 17
            .Columns("B").ColumnWidth = 6
            .Columns("C:Z").ColumnWidth = 12
            .Cells(1, 1).Value = "Брой"
            .Cells(1, 2).Value = "Име"
            .Cells(1, 3).Value = "Visibility"
            .Cells(1, 4).Value = "Светлинен_поток"
            .Cells(1, 5).Value = "IP"
            .Cells(1, 6).Value = "Монтаж"
            .Cells(1, 7).Value = "Мощност"
            .Cells(1, 8).Value = "Табло"
            .Cells(1, 9).Value = "Токов кръг"
            .Cells(1, 10).Value = "Обща мощност"
        End With
        '                    
        ' Настройва ws3 - Табла
        '  
        clearТабла(ws3)
        '                    
        ' Настройва ws1 - ПИЦ
        '
        clearPIC(ws1)
        '
        '
        '
        wb.SaveAs(nameExcel)
        objExcel.Visible = vbTrue
        Return wb
    End Function
    Private Sub Excel_Kol_smetka_Razdel(ws As excel.Worksheet, Text As String, Range_A As String, Range_B As String)
        Dim cells_Range As String = Range_A & ":" & Range_B
        Dim Red As Integer
        Red = Val(Mid(Range_B, 2, Len(Range_B)))
        Dim RowHeight As Double = 18.75
        Dim FontSize As Integer = 14
        If Mid(Range_A, 1, 1) = "B" Then
            RowHeight = IIf(Len(ws.Range(Range_A).Value) > 67, 2 * 15.75, 15.75)
            FontSize = 12
        End If
        With ws
            With .Range(cells_Range)
                .RowHeight = RowHeight
                .VerticalAlignment = excel.XlVAlign.xlVAlignCenter
                .HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                With .Interior
                    .Pattern = excel.XlPattern.xlPatternSolid
                    .PatternColorIndex = 24
                    .ThemeColor = excel.XlThemeColor.xlThemeColorAccent6
                    .TintAndShade = 0.7
                    .PatternTintAndShade = 0
                End With
                .Font.Bold = vbTrue
                .Font.Size = FontSize
                .WrapText = vbTrue
                .Font.ColorIndex = excel.XlColorIndex.xlColorIndexAutomatic
                .Borders(excel.XlBordersIndex.xlDiagonalDown).LineStyle = excel.XlLineStyle.xlLineStyleNone
                .Borders(excel.XlBordersIndex.xlDiagonalUp).LineStyle = excel.XlLineStyle.xlLineStyleNone
                .Borders(excel.XlBordersIndex.xlInsideHorizontal).LineStyle = excel.XlLineStyle.xlLineStyleNone
                .Borders(excel.XlBordersIndex.xlInsideVertical).LineStyle = excel.XlLineStyle.xlLineStyleNone
                With .Borders(excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = excel.XlLineStyle.xlContinuous
                    .ColorIndex = excel.XlColorIndex.xlColorIndexAutomatic
                    .TintAndShade = 0
                    .Weight = excel.XlBorderWeight.xlMedium
                End With
                With .Borders(excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = excel.XlLineStyle.xlContinuous
                    .ColorIndex = excel.XlColorIndex.xlColorIndexAutomatic
                    .TintAndShade = 0
                    .Weight = excel.XlBorderWeight.xlMedium
                End With
                With .Borders(excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = excel.XlLineStyle.xlContinuous
                    .ColorIndex = excel.XlColorIndex.xlColorIndexAutomatic
                    .TintAndShade = 0
                    .Weight = excel.XlBorderWeight.xlMedium
                End With
                With .Borders(excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = excel.XlLineStyle.xlContinuous
                    .ColorIndex = excel.XlColorIndex.xlColorIndexAutomatic
                    .TintAndShade = 0
                    .Weight = excel.XlBorderWeight.xlMedium
                End With
            End With
            .Cells(Red, 2).Value = Text
        End With
    End Sub
    Private Sub SaveToolStripButton_Click(sender As Object, e As EventArgs) Handles SaveToolStripButton.Click,
                                                                                    SaveToolStripMenuItem.Click
        If IsNothing(excel_Workbook) Then
            Exit Sub
        End If
        excel_Workbook.Save()
    End Sub
    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub
    Private Sub Form1_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        If IsNothing(excel_Workbook) Then
            Exit Sub
        End If
        Try
            excel_Workbook.Save()
            excel_Workbook.Close()
            excel_Workbook = Nothing
            'excel_Workbook.Quit()
        Catch ex As Exception
            MsgBox("Файла вече е затворен")
        End Try
    End Sub
    Private Sub OpenToolStripButton_Click(sender As Object, e As EventArgs) Handles OpenToolStripButton.Click,
                                                                                    OpenToolStripMenuItem.Click

        fullName = Application.DocumentManager.MdiActiveDocument.Name
        filePath = Mid(fullName, 1, InStrRev(fullName, "\"))
        fileName = Mid(fullName, InStrRev(fullName, "\") + 1, Len(fullName) - 6)

        OpenFileDialog1.InitialDirectory = filePath
        OpenFileDialog1.Filter = "Excel files (*.xls or *.xlsx)|*.xls;*.xlsx"
        OpenFileDialog1.FileName = "KS__.xlsx"
        If OpenFileDialog1.ShowDialog <> Windows.Forms.DialogResult.OK Then
            Exit Sub
        End If
        nameExcel = OpenFileDialog1.FileName
        '
        ' Проверява дали EXCEL е отворен
        '
        Dim stream As FileStream = Nothing
        Try
            stream = File.Open(nameExcel, FileMode.Open, FileAccess.ReadWrite, FileShare.None)
            stream.Close()
        Catch ex As Exception
            MsgBox("Отворен е файл с име : " + Chr(13) + Chr(13) +
                   nameExcel + Chr(13) + Chr(13) +
                   "Моля затворете го преди да продължите!")
            Exit Sub
        End Try
        '
        'Get all currently running process Ids for Excel applications
        '
        NoMyExcelProcesses = Process.GetProcessesByName("Excel")

        Dim objExcel As excel.Application = New excel.Application()
        excel_Workbook = objExcel.Workbooks.Open(nameExcel)

        Call Open_WS()

        objExcel.Visible = True
        MsgBox("Оправи EXCEL!!!")
        NewToolStripButton.Enabled = False
        OpenToolStripButton.Enabled = False
        SaveToolStripButton.Enabled = True
        SplitContainer1.Enabled = True
        ProgressBar_Extrat.Minimum = 0

        Бр_Етажи = wsKoef.Cells(1, 2).Value
        H_Етаж = wsKoef.Cells(2, 2).Value
        Kz = wsKoef.Cells(3, 2).Value
        Кабел_Розетка = wsKoef.Cells(4, 2).Value
        Кабел_кутия = wsKoef.Cells(5, 2).Value
        H_Контакт = wsKoef.Cells(6, 2).Value
        H_Ключ = wsKoef.Cells(7, 2).Value

        Брой_Етажи.Value = Бр_Етажи
        Височина_Етажи.Value = H_Етаж
        Koef_Zapas.Value = Kz
        Kabel_Rozetka.Value = Кабел_Розетка
        Kabel_kutiq.Value = Кабел_кутия
        Wiso4ina_Kontakti.Value = H_Контакт
        Wiso4ina_Kl.Value = H_Ключ

        koefYes = vbTrue ' МОЖЕ ДА МИНАВА ПРЕЗ ПРОЦЕДУРАТА СЛЕД КАТО ПРОЧЕТЕ ВСИЧКИ КОЕФИЦИЕНТИ !!!!!
    End Sub
    Private Sub Button_Erase_Click(sender As Object, e As EventArgs) Handles Button_Erase.Click
        With wsKol_Smetka
            With .Range("A6:D1000")
                .Clear()
                .Value = ""
                .Font.Name = "Cambria"
                .Font.Bold = vbFalse
                .Font.Size = 12
                .RowHeight = 15.75
                .WrapText = vbTrue
                .Font.ColorIndex = excel.XlColorIndex.xlColorIndexAutomatic
                .Borders(excel.XlBordersIndex.xlDiagonalDown).LineStyle = excel.XlLineStyle.xlLineStyleNone
                .Borders(excel.XlBordersIndex.xlDiagonalUp).LineStyle = excel.XlLineStyle.xlLineStyleNone
            End With
        End With
    End Sub
    Private Sub Button_Контакти_Вземи_Click(sender As Object, e As EventArgs) Handles Button_Вземи_Контакти.Click
        If IsNothing(excel_Workbook) Then
            MsgBox("Първо да беше отворил файла?")
            Exit Sub
        End If
        Me.Visible = vbFalse
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim edt As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database
        Dim SelectedSet = cu.GetObjects("INSERT", "Изберете блок")
        Dim arrBlock(500) As strKontakt
        If SelectedSet Is Nothing Then
            MsgBox("Няма маркиран нито един блок.")
            Me.Visible = vbTrue
            Exit Sub
        End If
        Dim blkRecId As ObjectId = ObjectId.Null
        Dim index As Integer = 0
        Dim index_Red As Integer = 2
        clearKa4vane(wsKa4vane)
        Me.Visible = vbTrue
        ProgressBar_Extrat.Maximum = SelectedSet.Count
        ProgressBar_Extrat.Value = ProgressBar_Extrat.Minimum
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                For Each sObj As SelectedObject In SelectedSet

                    ProgressBar_Extrat.Value += 1

                    blkRecId = sObj.ObjectId

                    Dim acBlkRef As BlockReference = DirectCast(acTrans.GetObject(blkRecId, OpenMode.ForRead), BlockReference)
                    Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                    Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection
                    Dim Visibility As String = ""

                    For Each prop As DynamicBlockReferenceProperty In props
                        If prop.PropertyName = "Visibility1" Then Visibility = prop.Value
                        If prop.PropertyName = "Visibility" Then Visibility = prop.Value
                        If prop.PropertyName = "Тип" Then Visibility = prop.Value
                    Next
                    Dim blName As String = (CType(acBlkRef.DynamicBlockTableRecord.GetObject(OpenMode.ForRead), BlockTableRecord)).Name
                    Dim iVisib As Integer = -1
                    Dim strWis As String = ""
                    Dim strLamp_Power As String = ""
                    Dim strLED_Lamp_Montav As String = ""
                    Dim strLED_Lamp_SW_Potok As String = ""
                    Dim strLED_Lamp_TIP As String = ""
                    Select Case blName
                        Case "Контакт"
                            For Each objID As ObjectId In attCol
                                Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                                Dim acAttRef As AttributeReference = dbObj
                                If acAttRef.Tag = "ВИС" Then strWis = acAttRef.TextString
                            Next
                            iVisib = Array.FindIndex(arrBlock, Function(f) f.blName = blName And
                                                                           f.blVisibility = Visibility 'And f.blWis = strWis
                                                                           )
                        Case "Авария", "Авария_100"
                            iVisib = Array.FindIndex(arrBlock, Function(f) f.blName = blName)
                        Case "Ключ_знак", "Ключ_знак_WIFI", "Ключ_квадрат",
                            "бойлерно табло", "Розетка_1", "Датчик_ПАБ", "Домофон",
                            "Високоговорител", "Аудио система", "Камери", "СОТ"
                            iVisib = Array.FindIndex(arrBlock, Function(f) f.blName = blName And
                                                                           f.blVisibility = Visibility)
                        Case "Бойлер", "Вентилации"
                            For Each objID As ObjectId In attCol
                                Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                                Dim acAttRef As AttributeReference = dbObj
                                If acAttRef.Tag = "МОЩНОСТ" Then strWis = acAttRef.TextString
                            Next
                            iVisib = Array.FindIndex(arrBlock, Function(f) f.blName = blName And
                                                                           f.blVisibility = Visibility And
                                                                           f.blLamp_Power = strWis)
                        Case "LED_луна"
                            If Visibility = "Драйвер" Then
                                For Each objID As ObjectId In attCol
                                    Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                                    Dim acAttRef As AttributeReference = dbObj
                                    If acAttRef.Tag = "MO6T_TRAFO" Then strWis = acAttRef.TextString
                                Next
                                iVisib = Array.FindIndex(arrBlock, Function(f) f.blName = blName And
                                                                           f.blVisibility = Visibility And
                                                                           f.blMO6T_TRAFO = strWis
                                                                           )
                            Else
                                For Each objID As ObjectId In attCol
                                    Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                                    Dim acAttRef As AttributeReference = dbObj
                                    If acAttRef.Tag = "МОЩНОСТ" Then strWis = acAttRef.TextString
                                Next
                                iVisib = Array.FindIndex(arrBlock, Function(f) f.blName = blName And
                                                                           f.blVisibility = Visibility And
                                                                           f.blLamp_Power = strWis
                                                                           )
                            End If
                        Case "Линия МХЛ - 220V", "Плафони", "Металхаогенна лампа", "Прожектор", "Полилей"
                            For Each objID As ObjectId In attCol
                                Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                                Dim acAttRef As AttributeReference = dbObj
                                If acAttRef.Tag = "МОЩНОСТ" Then strWis = acAttRef.TextString
                            Next
                            iVisib = Array.FindIndex(arrBlock, Function(f) f.blName = blName And
                                                                           f.blVisibility = Visibility And
                                                                           f.blLamp_Power = strWis
                                                                           )

                        Case "LED_lenta"
                            For Each prop As DynamicBlockReferenceProperty In props
                                If prop.PropertyName = "Дължина" Then strWis = prop.Value
                            Next
                            iVisib = Array.FindIndex(arrBlock, Function(f) f.blName = blName)
                        Case "LED_ULTRALUX", "LED_ULTRALUX_100"
                            For Each objID As ObjectId In attCol
                                Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                                Dim acAttRef As AttributeReference = dbObj
                                If acAttRef.Tag = "LED" Then strLamp_Power = acAttRef.TextString
                            Next
                            For Each prop As DynamicBlockReferenceProperty In props
                                If prop.PropertyName = "Монтаж" Then strLED_Lamp_Montav = prop.Value
                                If prop.PropertyName = "Св_поток" Then strLED_Lamp_SW_Potok = prop.Value
                                If prop.PropertyName = "Тип" Then strLED_Lamp_TIP = prop.Value
                                If prop.PropertyName = "Distance1" Then strWis = prop.Value

                            Next
                            iVisib = Array.FindIndex(arrBlock, Function(f) f.blName = blName And
                                                                           f.blLamp_Power = strLamp_Power And
                                                                           f.blLED_Lamp_Montav = strLED_Lamp_Montav And
                                                                           f.blLED_Lamp_SW_Potok = strLED_Lamp_SW_Potok And
                                                                           f.blLED_Lamp_TIP = strLED_Lamp_TIP
                                                                           )
                            Visibility = ""
                        Case "LED_DENIMA"
                            For Each objID As ObjectId In attCol
                                Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                                Dim acAttRef As AttributeReference = dbObj
                                If acAttRef.Tag = "LED" Then strLamp_Power = acAttRef.TextString
                            Next
                            For Each prop As DynamicBlockReferenceProperty In props
                                If prop.PropertyName = "Лампа" Then strLED_Lamp_Montav = prop.Value
                                If prop.PropertyName = "Светлинен_поток" Then strLED_Lamp_SW_Potok = prop.Value
                            Next
                            iVisib = Array.FindIndex(arrBlock, Function(f) f.blName = blName And
                                                                           f.blLamp_Power = strLamp_Power And
                                                                           f.blLED_Lamp_Montav = strLED_Lamp_Montav And
                                                                           f.blLED_Lamp_SW_Potok = strLED_Lamp_SW_Potok
                                                                           )

                            Visibility = ""
                        Case "Луминисцентна лампа"
                        Case "Качване"
                            Dim ka4vane As strКачване
                            For Each objID As ObjectId In attCol
                                Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                                Dim acAttRef As AttributeReference = dbObj
                                If acAttRef.Tag = "KOTA_1" Then ka4vane.KOTA_1 = acAttRef.TextString
                                If acAttRef.Tag = "KOTA_2" Then ka4vane.KOTA_2 = acAttRef.TextString
                                If acAttRef.Tag = "ТРЪБА_1" Then ka4vane.ТРЪБА_1 = acAttRef.TextString
                                If acAttRef.Tag = "ТРЪБА_2" Then ka4vane.ТРЪБА_2 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_0" Then ka4vane.Kabel_d_0 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_1" Then ka4vane.Kabel_d_1 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_2" Then ka4vane.Kabel_d_2 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_3" Then ka4vane.Kabel_d_3 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_4" Then ka4vane.Kabel_d_4 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_5" Then ka4vane.Kabel_d_5 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_6" Then ka4vane.Kabel_d_6 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_7" Then ka4vane.Kabel_d_7 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_8" Then ka4vane.Kabel_d_8 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_9" Then ka4vane.Kabel_d_9 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_10" Then ka4vane.Kabel_d_10 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_0" Then ka4vane.Kabel_g_0 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_1" Then ka4vane.Kabel_g_1 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_2" Then ka4vane.Kabel_g_2 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_3" Then ka4vane.Kabel_g_3 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_4" Then ka4vane.Kabel_g_4 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_5" Then ka4vane.Kabel_g_5 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_6" Then ka4vane.Kabel_g_6 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_7" Then ka4vane.Kabel_g_7 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_8" Then ka4vane.Kabel_g_8 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_9" Then ka4vane.Kabel_g_9 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_10" Then ka4vane.Kabel_g_10 = acAttRef.TextString
                            Next
                            With wsKa4vane
                                .Cells(index_Red, 1) = ka4vane.KOTA_1
                                .Cells(index_Red, 2) = ka4vane.ТРЪБА_1
                                .Cells(index_Red, 3) = ka4vane.Kabel_d_0
                                .Cells(index_Red, 4) = ka4vane.Kabel_d_1
                                .Cells(index_Red, 5) = ka4vane.Kabel_d_2
                                .Cells(index_Red, 6) = ka4vane.Kabel_d_3
                                .Cells(index_Red, 7) = ka4vane.Kabel_d_6
                                .Cells(index_Red, 8) = ka4vane.Kabel_d_5
                                .Cells(index_Red, 9) = ka4vane.Kabel_d_8
                                .Cells(index_Red, 10) = ka4vane.Kabel_d_7
                                .Cells(index_Red, 11) = ka4vane.Kabel_d_4
                                .Cells(index_Red, 12) = ka4vane.Kabel_d_9
                                .Cells(index_Red, 13) = ka4vane.Kabel_d_10
                                .Cells(index_Red, 14) = ka4vane.KOTA_2
                                .Cells(index_Red, 15) = ka4vane.ТРЪБА_2
                                .Cells(index_Red, 16) = ka4vane.Kabel_g_0
                                .Cells(index_Red, 17) = ka4vane.Kabel_g_1
                                .Cells(index_Red, 18) = ka4vane.Kabel_g_2
                                .Cells(index_Red, 19) = ka4vane.Kabel_g_3
                                .Cells(index_Red, 20) = ka4vane.Kabel_g_4
                                .Cells(index_Red, 21) = ka4vane.Kabel_g_5
                                .Cells(index_Red, 22) = ka4vane.Kabel_g_6
                                .Cells(index_Red, 23) = ka4vane.Kabel_g_7
                                .Cells(index_Red, 24) = ka4vane.Kabel_g_8
                                .Cells(index_Red, 25) = ka4vane.Kabel_g_9
                                .Cells(index_Red, 26) = ka4vane.Kabel_g_10
                                index_Red += 1
                            End With
                            Continue For
                        Case Else
                            Continue For
                    End Select
                    If iVisib = -1 Then
                        arrBlock(index).blVisibility = Visibility
                        arrBlock(index).blName = blName
                        arrBlock(index).count = 1
                        Select Case blName
                            Case "Контакт"
                                arrBlock(index).blWis = strWis
                            Case "LED_луна"
                                If Visibility = "Драйвер" Then
                                    arrBlock(index).blMO6T_TRAFO = strWis
                                Else
                                    arrBlock(index).blLamp_Power = strWis
                                End If
                            Case "Линия МХЛ - 220V", "Плафони", "Металхаогенна лампа", "Прожектор", "Полилей"
                                arrBlock(index).blLamp_Power = strWis
                            Case "LED_lenta"
                                arrBlock(index).blLED_Lenta_Length = strWis
                            Case "LED_ULTRALUX", "LED_ULTRALUX_100"
                                arrBlock(index).blLamp_Power = strLamp_Power
                                arrBlock(index).blLED_Lamp_SW_Potok = strLED_Lamp_SW_Potok
                                arrBlock(index).blLED_Lamp_Montav = strLED_Lamp_Montav
                                arrBlock(index).blLED_Lamp_TIP = strLED_Lamp_TIP
                                arrBlock(index).blWis = strWis
                            Case "LED_DENIMA"
                                arrBlock(index).blLamp_Power = strLamp_Power
                                arrBlock(index).blLED_Lamp_SW_Potok = strLED_Lamp_SW_Potok
                                arrBlock(index).blLED_Lamp_Montav = strLED_Lamp_Montav
                            Case "Бойлер", "Вентилации"
                                arrBlock(index).blLamp_Power = strWis
                            Case "Луминисцентна лампа"

                        End Select
                        index += 1
                    Else
                        arrBlock(iVisib).count = arrBlock(iVisib).count + 1
                        If blName = "LED_lenta" Then
                            arrBlock(iVisib).blLED_Lenta_Length = Str(Val(arrBlock(iVisib).blLED_Lenta_Length) + Val(strWis))
                        End If
                    End If
                Next
                Call Button_Изчисти_контакти_Click(sender, e)
                index = 2
                ProgressBar_Extrat.Value = ProgressBar_Extrat.Minimum
                For Each iarrBlock In arrBlock
                    If iarrBlock.count = 0 Then Exit For
                    ProgressBar_Extrat.Value += 1
                    With wsKontakti
                        .Cells(index, 1) = iarrBlock.count
                        .Cells(index, 2) = iarrBlock.blName
                        .Cells(index, 3) = iarrBlock.blWis
                        .Cells(index, 4) = iarrBlock.blVisibility
                        .Cells(index, 5) = iarrBlock.blMO6T_TRAFO
                        .Cells(index, 6) = iarrBlock.blLamp_Power
                        .Cells(index, 7) = iarrBlock.blLED_Lenta_Length
                        .Cells(index, 8) = iarrBlock.blLED_Lamp_SW_Potok
                        .Cells(index, 9) = iarrBlock.blLED_Lamp_TIP
                        .Cells(index, 10) = iarrBlock.blLED_Lamp_Montav
                    End With
                    ' #################################################################################################################################
                    Dim broj_elementi As Integer = iarrBlock.count
                    Select Case iarrBlock.blName
                        Case "LED_lenta", "LED_ULTRALUX", "LED_ULTRALUX_100", "LED_луна",
                             "Линия МХЛ - 220V", "Плафони", "Прожектор"
                            wsLines.Cells(2, 12).Value = wsLines.Cells(2, 12).Value + iarrBlock.count
                        Case "Авария_100", "Авария"
                            wsLines.Cells(2, 14).Value = wsLines.Cells(2, 14).Value + iarrBlock.count
                        Case "Ключ_знак"
                            Select Case iarrBlock.blVisibility
                                Case "Еднопозиционен", "Еднопозиционен - противовлажен", "Еднопозиционен - светещ"
                                    wsLines.Cells(2, 8).Value = wsLines.Cells(2, 8).Value + broj_elementi
                                Case "Двупозиционен", "Двупозиционен - противовлажен", "Двупозиционен - светещ",
                             "Деветор", "Девятор - противовлажен", "Девятор светещ"
                                    wsLines.Cells(2, 9).Value = wsLines.Cells(2, 9).Value + broj_elementi
                                Case "Кръстат", "Кръстат - противовлажен", "Кръстат светещ",
                             "Трипозиционен", "Трипозиционен - противовлажен", "Трипозиционен - светещ"
                                    wsLines.Cells(2, 10).Value = wsLines.Cells(2, 10).Value + broj_elementi
                            End Select
                        Case "Ключ_знак_WIFI"
                            Select Case iarrBlock.blVisibility
                                Case "Еднопозиционен - WiFi", "Еднопозиционен - Радио", "Еднопозиционен - Сенсор"
                                    wsLines.Cells(2, 8).Value = wsLines.Cells(2, 8).Value + broj_elementi
                                Case "Двупозиционен - WiFi", "Двупозиционен - Радио", "Двупозиционен - Сенсор"
                                    wsLines.Cells(2, 9).Value = wsLines.Cells(2, 9).Value + broj_elementi
                                Case "Трипозиционен - WiFi", "Трипозиционен - Радио", "Трипозиционен - Сенсор"
                                    wsLines.Cells(2, 10).Value = wsLines.Cells(2, 10).Value + broj_elementi
                                Case "Четирипозиционен - WiFi", "Четирипозиционен - Радио", "Четирипозиционен - Сенсор"
                                    wsLines.Cells(2, 11).Value = wsLines.Cells(2, 11).Value + broj_elementi
                            End Select
                        Case "Ключ_квадрат"
                            Select Case iarrBlock.blVisibility
                                Case "Звънец", "Звънец светещ", "Лихт бутон единичен",
                                     "Лихт бутон единичен светещ", "Стълбищен бутон", "С въженце",
                                     "Стълбищен бутон светещ", "Чипкарта", "Чипкарта светещ", "Ключ управление"
                                    wsLines.Cells(2, 8).Value = wsLines.Cells(2, 8).Value + broj_elementi
                                Case "Димер_обикновен", "Димер_сензорен", "ДКУ", "Завеси", "Щори", "Сензор",
                                      "Лихт бутон двоен", "Лихт бутон двоен светещ", "Регулатор температура"
                                    wsLines.Cells(2, 9).Value = wsLines.Cells(2, 9).Value + broj_elementi
                                Case "Лихт бутон троен", "Лихт бутон троен светещ"
                                    wsLines.Cells(2, 10).Value = wsLines.Cells(2, 10).Value + broj_elementi
                            End Select
                        Case "Контакт"
                            Select Case iarrBlock.blVisibility
                                Case "За монтаж в канал", "С детска защита - противовлажен",
                                     "С детска защита", "Евроамерикански стандарт", "Монифазен - IP 54",
                                     "Тригнездов - противовлажен", "Двугнездов - противовлажен",
                                     "Обикновен - противовлажен", "Тригнездов", "Двугнездов", "Обикновен",
                                     "1xU", "2xU"
                                    wsLines.Cells(2, 16).Value = wsLines.Cells(2, 16).Value + broj_elementi
                                Case "Трифазен - противовлажен", "Трифазен - IP 54", "Трифазен", "ТР+2МФ"
                                    wsLines.Cells(2, 17).Value = wsLines.Cells(2, 17).Value + broj_elementi
                                Case "Усилен", "Твърда връзка"
                                    wsLines.Cells(2, 18).Value = wsLines.Cells(2, 18).Value + broj_elementi
                            End Select
                        Case "Полилей"
                            Select Case iarrBlock.blVisibility
                                Case "1х60 - Кръгла", "1х60 - Рошава", "1х60 - Индийски",
                                     "2х60 - Кръгла", "2х60 - Рошава"
                                    wsLines.Cells(2, 12).Value = wsLines.Cells(2, 12).Value + iarrBlock.count
                                Case "3х60 - Кръгла", "3х60 - Рошава",
                                     "4х60 - Кръгла", "4х60 - Рошава"
                                    wsLines.Cells(2, 13).Value = wsLines.Cells(2, 13).Value + iarrBlock.count
                            End Select
                        Case "Металхаогенна лампа"
                            Select Case iarrBlock.blVisibility
                                Case "1х35 - Дъга", "1х35 - Кръг", "1х35 - Право",
                                     "1х35 - 90°", "1х35 - за картина", "2х35 - за картина",
                                     "2х35 - Дъга", "2х35 - Кръг", "2х35 - Право"
                                    wsLines.Cells(2, 12).Value = wsLines.Cells(2, 12).Value + iarrBlock.count
                                Case "3х35 - Дъга", "3х35 - Кръг", "3х35 - Право",
                                     "4х35 - Дъга", "4х35 - Кръг", "4х35 - Право"
                                    wsLines.Cells(2, 13).Value = wsLines.Cells(2, 13).Value + iarrBlock.count
                            End Select
                        Case "бойлерно табло"
                            wsLines.Cells(2, 28).Value = wsLines.Cells(2, 28).Value + iarrBlock.count
                        Case "Бойлер"
                            Select Case iarrBlock.blVisibility
                                Case "Сешоар с контакт", "Изход газ"
                                    wsLines.Cells(2, 20).Value = wsLines.Cells(2, 16).Value + iarrBlock.count
                                Case "ПВ", "Изход 3p"
                                    Dim I_n As Double
                                    Dim Power As Double
                                    ' Проверяваме дали стойността е число
                                    If Not Double.TryParse(iarrBlock.blLamp_Power, Power) Then
                                        ' Ако не е число, задаваме стойност 2999
                                        Power = 16
                                    End If
                                    I_n = Power / (Math.Sqrt(3) * 380 * 0.8 * 0.8)
                                    Select Case I_n
                                        Case < 17
                                            wsLines.Cells(2, 22).Value =
                                            wsLines.Cells(2, 22).Value + broj_elementi
                                        Case < 26
                                            wsLines.Cells(2, 23).Value =
                                            wsLines.Cells(2, 23).Value + broj_elementi
                                        Case < 33
                                            wsLines.Cells(2, 24).Value =
                                            wsLines.Cells(2, 24).Value + broj_elementi
                                        Case < 43
                                            wsLines.Cells(2, 25).Value =
                                            wsLines.Cells(2, 25).Value + broj_elementi
                                        Case < 65
                                            wsLines.Cells(2, 26).Value =
                                            wsLines.Cells(2, 26).Value + broj_elementi
                                        Case Else
                                            wsLines.Cells(2, 27).Value =
                                            wsLines.Cells(2, 27).Value + broj_elementi
                                    End Select
                                Case "Изход 1p", "Проточен", "Бойлер кухня", "Вертикален", "Хоризонтален"
                                    wsLines.Cells(2, 21).Value = wsLines.Cells(2, 21).Value + iarrBlock.count
                            End Select
                        Case "Вентилации"
                            Select Case iarrBlock.blVisibility
                                Case "Вентилатор - кръг - баня", "Вентилатор - правоъг", "Вентилатор - кръг"
                                    wsLines.Cells(2, 12).Value = wsLines.Cells(2, 12).Value + iarrBlock.count
                                Case "Вентилатор - канален 1P", "Вентилатор - прозоречен 1P", "Конвектор - АСТ",
                                      "Kонвектор - касетъчен", "Вентилатор - кръг - стенен", "Kонвектор", "Климатик_вътре"
                                    wsLines.Cells(2, 16).Value = wsLines.Cells(2, 16).Value + iarrBlock.count
                                Case "Вентилатор - канален 3P", "Вентилатор - прозоречен 3P",
                                     "Нагревател", "Горелка"
                                    Dim I_n As Double
                                    I_n = iarrBlock.blLamp_Power / (Math.Sqrt(3) * 380 * 0.8 * 0.8)
                                    Select Case I_n
                                        Case < 17
                                            wsLines.Cells(2, 22).Value =
                                            wsLines.Cells(2, 22).Value + broj_elementi
                                        Case < 26
                                            wsLines.Cells(2, 23).Value =
                                            wsLines.Cells(2, 23).Value + broj_elementi
                                        Case < 33
                                            wsLines.Cells(2, 24).Value =
                                            wsLines.Cells(2, 24).Value + broj_elementi
                                        Case < 43
                                            wsLines.Cells(2, 25).Value =
                                            wsLines.Cells(2, 25).Value + broj_elementi
                                        Case < 65
                                            wsLines.Cells(2, 26).Value =
                                            wsLines.Cells(2, 26).Value + broj_elementi
                                        Case Else
                                            wsLines.Cells(2, 27).Value =
                                            wsLines.Cells(2, 27).Value + broj_elementi
                                    End Select
                            End Select
                        Case "Луминисцентна лампа"
                        Case "LED_DENIMA"
                    End Select
                    ' #################################################################################################################################
                    index += 1
                Next
                acTrans.Commit()
                With wsKontakti
                    .Cells.Sort(Key1:= .Range("B2"),
                                Order1:=excel.XlSortOrder.xlAscending,
                                Header:=excel.XlYesNoGuess.xlYes,
                                OrderCustom:=1, MatchCase:=False,
                                Orientation:=excel.Constants.xlTopToBottom,
                                DataOption1:=excel.XlSortDataOption.xlSortTextAsNumbers,
                                Key2:= .Range("D2"),
                                Order2:=excel.XlSortOrder.xlAscending,
                                DataOption2:=excel.XlSortDataOption.xlSortTextAsNumbers
                                )
                End With
            Catch ex As Exception
                MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
                acTrans.Abort()
            End Try
        End Using
        ProgressBar_Extrat.Value = ProgressBar_Extrat.Minimum
    End Sub
    Private Sub clearKонтакти(ws As excel.Worksheet, ws_Name As String)
        With ws
            With .Range("A:Z")
                .Clear()
                .Value = ""
            End With
            With .Range("A:Z")
                .Font.Name = "Cambria"
                .Font.Size = 12
                .WrapText = vbTrue
            End With
            .Columns("A").ColumnWidth = 6
            .Columns("B").ColumnWidth = 25
            .Columns("C").ColumnWidth = 11
            .Columns("D").ColumnWidth = 41
            .Columns("E").ColumnWidth = 11
            .Columns("F").ColumnWidth = 13
            .Columns("G").ColumnWidth = 13
            .Columns("H").ColumnWidth = 13
            .Columns("I").ColumnWidth = 13
            .Columns("J").ColumnWidth = 16
            .Columns("K:Z").ColumnWidth = 12

            .Range("A1").Value = "Брой"
            .Range("B1").Value = "Име"
            .Range("C1").Value = "Височина"
            .Range("D1").Value = "Visibility"
            .Range("E1").Value = "Мощност трафо"
            .Range("F1").Value = "Мощност лампа"
            .Range("G1").Value = "Дължина LED лента"
            .Range("H1").Value = "LED Лампа Св_поток"
            .Range("I1").Value = "LED Лампа IP"
            .Range("J1").Value = "LED Лампа Монтаж"

            .Range("L1").Value = "Брой връзки до 2,5"
            .Range("M1").Value = "Брой връзки до 16"
            .Range("N1").Value = "Брой връзки над 16"

        End With
        If ws_Name = "GetExcel" Then Exit Sub
        For i As Integer = 7 To 29
            wsLines.Cells(2, i).Value = 0
        Next
    End Sub
    Private Sub Button_Изчисти_контакти_Click(sender As Object, e As EventArgs) Handles Button_Изчисти_контакти.Click
        clearKонтакти(wsKontakti, "Контакти")
    End Sub
    Private Sub Button_Генератор_Контакти_Click(sender As Object, e As EventArgs) Handles Button_Генератор_Контакти.Click
        Dim index As Integer = 0
        Dim i As Integer
        ProgressBar_Extrat.Maximum = ProgressBar_Maximum
        For i = 6 To 1000
            If wsKol_Smetka.Cells(i, 2).Value = "Раздел" Or wsKol_Smetka.Cells(i, 2).Value = "" Then
                index = i
                Exit For
            End If
            ProgressBar_Extrat.Value = i
        Next
        For i = 2 To 1000
            If wsKontakti.Cells(i, 2).Value = "" Then
                Exit For
            End If
        Next
        ProgressBar_Extrat.Maximum = i
        Dim Swyzwane_1_5 As Integer = 0
        Dim Swyzwane_2_5 As Integer = 0
        Dim Swyzwane_16 As Integer = 0
        Dim Text_Dostawka As String = ""
        Dim broj_elementi As Integer = 0
        Dim Kontakt As Integer = 0
        Dim Zs As Integer = 0
        Dim Konzola_1 As Integer = 0
        Dim Konzola_2 As Integer = 0
        Dim Konzola_3 As Integer = 0
        Dim Брой_Връзки_2 As Integer = 0
        Dim Брой_Връзки_16 As Integer = 0
        Dim Брой_Връзки_63 As Integer = 0

        Call Excel_Kol_smetka_Razdel(wsKol_Smetka, "ЕЛ. ИНСТАЛАЦИИ ЗА ОСВЕТЛЕНИЕ И КОНТАКТИ", "A" & Trim(index.ToString), "D" & Trim(index.ToString))
        index += 1
        Dim Avariq As Integer = 0
        For i = 2 To 10000
            If Trim(wsKontakti.Cells(i, 2).Value) = "" Then Exit For
            ProgressBar_Extrat.Value = i
            Text_Dostawka = ""

            broj_elementi = wsKontakti.Cells(i, 1).Value

            wsKontakti.Range("L" & i.ToString).Value = 0
            wsKontakti.Range("M" & i.ToString).Value = 0
            wsKontakti.Range("N" & i.ToString).Value = 0

            Select Case wsKontakti.Cells(i, 2).Value
                Case "LED_lenta"
                    broj_elementi = wsKontakti.Cells(i, 7).Value / 100
                    Text_Dostawka = "едноцветна гъвкава LED лента;"
                    Text_Dostawka = Text_Dostawka + " 30 led/м, тип 5050, захр. 12V,"
                    Text_Dostawka = Text_Dostawka + " мощност 14,4W/m;"
                    Text_Dostawka = Text_Dostawka + " светлинен поток 450lm/m;"
                    Text_Dostawka = Text_Dostawka + " цветна температура 3000К;"
                    Text_Dostawka = Text_Dostawka + "комплект с профили за монтаж"
                    wsKontakti.Range("L" & i.ToString).Value = 2 * broj_elementi
                Case "LED_ULTRALUX", "LED_ULTRALUX_100"
                    wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                    Text_Dostawka = "светодиодно, осветително тяло"
                    Select Case wsKontakti.Cells(i, 10).Value
                        Case "Гипсокартон"
                            Text_Dostawka = Text_Dostawka + " за монтаж на гипсокартон, "
                        Case "Повърхностен"
                            Text_Dostawka = Text_Dostawka + " за монтаж на таван, "
                        Case "Растерен"
                            Text_Dostawka = Text_Dostawka + " за монтаж на растерен таван, "
                        Case "Авариен"
                            Text_Dostawka = Text_Dostawka + " с авариен модул, "
                        Case "Промишлен"
                            Text_Dostawka = Text_Dostawka + " промишлен тип, "
                        Case "Мебел"
                            Text_Dostawka = Text_Dostawka + " за монтаж под кухненски шкаф, "
                        Case Else
                            Text_Dostawka = "НЕИЗВЕСТЕН ЕЛЕМЕНТ --> " &
                                            wsKontakti.Cells(i, 2).Value & " - " &
                                            wsKontakti.Cells(i, 4).Value
                    End Select

                    Text_Dostawka = Text_Dostawka + Str(wsKontakti.Cells(i, 6).Value) &
                                                    "W, " &
                                                    wsKontakti.Cells(i, 9).Value &
                                                    ", ЕПРА, Ф=" &
                                                    wsKontakti.Cells(i, 8).Value &
                                                    ", Тц=6000К, Ra>80" &
                                                    If(wsKontakti.Cells(i, 3).Value = 60, ", размер 120х30 см", "") &
                                                    ", комплект крепежни елементи"
                Case "LED_луна"
                    wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                    Select Case wsKontakti.Cells(i, 4).Value
                        Case "Лед луна противовлажна"
                            Text_Dostawka = "влагозащитена LED луна; IP 65"

                        Case "Лед луна"
                            Text_Dostawka = "LED луна"
                        Case "Драйвер"
                            Text_Dostawka = "захранване за LED осветително тяло; IN:110-240V,OUT:12VDC;" &
                                            Str(wsKontakti.Cells(i, 5).Value) & "W"
                        Case Else
                            Text_Dostawka = "НЕИЗВЕСТЕН ЕЛЕМЕНТ --> " & wsKontakti.Cells(i, 2).Value & " - " & wsKontakti.Cells(i, 4).Value
                    End Select
                Case "Авария_100", "Авария"
                    wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                    Text_Dostawka = "oсветително тяло за евакуационно осветление с вградени акумулаторни батерии; окомплектовано с LED 4W за монтаж на стена"

                   ' "Захранващ авариен модул за LED осветление с батериен блок LiFePO4, 12.8V, 4000 mAh"

                Case "Ключ_знак"
                    Konzola_1 = Konzola_1 + broj_elementi
                    Select Case wsKontakti.Cells(i, 4).Value
                        Case "Еднопозиционен"
                            Text_Dostawka = "ключ еднополюсен; 10А"
                            wsKontakti.Range("L" & i.ToString).Value = 2 * broj_elementi
                        Case "Еднопозиционен - противовлажен"
                            Text_Dostawka = "ключ еднополюсен; противовлажен; 10А"
                            wsKontakti.Range("L" & i.ToString).Value = 2 * broj_elementi
                        Case "Еднопозиционен - светещ"
                            Text_Dostawka = "ключ еднополюсен; светещ; 10А"
                            wsKontakti.Range("L" & i.ToString).Value = 2 * broj_elementi
                        Case "Двупозиционен"
                            Text_Dostawka = "ключ сериен; 10А"
                            wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                        Case "Двупозиционен - противовлажен"
                            wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                            Text_Dostawka = "ключ сериен; противовлажен; 10А"
                        Case "Двупозиционен - светещ"
                            wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                            Text_Dostawka = "ключ сериен; светещ; 10А"
                        Case "Трипозиционен"
                            wsKontakti.Range("L" & i.ToString).Value = 4 * broj_elementi
                            Text_Dostawka = "ключ триполюсен; 10А"
                        Case "Трипозиционен - противовлажен"
                            wsKontakti.Range("L" & i.ToString).Value = 4 * broj_elementi
                            Text_Dostawka = "ключ триполюсен; противовлажен; 10А"
                        Case "Трипозиционен - светещ"
                            wsKontakti.Range("L" & i.ToString).Value = 4 * broj_elementi
                            Text_Dostawka = "ключ триполюсен; светещ; 10А"
                        Case "Деветор"
                            wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                            Text_Dostawka = "ключ дивиаторен; 10А"
                        Case "Девятор - противовлажен"
                            wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                            Text_Dostawka = "ключ дивиаторен; противовлажен; 10А"
                        Case "Девятор светещ"
                            wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                            Text_Dostawka = "ключ дивиаторен; светещ; 10А"
                        Case "Кръстат"
                            wsKontakti.Range("L" & i.ToString).Value = 4 * broj_elementi
                            Text_Dostawka = "ключ кръстат; 10А"
                        Case "Кръстат - противовлажен"
                            wsKontakti.Range("L" & i.ToString).Value = 4 * broj_elementi
                            Text_Dostawka = "ключ кръстат; противовлажен; 10А"
                        Case "Кръстат светещ"
                            wsKontakti.Range("L" & i.ToString).Value = 4 * broj_elementi
                            Text_Dostawka = "ключ кръстат; светещ; 10А"
                        Case Else
                            Text_Dostawka = "НЕИЗВЕСТЕН ЕЛЕМЕНТ --> " & wsKontakti.Cells(i, 2).Value & " - " & wsKontakti.Cells(i, 4).Value
                    End Select
                Case "Ключ_знак_WIFI"
                    Konzola_1 = Konzola_1 + broj_elementi
                    Select Case wsKontakti.Cells(i, 4).Value
                        Case "Четирипозиционен - WiFi"
                            wsKontakti.Range("L" & i.ToString).Value = 5 * broj_elementi
                            Text_Dostawka = "ключ с WiFi управление; четири канала; 400W"
                        Case "Трипозиционен - WiFi"
                            wsKontakti.Range("L" & i.ToString).Value = 4 * broj_elementi
                            Text_Dostawka = "ключ с WiFi управление; три канала; 400W"
                        Case "Двупозиционен - WiFi"
                            wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                            Text_Dostawka = "ключ с WiFi управление; два канала; 400W"
                        Case "Еднопозиционен - WiFI"
                            wsKontakti.Range("L" & i.ToString).Value = 2 * broj_elementi
                            Text_Dostawka = "ключ с WiFi управление; един канал; 400W"
                        Case "Четирипозиционен - Радио"
                            wsKontakti.Range("L" & i.ToString).Value = 5 * broj_elementi
                            Text_Dostawka = "ключ с RF управление; четири канала; 400W"
                        Case "Трипозиционен - Радио"
                            wsKontakti.Range("L" & i.ToString).Value = 4 * broj_elementi
                            Text_Dostawka = "ключ с RF управление; три канала; 400W"
                        Case "Двупозиционен - Радио"
                            wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                            Text_Dostawka = "ключ с RF управление; два канала; 400W"
                        Case "Еднопозиционен - Радио"
                            wsKontakti.Range("L" & i.ToString).Value = 2 * broj_elementi
                            Text_Dostawka = "ключ с RF управление; един канал; 400W"
                        Case "Еднопозиционен - Сенсор"
                            wsKontakti.Range("L" & i.ToString).Value = 2 * broj_elementi
                            Text_Dostawka = "ключ със сензорно управление; един канал; 400W"
                        Case "Двупозиционен - Сенсор"
                            wsKontakti.Range("L" & i.ToString).Value = 2 * broj_elementi
                            Text_Dostawka = "ключ със сензорно управление; два канала; 400W"
                        Case "Трипозиционен - Сенсор"
                            wsKontakti.Range("L" & i.ToString).Value = 4 * broj_elementi
                            Text_Dostawka = "ключ със сензорно управление; три канала; 400W"
                        Case "Четирипозиционен - Сенсор"
                            wsKontakti.Range("L" & i.ToString).Value = 5 * broj_elementi
                            Text_Dostawka = "ключ със сензорно управление; четири канала; 400W"
                        Case Else
                            Text_Dostawka = "НЕИЗВЕСТЕН ЕЛЕМЕНТ --> " & wsKontakti.Cells(i, 2).Value & " - " & wsKontakti.Cells(i, 4).Value
                    End Select
                Case "Ключ_квадрат"
                    Konzola_1 = Konzola_1 + broj_elementi
                    Select Case wsKontakti.Cells(i, 4).Value
                        Case "Димер_обикновен"
                            Text_Dostawka = "димер"
                            wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                        Case "Димер_сензорен"
                            wsKontakti.Range("L" & i.ToString).Value = 1 * broj_elementi
                            Text_Dostawka = "димер със сензорно управление"
                        Case "ДКУ"
                            wsKontakti.Range("L" & i.ToString).Value = 2 * broj_elementi
                            Text_Dostawka = "двубутонна кнопка управление"
                        Case "Завеси"
                            wsKontakti.Range("L" & i.ToString).Value = 2 * broj_elementi
                            Text_Dostawka = "ключ за управление на завеси"
                        Case "Звънец"
                            wsKontakti.Range("L" & i.ToString).Value = 1 * broj_elementi
                            Text_Dostawka = "бутон за звънец"
                        Case "Звънец светещ"
                            wsKontakti.Range("L" & i.ToString).Value = 1 * broj_elementi
                            Text_Dostawka = "бутон за звънец; светещ"
                        Case "Лихт бутон двоен"
                            wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                            Text_Dostawka = "лихт бутон; двоен"
                        Case "Лихт бутон двоен светещ"
                            wsKontakti.Range("L" & i.ToString).Value = 2 * broj_elementi
                            Text_Dostawka = "лихт бутон; двоен; светещ"
                        Case "Лихт бутон единичен"
                            wsKontakti.Range("L" & i.ToString).Value = 1 * broj_elementi
                            Text_Dostawka = "лихт бутон"
                        Case "Лихт бутон единичен светещ"
                            wsKontakti.Range("L" & i.ToString).Value = 1 * broj_elementi
                            Text_Dostawka = "лихт бутон; светещ"
                        Case "Лихт бутон троен"
                            wsKontakti.Range("L" & i.ToString).Value = 4 * broj_elementi
                            Text_Dostawka = "лихт бутон; троен"
                        Case "Лихт бутон троен светещ"
                            wsKontakti.Range("L" & i.ToString).Value = 4 * broj_elementi
                            Text_Dostawka = "лихт бутон; троен; светещ"
                        Case "Регулатор температура"
                            Text_Dostawka = "терморегулатор"
                        Case "С въженце"
                            wsKontakti.Range("L" & i.ToString).Value = 2 * broj_elementi
                            Text_Dostawka = "бутон; с въженце"
                        Case "Сензор"
                            wsKontakti.Range("L" & i.ToString).Value = 2 * broj_elementi
                            Text_Dostawka = "датчик за движение; за вграждане в розетка"
                        Case "Стълбищен бутон"
                            wsKontakti.Range("L" & i.ToString).Value = 2 * broj_elementi
                            Text_Dostawka = "стълбищен бутон"
                        Case "Стълбищен бутон светещ"
                            wsKontakti.Range("L" & i.ToString).Value = 2 * broj_elementi
                            Text_Dostawka = "стълбищен бутон; светещ"
                        Case "Чипкарта"
                            Text_Dostawka = "чипкарта"
                            wsKontakti.Range("L" & i.ToString).Value = 2 * broj_elementi
                        Case "Чипкарта светещ"
                            Text_Dostawka = "чипкарта; светеща"
                            wsKontakti.Range("L" & i.ToString).Value = 2 * broj_elementi
                        Case "Щори"
                            Text_Dostawka = "ключ за управление на щори"
                            wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                        Case "Ключ управление"
                            Text_Dostawka = "Ключ управление"
                            wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                        Case Else
                            Text_Dostawka = "НЕИЗВЕСТЕН ЕЛЕМЕНТ --> " & wsKontakti.Cells(i, 2).Value & " - " & wsKontakti.Cells(i, 4).Value
                    End Select
                Case "Контакт"
                    Select Case wsKontakti.Cells(i, 4).Value
                        Case "Трифазен - противовлажен"
                            Text_Dostawka = "контакт трифазен;25А; противовлажен"
                            Zs = Zs + 3 * broj_elementi
                            wsKontakti.Range("M" & i.ToString).Value = 5 * broj_elementi
                        Case "Трифазен - IP 54"
                            wsKontakti.Range("M" & i.ToString).Value = 5 * broj_elementi
                            Text_Dostawka = "контакт трифазен;25А; евро индустриален; 3p+PE+N; IP 54"
                            Zs = Zs + 3 * broj_elementi
                        Case "Трифазен"
                            wsKontakti.Range("M" & i.ToString).Value = 5 * broj_elementi
                            Text_Dostawka = "контакт трифазен;25А"
                            Zs = Zs + 3 * broj_elementi
                        Case "За монтаж в канал"
                            wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                            Text_Dostawka = "контакт 'Шуко' едногнездов; за монтаж в кабелен канал"
                            Zs = Zs + 1 * broj_elementi
                        Case "С детска защита - противовлажен"
                            wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                            Text_Dostawka = "контакт 'Шуко' едногнездов; детска защита; противовлажен"
                            Konzola_1 = Konzola_1 + broj_elementi
                            Zs = Zs + 1 * broj_elementi
                        Case "С детска защита"
                            wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                            Text_Dostawka = "контакт 'Шуко' едногнездов; детска защита"
                            Konzola_1 = Konzola_1 + broj_elementi
                            Zs = Zs + 1 * broj_elementi
                        Case "Евроамерикански стандарт"
                            wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                            Text_Dostawka = "Контактен излаз евроамерикански стандарт"
                            Konzola_1 = Konzola_1 + broj_elementi
                            Zs = Zs + 1 * broj_elementi
                        Case "Монифазен - IP 54"
                            wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                            Text_Dostawka = "контакт еднофазен;25А; евро индустриален; 1p+PE+N; IP 54"
                            Zs = Zs + 1 * broj_elementi
                        Case "Твърда връзка"
                            wsKontakti.Range("M" & i.ToString).Value = 3 * broj_elementi
                            Text_Dostawka = "контакт;твърда връзка; скрит монтаж; 25А"
                            Zs = Zs + 1 * broj_elementi
                        Case "Усилен"
                            wsKontakti.Range("M" & i.ToString).Value = 3 * broj_elementi
                            Text_Dostawka = "Контакт;усилен;1p+PE+N; 25А"
                            Zs = Zs + 1 * broj_elementi
                        Case "Тригнездов - противовлажен"
                            wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                            Text_Dostawka = "контакт 'Шуко'; тригнездов; противовлажен"
                            Zs = Zs + 3 * broj_elementi
                            wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                        Case "Двугнездов - противовлажен"
                            Text_Dostawka = "контакт 'Шуко'; двугнездов; противовлажен"
                            Zs = Zs + 2 * broj_elementi
                            wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                        Case "Обикновен - противовлажен"
                            wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                            Text_Dostawka = "контакт 'Шуко'; противовлажен"
                            Zs = Zs + 1 * broj_elementi
                        Case "Тригнездов"
                            wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                            Kontakt = Kontakt + 3 * broj_elementi
                            Konzola_3 = Konzola_3 + broj_elementi
                            Zs = Zs + 3 * broj_elementi
                        Case "Двугнездов"
                            wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                            Kontakt = Kontakt + 2 * broj_elementi
                            Konzola_2 = Konzola_2 + broj_elementi
                            Zs = Zs + 2 * broj_elementi
                        Case "Обикновен"
                            wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                            Kontakt = Kontakt + 1 * broj_elementi
                            Konzola_1 = Konzola_1 + broj_elementi
                            Zs = Zs + 1 * broj_elementi
                        Case "1xU"
                            wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                            Text_Dostawka = "контакт 'Шуко' едногнездов; USB зарядно тип A; ток 3A;"
                            Zs = Zs + 1 * broj_elementi
                        Case "2xU"
                            wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                            Text_Dostawka = "контакт 'Шуко' двугнездов; USB зарядно тип A+C; ток 3A;"
                            Zs = Zs + 1 * broj_elementi
                        Case "ТР+2МФ"
                            wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                            Text_Dostawka = "контакт трифазен модул, трифазен + 2 монофазни контакта, открит монтаж"
                            Zs = Zs + 3 * broj_elementi
                        Case Else
                            Text_Dostawka = "НЕИЗВЕСТЕН ЕЛЕМЕНТ --> " & wsKontakti.Cells(i, 2).Value & " - " & wsKontakti.Cells(i, 4).Value
                    End Select
                Case "Линия МХЛ - 220V"
                    wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                    Select Case wsKontakti.Cells(i, 4).Value
                        Case "1х26 - квадрат IP 54", "1х26 - IP 54"
                            Text_Dostawka = "осветително тяло LED; " & wsKontakti.Cells(i, 6).Value & "W; противовлажено IP54; 220V; модел по избор на архитекта"
                        Case "1х26 - квадрат", "1x26 - без решетка", "1x26-с решетка"
                            Text_Dostawka = "осветително тяло LED; " & wsKontakti.Cells(i, 6).Value & "W; IP20; 220V; модел по избор на архитекта"
                        Case "1х26 - квадрат IP 54 датчик"
                            Text_Dostawka = "осветително тяло LED; " & wsKontakti.Cells(i, 6).Value & "W; противовлажено IP54; датчик за движение; 220V; модел по избор на архитекта"
                        Case "1х26 - квадрат датчик"
                            Text_Dostawka = "осветително тяло LED; " & wsKontakti.Cells(i, 6).Value & "W; IP20; датчик за движение; 220V; модел по избор на архитекта"
                    End Select
                Case "Плафони"
                    wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                    Select Case wsKontakti.Cells(i, 4).Value
                        Case "текст"
                        Case "Лампион - рошав", "Лампион"
                            Text_Dostawka = "лампион; " &
                                wsKontakti.Cells(i, 6).Value &
                                "W;-модел по избор на архитекта"
                        Case "Настолна лампа - рошава", "Настолна лампа"
                            Text_Dostawka = "настолна лампа; " &
                                wsKontakti.Cells(i, 6).Value &
                                "W; модел по избор на архитекта"
                        Case "Пендел - противовлажен с датчик"
                            Text_Dostawka = "осветително тяло пендел; " &
                                wsKontakti.Cells(i, 6).Value &
                                "W; IP54; противовлажен; с вграден датчик за движение; модел по избор на архитекта"
                        Case "Пендел с датчик"
                            Text_Dostawka = "осветително тяло пендел; " &
                                wsKontakti.Cells(i, 6).Value &
                                "W; IP20; противовлажен; с вграден датчик за движение; модел по избор на архитекта"
                        Case "Пендел"
                            Text_Dostawka = "осветително тяло пендел; " &
                                wsKontakti.Cells(i, 6).Value &
                                "W; IP54; противовлажен; модел по избор на архитекта"
                        Case "Пендел - противовлажен"
                            Text_Dostawka = "осветително тяло пендел; " &
                                wsKontakti.Cells(i, 6).Value &
                                "W; IP20; модел по избор на архитекта"
                        Case "Плафон - противовлажен с датчик"
                            Text_Dostawka = "осветително тяло плафон; " &
                                wsKontakti.Cells(i, 6).Value &
                                "W; IP54; противовлажен; с вграден датчик за движение; модел по избор на архитекта"
                        Case "Плафон - противовлажен"
                            Text_Dostawka = "осветително тяло плафон; " &
                                wsKontakti.Cells(i, 6).Value &
                                "W; IP54; противовлажен; модел по избор на архитекта"
                        Case "Плафон с датчик"
                            Text_Dostawka = "осветително тяло плафон; " &
                                wsKontakti.Cells(i, 6).Value &
                                "W; IP20; с вграден датчик за движение; модел по избор на архитекта"
                        Case "Плафон"
                            Text_Dostawka = "осветително тяло плафон; " &
                                wsKontakti.Cells(i, 6).Value &
                                "W; IP20; модел по избор на архитекта"
                        Case "Аплик - противовлажен с датчик", "Аплик - Рошав - противовлажен с датчик"
                            Text_Dostawka = "осветително тяло аплик за стена; " &
                                wsKontakti.Cells(i, 6).Value &
                                "W; IP54; противовлажен; с вграден датчик за движение; модел по избор на архитекта"
                        Case "Аплик - противовлажен", "Аплик - Рошав - противовлажен"
                            Text_Dostawka = "осветително тяло аплик за стена; " &
                                wsKontakti.Cells(i, 6).Value &
                                "W; IP54; противовлажен; модел по избор на архитекта"
                        Case "Аплик с датчик", "Аплик - Рошав с датчик"
                            Text_Dostawka = "осветително тяло аплик за стена; " &
                                wsKontakti.Cells(i, 6).Value &
                                "W IP20; с вграден датчик за движение; модел по избор на архитекта"
                        Case "Аплик", "Аплик - Рошав"
                            Text_Dostawka = "осветително тяло аплик за стена; " &
                                wsKontakti.Cells(i, 6).Value &
                                "W; IP20; модел по избор на архитекта"
                        Case "Фотодатчик"
                            Text_Dostawka = "фотоелемент за стенен монтаж; IP54; 2 до 100 Lx"
                        Case "Датчик 360°"
                            Text_Dostawka = "датчик за движение;"
                            Text_Dostawka = Text_Dostawka + " за монтаж на таван;"
                            Text_Dostawka = Text_Dostawka + " обхват: 360°;"
                            Text_Dostawka = Text_Dostawka + " осветеност: 2-2000lx;"
                            Text_Dostawka = Text_Dostawka + " обхват: 4-5m"
                        Case "Датчик насочен"
                            Text_Dostawka = "датчик за движение;"
                            Text_Dostawka = Text_Dostawka + " за монтаж на стена;"
                            Text_Dostawka = Text_Dostawka + " насочен: 165°;"
                            Text_Dostawka = Text_Dostawka + " осветеност: 2-2000lx;"
                            Text_Dostawka = Text_Dostawka + " обхват: 10-12m"
                        Case "Бански аплик ЛЕД", "Бански аплик"
                            Text_Dostawka = "осветително тяло аплик за баня; " &
                                wsKontakti.Cells(i, 6).Value &
                                "W IP54; противовлажен; с вграден ключ; модел по избор на архитекта"
                        Case "Общо означение"
                            Text_Dostawka = "осветително тяло; " & wsKontakti.Cells(i, 6).Value & "W; модел по избор на архитекта"
                        Case "Фасадно"
                            Text_Dostawka = "Фасадно осветително тяло;"
                            Text_Dostawka += "IP54; противовлажен; модел по избор на архитекта"
                        Case "Фасадно с датчик"
                            Text_Dostawka = "Фасадно осветително тяло;"
                            Text_Dostawka += "с вграден датчик за движение;"
                            Text_Dostawka += "IP54; противовлажен;"
                            Text_Dostawka += "модел по избор на архитекта"
                        Case Else
                            Text_Dostawka = "НЕИЗВЕСТЕН ЕЛЕМЕНТ --> " & wsKontakti.Cells(i, 2).Value & " - " & wsKontakti.Cells(i, 4).Value
                    End Select
                Case "Полилей"
                    wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                    Select Case wsKontakti.Cells(i, 4).Value
                        Case "1х60 - Индийски"
                            Text_Dostawka = "плафониера с дистанционно управление; " &
                                "една лампа; " &
                                wsKontakti.Cells(i, 6).Value &
                                "W; IP20; модел по избор на архитекта"
                        Case "1х60 - Кръгла", "1х60 - Рошава"
                            Text_Dostawka = "осветително тяло полилей; " &
                                "една лампа; " &
                                wsKontakti.Cells(i, 6).Value &
                                "W; IP20; модел по избор на архитекта"
                        Case "2х60 - Кръгла", "2х60 - Рошава"
                            Text_Dostawka = "осветително тяло полилей; " &
                                "две лампи; " &
                                wsKontakti.Cells(i, 6).Value &
                                "W; IP20; модел по избор на архитекта"
                        Case "3х60 - Кръгла", "3х60 - Рошава"
                            Text_Dostawka = "осветително тяло полилей; " &
                                "три лампи; " &
                                wsKontakti.Cells(i, 6).Value &
                                "W; IP20; модел по избор на архитекта"
                        Case "4х60 - Кръгла", "4х60 - Рошава"
                            Text_Dostawka = "осветително тяло полилей; " &
                                "четири лампи; " &
                                wsKontakti.Cells(i, 6).Value &
                                "W; IP20; модел по избор на архитекта"
                        Case Else
                            Text_Dostawka = "НЕИЗВЕСТЕН ЕЛЕМЕНТ --> " & wsKontakti.Cells(i, 2).Value & " - " & wsKontakti.Cells(i, 4).Value
                    End Select
                Case "Металхаогенна лампа"
                    wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                    Select Case wsKontakti.Cells(i, 4).Value
                        Case "1х35 - Дъга", "1х35 - Кръг", "1х35 - Право"
                            Text_Dostawka = "осветитално тяло със спотове; " &
                                "една лампа; " &
                                wsKontakti.Cells(i, 6).Value &
                                "W; IP20; модел по избор на архитекта"
                        Case "1х35 - 90°"
                            Text_Dostawka = "осветитално тяло със спотове; " &
                                "една лампа; " &
                                "за монтаж на стена; " &
                                wsKontakti.Cells(i, 6).Value &
                                "W; IP20; модел по избор на архитекта"
                        Case "1х35 - за картина"
                            Text_Dostawka = "осветитално тяло със спотове; " &
                                "една лампа; " &
                                "за монтаж над картина; " &
                                wsKontakti.Cells(i, 6).Value &
                                "W; IP20; модел по избор на архитекта"
                        Case "2х35 - за картина"
                            Text_Dostawka = "осветитално тяло със спотове; " &
                                "две лампи; " &
                                "за монтаж над картина; " &
                                wsKontakti.Cells(i, 6).Value &
                                "W; IP20; модел по избор на архитекта"
                        Case "2х35 - Дъга", "2х35 - Кръг", "2х35 - Право"
                            Text_Dostawka = "осветитално тяло със спотове; " &
                                "две лампи; " &
                                wsKontakti.Cells(i, 6).Value &
                                "W; IP20; модел по избор на архитекта"
                        Case "3х35 - Дъга", "3х35 - Кръг", "3х35 - Право"
                            Text_Dostawka = "осветитално тяло със спотове; " &
                                "три лампи; " &
                                wsKontakti.Cells(i, 6).Value &
                                "W; IP20; модел по избор на архитекта"
                        Case "4х35 - Дъга", "4х35 - Кръг", "4х35 - Право"
                            Text_Dostawka = "осветитално тяло със спотове; " &
                                "черири лампи; " &
                                wsKontakti.Cells(i, 6).Value &
                                "W; IP20; модел по избор на архитекта"
                        Case Else
                            Text_Dostawka = "НЕИЗВЕСТЕН ЕЛЕМЕНТ --> " & wsKontakti.Cells(i, 2).Value & " - " & wsKontakti.Cells(i, 4).Value
                    End Select
                Case "Прожектор"
                    wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                    Select Case wsKontakti.Cells(i, 4).Value
                        Case "МХЛ - кръгла", "МХЛ"
                            Text_Dostawka = "улично осветително тяло; " &
                                wsKontakti.Cells(i, 6).Value &
                                "W; IP65; 100lm/W; Тц=4200K; Ra>80"
                        Case "МЛХ - с датчик"
                            Text_Dostawka = "улично осветително тяло; " &
                                wsKontakti.Cells(i, 6).Value &
                                "W; с вграден датчик за движение;" &
                                " IP65; 100lm/W; Тц=4200K; Ra>80IP54"
                        Case Else
                            Text_Dostawka = "НЕИЗВЕСТЕН ЕЛЕМЕНТ --> " & wsKontakti.Cells(i, 2).Value & " - " & wsKontakti.Cells(i, 4).Value
                    End Select
                Case "бойлерно табло"
                    wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                    Select Case wsKontakti.Cells(i, 4).Value
                        Case "С два контакта и един ключ"
                            Text_Dostawka = "бойлерно табло с два контакта и един ключ; 25A"
                        Case "С два ключа и контакт"
                            Text_Dostawka = "бойлерно табло с два ключа и контакт; 25A"
                        Case "Ключ и контакт"
                            Text_Dostawka = "бойлерно табло с ключ и контакт; 25A"
                        Case "Само ключ"
                            Text_Dostawka = "бойлерно табло; 25A"
                        Case Else
                            Text_Dostawka = "НЕИЗВЕСТЕН ЕЛЕМЕНТ --> " & wsKontakti.Cells(i, 2).Value & " - " & wsKontakti.Cells(i, 4).Value
                    End Select
                Case "Бойлер"
                    Select Case wsKontakti.Cells(i, 4).Value
                        Case "Сешоар с контакт"
                            wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                            Text_Dostawka = "Сешоар за ръце; " &
                                "вграден контакт;" &
                                wsKontakti.Cells(i, 6).Value &
                                "W"
                        Case "Сешоар"
                            wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                            Text_Dostawka = "Сешоар за ръце; " &
                                wsKontakti.Cells(i, 6).Value &
                                "W"
                        Case "Изход газ"
                            wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                            With wsKol_Smetka
                                .Cells(index, 2).Value = "Доставка и монтаж на " &
                                    "автоматичен прекъсвач; E60; крива: С;брой полюси: 1;" &
                                    "номинален ток: 6А"
                                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                                .Cells(index, 3).Value = "бр."
                                .Cells(index, 4).Value = broj_elementi
                                index += 1
                                .Cells(index, 2).Value = "Доставка и монтаж на " &
                                    "кутия за автоматичен прекъсвач; брой модули: 2; брой редве:1"
                                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                                .Cells(index, 3).Value = "бр."
                                .Cells(index, 4).Value = broj_elementi
                                Text_Dostawka = ""
                                index += 1
                            End With
                        Case "ПВ"
                            Text_Dostawka = "пускател въздушен; " &
                                wsKontakti.Cells(i, 6).Value
                        Case "Изход 1p", "Проточен", "Бойлер кухня",
                             "Вертикален", "Хоризонтален"
                            wsKontakti.Range("M" & i.ToString).Value = 3 * broj_elementi
                        Case "Изход 3p"
                            Dim value As Double
                            ' Проверяваме дали стойността е число
                            If Not Double.TryParse(wsKontakti.Range("F" & i.ToString).Value, value) Then
                                ' Ако не е число, задаваме стойност 2999
                                value = 299
                            End If
                            ' Изпълняваме Select Case с проверената или зададената стойност
                            Dim rezultat As Integer = 5 * broj_elementi
                            Select Case value
                                Case Is < 3000
                                    wsKontakti.Range("L" & i.ToString).Value = rezultat
                                Case 3000 To 29999
                                    wsKontakti.Range("M" & i.ToString).Value = rezultat
                                Case Else
                                    wsKontakti.Range("N" & i.ToString).Value = rezultat
                            End Select

                        Case "Линии", "Само текст"
                        Case Else
                            Text_Dostawka = "НЕИЗВЕСТЕН ЕЛЕМЕНТ --> " & wsKontakti.Cells(i, 2).Value & " - " & wsKontakti.Cells(i, 4).Value
                    End Select
                Case "Вентилации"
                    wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                    Select Case wsKontakti.Cells(i, 4).Value
                        Case "Вентилатор - кръг - баня", "Вентилатор - правоъг", "Вентилатор - кръг"
                            Text_Dostawka = "вентилатор за баня; " &
                                wsKontakti.Cells(i, 6).Value &
                                "W"
                            wsKontakti.Range("L" & i.ToString).Value = 3 * broj_elementi
                        Case "Вентилатор - канален 1P", "Вентилатор - прозоречен 1P",
                             "Линии", "Конвектор - АСТ", "Kонвектор - касетъчен",
                             "Вентилатор - кръг - стенен", "Kонвектор", "Климатик_вътре",
                             "Вентилатор - канален 3P", "Вентилатор - прозоречен 3P", "Нагревател", "Горелка"
                        Case Else
                            Text_Dostawka = "НЕИЗВЕСТЕН ЕЛЕМЕНТ --> " & wsKontakti.Cells(i, 2).Value & " - " & wsKontakti.Cells(i, 4).Value
                    End Select
                Case "Луминисцентна лампа"
                Case "LED_DENIMA"
                Case Else
                    Text_Dostawka = "НЕИЗВЕСТЕН ЕЛЕМЕНТ --> " & wsKontakti.Cells(i, 2).Value & " - " & wsKontakti.Cells(i, 4).Value
            End Select
            If Trim(Text_Dostawka) <> "" Then
                With wsKol_Smetka
                    .Cells(index, 2).Value = "Доставка и монтаж на " + Text_Dostawka
                    .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                    .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                    .Cells(index, 3).Value = IIf(wsKontakti.Cells(i, 2).Value = "LED_lenta", "m", "бр.")
                    .Cells(index, 4).Value = broj_elementi
                End With
                index += 1
            End If
            Брой_Връзки_2 = Брой_Връзки_2 + wsKontakti.Range("L" & i.ToString).Value
            Брой_Връзки_16 = Брой_Връзки_16 + wsKontakti.Range("M" & i.ToString).Value
            Брой_Връзки_63 = Брой_Връзки_63 + wsKontakti.Range("N" & i.ToString).Value
        Next
        If Kontakt > 0 Then
            Text_Dostawka = "контакт единичен 'Шуко'; механизъм; скрит монтаж; 16А; IP20"
            With wsKol_Smetka
                .Cells(index, 2).Value = "Доставка и монтаж на " + Text_Dostawka
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "бр."
                .Cells(index, 4).Value = Kontakt
            End With
            index += 1
        End If
        If Konzola_1 > 0 Then
            Text_Dostawka = "единична конзола"
            With wsKol_Smetka
                .Cells(index, 2).Value = "Доставка и монтаж на " + Text_Dostawka
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "бр."
                .Cells(index, 4).Value = Konzola_1
            End With
            index += 1
        End If
        If Konzola_2 > 0 Then
            Text_Dostawka = "двойна конзола"
            With wsKol_Smetka
                .Cells(index, 2).Value = "Доставка и монтаж на " + Text_Dostawka
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "бр."
                .Cells(index, 4).Value = Konzola_2
            End With
            index += 1
        End If
        If Konzola_3 > 0 Then
            Text_Dostawka = "тройна конзола"
            With wsKol_Smetka
                .Cells(index, 2).Value = "Доставка и монтаж на " + Text_Dostawka
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "бр."
                .Cells(index, 4).Value = Konzola_3
            End With
            index += 1
        End If
        If wsKoef.Cells(10, 2).Value > 0 Then
            Text_Dostawka = "Доставка и монтаж на разклонителни кутии"
            With wsKol_Smetka
                .Cells(index, 2).Value = Text_Dostawka
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "бр."
                .Cells(index, 4).Value = wsKoef.Cells(10, 2).Value
            End With
            index += 1
        End If
        If (Konzola_1 + Konzola_2 + Konzola_3 + wsKoef.Cells(10, 2).Value) > 0 Then
            Text_Dostawka = "Направа на отвор в стената за монтаж на конзоли и разклонителни кутии"
            With wsKol_Smetka
                .Cells(index, 2).Value = Text_Dostawka
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "бр."
                .Cells(index, 4).Value = Konzola_1 +
                                         Konzola_2 * 3 +
                                         Konzola_3 * 3 +
                                         wsKoef.Cells(10, 2).Value
            End With
            index += 1
        End If
        With wsKol_Smetka
            If Брой_Връзки_2 > 0 Then
                Text = "Свързване проводник към съоръжение до 2,5мм²"
                .Cells(index, 2).Value = Text
                .Cells(index, 3).Value = "бр."
                .Cells(index, 4).Value = Брой_Връзки_2
                index += 1
            End If
            If Брой_Връзки_16 > 0 Then
                Text = "Свързване проводник към съоръжение до 16мм²"
                .Cells(index, 2).Value = Text
                .Cells(index, 3).Value = "бр."
                .Cells(index, 4).Value = Брой_Връзки_16
                index += 1
            End If
            If Брой_Връзки_63 > 0 Then
                Text = "Свързване проводник към съоръжение до 16мм²"
                .Cells(index, 2).Value = Text
                .Cells(index, 3).Value = "бр."
                .Cells(index, 4).Value = Брой_Връзки_63
                index += 1
            End If
        End With
        Calc_Ka4vane_Silova()
        index = Kol_Smetka_Kabeli(index, vbTrue, "СИЛОВА")
        If Zs > 0 Then
            wsKoef.Cells(4, 5).Value = Zs
        End If
        ProgressBar_Extrat.Value = ProgressBar_Extrat.Minimum
    End Sub
    Private Sub Calc_Ka4vane_Silova()
        Dim Kabel_Sum(100) As strKabel
        With wsKa4vane
            Dim kabel_Ot(100) As strKabel
            Dim kabel_Ky(100) As strKabel
            Dim br_Ot As Integer = 0
            Dim br_Ky As Integer = 0
            For i = 2 To 500
                If Len(Trim(.Range("A" & i.ToString).Value &
                        .Range("N" & i.ToString).Value)) = 0 Then Exit For

                For j = 3 To 13
                    Dim vaCells As String = .Cells(i, j).Value
                    Dim poz As Integer = InStr(vaCells, "л.")
                    Dim br_kab As Integer = Val(Mid(vaCells, 1, poz))
                    If br_kab = 0 Then Continue For
                    kabel_Ot(br_Ot).blType = Trim(Mid(vaCells, poz + 2, Len(vaCells)))
                    kabel_Ot(br_Ot).blLength = Val(.Range("A" & i.ToString).Value) * br_kab
                    kabel_Ot(br_Ot).blPol = Trim(.Range("B" & i.ToString).Value)
                    br_Ot += 1
                Next

                For j = 16 To 26
                    Dim vaCells As String = .Cells(i, j).Value
                    Dim poz As Integer = InStr(vaCells, "л.")
                    Dim br_kab As Integer = Val(Mid(vaCells, 1, poz))
                    If br_kab = 0 Then Continue For
                    kabel_Ky(br_Ky).blType = Trim(Mid(vaCells, poz + 2, Len(vaCells)))
                    kabel_Ky(br_Ky).blLength = Val(.Range("N" & i.ToString).Value) * br_kab
                    kabel_Ky(br_Ky).blPol = Trim(.Range("O" & i.ToString).Value)
                    br_Ky += 1
                Next
            Next
            Dim indexSort As Integer = 0
            For i = 0 To UBound(kabel_Ot)
                If kabel_Ot(i).blLength = 0 Then Exit For
                Dim iVisib As Integer = -1
                Dim strType As String = kabel_Ot(i).blType
                Dim strPol As String = kabel_Ot(i).blPol
                iVisib = Array.FindIndex(Kabel_Sum, Function(f) f.blType = strType And
                                               f.blPol = strPol)

                If iVisib = -1 Then
                    Kabel_Sum(indexSort).blType = kabel_Ot(i).blType
                    Kabel_Sum(indexSort).blPol = kabel_Ot(i).blPol
                    Kabel_Sum(indexSort).blLength = kabel_Ot(i).blLength
                    indexSort += 1
                Else
                    Kabel_Sum(iVisib).blLength += kabel_Ot(i).blLength
                End If
            Next
            For i = 0 To UBound(kabel_Ky)
                If kabel_Ky(i).blLength = 0 Then Exit For
                Dim iVisib As Integer = -1
                Dim strType As String = kabel_Ky(i).blType
                Dim strPol As String = kabel_Ky(i).blPol
                iVisib = Array.FindIndex(Kabel_Sum, Function(f) f.blType = strType And
                                               f.blPol = strPol)

                If iVisib = -1 Then
                    Kabel_Sum(indexSort).blType = kabel_Ky(i).blType
                    Kabel_Sum(indexSort).blPol = kabel_Ky(i).blPol
                    Kabel_Sum(indexSort).blLength = kabel_Ky(i).blLength
                    indexSort += 1
                Else
                    Kabel_Sum(iVisib).blLength += kabel_Ky(i).blLength
                End If
            Next
        End With
        '
        ' Преобразува кабелите и типовете ако са в тръби
        '
        For i = 0 To UBound(Kabel_Sum)
            If Kabel_Sum(i).blType = Nothing Then Exit For

            If Kabel_Sum(i).blPol = "Тръби" Then
                Kabel_Sum(i).blPol = cu.SET_line_Type(Kabel_Sum(i).blType)
                Kabel_Sum(i).blPol = cu.GET_line_Type(Kabel_Sum(i).blPol, True)
            End If

            Kabel_Sum(i).blType = cu.line_Layer(Kabel_Sum(i).blType)
        Next
        '
        ' Търси тип кабели и го добавя
        '
        For i = 0 To UBound(Kabel_Sum)
            If Kabel_Sum(i).blType = Nothing Then Exit For
            Dim strType As String = Kabel_Sum(i).blType
            Dim strPol As String = Kabel_Sum(i).blPol
            Dim boKabel As Boolean = True   ' променлива дали е добавен кабел
            '
            ' Проверява таблица в EXCEL за наличие на тип кабел и записва дължината на качването
            '
            For j = 3 To 500
                Dim ssss = wsLines.Range("A" & j.ToString).Value
                Dim sss = wsLines.Range("B" & j.ToString).Value
                If ssss = "" Then Exit For
                If ssss = "СИЛОВА-КАБЕЛ" Then
                    If sss = Trim(strType) Then
                        wsLines.Range("BG" & j.ToString).Value = Kabel_Sum(i).blLength / 2
                        '
                        '   намира първия кабел и го записва                        '   
                        '
                        boKabel = False
                        Exit For
                    End If
                End If
            Next
            '
            '   Ако не е намерен тип кабел добавя този тип в таблица в EXCEL
            '
            If boKabel Then
                For j = 3 To 500
                    Dim ssss As String = wsLines.Range("A" & j.ToString).Value
                    If ssss = "" Then
                        wsLines.Range("A" & j.ToString).Value = "СИЛОВА-КАБЕЛ"
                        wsLines.Range("B" & j.ToString).Value = Trim(Kabel_Sum(i).blType)
                        wsLines.Range("C" & j.ToString).Value = Kabel_Sum(i).blLength / 2

                        Dim formu As String = ""
                        Dim rang As String = ""
                        Dim colum As String = ""

                        formu = "=sum(G3:BZ3)+C3"
                        rang = "D3:D" & Trim((j).ToString)
                        wsLines.Range(rang).Formula = formu

                        formu = "=D3*$E$2"
                        rang = "E3:E" & Trim((j).ToString)
                        wsLines.Range(rang).Formula = formu

                        formu = "=INT(E3/10+$F$2)*10"
                        rang = "F3:F" & Trim((j).ToString)
                        wsLines.Range(rang).Formula = formu

                        formu = 1.3
                        rang = "E2:E2"
                        wsLines.Range(rang).Formula = formu

                        formu = 1
                        rang = "F2:F2"
                        wsLines.Range(rang).Formula = formu
                        Exit For
                    End If
                Next
            End If
        Next
        '
        ' Записва начина на полагане
        '
        For i = 0 To UBound(Kabel_Sum)
            If Kabel_Sum(i).blType = Nothing Then Exit For
            Dim strPol As String = Kabel_Sum(i).blPol
            Select Case strPol
                Case "Тръби"
                Case "Скара"
                Case "Канал"
            End Select
        Next
        '
        ' Търси начин на полагане и го добавя
        '
        For i = 0 To UBound(Kabel_Sum)
            If Kabel_Sum(i).blType = Nothing Then Exit For
            Dim strType As String = Kabel_Sum(i).blType
            Dim strPol As String = Kabel_Sum(i).blPol
            Dim boKabel As Boolean = True   ' променлива дали е добавен кабел
            '
            ' Проверява таблица в EXCEL за наличие на начин на полагане и записва дължината на качването
            '
            For j = 3 To 500
                Dim ssss = wsLines.Range("A" & j.ToString).Value
                Dim sss = wsLines.Range("B" & j.ToString).Value
                If ssss = "" Then Exit For
                If ssss = "СИЛОВА-ПОЛАГАНЕ" Then
                    If sss = Trim(strPol) Then
                        wsLines.Range("BG" & j.ToString).Value = Kabel_Sum(i).blLength / 2
                        '
                        '   намира първия кабел и го записва                        '   
                        '
                        boKabel = False
                        Exit For
                    End If
                End If
            Next
            '
            '   Ако не е намерен начин на полагане ги добавя в таблица в EXCEL
            '
            If boKabel Then
                For j = 3 To 500
                    Dim ssss As String = wsLines.Range("A" & j.ToString).Value
                    If ssss = "" Then
                        wsLines.Range("A" & j.ToString).Value = "СИЛОВА-ПОЛАГАНЕ"
                        wsLines.Range("B" & j.ToString).Value = Trim(Kabel_Sum(i).blPol)
                        wsLines.Range("C" & j.ToString).Value = Kabel_Sum(i).blLength / 2

                        Dim formu As String = ""
                        Dim rang As String = ""
                        Dim colum As String = ""

                        formu = "=sum(G3:BZ3)+C3"
                        rang = "D3:D" & Trim((j).ToString)
                        wsLines.Range(rang).Formula = formu

                        formu = "=D3*$E$2"
                        rang = "E3:E" & Trim((j).ToString)
                        wsLines.Range(rang).Formula = formu

                        formu = "=INT(E3/10+$F$2)*10"
                        rang = "F3:F" & Trim((j).ToString)
                        wsLines.Range(rang).Formula = formu

                        formu = 1.3
                        rang = "E2:E2"
                        wsLines.Range(rang).Formula = formu

                        formu = 1
                        rang = "F2:F2"
                        wsLines.Range(rang).Formula = formu
                        Exit For
                    End If
                Next
            End If
        Next
        '
        '  Сортира EXCEL листа
        ' 
        With wsLines
            .Cells.Sort(Key1:= .Range("A2"),
                        Order1:=excel.XlSortOrder.xlAscending,
                        Header:=excel.XlYesNoGuess.xlYes,
                        OrderCustom:=1, MatchCase:=False,
                        Orientation:=excel.Constants.xlTopToBottom,
                        DataOption1:=excel.XlSortDataOption.xlSortTextAsNumbers,
                        Key2:= .Range("B2"),
                        Order2:=excel.XlSortOrder.xlAscending,
                        DataOption2:=excel.XlSortDataOption.xlSortTextAsNumbers
                        )
        End With
    End Sub
    Private Sub Button_Номерирай_Click(sender As Object, e As EventArgs) Handles Button_Номерирай.Click,
                                                                                 Button_Свий_Лист.Click
        Dim Index As Integer = 6
        Dim Nomer As Integer = 1
        Dim Rang As String
        ProgressBar_Extrat.Maximum = ProgressBar_Maximum
        Dim _RowHeight As Double = 15.75
        Dim Red_RowHeight As Double = 1
        _RowHeight = IIf(sender.name = "Button_Номерирай", 15.75, 15)
        With wsKol_Smetka
            Do
                If wsKol_Smetka.Cells(Index, 2).Interior.ThemeColor <>
                excel.XlThemeColor.xlThemeColorAccent6 Then
                    Dim sss As String = Mid(wsKol_Smetka.Cells(Index, 2).Value, 1, 3)
                    If sss <> " - " Then
                        wsKol_Smetka.Cells(Index, 1).Value = Nomer
                        Nomer += 1
                    End If
                Else
                    wsKol_Smetka.Cells(Index, 1).Value = ""
                End If
                Select Case Len(wsKol_Smetka.Cells(Index, 2).Value)
                    Case <= 67
                        Red_RowHeight = 1
                    Case <= 125
                        Red_RowHeight = 2
                    Case > 125
                        Red_RowHeight = 3
                End Select
                wsKol_Smetka.Cells(Index, 2).RowHeight = Red_RowHeight * _RowHeight
                ProgressBar_Extrat.Value = Index
                Index += 1
            Loop Until wsKol_Smetka.Cells(Index, 2).Value = ""
        End With
        With wsKol_Smetka
            With .Range("A:D")
                .Font.Name = "Cambria"
                .Font.Size = 12
                .WrapText = vbTrue
                .VerticalAlignment = excel.XlVAlign.xlVAlignCenter
                .HorizontalAlignment = excel.XlHAlign.xlHAlignCenter
                .Font.Bold = vbFalse
                .Font.ColorIndex = excel.XlColorIndex.xlColorIndexAutomatic
                .Borders(excel.XlBordersIndex.xlDiagonalDown).LineStyle = excel.XlLineStyle.xlLineStyleNone
                .Borders(excel.XlBordersIndex.xlDiagonalUp).LineStyle = excel.XlLineStyle.xlLineStyleNone
            End With
            .Columns("A:A").ColumnWidth = 6
            .Columns("B:B").ColumnWidth = 67
            .Columns("C:C").ColumnWidth = 7
            .Columns("D:D").ColumnWidth = 7
            With .Range("A1:D1")
                .MergeCells = vbTrue
                .Font.Name = "Cambria"
                .Font.Size = 14
                .WrapText = vbTrue
                .VerticalAlignment = excel.XlVAlign.xlVAlignCenter
                .HorizontalAlignment = excel.XlHAlign.xlHAlignCenter
                .Font.Bold = vbTrue
                .Font.ColorIndex = excel.XlColorIndex.xlColorIndexAutomatic
            End With
            With .Range("A2:D2")
                .MergeCells = vbTrue
                .Font.Name = "Cambria"
                .Font.Size = 12
                .WrapText = vbTrue
                .VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Font.Bold = vbFalse
                .Characters(1, 6).Font.Bold = vbTrue
                .Font.ColorIndex = excel.XlColorIndex.xlColorIndexAutomatic
                Select Case Len(wsKol_Smetka.Cells(2, 1).Value)
                    Case <= 85
                        Red_RowHeight = 1
                    Case <= 170
                        Red_RowHeight = 2
                    Case <= 260
                        Red_RowHeight = 3
                    Case <= 350
                        Red_RowHeight = 4
                    Case <= 440
                        Red_RowHeight = 5
                    Case <= 520
                        Red_RowHeight = 6
                    Case > 520
                        Red_RowHeight = 7
                End Select
                .RowHeight = _RowHeight * Red_RowHeight
            End With
            With .Range("A4:D4")
                .Font.Size = 12
                .Font.Bold = vbTrue
            End With
            Call Excel_Kol_smetka_Razdel(wsKol_Smetka, "ЧАСТ ЕЛЕКТРО", "A5", "D5")
            .Range("A5:D5").VerticalAlignment = excel.XlVAlign.xlVAlignCenter
            .Range("A5:D5").HorizontalAlignment = excel.XlHAlign.xlHAlignCenter
            With .Range("A5:D5").Interior
                .Pattern = excel.XlPattern.xlPatternSolid
                .PatternColorIndex = 24
                .ThemeColor = excel.XlThemeColor.xlThemeColorAccent4
                .TintAndShade = 0.4
                .PatternTintAndShade = 0
            End With
            Rang = "A7:" & "A" & Trim(Str(Index))
            With .Range(Rang)
                .VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .HorizontalAlignment = excel.XlHAlign.xlHAlignCenter
            End With
            Rang = "B7:" & "B" & Trim(Str(Index))
            With .Range(Rang)
                .VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
            End With
            Rang = "C7:" & "C" & Trim(Str(Index))
            With .Range(Rang)
                .VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .HorizontalAlignment = excel.XlHAlign.xlHAlignCenter
            End With
            Rang = "D7:" & "D" & Trim(Str(Index))
            With .Range(Rang)
                .VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .HorizontalAlignment = excel.XlHAlign.xlHAlignRight
                .IndentLevel = 1
            End With
            Rang = "A7:" & "D" & Trim(Str(Index - 1))
            With .Range(Rang)
                With .Borders(excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = excel.XlLineStyle.xlContinuous
                    .ColorIndex = excel.XlColorIndex.xlColorIndexAutomatic
                    .TintAndShade = 0
                    .Weight = excel.XlBorderWeight.xlMedium
                End With
                With .Borders(excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = excel.XlLineStyle.xlContinuous
                    .ColorIndex = excel.XlColorIndex.xlColorIndexAutomatic
                    .TintAndShade = 0
                    .Weight = excel.XlBorderWeight.xlMedium
                End With
                With .Borders(excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = excel.XlLineStyle.xlContinuous
                    .ColorIndex = excel.XlColorIndex.xlColorIndexAutomatic
                    .TintAndShade = 0
                    .Weight = excel.XlBorderWeight.xlMedium
                End With
                With .Borders(excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = excel.XlLineStyle.xlContinuous
                    .ColorIndex = excel.XlColorIndex.xlColorIndexAutomatic
                    .TintAndShade = 0
                    .Weight = excel.XlBorderWeight.xlMedium
                End With
                With .Borders(excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = excel.XlLineStyle.xlDot
                    .ColorIndex = excel.XlColorIndex.xlColorIndexAutomatic
                    .TintAndShade = 0
                    .Weight = excel.XlBorderWeight.xlThin
                End With
                With .Borders(excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = excel.XlLineStyle.xlContinuous
                    .ColorIndex = excel.XlColorIndex.xlColorIndexAutomatic
                    .TintAndShade = 0
                    .Weight = excel.XlBorderWeight.xlThin
                End With
            End With
            Index += 1
            Rang = "A" & Trim(Str(Index)) & ":" & "D" & Trim(Str(Index))
            .Range(Rang).MergeCells = vbTrue
            Dim Text As String = ""
            Text = "Продуктите трябва да съответстват на европейските технически спецификации при"
            Text = Text + " спазване изискванията на Регламент № 305/2011 на Европейския парламент и"
            Text = Text + " на Съвета за определяне на хармонизирани условия за предлагане на пазара на"
            Text = Text + " строителни продукти и за отмяна на Директива 89/106/ЕИО на Съвета и"
            Text = Text + " чл. 5, ал. 1 от НСИСОССП – придружен с маркировка СЕ и с"
            Text = Text + " прилагане на декларация за експлоатационните показатели на продукта и"
            Text = Text + " указания за прилагане, изготвени на български език."
            With .Cells(Index, 1)
                .Value = Text
                .RowHeight = _RowHeight * 6
                .VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Font.Bold = vbFalse
                .Font.Size = 12
            End With
            Index += 2
            Text = "ЗАБЕЛЕЖКА: Навсякъде, където са посочени изрично конкретни продукти с конкретни търговски марки"
            Text = Text + " следва да разбира и оферира не задължително посочените продукти, а равностойни,"
            Text = Text + " еквивалентни, със същите или по-добри параметри от посочените,"
            Text = Text + " като се спазват всички изисквания на действащите нормативни документи и"
            Text = Text + " съответстват на решението на проектанта и избраната технология!"
            Rang = "A" & Trim(Str(Index)) & ":" & "D" & Trim(Str(Index))
            .Range(Rang).MergeCells = vbTrue
            Rang = "A" & Trim(Str(Index)) & ":" & "A" & Trim(Str(Index))
            With .Range(Rang)
                .Value = Text
                .RowHeight = _RowHeight * 5
                .VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Font.Bold = vbFalse
                .Font.Underline = excel.XlUnderlineStyle.xlUnderlineStyleNone
                .Font.Size = 12
                .Characters(1, 10).Font.Bold = vbTrue
                .Characters(1, 10).Font.Underline = excel.XlUnderlineStyle.xlUnderlineStyleSingle
            End With
            Index += 3
            Rang = "A" & Trim(Str(Index)) & ":" & "D" & Trim(Str(Index + 1))
            .Range(Rang).HorizontalAlignment = excel.XlHAlign.xlHAlignCenterAcrossSelection
            With .Cells(Index, 4)
                .Value = "Проектант: ......................................................."
                .RowHeight = _RowHeight
                .VerticalAlignment = excel.XlVAlign.xlVAlignCenter
                .HorizontalAlignment = excel.XlHAlign.xlHAlignRight
                .Font.Bold = vbFalse
                .Font.Size = 12
                .WrapText = vbFalse
            End With
            Index += 1
            With .Cells(Index, 4)
                .Value = "/" & ComboBox_Proektant.SelectedItem & "/"
                .RowHeight = _RowHeight
                .VerticalAlignment = excel.XlVAlign.xlVAlignCenter
                .HorizontalAlignment = excel.XlHAlign.xlHAlignRight
                .Font.Bold = vbFalse
                .WrapText = vbFalse
                .Font.Size = 12
            End With
            Index += 1
            ProgressBar_Extrat.Maximum = Index
            For i As Integer = 6 To Index
                If wsKol_Smetka.Cells(i, 2).Interior.ThemeColor =
                        excel.XlThemeColor.xlThemeColorAccent6 Then
                    If wsKol_Smetka.Cells(i, 1).Interior.ThemeColor <>
                            excel.XlThemeColor.xlThemeColorAccent6 Then
                        Call Excel_Kol_smetka_Razdel(
                            wsKol_Smetka,
                            wsKol_Smetka.Cells(i, 2).Value,
                            "B" & Trim(i.ToString),
                            "D" & Trim(i.ToString)
                            )
                    Else
                        Call Excel_Kol_smetka_Razdel(
                            wsKol_Smetka,
                            wsKol_Smetka.Cells(i, 2).Value,
                            "A" & Trim(i.ToString),
                            "D" & Trim(i.ToString)
                            )
                    End If
                End If
                ProgressBar_Extrat.Value = i
            Next
        End With
        ProgressBar_Extrat.Value = ProgressBar_Extrat.Minimum
    End Sub
    Private Sub ExcelUtilForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SaveToolStripButton.Enabled = vbFalse
        SplitContainer1.Enabled = vbFalse
        If InStr(Application.DocumentManager.MdiActiveDocument.Name, "_PV") Then
            File_PV = True
        End If
    End Sub
    Private Sub Button_Вземи_Интернет__Click(sender As Object, e As EventArgs) Handles Button_Вземи_РОЗЕТКИ1.Click
        If IsNothing(excel_Workbook) Then
            MsgBox("Първо да беше отворил файла?")
            Exit Sub
        End If
        Me.Visible = vbFalse
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim edt As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database
        Dim SelectedSet = cu.GetObjects("INSERT", "Изберете блок")
        Dim arrBlock(500) As strKontakt
        Dim Качване(100) As strКачване

        If SelectedSet Is Nothing Then
            MsgBox("Няма маркиран нито един блок.")
            Me.Visible = vbTrue
            Exit Sub
        End If
        Dim blkRecId As ObjectId = ObjectId.Null
        Dim index As Integer = 0
        Dim index_Качване As Integer = 0
        Dim index_Red As Integer = 0

        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                For Each sObj As SelectedObject In SelectedSet
                    blkRecId = sObj.ObjectId

                    Dim acBlkRef As BlockReference = DirectCast(acTrans.GetObject(blkRecId, OpenMode.ForRead), BlockReference)
                    Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                    Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection
                    Dim Visibility As String = ""
                    Dim strRACH_NAIMENO As String = ""
                    Dim strRACH_UNIT As String = ""
                    Dim strRACH_Wiso4ina As String = ""

                    Dim strKOTA_1 As String = ""
                    Dim strKOTA_2 As String = ""
                    Dim strТРЪБА_1 As String = ""
                    Dim strТРЪБА_2 As String = ""
                    Dim strKabel_d_0 As String = ""
                    Dim strKabel_d_1 As String = ""
                    Dim strKabel_d_2 As String = ""
                    Dim strKabel_d_3 As String = ""
                    Dim strKabel_d_4 As String = ""
                    Dim strKabel_d_5 As String = ""
                    Dim strKabel_d_6 As String = ""
                    Dim strKabel_d_7 As String = ""
                    Dim strKabel_d_8 As String = ""
                    Dim strKabel_d_9 As String = ""
                    Dim strKabel_d_10 As String = ""
                    Dim strKabel_g_0 As String = ""
                    Dim strKabel_g_1 As String = ""
                    Dim strKabel_g_2 As String = ""
                    Dim strKabel_g_3 As String = ""
                    Dim strKabel_g_4 As String = ""
                    Dim strKabel_g_5 As String = ""
                    Dim strKabel_g_6 As String = ""
                    Dim strKabel_g_7 As String = ""
                    Dim strKabel_g_8 As String = ""
                    Dim strKabel_g_9 As String = ""
                    Dim strKabel_g_10 As String = ""

                    For Each prop As DynamicBlockReferenceProperty In props
                        If prop.PropertyName = "Visibility" Then Visibility = prop.Value
                        If prop.PropertyName = "Розеток в ряд" Then Visibility = prop.Value
                        If prop.PropertyName = "В юнитах" Then strRACH_UNIT = prop.Value
                        If prop.PropertyName = "Высота" Then strRACH_Wiso4ina =
                            (Int(Val(prop.Value) / 44.45) + 1).ToString
                    Next

                    For Each objID As ObjectId In attCol
                        Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                        Dim acAttRef As AttributeReference = dbObj
                        If acAttRef.Tag = "НАИМЕНОВАНИЕ" Then strRACH_NAIMENO = acAttRef.TextString
                    Next

                    Dim blName As String = (CType(acBlkRef.DynamicBlockTableRecord.GetObject(OpenMode.ForRead), BlockTableRecord)).Name

                    Dim iVisib As Integer = -1
                    Select Case blName
                        Case "19'' шкаф"
                            iVisib = Array.FindIndex(arrBlock, Function(f) f.blName = blName And
                                                                           f.blRACH_UNIT = strRACH_UNIT)
                        Case "19'' шаблон"
                            iVisib = Array.FindIndex(arrBlock, Function(f) f.blName = blName And
                                                                           f.blRACH_Wiso4ina = strRACH_Wiso4ina)
                        Case "19'' CD player",
                             "19'' комм.панель RJ45",
                             "19'' Switch",
                             "Пач 24",
                             "Розетка_1",
                             "Табло_Ново",
                             "19'' блок розеток",
                             "19'' 24-Диска",
                             "19'' DVR",
                             "19'' вентиляторный модуль",
                             "19'' ДИН шина",
                             "19'' ИБП",
                             "19'' сервер",
                             "19'' Телефонна централа"
                            iVisib = Array.FindIndex(arrBlock, Function(f) f.blName = blName And
                                                                           f.blVisibility = Visibility)
                        Case "Кабел"
                            Continue For
                        Case "Качване"
                            For Each objID As ObjectId In attCol
                                Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                                Dim acAttRef As AttributeReference = dbObj
                                If acAttRef.Tag = "KOTA_1" Then strKOTA_1 = acAttRef.TextString
                                If acAttRef.Tag = "KOTA_2" Then strKOTA_2 = acAttRef.TextString
                                If acAttRef.Tag = "ТРЪБА_1" Then strТРЪБА_1 = acAttRef.TextString
                                If acAttRef.Tag = "ТРЪБА_2" Then strТРЪБА_2 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_0" Then strKabel_d_0 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_1" Then strKabel_d_1 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_2" Then strKabel_d_2 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_3" Then strKabel_d_3 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_4" Then strKabel_d_4 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_5" Then strKabel_d_5 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_6" Then strKabel_d_6 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_7" Then strKabel_d_7 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_8" Then strKabel_d_8 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_9" Then strKabel_d_9 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_10" Then strKabel_d_10 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_0" Then strKabel_g_0 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_1" Then strKabel_g_1 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_2" Then strKabel_g_2 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_3" Then strKabel_g_3 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_4" Then strKabel_g_4 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_5" Then strKabel_g_5 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_6" Then strKabel_g_6 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_7" Then strKabel_g_7 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_8" Then strKabel_g_8 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_9" Then strKabel_g_9 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_10" Then strKabel_g_10 = acAttRef.TextString
                            Next
                        Case Else
                            Continue For
                    End Select
                    If iVisib = -1 Then
                        arrBlock(index).count = 1
                        arrBlock(index).blName = blName
                        arrBlock(index).blVisibility = Visibility
                        arrBlock(index).blRACH_NAIMENO = strRACH_NAIMENO
                        arrBlock(index).blRACH_UNIT = strRACH_UNIT
                        arrBlock(index).blRACH_Wiso4ina = strRACH_Wiso4ina
                        If blName = "Качване" Then
                            Качване(index_Качване).KOTA_1 = strKOTA_1
                            Качване(index_Качване).KOTA_2 = strKOTA_2
                            Качване(index_Качване).ТРЪБА_1 = strТРЪБА_1
                            Качване(index_Качване).ТРЪБА_2 = strТРЪБА_2
                            Качване(index_Качване).Kabel_d_0 = strKabel_d_0
                            Качване(index_Качване).Kabel_d_1 = strKabel_d_1
                            Качване(index_Качване).Kabel_d_2 = strKabel_d_2
                            Качване(index_Качване).Kabel_d_3 = strKabel_d_3
                            Качване(index_Качване).Kabel_d_6 = strKabel_d_6
                            Качване(index_Качване).Kabel_d_7 = strKabel_d_7
                            Качване(index_Качване).Kabel_d_4 = strKabel_d_4
                            Качване(index_Качване).Kabel_d_5 = strKabel_d_5
                            Качване(index_Качване).Kabel_d_8 = strKabel_d_8
                            Качване(index_Качване).Kabel_d_9 = strKabel_d_9
                            Качване(index_Качване).Kabel_d_10 = strKabel_d_10
                            Качване(index_Качване).Kabel_g_0 = strKabel_g_0
                            Качване(index_Качване).Kabel_g_1 = strKabel_g_1
                            Качване(index_Качване).Kabel_g_2 = strKabel_g_2
                            Качване(index_Качване).Kabel_g_3 = strKabel_g_3
                            Качване(index_Качване).Kabel_g_4 = strKabel_g_4
                            Качване(index_Качване).Kabel_g_5 = strKabel_g_5
                            Качване(index_Качване).Kabel_g_6 = strKabel_g_6
                            Качване(index_Качване).Kabel_g_7 = strKabel_g_7
                            Качване(index_Качване).Kabel_g_8 = strKabel_g_8
                            Качване(index_Качване).Kabel_g_9 = strKabel_g_9
                            Качване(index_Качване).Kabel_g_10 = strKabel_g_10
                            index_Качване += 1
                        End If
                        index += 1
                    Else
                        arrBlock(iVisib).count = arrBlock(iVisib).count + 1
                    End If
                Next
                index_Качване = 0
                wsLines.Cells(2, 35).Value = 0
                wsLines.Cells(2, 36).Value = 0
                wsLines.Cells(2, 37).Value = 0
                index_Red = 2
                Dim dylvina1 As Double = 0
                Dim dylvina2 As Double = 0
                Dim br_linii As Double = 0
                Dim Kabel_Internet As Double = 0
                Dim Kabel_Televiziq As Double = 0
                Dim Kabel_telefon As Double = 0

                For Each iarrBlock In arrBlock
                    If iarrBlock.count = 0 Then Exit For
                    With wsInternet
                        .Cells(index_Red, 1) = iarrBlock.count
                        .Cells(index_Red, 2) = iarrBlock.blName
                        .Cells(index_Red, 3) = iarrBlock.blVisibility
                        .Cells(index_Red, 4) = iarrBlock.blRACH_NAIMENO
                        .Cells(index_Red, 5) = iarrBlock.blRACH_UNIT
                        .Cells(index_Red, 6) = iarrBlock.blRACH_Wiso4ina
                        If iarrBlock.blName = "Качване" Then
                            .Cells(index_Red, 11) = Качване(index_Качване).KOTA_1
                            .Cells(index_Red, 12) = Качване(index_Качване).ТРЪБА_1
                            .Cells(index_Red, 13) = Качване(index_Качване).Kabel_d_0
                            .Cells(index_Red, 14) = Качване(index_Качване).Kabel_d_1
                            .Cells(index_Red, 15) = Качване(index_Качване).Kabel_d_2
                            .Cells(index_Red, 16) = Качване(index_Качване).Kabel_d_3
                            .Cells(index_Red, 17) = Качване(index_Качване).Kabel_d_6
                            .Cells(index_Red, 18) = Качване(index_Качване).Kabel_d_5
                            .Cells(index_Red, 19) = Качване(index_Качване).Kabel_d_8
                            .Cells(index_Red, 20) = Качване(index_Качване).Kabel_d_7
                            .Cells(index_Red, 21) = Качване(index_Качване).Kabel_d_4
                            .Cells(index_Red, 22) = Качване(index_Качване).Kabel_d_9
                            .Cells(index_Red, 23) = Качване(index_Качване).Kabel_d_10
                            .Cells(index_Red, 24) = Качване(index_Качване).KOTA_2
                            .Cells(index_Red, 25) = Качване(index_Качване).ТРЪБА_2
                            .Cells(index_Red, 26) = Качване(index_Качване).Kabel_g_0
                            .Cells(index_Red, 27) = Качване(index_Качване).Kabel_g_1
                            .Cells(index_Red, 28) = Качване(index_Качване).Kabel_g_2
                            .Cells(index_Red, 29) = Качване(index_Качване).Kabel_g_3
                            .Cells(index_Red, 30) = Качване(index_Качване).Kabel_g_4
                            .Cells(index_Red, 31) = Качване(index_Качване).Kabel_g_5
                            .Cells(index_Red, 32) = Качване(index_Качване).Kabel_g_6
                            .Cells(index_Red, 33) = Качване(index_Качване).Kabel_g_7
                            .Cells(index_Red, 34) = Качване(index_Качване).Kabel_g_8
                            .Cells(index_Red, 35) = Качване(index_Качване).Kabel_g_9
                            .Cells(index_Red, 36) = Качване(index_Качване).Kabel_g_10
                            index_Качване += 1

                            dylvina1 = Val(.Cells(index_Red, 11).value)
                            dylvina2 = Val(.Cells(index_Red, 24).value)
                            For i = 11 To 36

                                br_linii = InStr(.Cells(index_Red, i).value, "л.")
                                If br_linii = 0 Then Continue For
                                br_linii = Mid(.Cells(index_Red, i).value, 1, br_linii - 1)
                                If InStr(.Cells(index_Red, i).value, "RG6/64") > 1 Then
                                    .Cells(index_Red, 9).value = IIf(i < 24, br_linii * dylvina1, br_linii * dylvina2) + .Cells(index_Red, 9).value
                                    Kabel_Televiziq = Kabel_Televiziq + .Cells(index_Red, 9).value
                                End If
                                If InStr(.Cells(index_Red, i).value, "FTP") > 1 Then
                                    .Cells(index_Red, 8).value = IIf(i < 24, br_linii * dylvina1, br_linii * dylvina2) + .Cells(index_Red, 8).value
                                    Kabel_Internet = Kabel_Internet + .Cells(index_Red, 8).value
                                End If
                                If InStr(.Cells(index_Red, i).value, "ПТПВ") > 1 Then
                                    .Cells(index_Red, 10).value = IIf(i < 24, br_linii * dylvina1, br_linii * dylvina2) + .Cells(index_Red, 10).value
                                    Kabel_telefon = Kabel_telefon + .Cells(index_Red, 10).value
                                End If
                            Next
                        End If
                        If iarrBlock.blName = "Розетка_1" And iarrBlock.blVisibility = "TV" Then
                            wsLines.Cells(2, 30).Value = iarrBlock.count
                        End If
                        If iarrBlock.blName = "Розетка_1" And iarrBlock.blVisibility = "@" Then
                            wsLines.Cells(2, 31).Value = iarrBlock.count
                        End If
                        If iarrBlock.blName = "Розетка_1" And iarrBlock.blVisibility = "WiFi" Then
                            wsLines.Cells(2, 32).Value = iarrBlock.count
                        End If
                        If iarrBlock.blName = "Розетка_1" And iarrBlock.blVisibility = "T" Then
                            wsLines.Cells(2, 34).Value = iarrBlock.count
                        End If
                        If iarrBlock.blName = "Табло_Ново" And iarrBlock.blVisibility = "Слаботоково табло" Then
                            wsLines.Cells(2, 33).Value = iarrBlock.count
                        End If
                    End With
                    index_Red += 1
                Next

                wsLines.Cells(2, 36).Value = Kabel_Televiziq / 2
                wsLines.Cells(2, 35).Value = Kabel_Internet / 2
                wsLines.Cells(2, 37).Value = Kabel_telefon / 2
                acTrans.Commit()
                With wsInternet
                    .Cells.Sort(Key1:= .Range("B2"),
                                Order1:=excel.XlSortOrder.xlAscending,
                                Header:=excel.XlYesNoGuess.xlYes,
                                OrderCustom:=1, MatchCase:=False,
                                Orientation:=excel.Constants.xlTopToBottom,
                                DataOption1:=excel.XlSortDataOption.xlSortTextAsNumbers,
                                Key2:= .Range("D2"),
                                Order2:=excel.XlSortOrder.xlAscending,
                                DataOption2:=excel.XlSortDataOption.xlSortTextAsNumbers
                                )
                End With
            Catch ex As Exception
                MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
                acTrans.Abort()
            End Try
        End Using

        Me.Visible = vbTrue
    End Sub
    Private Sub Button_Генерирай_Интернет_Click(sender As Object, e As EventArgs) Handles Button_Генерирай_Интернет.Click
        Dim index As Integer = 0
        Dim index_internet As Integer = 0
        Dim index_Kabel As Integer = 0
        Dim i As Integer
        ProgressBar_Extrat.Maximum = ProgressBar_Maximum
        For i = 6 To 1000
            If wsKol_Smetka.Cells(i, 2).Value = "Раздел" Or wsKol_Smetka.Cells(i, 2).Value = "" Then
                index = i
                Exit For
            End If
            ProgressBar_Extrat.Value = i
        Next
        For i = 2 To 1000
            If wsInternet.Cells(i, 2).Value = "" Then
                Exit For
            End If
        Next
        index_internet = i
        ProgressBar_Extrat.Maximum = i
        Dim Text_Dostawka As String = ""
        Dim broj_elementi As Integer = 0
        Dim broj_TV As Integer = 0
        Dim broj_In As Integer = 0
        Dim broj_Tel As Integer = 0
        Dim broj_HDMI As Integer = 0
        Dim broj_USB As Integer = 0

        Call Excel_Kol_smetka_Razdel(wsKol_Smetka, "СЛАБОТОКОВИ ИНСТАЛАЦИИ", "A" & Trim(index.ToString), "D" & Trim(index.ToString))
        index += 1
        Call Excel_Kol_smetka_Razdel(wsKol_Smetka, "ИНСТАЛAЦИЯ ИНТЕРНЕТ И ТЕЛЕВИЗИЯ", "B" & Trim(index.ToString), "D" & Trim(index.ToString))
        Dim Kabel(1000) As strKabel
        index += 1
        For i = 2 To 10000
            If Trim(wsInternet.Cells(i, 2).Value) = "" Then Exit For
            ProgressBar_Extrat.Value = i
            Text_Dostawka = ""
            broj_elementi = wsInternet.Cells(i, 1).Value
            If wsInternet.Cells(i, 2).Value = "Табло_Ново" Then
                With wsKol_Smetka
                    .Cells(index, 2).Value = "Доставка и монтаж на слаботоково (комуникациионно) табло; за вграждане; 2 реда"
                    .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                    .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                    .Cells(index, 3).Value = "бр."
                    .Cells(index, 4).Value = broj_elementi
                End With
                index += 1
                With wsKol_Smetka
                    .Cells(index, 2).Value = "Направа нишa за слаботоково таблo"
                    .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                    .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                    .Cells(index, 3).Value = "бр."
                    .Cells(index, 4).Value = broj_elementi
                End With
                index += 1
                Exit For
            End If
        Next
        For i = 2 To 10000
            If Trim(wsInternet.Cells(i, 2).Value) = "" Then Exit For
            ProgressBar_Extrat.Value = i
            Text_Dostawka = ""
            broj_elementi = wsInternet.Cells(i, 1).Value
            Select Case wsInternet.Cells(i, 2).Value
                Case "Розетка_1"
                    Select Case wsInternet.Cells(i, 3).Value
                        Case "WiFi"
                            Text_Dostawka = "безжичен рутер"
                            broj_In = broj_In + broj_elementi ' Натрупва интернет розетките
                        Case "TV"
                            Text_Dostawka = "телевизионна розетка с 1 гнездо; за вграждане"
                            broj_TV = broj_TV + broj_elementi ' Натрупва телевизионните розетки
                        Case "T"
                            Text_Dostawka = "телефонна розетка с 1 гнездо; букса RJ11; за вграждане"
                            broj_Tel = broj_Tel + broj_elementi ' Натрупва телефонните розетки
                        Case "@"
                            broj_In = broj_In + broj_elementi ' Натрупва интернет розетките
                        Case "HDMI"
                            Text_Dostawka = "HDMI розетка с 1 гнездо; букса HDMI; за вграждане"
                            broj_HDMI = broj_HDMI + broj_elementi ' Натрупва HDMI розетките
                        Case "USB"
                            Text_Dostawka = "USB розетка с 1 гнездо; букса USB; за вграждане"
                            broj_USB = broj_USB + broj_elementi ' Натрупва USB розетките
                        Case "Access point"
                            Text_Dostawka = "Wi-Fi точка на достъп (AP)"
                            broj_In = broj_In + broj_elementi ' Натрупва интернет розетките
                        Case "Router"
                            Text_Dostawka = "рутер (без Wi-Fi)"
                            broj_In = broj_In + broj_elementi ' Натрупва интернет розетките
                    End Select
                Case "КАБЕЛ"
                    Kabel(index_Kabel).blType = wsInternet.Cells(i, 3).Value
                    Kabel(index_Kabel).blPol = wsInternet.Cells(i, 4).Value
                    Kabel(index_Kabel).blLength = wsInternet.Cells(i, 5).Value
                    index_Kabel += 1
                Case "Качване", "19'' 24-Диска", "19'' CD player", "19'' DVR",
                     "19'' Switch", "19'' Телефонна централа", "19'' сервер",
                     "19'' заземител", "19'' блок розеток", "19'' вентиляторный модуль",
                     "19'' ДИН шина", "19'' ИБП", "19'' комм.панель RJ45", "19'' шаблон",
                     "19'' шкаф", "Табло_Ново"
                    Continue For
                Case Else
                    Text_Dostawka = "НЕИЗВЕСТЕН ЕЛЕМЕНТ --> " & wsInternet.Cells(i, 2).Value & " - " & wsInternet.Cells(i, 4).Value
            End Select
            If Trim(Text_Dostawka) <> "" Then
                index = AddItemToSheet(index,
                               "Доставка и монтаж на",
                               Text_Dostawka,
                               broj_elementi)
            End If
        Next
        index = Kol_Smetka_Kabeli(index, vbFalse, "ИНТЕРНЕТ")
        If broj_TV > 0 Then
            index = AddItemToSheet(index,
               "Доставка и монтаж на",
               "Сплитер 1 вход - 6 изхода, за кабелни системи",
               broj_TV)
        End If
        If broj_In > 0 Then
            index = AddItemToSheet(index,
               "Доставка и монтаж на",
               "интернет розетка с 1 гнездо; букса RJ45; за вграждане; категория: 5e",
               broj_In)
        End If
        If (broj_TV + broj_In + broj_Tel + broj_HDMI + broj_USB) > 0 Then
            index = AddItemToSheet(index,
               "Направа на отвор в стената за монтаж на розетки",
               "",
               broj_TV + broj_In + broj_Tel + broj_HDMI + broj_USB)
        End If
        If (broj_TV + broj_In + broj_Tel + broj_HDMI + broj_USB) > 0 Then
            index = AddItemToSheet(index,
               "Доставка и монтаж на единична конзола",
               "",
               broj_TV + broj_In + broj_Tel + broj_HDMI + broj_USB)
        End If
        ProgressBar_Extrat.Value = ProgressBar_Extrat.Minimum
        For i = 2 To 10000
            If Trim(wsInternet.Cells(i, 2).Value) = "" Then Exit Sub
            If wsInternet.Cells(i, 2).Value = "19'' шкаф" Then Exit For
        Next
        Call Excel_Kol_smetka_Razdel(wsKol_Smetka, "АКТИВНО ОБОРУДВАНЕ НА RACK ШКАФ", "B" & Trim(index.ToString), "D" & Trim(index.ToString))
        index += 1
        For i = 2 To 10000
            If Trim(wsInternet.Cells(i, 2).Value) = "" Then Exit For
            ProgressBar_Extrat.Value = i
            Text_Dostawka = ""
            broj_elementi = wsInternet.Cells(i, 1).Value
            Select Case wsInternet.Cells(i, 2).Value
                Case "19'' 24-Диска"
                    Text_Dostawka = "мрежово хранилище за данни NAS; за монтаж в 19'' RACK шкаф;" &
                        "поддържа до 12 диска 2,5''/3,5'' SATA;"
                Case "19'' CD player"
                    Select Case wsInternet.Cells(i, 3).Value
                        Case "Заряден токоизправител"
                        Case "Оповестителна центала VM 3240-VA"
                            Text_Dostawka = "цифров контролер за оповестяване" & vbCrLf &
                                            " - сертифициран по EN54" & vbCrLf &
                                            " - 6 зонален усилвател 240W;" & vbCrLf &
                                            " - сертифициран по EN54" & vbCrLf &
                                            " - Номинална изходяща мощност: 240W;" & vbCrLf &
                                            " - Честотна характеристика  50 – 20000Hz. ± 3Db (при 1/3 от номиналната мощност);" & vbCrLf &
                                            " - Памет за съобщения : 64МВ, 48kHz"
                        Case "Усилвател VM3240E"
                            Text_Dostawka = "усилвател" & vbCrLf &
                                            " - сертифициран по EN54" & vbCrLf &
                                            " - Номинална изходяща мощност: 240W;" & vbCrLf &
                                            " - Честотна характеристика  50 – 20000Hz. ± 3Db (при 1/3 от номиналната мощност);"
                        Case "Усилвател VP - 2241"
                            Text_Dostawka = "усилвател 240 W x 1 канал" & vbCrLf &
                                            " - изходяща мощност 240W;" & vbCrLf &
                                            " - сертифициран по EN54"
                        Case "CD Player"
                            Text_Dostawka = "източник на звук" & vbCrLf &
                                            " - DVD, CD, Mp3, USB;" & vbCrLf &
                                            " - сертифициран по EN54"
                    End Select
                Case "19'' DVR"
                    Text_Dostawka = "16 канален мрежов видеорекордер; поддържа 2бр. SATA диска до 6 TB всеки"
                Case "19'' Switch"
                    Text_Dostawka = "суич; " & wsInternet.Cells(i, 3).Value & "; 10/100/1000 Mbps; за монтаж в комуникационен шкаф"
                Case "19'' Телефонна централа"
                    Text_Dostawka = "IP телефонна централа; за монтаж в 19'' комуникационен шкаф"
                Case "19'' сервер"
                    Text_Dostawka = "сървър; за монтаж в 19'' комуникационен шкаф"
                Case "Качване", "Розетка_1", "Табло_Ново", "ELEKTRO",
                     "коакс.кабел RG6/64", "FTP 4x2х24AWG", "ПТПВ 1х2х0,5мм²",
                     "19'' заземител", "19'' блок розеток", "19'' вентиляторный модуль",
                     "19'' ДИН шина", "19'' ИБП", "19'' комм.панель RJ45", "19'' шаблон",
                     "19'' шкаф"
                    Continue For
                Case Else
                    Text_Dostawka = "НЕИЗВЕСТЕН ЕЛЕМЕНТ --> " & wsInternet.Cells(i, 2).Value & " - " & wsInternet.Cells(i, 4).Value
            End Select
            If Trim(Text_Dostawka) <> "" Then
                With wsKol_Smetka
                    .Cells(index, 2).Value = "Доставка и монтаж на " + Text_Dostawka
                    .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                    .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                    .Cells(index, 3).Value = "бр."
                    .Cells(index, 4).Value = broj_elementi
                End With
                index += 1
            End If
        Next
        Call Excel_Kol_smetka_Razdel(wsKol_Smetka, "ПАСИВНО ОБОРУДВАНЕ НА RACK ШКАФ", "B" & Trim(index.ToString), "D" & Trim(index.ToString))
        index += 1
        Dim Aranv_hor As Integer = 0
        Dim Aranv_vert As Integer = 0
        For i = 2 To 10000
            If Trim(wsInternet.Cells(i, 2).Value) = "" Then Exit For
            ProgressBar_Extrat.Value = i
            Text_Dostawka = ""
            broj_elementi = wsInternet.Cells(i, 1).Value
            Select Case wsInternet.Cells(i, 2).Value
                Case "19'' 24-Диска"
                    Text_Dostawka = "твърд диск за вграждане във мрежово хранилище за данни NAS:" & vbCrLf &
                        "- капацитет: 4TB"
                    broj_elementi = 4
                Case "19'' заземител"
                    Text_Dostawka = "комплект за заземяване:" & vbCrLf &
                        "- 10 точки"
                Case "19'' CD player"
                    Select Case wsInternet.Cells(i, 3).Value
                        Case "Заряден токоизправител"
                            Text_Dostawka = "заряден токоизправител:" & vbCrLf &
                                            " - сертифициран по EN54" & vbCrLf &
                                            " - захранване: 220V/AC;" & vbCrLf &
                                            " - изходно напрежежие:27,3V/DC;"
                            With wsKol_Smetka
                                .Cells(index, 2).Value = "Доставка и монтаж на " + Text_Dostawka
                                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                                .Cells(index, 3).Value = "бр."
                                .Cells(index, 4).Value = broj_elementi
                            End With
                            index += 1
                            broj_elementi = broj_elementi * 2
                            Text_Dostawka = "акумолаторна батерия:" & vbCrLf &
                                "- 12V; 65Ah;" & vbCrLf &
                                "- сертифициран по EN54, препоръчана от производителя на апаратурата"
                        Case "Оповестителна центала VM 3240-VA", "Усилвател VM3240E",
                             "Усилвател VP - 2241", "CD Player"
                    End Select
                Case "19'' DVR"
                    Text_Dostawka = "монитор: " & vbCrLf &
                                    "- Размер на екрана: 21.5'' - 54.6 см" & vbCrLf &
                                    "- Резолюция: 1920x1080 при 75 Hz;" & vbCrLf &
                                    "- Съотношение на картината: 16:9;" & vbCrLf &
                                    "- Видео входове: аналогов VGA и цифров HDMI;" & vbCrLf &
                                    "- Комплект: HDMI кабел, захранващ кабел."
                    With wsKol_Smetka
                        .Cells(index, 2).Value = "Доставка и монтаж на " + Text_Dostawka
                        .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                        .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                        .Cells(index, 3).Value = "бр."
                        .Cells(index, 4).Value = broj_elementi
                    End With
                    index += 1
                    broj_elementi = broj_elementi * 2
                    Text_Dostawka = "твърд диск за вграждане във видеорекордер:" & vbCrLf &
                                    "- капацитет: 6 TB"
                Case "19'' блок розеток"
                    Text_Dostawka = "разклонител с филтър:" & vbCrLf &
                        "- за монтаж в 19'' RACK шкаф;" & vbCrLf &
                        "- " & wsInternet.Cells(i, 3).Value & "x230V контакта тип Шуко;" & vbCrLf &
                        "- с ключ и защита от пренапрежение"
                Case "19'' вентиляторный модуль"
                    Text_Dostawka = "вентилаторен модул:" & vbCrLf &
                        "- за монтаж в 19'' RACK шкаф;" & vbCrLf &
                        "- 4 вентилатора; термостат"
                Case "19'' ДИН шина"
                    Text_Dostawka = "панел с DIN шина и заден покривен капак;"
                Case "19'' ИБП"
                    Text_Dostawka = "непрекъсваемо токозахранващ източник (UPS):" & vbCrLf &
                        "- капацитет: 6000VA/4800W"
                    With wsKol_Smetka
                        .Cells(index, 2).Value = "Доставка и монтаж на " + Text_Dostawka
                        .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                        .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                        .Cells(index, 3).Value = "бр."
                        .Cells(index, 4).Value = broj_elementi
                    End With
                    index += 1
                    Text_Dostawka = "комплект планки за монтаж на UPS в 19'' комуникационен шкаф"
                Case "19'' комм.панель RJ45"
                    Text_Dostawka = "пач панел: " & wsInternet.Cells(i, 3).Value & ";" & vbCrLf &
                        "- екраниран;" & vbCrLf &
                        "- категория: 6;" & vbCrLf &
                        "- оборудван"
                    Aranv_hor = Aranv_hor + broj_elementi
                Case "19'' шаблон"
                    Text_Dostawka = "разделителен панел; " & wsInternet.Cells(i, 6).Value & "U"
                Case "19'' шкаф"
                    Text_Dostawka = "19'' комуникационен шкаф (RACK); " & wsInternet.Cells(i, 5).Value & "U"
                    Aranv_vert = Aranv_vert + wsInternet.Cells(i, 5).Value / 4
                Case "Качване", "Розетка_1", "Табло_Ново", "ELEKTRO",
                     "коакс.кабел RG6/64", "FTP 4x2х24AWG", "ПТПВ 1х2х0,5мм²",
                     "19'' Switch", "19'' Телефонна централа", "19'' сервер",
                     "19'' 24-Диска"
                    Continue For
                Case Else
                    Text_Dostawka = "НЕИЗВЕСТЕН ЕЛЕМЕНТ --> " & wsInternet.Cells(i, 2).Value & " - " & wsInternet.Cells(i, 4).Value
            End Select
            If Trim(Text_Dostawka) <> "" Then
                With wsKol_Smetka
                    .Cells(index, 2).Value = "Доставка и монтаж на " + Text_Dostawka
                    .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                    .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                    .Cells(index, 3).Value = "бр."
                    .Cells(index, 4).Value = broj_elementi
                End With
                index += 1
            End If
        Next
        With wsKol_Smetka
            .Cells(index, 2).Value = "Доставка и монтаж на аранжиращ панел; хоризонтален"
            .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
            .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
            .Cells(index, 3).Value = "бр."
            .Cells(index, 4).Value = Aranv_hor
        End With
        index += 1
        With wsKol_Smetka
            .Cells(index, 2).Value = "Доставка и монтаж на аранжиращa скоба"
            .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
            .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
            .Cells(index, 3).Value = "бр."
            .Cells(index, 4).Value = Aranv_hor
        End With
        index += 1
        With wsKol_Smetka
            .Cells(index, 2).Value = "Комплексно изпитване на системата"
            .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
            .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
            .Cells(index, 3).Value = "бр."
            .Cells(index, 4).Value = 1
        End With
        index += 1
        With wsKol_Smetka
            .Cells(index, 2).Value = "Обучение на персонал за работа със системата"
            .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
            .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
            .Cells(index, 3).Value = "бр."
            .Cells(index, 4).Value = 1
        End With
        index += 1
        ProgressBar_Extrat.Value = ProgressBar_Extrat.Minimum
    End Sub
    ''' <summary>
    ''' Добавя ред в количествената сметка (Excel) за определен елемент.
    ''' </summary>
    ''' <param name="currentRow">Текущият индекс на реда в Excel. Функцията връща увеличен индекс.</param>
    ''' <param name="prefixText">Текст, който се поставя пред описанието на елемента (например "Доставка и монтаж на").</param>
    ''' <param name="itemDescription">Текстовото описание на елемента.</param>
    ''' <param name="quantity">Броят на елементите.</param>
    ''' <returns>Връща новия индекс (след добавянето на реда).</returns>
    Function AddItemToSheet(currentRow As Integer,
                        prefixText As String,
                        itemDescription As String,
                        quantity As Integer) As Integer
        ' Добавяме ред в Excel (wsKol_Smetka)
        With wsKol_Smetka
            ' Колона 2: Комбинираме текста и описанието на елемента
            .Cells(currentRow, 2).Value = prefixText & " " & itemDescription
            .Cells(currentRow, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
            .Cells(currentRow, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
            ' Колона 3: Единица (брой)
            .Cells(currentRow, 3).Value = "бр."
            ' Колона 4: Количество
            .Cells(currentRow, 4).Value = quantity
        End With
        ' Увеличаваме индекса за следващия ред
        currentRow += 1
        ' Връщаме новия индекс
        Return currentRow
    End Function
    Private Sub Button_Изчисти_Интернет_Click(sender As Object, e As EventArgs) Handles Button_Изчисти_Интернет.Click
        clearInternet(wsInternet)
    End Sub
    Private Sub Button_Вземи_КАБЕЛИ_Click(sender As Object, e As EventArgs) Handles Button_Слаботокови_КАБЕЛИ.Click,
                                                                                    Button_силова_КАБЕЛИ.Click,
                                                                                    Button_Вземи_ПИЦ_КАБЕЛИ.Click,
                                                                                    Button_Вземи_ДОМОФ_КАБЕЛИ.Click,
                                                                                    Button_МЪЛНИЯ_КАБЕЛИ.Click,
                                                                                    Button_ВЪНШНО_КАБЕЛИ.Click,
                                                                                    Button_Вземи_ВИДЕО_КАБЕЛИ.Click,
                                                                                    Button_Вземи_СОТ_КАБЕЛИ.Click,
                                                                                    Button_ФОТОВОЛТАИЦИ_КАБЕЛИ.Click
        Me.Visible = vbFalse
        Dim cu As CommonUtil = New CommonUtil()
        Dim ss = cu.GetObjects("LINE", "Изберете Линия")
        Dim i, index As Integer
        Dim Инсталация As String = ""
        Me.Visible = vbTrue
        If ss Is Nothing Then
            MsgBox("Няма маркиран линия в слой 'EL'.")
            Exit Sub
        End If
        Dim Kabel(1000, 2) As String

        Kabel = cu.GET_LINE_TYPE_KABEL(Kabel, ss, vbFalse)
        For index = 2 To 1000
            If wsLines.Cells(index, 1).Value = "" Then Exit For
        Next
        Dim masType() As strKabel = Kol_smetka_Kabel(Kabel, vbTrue)
        Dim formu As String = ""
        Dim rang As String = ""
        Dim colum As String = ""
        Dim Kabel_fec As Double = 1
        Dim boEL_2x1_5 As Boolean = True
        Dim boEL_4x1_5 As Boolean = True

        For i = 0 To UBound(masType)
            If masType(i).blLength = 0 Then Exit For
            If masType(i).blType = "EL__DIM" Or
               masType(i).blType = "ELEKTRO" Or
               masType(i).blType = "EL-DIM" Then
                Continue For
            End If
            If masType(i).blType = "СВТ2x1,5mm²" Then
                boEL_2x1_5 = False
            End If
            If masType(i).blType = "СВТ4x1,5mm²" Then
                boEL_4x1_5 = False
            End If
            '
            formu = ""
            rang = ""
            colum = ""
            '
            Select Case sender.name
                Case "Button_ФОТОВОЛТАИЦИ_КАБЕЛИ"
                    '
                    'Кабели за ВЪНШНО захранване
                    '
                    rang = "ФОТОВОЛТАИЦИ"
                    Инсталация = "ФОТОВОЛТАИЦИ"
                    Kabel_fec = 10
                Case "Button_ВЪНШНО_КАБЕЛИ"
                    '
                    'Кабели за ВЪНШНО захранване
                    '
                    rang = "ВЪНШНО"
                    Инсталация = "ВЪНШНО"
                Case "Button_силова_КАБЕЛИ"
                    '                      
                    ' СИЛОВИ КАБЕЛИ
                    '
                    Инсталация = "СИЛОВА"
                    Select Case masType(i).blType
                        Case "H1Z2Z2-K 1/1.5kV 1x4.0мм²"
                            rang = colum & Trim(index.ToString)
                        Case "СВТ2x1,5mm²"
                            colum = "$H$"
                            formu = "=" & colum & "2*" & Кабел_кутия.ToString &
                                    "+" & colum & "2*" & Str(H_Етаж - H_Ключ) &
                                    "+" & colum & "2*" & Str(Кабел_Розетка)
                            rang = colum & Trim(index.ToString)
                            wsLines.Range(rang).Formula = formu
                        Case "СВТ3x1,5mm²"
                            colum = "$I$"
                            formu = "=" & colum & "2*" & Кабел_кутия.ToString &
                                    "+" & colum & "2*" & Str(H_Етаж - H_Ключ) &
                                    "+" & colum & "2*" & Str(3 * Кабел_Розетка)
                            rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                            wsLines.Range(rang).Formula = formu
                            colum = "$L$"
                            formu = "=" & colum & "2*" & Кабел_кутия.ToString
                            rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                            wsLines.Range(rang).Formula = formu
                            colum = "$N$"
                            formu = "=" & colum & "2*" & Кабел_кутия.ToString
                            rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                            wsLines.Range(rang).Formula = formu
                        Case "СВТ3x2,5mm²"
                            colum = "$P$"
                            formu = "=" & colum & "2*" & Кабел_кутия.ToString &
                                    "+" & colum & "2*" & Str(H_Етаж - H_Контакт) &
                                    "+" & colum & "2*" & Str(3 * Кабел_Розетка)
                            rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                            wsLines.Range(rang).Formula = formu
                            colum = "$T$"
                            formu = "=" & colum & "2*" & Кабел_кутия.ToString & "+" & colum & "2*" & Str(H_Етаж - H_Контакт)
                            rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                            wsLines.Range(rang).Formula = formu
                        Case "СВТ3x4,0mm²"
                            colum = "$U$"
                            formu = "=" & colum & "2*" & Кабел_кутия.ToString &
                                    "+" & colum & "2*" & Str(H_Етаж - H_Контакт)
                            rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                            wsLines.Range(rang).Formula = formu
                            colum = "$R$"
                            formu = "=" & colum & "2*" & Кабел_кутия.ToString &
                                    "+" & colum & "2*" & Str(H_Етаж - H_Контакт)
                            rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                            wsLines.Range(rang).Formula = formu
                            colum = "$AB$"
                            formu = "=" & colum & "2*" & Кабел_кутия.ToString &
                                    "+" & colum & "2*" & Str(H_Етаж - H_Ключ)
                            rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                            wsLines.Range(rang).Formula = formu
                        Case "СВТ4x1,5mm²"
                            colum = "$J$"
                            formu = "=" & colum & "2*" & Кабел_кутия.ToString &
                                    "+" & colum & "2*" & Str(H_Етаж - H_Ключ) &
                                    "+" & colum & "2*" & Str(3 * Кабел_Розетка)
                            rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                            wsLines.Range(rang).Formula = formu
                            colum = "$M$"
                            formu = "=" & colum & "2*" & Кабел_кутия.ToString
                            rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                            wsLines.Range(rang).Formula = formu
                        Case "СВТ5x1,5mm²"
                            colum = "$K$"
                            formu = "=" & colum & "2*" & Кабел_кутия.ToString &
                                    "+" & colum & "2*" & Str(H_Етаж - H_Ключ) &
                                    "+" & colum & "2*" & Str(3 * Кабел_Розетка)
                            rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                            wsLines.Range(rang).Formula = formu
                        Case "СВТ5x2,5mm²"
                            colum = "$Q$"
                            formu = "=" & colum & "2*" & Кабел_кутия.ToString &
                                    "+" & colum & "2*" & Str(H_Етаж - H_Контакт) &
                                    "+" & colum & "2*" & Str(3 * Кабел_Розетка)
                            rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                            wsLines.Range(rang).Formula = formu
                            colum = "$V$"
                            formu = "=" & colum & "2*" & Кабел_кутия.ToString &
                                    "+" & colum & "2*" & Str(H_Етаж - H_Контакт)
                            rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                            wsLines.Range(rang).Formula = formu
                        Case "СВТ5x4,0mm²"
                            colum = "$W$"
                            formu = "=" & colum & "2*" & Кабел_кутия.ToString &
                                    "+" & colum & "2*" & Str(H_Етаж - H_Контакт)
                            rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                            wsLines.Range(rang).Formula = formu
                        Case "СВТ5x6,0mm²"
                            colum = "$X$"
                            formu = "=" & colum & "2*" & Кабел_кутия.ToString &
                                    "+" & colum & "2*" & Str(H_Етаж - H_Контакт)
                            rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                            wsLines.Range(rang).Formula = formu
                        Case "СВТ5x10mm²"
                            colum = "$Y$"
                            formu = "=" & colum & "2*" & Кабел_кутия.ToString &
                                    "+" & colum & "2*" & Str(H_Етаж - H_Контакт)
                            rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                            wsLines.Range(rang).Formula = formu
                        Case "СВТ5x16mm²"
                            colum = "$Z$"
                            formu = "=" & colum & "2*" & Кабел_кутия.ToString &
                                    "+" & colum & "2*" & Str(H_Етаж - H_Контакт)
                            rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                            wsLines.Range(rang).Formula = formu
                        Case "СВТ5x25mm²", "СВТ5x35mm²"
                            colum = "$Z$"
                            formu = "=" & colum & "2*" & Кабел_кутия.ToString &
                                    "+" & colum & "2*" & Str(H_Етаж - H_Контакт)
                            rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                            wsLines.Range(rang).Formula = formu
                        Case "NHXH FE180/Е30 2x1,5mm²"
                        Case "NHXH FE180/Е30 3x1,5mm²"
                            colum = "$N$"
                            formu = "=2*" & colum & "2*" & Кабел_кутия.ToString &
                                    "+" & colum & "2*" & Str(H_Етаж - 2)
                            rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                            wsLines.Range(rang).Formula = formu
                        Case "NHXH FE180/Е30 3x2,5mm²"
                        Case "NHXH FE180/Е30 3x4,0mm²"
                        Case "NHXH FE180/Е30 3x6,0mm²"
                        Case "NHXH FE180/Е30 4x1,5mm²"
                        Case "NHXH FE180/Е30 4x6,0mm²"
                        Case "NHXH FE180/Е30 5x1,5mm²"
                        Case "NHXH FE180/Е30 5x2,5mm²"
                        Case "NHXH FE180/Е30 5x4,0mm²"
                        Case "NHXH FE180/Е30 5x6,0mm²"
                    End Select
                Case "Button_Слаботокови_КАБЕЛИ"
                    '
                    ' Слаботокови кабели - ИНТЕРНЕТ
                    '
                    Инсталация = "ИНТЕРНЕТ"
                    Select Case masType(i).blType
                        Case "ПТПВ 1х2х0,5мм²"
                            colum = "$AH$"
                            formu = "=2*" & colum & "2*" & Кабел_кутия.ToString &
                                "+" & colum & "2*" & Str(H_Етаж - H_Контакт)
                            rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                            wsLines.Range(rang).Formula = formu
                            colum = "$AK$"
                            formu = "=" & colum & "2"
                            rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                            wsLines.Range(rang).Formula = formu
                        Case "FTP 4x2х24AWG"
                            colum = "$AE$"
                            formu = "=2*" & colum & "2*" & Кабел_кутия.ToString &
                                       "+" & colum & "2*" & Str(H_Етаж - H_Контакт)
                            rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                            wsLines.Range(rang).Formula = formu
                            colum = "$AF$"
                            formu = "=2*" & colum & "2*" & Кабел_кутия.ToString &
                                    "+" & colum & "2*" & Str(H_Етаж - H_Контакт)
                            rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                            wsLines.Range(rang).Formula = formu
                            colum = "$AG$"
                            formu = "=2*" & colum & "2*" & Кабел_кутия.ToString &
                                    "+" & colum & "2*" & Str(H_Етаж - H_Контакт)
                            rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                            wsLines.Range(rang).Formula = formu
                            colum = "$AI$"
                            formu = "=" & colum & "2"
                            rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                            wsLines.Range(rang).Formula = formu
                        Case "коакс.кабел RG6/64"
                            colum = "$AD$"
                            formu = "=2*" & colum & "2*" & Кабел_кутия.ToString &
                                "+" & colum & "2*" & Str(H_Етаж - H_Контакт)
                            rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                            wsLines.Range(rang).Formula = formu
                            colum = "$AJ$"
                            formu = "=" & colum & "2"
                            rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                            wsLines.Range(rang).Formula = formu
                    End Select
                    '
                    ' Слаботокови кабели - ПОЖАРОИЗВЕСТЯВАНЕ
                    '
                Case "Button_Вземи_ПИЦ_КАБЕЛИ"
                    Инсталация = "ПИЦ"
                    Select Case masType(i).blType
                        Case "FS 3x0,50mm²"
                            colum = "$AM$"
                            formu = "=" & colum & "2" &
                                    IIf(RadioButton_Адресируема.Checked, "*2*", "*") &
                                    Кабел_кутия.ToString
                            rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                            wsLines.Range(rang).Formula = formu

                            colum = "$AQ$"
                            formu = "=" & colum & "2"                           ' Взима само дължината
                            rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                            wsLines.Range(rang).Formula = formu

                        Case "FS 2x0,50mm²"
                            colum = "$AN$"
                            formu = "=2*" & colum & "2*" & Кабел_кутия.ToString &   ' кабел в лампата
                                "+" & colum & "2*" & (H_Етаж - 2).ToString      ' Слизане до датчика
                            rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                            wsLines.Range(rang).Formula = formu

                            colum = "$AM$"
                            formu = "=" & colum & "2" &
                                        IIf(RadioButton_Адресируема.Checked, "*2*", "*") &
                            Кабел_кутия.ToString
                            rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                            wsLines.Range(rang).Formula = formu

                            colum = "$AO$"
                            formu = "=2*" & colum & "2*" & Кабел_кутия.ToString &   ' Кабел в пожароизвестителя или изпълнителното устройство
                             "+" & colum & "2*" & (H_Етаж - H_Контакт).ToString                ' Кабел от тавана до пожароизвестителя
                            rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                            wsLines.Range(rang).Formula = formu

                            colum = "$AP$"
                            formu = "=" & colum & "2"                           ' Взима само дължината
                            rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                            wsLines.Range(rang).Formula = formu
                    End Select
                Case "Button_Вземи_ДОМОФ_КАБЕЛИ"
                    '
                    ' Слаботокови кабели - ДОМОФОННА
                    '
                    Инсталация = "ДОМОФ"
                    If masType(i).blType = "FTP 4x2х24AWG" Then
                        colum = "$AS$"
                        formu = "=2*" & colum & "2*" & Кабел_кутия.ToString &
                                "+2*" & colum & "2*" & Str(H_Етаж - H_Контакт)
                        rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                        wsLines.Range(rang).Formula = formu

                        colum = "$AT$"
                        formu = "=" & colum & "2"                           ' Взима само дължината
                        rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                        wsLines.Range(rang).Formula = formu
                    End If
                Case "Button_Вземи_ВИДЕО_КАБЕЛИ"
                    Инсталация = "ВИДЕО"
                    '
                    'Кабели за ВИДЕОНАБЛЮДЕНИЕ
                    '
                    If masType(i).blType = "FTP 4x2х24AWG" Then
                        colum = "$AZ$"
                        formu = "=2*" & colum & "2*" & Кабел_кутия.ToString &
                                "+2*" & colum & "2*" & Str(H_Етаж - H_Контакт)
                        rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                        wsLines.Range(rang).Formula = formu

                        colum = "$BA$"
                        formu = "=" & colum & "2"                           ' Взима само дължината
                        rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                        wsLines.Range(rang).Formula = formu
                    End If
                Case "Button_Вземи_СОТ_КАБЕЛИ"
                    Инсталация = "СОТ"
                    '
                    'Кабели за СОТ
                    '
                    If masType(i).blType = "CAB/6/WH 6х25SWG" Then
                        colum = "$BC$"
                        formu = "=2*" & colum & "2*" & Кабел_кутия.ToString &   ' За датик и центала
                                "+" & colum & "2*" & (H_Етаж - 2.2).ToString &  ' Слизане до датчик
                                "+" & colum & "2*" & (H_Етаж - 1.5).ToString    ' Слизане до центала
                        rang = colum & Trim(index.ToString)
                        wsLines.Range(rang).Formula = formu
                        colum = "$BD$"
                        formu = "=2*" & colum & "2*" & Кабел_кутия.ToString &   ' За датчик и центала
                                "+" & colum & "2*" & (H_Етаж - 1.5).ToString &  ' Слизане до датчик
                                "+" & colum & "2*" & (H_Етаж - 1.5).ToString    ' Слизане до центала
                        rang = colum & Trim(index.ToString)
                        wsLines.Range(rang).Formula = formu
                        colum = "$BE$"
                        formu = "=" & colum & "2"                               ' Взима само дължината за качване
                        rang = colum & Trim(index.ToString)
                        wsLines.Range(rang).Formula = formu
                    End If
                Case "Button_МЪЛНИЯ_КАБЕЛИ"
                    '
                    ' ЗАЗЕМИТЕЛНА МЪЛНИЕЗАЩИТНА
                    '
                    Инсталация = "МЪЛНИЯ"
                    If masType(i).blType = "AlMgSi Ф8мм" Then
                        colum = "$AX$"
                        formu = "=0"
                        rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                        wsLines.Range(rang).Formula = formu
                    End If
            End Select
            With wsLines
                .Cells(index, 1) = Инсталация + "-КАБЕЛ"
                .Cells(index, 2) = masType(i).blType
                .Cells(index, 3) = masType(i).blLength / 100 / Kabel_fec
            End With
            index += 1
        Next
        If sender.name = "Button_силова_КАБЕЛИ" Then
            If boEL_2x1_5 Then
                If wsLines.Cells(2, 8).Value > 0 Then
                    colum = "$H$"
                    formu = "=" & colum & "2*" & Кабел_кутия.ToString &
                        "+" & colum & "2*" & Str(H_Етаж - H_Ключ) &
                        "+" & colum & "2*" & Str(Кабел_Розетка)
                    rang = colum & Trim(index.ToString)
                    wsLines.Range(rang).Formula = formu
                    With wsLines
                        .Cells(index, 1) = Инсталация + "-КАБЕЛ"
                        .Cells(index, 2) = "СВТ2x1,5mm²"
                        .Cells(index, 3) = 0
                    End With
                    index += 1
                End If
            End If

            If boEL_4x1_5 Then
                If (wsLines.Cells(2, 10).Value + wsLines.Cells(2, 13).Value) > 0 Then
                    colum = "$J$"
                    formu = "=" & colum & "2*" & Кабел_кутия.ToString &
                        "+" & colum & "2*" & Str(H_Етаж - H_Ключ) &
                        "+" & colum & "2*" & Str(3 * Кабел_Розетка)
                    rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                    wsLines.Range(rang).Formula = formu
                    colum = "$M$"
                    formu = "=" & colum & "2*" & Кабел_кутия.ToString
                    rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                    wsLines.Range(rang).Formula = formu
                    With wsLines
                        .Cells(index, 1) = Инсталация + "-КАБЕЛ"
                        .Cells(index, 2) = "СВТ4x1,5mm²"
                        .Cells(index, 3) = 0
                    End With
                    index += 1
                End If
            End If
        End If
        Dim masPol() As strKabel = Kol_smetka_Kabel(Kabel, vbFalse)

        For i = 0 To UBound(masPol)
            If masPol(i).blLength = 0 Then Exit For
            formu = ""
            rang = ""
            colum = ""
            With wsLines
                Select Case masPol(i).blPol
                    Case "ByLayer"
                        Continue For
                    Case "изт. в HDPE тр.ф40/32mm"
                    Case "изт. в HDPE тр.ф50/41mm"
                    Case "изт. в HDPE тр.ф63/53mm"
                    Case "изт. в HDPE тр.ф75/61mm"
                    Case "изт. в HDPE тр.ф90/75mm"
                    Case "изт. в HDPE тр.ф110/94mm"
                    Case "изт. в HDPE тр.ф125/108mm"
                    Case "изт. в HDPE тр.ф140/121mm"
                    Case "изт. в HDPE тр.ф160/136mm"
                    Case "изт. в HDPE тр.ф200/170mm"
                    Case "изт. в PE тр.ф18,7/13,5mm"
                    Case "изт. в PE тр.ф21,2/16mm"
                    Case "изт. в PE тр.ф28,5/22,9mm"
                    Case "изт. в PE тр.ф34,5/28,4mm"
                    Case "изт. в PE тр.ф46,5/35,9mm"
                    Case "изт. в PVC тр.ф16,0/11,3mm"
                        Select Case sender.name
                            Case "Button_Слаботокови_КАБЕЛИ"
                                colum = "$AH$"
                                formu = "=" & colum & "2*" & Str(H_Етаж - H_Контакт)
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu
                                colum = "$AE$"
                                formu = "=" & colum & "2*" & Str(H_Етаж - H_Контакт)
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu
                                colum = "$AF$"
                                formu = "=" & colum & "2*" & Str(H_Етаж - H_Контакт)
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu
                                colum = "$AG$"
                                formu = "=" & colum & "2*" & Str(H_Етаж - H_Контакт)
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu
                                colum = "$AD$"
                                formu = "=" & colum & "2*" & Str(H_Етаж - H_Контакт)
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu
                                colum = "$AI$"
                                formu = "=" & colum & "2"
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu
                                colum = "$AJ$"
                                formu = "=" & colum & "2"
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu
                                colum = "$AK$"
                                formu = "=" & colum & "2"
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu
                                colum = "$AI$"
                                formu = "=" & colum & "2"
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu
                            Case "Button_Вземи_ПИЦ_КАБЕЛИ"
                                colum = "$AN$"
                                formu = "=" & colum & "2*" & Str(H_Етаж - 2)
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu
                                colum = "$AO$"
                                formu = "=" & colum & "2*" & Str(H_Етаж - 1.5)
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu
                                colum = "$AP$"
                                formu = "=" & colum & "2"
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu
                                colum = "$AQ$"
                                formu = "=" & colum & "2"
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu
                            Case "Button_Вземи_ДОМОФ_КАБЕЛИ"
                                colum = "$AS$"
                                formu = "=" & colum & "2*" & Str(H_Етаж - 2)
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu
                                colum = "$AT$"
                                formu = "=" & colum & "2"
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu
                            Case "Button_Вземи_ВИДЕО_КАБЕЛИ"
                                colum = "$AZ$"
                                formu = "=" & colum & "2*" & Str(H_Етаж - 0.5) + ' За RACK шкаф
                                        "+" & colum & "2*" & Str(H_Етаж - 2) ' За височина на камерата
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu
                                colum = "$BA$"
                                formu = "=" & colum & "2"
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu
                            Case "Button_Вземи_СОТ_КАБЕЛИ"
                                colum = "$BC$"
                                formu = "=" & colum & "2*" & Кабел_кутия.ToString &   ' За датик и центала
                                "+" & colum & "2*" & (H_Етаж - 2.2).ToString &  ' Слизане до датчик
                                "+" & colum & "2*" & (H_Етаж - 1.5).ToString    ' Слизане до центала
                                rang = colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu
                                colum = "$BD$"
                                formu = "=" & colum & "2*" & Кабел_кутия.ToString &   ' За датик и центала
                                "+" & colum & "2*" & (H_Етаж - 1.5).ToString &  ' Слизане до датчик
                                "+" & colum & "2*" & (H_Етаж - 1.5).ToString    ' Слизане до центала
                                rang = colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu
                                colum = "$BE$"
                                formu = "=" & colum & "2"                               ' Взима само дължината за качване
                                rang = colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu
                        End Select
                    Case "изт. в PVC тр.ф20,0/14,6mm"
                        colum = "$H$"
                        formu = "=" & colum & "2*" & (H_Етаж - H_Ключ).ToString
                        rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                        wsLines.Range(rang).Formula = formu
                        colum = "$I$"
                        formu = "=" & colum & "2*" & (H_Етаж - H_Ключ).ToString
                        rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                        wsLines.Range(rang).Formula = formu
                        colum = "$J$"
                        formu = "=" & colum & "2*" & (H_Етаж - H_Ключ).ToString
                        rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                        wsLines.Range(rang).Formula = formu
                        colum = "$K$"
                        formu = "=" & colum & "2*" & (H_Етаж - H_Ключ).ToString
                        rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                        wsLines.Range(rang).Formula = formu
                        colum = "$P$"
                        formu = "=" & colum & "2*" & (H_Етаж - H_Ключ).ToString
                        rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                        wsLines.Range(rang).Formula = formu
                        colum = "$T$"
                        formu = "=" & colum & "2*" & (H_Етаж - H_Ключ).ToString
                        rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                        wsLines.Range(rang).Formula = formu
                        colum = "$Q$"
                        formu = "=" & colum & "2*" & (H_Етаж - H_Ключ).ToString
                        rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                        wsLines.Range(rang).Formula = formu
                        colum = "$V$"
                        formu = "=" & colum & "2*" & (H_Етаж - H_Ключ).ToString
                        rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                        wsLines.Range(rang).Formula = formu
                    Case "изт. в PVC тр.ф25,0/18,5mm"
                        colum = "$R$"
                        formu = "=" & colum & "2*" & (H_Етаж - H_Ключ).ToString
                        rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                        wsLines.Range(rang).Formula = formu
                        colum = "$U$"
                        formu = "=" & colum & "2*" & (H_Етаж - H_Ключ).ToString
                        rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                        wsLines.Range(rang).Formula = formu
                        colum = "$W$"
                        formu = "=" & colum & "2*" & (H_Етаж - H_Ключ).ToString
                        rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                        wsLines.Range(rang).Formula = formu
                        colum = "$X$"
                        formu = "=" & colum & "2*" & (H_Етаж - H_Ключ).ToString
                        rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                        wsLines.Range(rang).Formula = formu
                    Case "изт. в PVC тр.ф32,0/24,3mm"
                        colum = "$Y$"
                        formu = "=" & colum & "2*" & (H_Етаж - H_Ключ).ToString
                        rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                        wsLines.Range(rang).Formula = formu
                        colum = "$Z$"
                        formu = "=" & colum & "2*" & (H_Етаж - H_Ключ).ToString
                        rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                        wsLines.Range(rang).Formula = formu
                    Case "изт. в PVC тр.ф40,0/31,2mm"
                    Case "изт. в PVC тр.ф50,0/39,6mm"
                    Case "изт. в каб.кан.100х40mm"
                    Case "изт. в каб.кан.100х60mm"
                    Case "изт. в каб.кан.12х12mm"
                    Case "изт. в каб.кан.16х16mm"
                    Case "изт. в каб.кан.20х20mm"
                        Select Case sender.name
                            Case "Button_Слаботокови_КАБЕЛИ"
                                colum = "$AH$"
                                formu = "=2*" & colum & "2*" & Кабел_кутия.ToString &
                                        "+" & colum & "2*" & Str(H_Етаж - H_Контакт)
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu
                                colum = "$AK$"
                                formu = "=" & colum & "2"
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu

                                colum = "$AE$"
                                formu = "=2*" & colum & "2*" & Кабел_кутия.ToString &
                                        "+" & colum & "2*" & Str(H_Етаж - H_Контакт)
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu
                                colum = "$AF$"
                                formu = "=2*" & colum & "2*" & Кабел_кутия.ToString &
                                        "+" & colum & "2*" & Str(H_Етаж - H_Контакт)
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu
                                colum = "$AG$"
                                formu = "=2*" & colum & "2*" & Кабел_кутия.ToString &
                                        "+" & colum & "2*" & Str(H_Етаж - H_Контакт)
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu
                                colum = "$AI$"
                                formu = "=" & colum & "2"
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu

                                colum = "$AD$"
                                formu = "=2*" & colum & "2*" & Кабел_кутия.ToString &
                                        "+" & colum & "2*" & Str(H_Етаж - H_Контакт)
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu
                                colum = "$AJ$"
                                formu = "=" & colum & "2"
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu
                            Case "Button_силова_КАБЕЛИ"
                                colum = "$H$"
                                formu = "=" & colum & "2*" & Кабел_кутия.ToString &
                                "+" & colum & "2*" & Str(H_Етаж - H_Ключ) &
                                "+" & colum & "2*" & Str(Кабел_Розетка)
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu

                                colum = "$I$"
                                formu = "=" & colum & "2*" & Кабел_кутия.ToString &
                                "+" & colum & "2*" & Str(H_Етаж - H_Ключ) &
                                "+" & colum & "2*" & Str(3 * Кабел_Розетка)
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu
                                colum = "$L$"
                                formu = "=" & colum & "2*" & Кабел_кутия.ToString
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu
                                colum = "$N$"
                                formu = "=" & colum & "2*" & Кабел_кутия.ToString
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu
                                colum = "$P$"
                                formu = "=" & colum & "2*" & Кабел_кутия.ToString &
                                "+" & colum & "2*" & Str(H_Етаж - H_Контакт) &
                                "+" & colum & "2*" & Str(3 * Кабел_Розетка)
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu
                                colum = "$T$"
                                formu = "=" & colum & "2*" & Кабел_кутия.ToString & "+" & colum & "2*" & Str(H_Етаж - H_Контакт)
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu
                                colum = "$U$"
                                formu = "=" & colum & "2*" & Кабел_кутия.ToString &
                                "+" & colum & "2*" & Str(H_Етаж - H_Контакт)
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu
                                colum = "$R$"
                                formu = "=" & colum & "2*" & Кабел_кутия.ToString &
                                "+" & colum & "2*" & Str(H_Етаж - H_Контакт)
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu
                                colum = "$AB$"
                                formu = "=" & colum & "2*" & Кабел_кутия.ToString &
                                "+" & colum & "2*" & Str(H_Етаж - H_Ключ)
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu
                                colum = "$J$"
                                formu = "=" & colum & "2*" & Кабел_кутия.ToString &
                                "+" & colum & "2*" & Str(H_Етаж - H_Ключ) &
                                "+" & colum & "2*" & Str(3 * Кабел_Розетка)
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu
                                colum = "$M$"
                                formu = "=" & colum & "2*" & Кабел_кутия.ToString
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu
                                colum = "$K$"
                                formu = "=" & colum & "2*" & Кабел_кутия.ToString &
                                "+" & colum & "2*" & Str(H_Етаж - H_Ключ) &
                                "+" & colum & "2*" & Str(3 * Кабел_Розетка)
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu
                                colum = "$Q$"
                                formu = "=" & colum & "2*" & Кабел_кутия.ToString &
                                "+" & colum & "2*" & Str(H_Етаж - H_Контакт) &
                                "+" & colum & "2*" & Str(3 * Кабел_Розетка)
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu
                                colum = "$V$"
                                formu = "=" & colum & "2*" & Кабел_кутия.ToString &
                                "+" & colum & "2*" & Str(H_Етаж - H_Контакт)
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu
                            Case "Button_Вземи_ПИЦ_КАБЕЛИ"
                                colum = "$AN$"
                                formu = "=" & colum & "2*" & (H_Етаж - 2.1).ToString
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu

                                colum = "$AO$"
                                formu = "=" & colum & "2*" & (H_Етаж - H_Контакт).ToString
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu

                                colum = "$AP$"
                                formu = "=" & colum & "2"                           ' Взима само дължината
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu

                                colum = "$AQ$"
                                formu = "=" & colum & "2"                           ' Взима само дължината
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu

                            Case "Button_Вземи_ДОМОФ_КАБЕЛИ"
                                colum = "$AS$"
                                formu = "=2*" & colum & "2*" & Кабел_кутия.ToString &
                                        "+" & colum & "2*" & Str(H_Етаж - H_Контакт)
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu

                                colum = "$AT$"
                                formu = "=" & colum & "2"                           ' Взима само дължината
                                rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                                wsLines.Range(rang).Formula = formu
                        End Select
                    Case "изт. в каб.кан.25х20mm"
                    Case "изт. в каб.кан.25х25mm"
                    Case "изт. в каб.кан.30х25mm"
                    Case "изт. в каб.кан.40х20mm"
                    Case "изт. в каб.кан.40х25mm"
                    Case "изт. в каб.кан.40х40mm"
                    Case "изт. в каб.кан.60х20mm"
                    Case "изт. в каб.кан.60х40mm"
                    Case "изт. в каб.кан.60х60mm"
                    Case "изт. в каб.кан.80х20mm"
                    Case "изт. в каб.кан.80х25mm"
                    Case "изт. в каб.кан.80х40mm"
                    Case "изт. в мет.тр.ф13,2/9,0mm"
                    Case "изт. в мет.тр.ф15,2/11,0mm"
                    Case "изт. в мет.тр.ф18,4/14,0mm"
                    Case "изт. в мет.тр.ф22,4/18,0mm"
                    Case "изт. в мет.тр.ф30,4/26,0mm"
                    Case "изт. в мет.тр.ф42,4/37,0mm"
                    Case "изт. в негор.PVC тр.ф16,0/10,7mm"
                        colum = "$N$"
                        formu = "=" & colum & "2*" & Str(H_Етаж - 2)
                        rang = colum & Trim(index.ToString) & ":" & colum & Trim(index.ToString)
                        wsLines.Range(rang).Formula = formu
                    Case "изт. в негор.PVC тр.ф20,0/14,1mm"
                    Case "изт. в негор.PVC тр.ф25,0/18,2mm"
                    Case "изт. в негор.PVC тр.ф32,0/24,3mm"
                    Case "изт. в негор.PVC тр.ф40,0/32,3mm"
                    Case "изт.въздушно"
                    Case "изт.по носещо въже"
                    Case "открито на ПКОМ скоби"
                    Case "положен в изкоп 0,8/0,5м в"
                    Case "положен по кабелна скара"
                    Case "скрито под мазилката"
                End Select
                .Cells(index, 1) = Инсталация + "-ПОЛАГАНЕ"
                .Cells(index, 2) = masPol(i).blPol
                .Cells(index, 3) = masPol(i).blLength / 100 / Kabel_fec
                index += 1
            End With
        Next
        formu = ""
        rang = ""
        colum = ""

        formu = "=sum(G3:BZ3)+C3"
        rang = "D3:D" & Trim((index - 1).ToString)
        wsLines.Range(rang).Formula = formu

        formu = "=D3*$E$2"
        rang = "E3:E" & Trim((index - 1).ToString)
        wsLines.Range(rang).Formula = formu

        formu = "=INT(E3/10+$F$2)*10"
        rang = "F3:F" & Trim((index - 1).ToString)
        wsLines.Range(rang).Formula = formu

        formu = 1.3
        rang = "E2:E2"
        wsLines.Range(rang).Formula = formu

        formu = 1
        rang = "F2:F2"
        wsLines.Range(rang).Formula = formu
        '
        '  Сортира EXCEL листа
        ' 
        With wsLines
            .Cells.Sort(Key1:= .Range("A2"),
                        Order1:=excel.XlSortOrder.xlAscending,
                        Header:=excel.XlYesNoGuess.xlYes,
                        OrderCustom:=1, MatchCase:=False,
                        Orientation:=excel.Constants.xlTopToBottom,
                        DataOption1:=excel.XlSortDataOption.xlSortTextAsNumbers,
                        Key2:= .Range("B2"),
                        Order2:=excel.XlSortOrder.xlAscending,
                        DataOption2:=excel.XlSortDataOption.xlSortTextAsNumbers
                        )
        End With
    End Sub
    ' Функция Kol_smetka_Kabel
    ' Входни параметри:
    ' Kabel - Масив, който съдържа данни за кабели.
    ' index - Булев параметър, който определя типа на сортиране.
    '         True - сортиране по тип на кабела.
    '         False - сортиране по начин на полагане на кабела.
    ' Изход:
    ' Връща масив от структури strKabel, който съдържа резултатите от изчисленията.
    Function Kol_smetka_Kabel(Kabel As Array, index As Boolean) As Array
        ' Обявяване на масив за сортираните кабели с максимален размер 500.
        Dim kabelSort(500) As strKabel

        ' Ако index е True, сортиране по тип на кабела
        If index Then
            Dim indexSort As Integer = 0
            ' Обхождане на входния масив Kabel
            For i = 0 To UBound(Kabel)
                ' Ако е достигнат празен елемент, прекратяване на цикъла
                If Kabel(i, 0) = Nothing Then Exit For
                Dim iVisib As Integer = -1
                ' Променлива, която държи текущия тип кабел
                Dim kabSort As String = Kabel(i, 0)
                ' Търсене на индекс на елемент с този тип в сортирания масив
                iVisib = Array.FindIndex(kabelSort, Function(f) f.blType = kabSort)
                ' Ако типът кабел не е намерен в сортирания масив
                If iVisib = -1 Then
                    ' Добавяне на нов тип кабел в сортирания масив
                    kabelSort(indexSort).blType = Kabel(i, 0)
                    kabelSort(indexSort).blLength = Kabel(i, 2)
                    indexSort += 1
                Else
                    ' Ако типът кабел е намерен, добавяне на дължината към съществуващия елемент
                    kabelSort(iVisib).blLength = kabelSort(iVisib).blLength + Kabel(i, 2)
                End If
            Next
            ' Ако index е False, сортиране по начин на полагане
        Else
            Dim indexSort As Integer = 0
            ' Обхождане на входния масив Kabel
            For i = 0 To UBound(Kabel)
                ' Ако е достигнат празен елемент, прекратяване на цикъла
                If Kabel(i, 2) = Nothing Then Exit For
                Dim iVisib As Integer = -1
                ' Променлива, която държи текущия начин на полагане на кабела
                Dim kabSort As String = Kabel(i, 1)
                ' Търсене на индекс на елемент с този начин на полагане в сортирания масив
                iVisib = Array.FindIndex(kabelSort, Function(f) f.blPol = kabSort)
                ' Ако начинът на полагане не е намерен в сортирания масив
                If iVisib = -1 Then
                    ' Добавяне на нов начин на полагане в сортирания масив
                    kabelSort(indexSort).blPol = Kabel(i, 1)
                    kabelSort(indexSort).blLength = Kabel(i, 2)
                    kabelSort(indexSort).blCount += 1
                    indexSort += 1
                Else
                    ' Ако начинът на полагане е намерен, добавяне на дължината към съществуващия елемент
                    kabelSort(iVisib).blLength = kabelSort(iVisib).blLength + Kabel(i, 2)
                    kabelSort(iVisib).blCount += 1
                End If
            Next
        End If
        ' Връщане на сортирания масив
        Return kabelSort
    End Function
    Private Sub clearDomofon(ws As excel.Worksheet)
        With ws
            With .Range("A:AK")
                .Clear()
                .Value = ""
            End With
            With .Range("A:AK")
                .Font.Name = "Cambria"
                .Font.Size = 12
                .WrapText = vbTrue
            End With
            .Columns("A").ColumnWidth = 6
            .Columns("B").ColumnWidth = 15
            .Columns("C").ColumnWidth = 40
            .Columns("D").ColumnWidth = 11
            .Columns("E").ColumnWidth = 21
            .Columns("K:AG").ColumnWidth = 30
            .Cells(1, 1).Value = "Брой"
            .Cells(1, 2).Value = "Име"
            .Cells(1, 3).Value = "Visibility"
            .Cells(1, 5).Value = "Качване"
        End With
        If IsNothing(wsLines) Then Exit Sub
        Select Case ws.Name
            Case "Домофонна"
                wsLines.Range("AS2").Value = 0
                wsLines.Range("AT2").Value = 0
            Case "Видеонаблюдение"
                wsLines.Range("AZ2").Value = 0
                wsLines.Range("BA2").Value = 0
            Case "СОТ"
                wsLines.Range("BC2").Value = 0
                wsLines.Range("BD2").Value = 0
                wsLines.Range("BE2").Value = 0
        End Select
    End Sub
    Private Sub clearКбелнаскара(ws As excel.Worksheet)
        With ws
            With .Range("A:AZ")
                .Clear()
                .Value = ""
            End With
            With .Range("A:Z")
                .Font.Name = "Cambria"
                .Font.Size = 12
                .WrapText = vbTrue
                .Style = "Comma"
            End With
            .Columns("A").ColumnWidth = 16
            .Columns("B").ColumnWidth = 9
            .Columns("C").ColumnWidth = 11
            .Columns("D").ColumnWidth = 21

            .Columns("E:Z").ColumnWidth = 14
            .Columns("S:AG").ColumnWidth = 5

            .Range("A1").Value = "Име"
            .Range("B1").Value = "Ширина"
            .Range("C1").Value = "Височина"
            .Range("D1").Value = "Дължина аутокад"
            .Range("E1").Value = "Брой"
            .Range("F1").Value = "Общо вертикални"
            .Range("G1").Value = "Общо х 1,1"
            .Range("H1").Value = "Общо с коефициент"
            .Range("J1").Value = "Качване"
            .Range("K1").Value = "Табла"
            .Range("M1").Value = "Ключове"
            .Range("N1").Value = "Контакти"
            .Range("O1").Value = "Датчици"
            .Range("P1").Value = "Качване 1"
            .Range("Q1").Value = "Качване 2"

            .Range("G2:R2").Value = 0
            .Range("G3:R3").Value = 0

            .Range("G2").Value = 1.1
            .Range("H2").Value = 1

        End With
        Kol_Smetka_Cells_Format(ws, "I1", "ОБЩО")
        Kol_Smetka_Cells_Format(ws, "L1", "СКАРИ")
        Kol_Smetka_Cells_Format(ws, "R1", "КАНАЛИ")

        ws.Outline.ShowLevels(0, 1)
    End Sub
    Private Sub clearInternet(ws As excel.Worksheet)
        With ws
            With .Range("A:AK")
                .Clear()
                .Value = ""
            End With
            With .Range("A:AK")
                .Font.Name = "Cambria"
                .Font.Size = 12
                .WrapText = vbTrue
            End With
            .Columns("A").ColumnWidth = 6
            .Columns("B").ColumnWidth = 15
            .Columns("C").ColumnWidth = 40
            .Columns("D").ColumnWidth = 25
            .Columns("E").ColumnWidth = 8
            .Columns("F").ColumnWidth = 11
            .Columns("G").ColumnWidth = 2
            .Columns("H").ColumnWidth = 21
            .Columns("I").ColumnWidth = 21
            .Columns("J").ColumnWidth = 21
            .Columns("K:AG").ColumnWidth = 30
            .Cells(1, 1).Value = "Брой"
            .Cells(1, 2).Value = "Име"
            .Cells(1, 3).Value = "Visibility"
            .Cells(1, 4).Value = "Наименование"
            .Cells(1, 5).Value = "Юнити"
            .Cells(1, 6).Value = "Височина"
            .Cells(1, 8).Value = "Качване FTP"
            .Cells(1, 9).Value = "Качване Телевизия"
            .Cells(1, 10).Value = "Качване Телефон"
        End With
        If IsNothing(wsLines) Then Exit Sub
        With wsLines
            .Cells(2, 30).Value = 0
            .Cells(2, 31).Value = 0
            .Cells(2, 32).Value = 0
            .Cells(2, 33).Value = 0
            .Cells(2, 34).Value = 0
            .Cells(2, 35).Value = 0
            .Cells(2, 36).Value = 0
            .Cells(2, 37).Value = 0
        End With
    End Sub
    Private Sub clearKabeli(ws As excel.Worksheet,      ' Лист в който да се изтрива
                            Rows1 As Integer,           ' начален ред за изтриване
                            Rows2 As Integer            ' краен ред за изтриване
                            )

        If Rows2 < Rows1 Then Exit Sub
        Dim Rows_Delete As String = Rows1.ToString + ":" + Rows2.ToString
        With ws
            With .Rows(Rows_Delete)
                .delete()
            End With
            With .Range("A:BZ")
                .Font.Name = "Cambria"
                .Font.Size = 12
                .WrapText = vbTrue
                .Style = "Comma"
            End With
            If Rows2 > 900 Then
                .Range("C:F").Group()
                .Range("H:N").Group()
                .Range("P:R").Group()
                .Range("T:AB").Group()
                .Range("AD:AK").Group()
                .Range("AM:AQ").Group()
                .Range("AS:AT").Group()
                .Range("AV:AV").Group()
                .Range("AX:AX").Group()
                .Range("AZ:BA").Group()
                .Range("BC:BE").Group()

                .Range("C2:BF2").Value = 0
            End If

            .Columns("A").ColumnWidth = 25
            .Columns("B").ColumnWidth = 40
            .Columns("C").ColumnWidth = 10
            .Columns("D:BZ").ColumnWidth = 15

            .Range("A2").Value = "Вид инсталация"
            .Range("B2").Value = "Тип кабел"
            .Range("C1").Value = "От Autocad"
            .Range("D1").Value = "Общо вертикални"
            .Range("E1").Value = "Общо с коефициент"
            .Range("F1").Value = "Общо закръглено"

            .Cells(1, 8).Value = "Kлюч двужилен"
            .Cells(1, 9).Value = "Kлюч трижилен"
            .Cells(1, 10).Value = "Kлюч четирижилен"
            .Cells(1, 11).Value = "Kлюч петрижилен"
            .Cells(1, 12).Value = "Осветител трижилен"
            .Cells(1, 13).Value = "Осветител четирижилен"
            .Cells(1, 14).Value = "Осветител ИЗХОД"

            .Cells(1, 16).Value = "Монофазен контакт"
            .Cells(1, 17).Value = "Трифазен контакт"
            .Cells(1, 18).Value = "Усилен контакт"

            .Cells(1, 20).Value = "Бойлер трижилен 2,5"
            .Cells(1, 21).Value = "Бойлер трижилен 4,0"
            .Cells(1, 22).Value = "Бойлер петрижилен 2,5"
            .Cells(1, 23).Value = "Бойлер петрижилен 4,0"
            .Cells(1, 24).Value = "Бойлер петрижилен 6,0"
            .Cells(1, 25).Value = "Бойлер петрижилен 10,0"
            .Cells(1, 26).Value = "Бойлер петрижилен 16,0"
            .Cells(1, 27).Value = "Бойлер виж чертежа"
            .Cells(1, 28).Value = "Бойлерно табло"

            .Cells(1, 30).Value = "Розетка телевизия"
            .Cells(1, 31).Value = "Розетка интернет"
            .Cells(1, 32).Value = "Розетка рутер"
            .Cells(1, 33).Value = "Розетка табло"
            .Cells(1, 34).Value = "Розетка Телефон"
            .Cells(1, 35).Value = "Качване FTP"
            .Cells(1, 36).Value = "Качване Телевизия"
            .Cells(1, 37).Value = "Качване Телефон"

            .Cells(1, 39).Value = "Датчик ПИЦ"
            .Cells(1, 40).Value = "Паралелен ПИЦ"
            .Cells(1, 41).Value = "Ръчен ПИЦ"
            .Cells(1, 42).Value = "Качване двужилен"
            .Cells(1, 43).Value = "Качване трижилен"

            .Range("AS1").Value = "Домофон"
            .Range("AT1").Value = "Домофон качване"

            .Range("AV1").Value = "Заземител"
            .Range("AX1").Value = "Заземител"

            .Range("AZ1").Value = "Видеокамера"
            .Range("BA1").Value = "Видео качване"

            .Range("BC1").Value = "Датчик СОТ"
            .Range("BD1").Value = "Клавиатура СОТ"
            .Range("BE1").Value = "СОТ качване"
            .Range("BG1").Value = "Силова качване"
            .Range("BH1").Value = "Силова качване"
        End With

        Kol_Smetka_Cells_Format(ws, "G1", "ОБЩО")
        Kol_Smetka_Cells_Format(ws, "O1", "ОСВЕТЛЕНИЕ")
        Kol_Smetka_Cells_Format(ws, "S1", "КОНТАКТИ")
        Kol_Smetka_Cells_Format(ws, "AC1", "БОЙЛЕРИ")
        Kol_Smetka_Cells_Format(ws, "AL1", "ИНТЕРНЕТ")
        Kol_Smetka_Cells_Format(ws, "AR1", "ПИЦ")
        Kol_Smetka_Cells_Format(ws, "AU1", "ДОМОФОН")
        Kol_Smetka_Cells_Format(ws, "AW1", "ЗАЩИТНА")
        Kol_Smetka_Cells_Format(ws, "AY1", "МЪЛНИЯ")
        Kol_Smetka_Cells_Format(ws, "BB1", "ВИДЕО")
        Kol_Smetka_Cells_Format(ws, "BF1", "СОТ")
        Kol_Smetka_Cells_Format(ws, "BG1", "Kачване")
        Kol_Smetka_Cells_Format(ws, "BH1", "Брой")

        ws.Outline.ShowLevels(0, 1)

    End Sub
    Private Sub clearKa4vane(ws As excel.Worksheet      ' Лист в който да се изтрива
                            )
        With ws
            With .Range("A:Z")
                .Clear()
                .Value = ""
                .ColumnWidth = 15
                .Font.Name = "Cambria"
                .Font.Size = 12
                .WrapText = vbTrue
            End With
            .Columns("A").ColumnWidth = 10
            .Columns("B").ColumnWidth = 15
            .Columns("N").ColumnWidth = 10
            .Columns("O").ColumnWidth = 15
            .Range("A1").Value = "Дължина"
            .Range("B1").Value = "Полагане"
            .Range("N1").Value = "Дължина"
            .Range("O1").Value = "Полагане"
        End With
    End Sub
    Private Sub Kol_Smetka_Cells_Format(ws As excel.Worksheet,
                                        Range As String,
                                        Text As String)
        With ws.Range(Range)
            .Value = Text
            .ColumnWidth = 3
            .HorizontalAlignment = excel.XlHAlign.xlHAlignGeneral
            .VerticalAlignment = excel.XlVAlign.xlVAlignBottom
            .Font.Bold = vbTrue
            .WrapText = True
            .Orientation = 90
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .MergeCells = False
        End With
    End Sub
    ''' <summary>
    ''' Обработва промяната на стойностите на контроли, свързани с изчисленията за кабели, контакти и етажи.
    ''' Тази процедура се извиква автоматично при промяна на следните контроли:
    ''' - Брой_Етажи: брой етажи в сградата
    ''' - Височина_Етажи: височина на всеки етаж
    ''' - Koef_Zapas: коефициент на запас за кабели
    ''' - Kabel_Rozetka: дължина на кабел за розетки
    ''' - Kabel_kutiq: дължина на кабел за разпределителни кутии
    ''' - Wiso4ina_Kontakti: височина на контакти
    ''' - Wiso4ina_Kl: височина на ключове
    ''' 
    ''' Процедурата актуализира съответните променливи и, ако Excel работният лист е наличен,
    ''' записва тези стойности в клетки за последващи изчисления или справки.
    ''' </summary>
    ''' <param name="sender">Контролът, който е предизвикал промяната.</param>
    ''' <param name="e">Аргументи на събитието (EventArgs).</param>
    Private Sub Брой_Етажи_ValueChanged(sender As Object, e As EventArgs) Handles Брой_Етажи.ValueChanged,
                                                                                  Височина_Етажи.ValueChanged,
                                                                                  Koef_Zapas.ValueChanged,
                                                                                  Kabel_Rozetka.ValueChanged,
                                                                                  Kabel_kutiq.ValueChanged,
                                                                                  Wiso4ina_Kontakti.ValueChanged,
                                                                                  Wiso4ina_Kl.ValueChanged

        '---------------------------------------------------------------
        ' Проверка дали коефициентите са вече заредени.
        ' Ако koefYes е False, това означава, че някои начални стойности все още
        ' не са прочетени, затова процедурата прекратява изпълнението си.
        ' Това предотвратява некоректни изчисления или грешки.
        '---------------------------------------------------------------
        If Not koefYes Then Exit Sub

        '---------------------------------------------------------------
        ' Присвояване на стойности от контролите на съответните променливи.
        ' Това прави променливите достъпни за изчисления в останалата част от програмата.
        '---------------------------------------------------------------
        Бр_Етажи = Брой_Етажи.Value        ' Запис на броя етажи
        H_Етаж = Височина_Етажи.Value     ' Височина на един етаж
        Kz = Koef_Zapas.Value              ' Коефициент за запас на кабели
        Кабел_Розетка = Kabel_Rozetka.Value ' Дължина на кабела за розетки
        Кабел_кутия = Kabel_kutiq.Value      ' Дължина на кабела за разпределителни кутии
        H_Контакт = Wiso4ina_Kontakti.Value  ' Височина на контактите
        H_Ключ = Wiso4ina_Kl.Value           ' Височина на ключовете

        '---------------------------------------------------------------
        ' Ако работният Excel лист wsKoef съществува, записваме стойностите в клетки.
        ' Това е важно, защото другите процедури може да използват тези клетки
        ' за изчисления, графики или за справки от потребителя.
        '---------------------------------------------------------------
        If Not IsNothing(wsKoef) Then
            wsKoef.Cells(1, 2).Value = Брой_Етажи.Value       ' Брой етажи → клетка B1
            wsKoef.Cells(2, 2).Value = Височина_Етажи.Value  ' Височина на етажите → клетка B2
            wsKoef.Cells(3, 2).Value = Koef_Zapas.Value       ' Коефициент на запас → клетка B3
            wsKoef.Cells(4, 2).Value = Kabel_Rozetka.Value    ' Кабел за розетки → клетка B4
            wsKoef.Cells(5, 2).Value = Kabel_kutiq.Value      ' Кабел за кутии → клетка B5
            wsKoef.Cells(6, 2).Value = Wiso4ina_Kontakti.Value ' Височина на контакти → клетка B6
            wsKoef.Cells(7, 2).Value = Wiso4ina_Kl.Value      ' Височина на ключове → клетка B7
        End If

        '---------------------------------------------------------------
        ' Потвърждаваме, че коефициентите са вече заредени и че може да се извършват
        ' изчисления и записи без риск от грешки.
        '---------------------------------------------------------------
        koefYes = vbTrue
    End Sub
    Private Sub Button_Изчисти_Силови_Кабели_Click(sender As Object, e As EventArgs) Handles Button_Изчисти_Силови_Кабели.Click,
                                                                                             Button_Изчисти_ПИЦ_Кабели.Click,
                                                                                             Button_Изчисти_Интернет_Кабели.Click,
                                                                                             Button_Изчисти_ДОМОФ_Кабели.Click,
                                                                                             Button_Изчисти_МЪЛНИЯ_Кабели.Click,
                                                                                             Button_Изчисти_ВИДЕО_Кабели.Click,
                                                                                             Button_Изчисти_СОТ_Кабели.Click,
                                                                                             Button_Изчисти_ВЪНШНО_Кабели.Click,
                                                                                             Button_Изчисти_ВЪНШНО_Траншея.Click,
                                                                                             Button_Изчисти_ФОТОВОЛТАИЦИ_Кабели.Click,
                                                                                             Button_Изчисти_ФОТОВОЛТАИЦИ_Траншея.Click

        Dim minRows As Integer = 10000
        Dim maxRows As Integer = 0
        Dim Rows As String = "3:3"
        Dim Инсталация As String = ""
        Select Case sender.name
            Case "Button_Изчисти_ФОТОВОЛТАИЦИ_Кабели"
                Инсталация = "ФОТОВОЛТАИЦИ"
            Case "Button_Изчисти_ВЪНШНО_Траншея", "Button_Изчисти_ФОТОВОЛТАИЦИ_Траншея"
                Инсталация = "ТРАНШЕЯ"
            Case "Button_Изчисти_Интернет_Кабели"
                Инсталация = "ИНТЕРНЕТ"
            Case "Button_Изчисти_Силови_Кабели"
                Инсталация = "СИЛОВА"
            Case "Button_Изчисти_ПИЦ_Кабели"
                Инсталация = "ПИЦ"
            Case "Button_Изчисти_ДОМОФ_Кабели"
                Инсталация = "ДОМОФ"
            Case "Button_Изчисти_МЪЛНИЯ_Кабели"
                Инсталация = "МЪЛНИЯ"
            Case "Button_Изчисти_ЗАЗЕМЛ"
                Инсталация = "ЗАЩИТНА"
            Case "Button_Изчисти_ВЪНШНО_Кабели_"
                Инсталация = "ВЪНШНО"
            Case "Button_Изчисти_ВИДЕО_Кабели"
                Инсталация = "ВИДЕО"
            Case "Button_Изчисти_СОТ_Кабели"
                Инсталация = "СОТ"
        End Select
        For i As Integer = 3 To 1000
            If wsLines.Cells(i, 1).value = "" Then Exit For
            If InStr(wsLines.Cells(i, 1).value, Инсталация) Then
                minRows = Math.Min(minRows, i)
                maxRows = Math.Max(maxRows, i)
            End If
        Next
        Call clearKabeli(wsLines, minRows, maxRows)
    End Sub
    '' Функция, която обработва различни видове инсталации (тръби, канали, крепежни елементи и др.).
    'Private Sub ОбработиИнсталация(wsLines As Worksheet, wsKol_Smetka As Worksheet, ByRef index As Integer, i As Integer,
    '                           ByRef Кабел_Тръба As Dictionary(Of String, Double),
    '                           ByRef Кабел_Kanal As Double, ByRef Кабел_Скара As Double,
    '                           ByRef Кабел_Конструкция As Double, ByRef Кабел_Креп_елем As Double,
    '                           ByRef Кабел_Въже As Double, ByRef Кабел_ПКОМ As Double, ByRef Кабел_Въздушно As Double)
    '    ' Извлича типа на кабела от текущия ред (например "PVC", "HDP" и т.н.).
    '    Dim типКабел As String = Mid(wsLines.Cells(i, 2).Value, 8, 3)
    '    ' Извлича дължината на кабела от текущия ред.
    '    Dim дължина As Double = wsLines.Cells(i, 6).Value
    '    ' Използва Select Case конструкция за различни типове кабели и тръби.
    '    Select Case типКабел
    '    ' Актуализира стойностите на тръбите в речника.
    '        Case "PVC", "PE ", "HDP", "мет", "нег"
    '            Кабел_Тръба(типКабел) += дължина
    '        Case Else
    '            ' Проверява за други видове инсталации и актуализира съответните променливи.
    '            If wsLines.Cells(i, 2).Value.ToString().Contains("въже") Then
    '                Кабел_Въже += дължина
    '            ElseIf wsLines.Cells(i, 2).Value.ToString().Contains("ПКОМ") Then
    '                Кабел_ПКОМ += дължина
    '            ElseIf wsLines.Cells(i, 2).Value.ToString().Contains("скара") Then
    '                Кабел_Скара += дължина
    '            ElseIf wsLines.Cells(i, 2).Value = "по конструкция" Then
    '                Кабел_Конструкция += дължина
    '            ElseIf wsLines.Cells(i, 2).Value = "крепежни елементи" Then
    '                Кабел_Креп_елем += дължина
    '            ElseIf wsLines.Cells(i, 2).Value = "изт. въздушно" Then
    '                Кабел_Въздушно += дължина
    '            ElseIf Mid(wsLines.Cells(i, 2).Value, 8, 4) = "каб." Then
    '                Кабел_Kanal += дължина
    '            End If
    '    End Select
    'End Sub


    ' Основната функция, която изчислява и записва различни видове кабелни инсталации в Excel лист.
    Private Function Kol_Smetka_Kabeli(index As Integer,                    ' Ред от който да започне записването.
                                       Kabel_Obuvka As Boolean,             ' Да проверява ли за кабелни обувки
                                       Инсталация As String                 ' Вид на инсталацията
                                       ) As Integer
        Dim Кабел_Общо As Double = 0
        Dim Кабел_Канал As Double = 0
        Dim Кабел_Тръба_PVC As Double = 0
        Dim Кабел_Тръба_PE As Double = 0
        Dim Кабел_Тръба_HDPE As Double = 0
        Dim Кабел_Тръба_MET As Double = 0
        Dim Кабел_Тръба_НЕГОР As Double = 0
        Dim Кабел_Въже As Double = 0
        Dim Кабел_ПКОМ As Double = 0
        Dim Кабел_Скара As Double = 0
        Dim Кабел_Конструкция As Double = 0
        Dim Кабел_Креп_елем As Double = 0
        Dim Кабел_Изкоп As Double = 0
        Dim Кабел_Въздушно As Double = 0
        Dim Тип_Инсталация As String = Инсталация & "-КАБЕЛ"
        Dim Вид_Инсталация As String = Инсталация & "-ПОЛАГАНЕ"


        ProgressBar_Extrat.Maximum = ProgressBar_Maximum
        '
        ' Изчислява общата дължина на кабелите и доставя кабелите.
        ' Ако сечението на кабела е по-голямо от 16mm² прави суха разделка и доставя кабелните обувки.
        '
        For i = 3 To 200
            If Trim(wsLines.Cells(i, 2).Value) = "" Then Exit For
            ProgressBar_Extrat.Value = i
            If wsLines.Cells(i, 1).Value <> Тип_Инсталация Then Continue For
            With wsKol_Smetka
                .Cells(index, 2).Value = "Доставка на кабел " + wsLines.Cells(i, 2).Value
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "m"
                .Cells(index, 4).Value = wsLines.Cells(i, 6).Value
                Dim Кабел As String = wsLines.Cells(i, 2).Value
                If Kabel_Obuvka = True And (Кабел.IndexOf("САВТ") <> -1 Or Кабел.IndexOf("СВТ") <> -1) Then
                    Dim Материал As String = IIf(Кабел.IndexOf("САВТ") <> -1, "Al", "Cu")
                    Dim Позиция_Х As Integer = Кабел.IndexOf("x")
                    Dim Позиция_Плюс As Integer = Кабел.IndexOf("+")
                    Dim Позиция_MM As Integer = Кабел.IndexOf("mm")
                    Dim Сечение_1 As Integer = 0
                    Dim Сечение_2 As Integer = 0
                    Dim Брой_Жила_1 As Integer = Val(Кабел.Substring(Позиция_Х - 1, 1))
                    If Позиция_Плюс > -1 Then
                        Сечение_1 = Val(Кабел.Substring(Позиция_Х + 1, Позиция_Плюс - Позиция_Х - 1))
                        Сечение_2 = Val(Кабел.Substring(Позиция_Плюс + 1, Позиция_MM - Позиция_Плюс - 1))
                    Else
                        Сечение_1 = Val(Кабел.Substring(Позиция_Х + 1, Позиция_MM - Позиция_Х - 1))
                    End If

                    If Сечение_1 > 16 Then
                        index += 1
                        .Cells(index, 2).Value = "Направа на суха разделка на кабел " + wsLines.Cells(i, 2).Value
                        .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                        .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                        .Cells(index, 3).Value = "бр."
                        .Cells(index, 4).Value = wsLines.Cells(i, 60).Value * 2
                        If Позиция_Плюс > -1 Then
                            index += 1
                            .Cells(index, 2).Value = "Доставка и монтаж на кабелна обувка " + Материал + " " + Сечение_1.ToString + "mm²"
                            .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                            .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                            .Cells(index, 3).Value = "бр."
                            .Cells(index, 4).Value = wsLines.Cells(i, 60).Value * 2 * Брой_Жила_1
                            index += 1
                            .Cells(index, 2).Value = "Доставка и монтаж на кабелна обувка " + Материал + " " + Сечение_2.ToString + "mm²"
                            .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                            .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                            .Cells(index, 3).Value = "бр."
                            .Cells(index, 4).Value = wsLines.Cells(i, 60).Value * 2
                        Else
                            index += 1
                            .Cells(index, 2).Value = "Доставка и монтаж на кабелна обувка " + Материал + " " + Сечение_1.ToString + "mm²"
                            .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                            .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                            .Cells(index, 3).Value = "бр."
                            .Cells(index, 4).Value = wsLines.Cells(i, 60).Value * 2 * Брой_Жила_1
                        End If
                    End If
                End If
            End With
            Кабел_Общо = Кабел_Общо + wsLines.Cells(i, 6).Value
            index += 1
        Next
        '
        ' Изчислява общата дължина на тръби
        '
        For i = 3 To 200
            If Trim(wsLines.Cells(i, 2).Value) = "" Then Exit For
            ProgressBar_Extrat.Value = i
            If wsLines.Cells(i, 1).Value <> Вид_Инсталация Then Continue For
            Dim tryba As String = ""
            Dim kab As String = wsLines.Cells(i, 2).value
            Dim dyl As String = wsLines.Cells(i, 6).Value
            Select Case Mid(wsLines.Cells(i, 2).Value, 8, 3)
                Case "PVC"
                    tryba = "PVC гофрирана тръба ø"
                    tryba = tryba + Mid(kab, 16, Len(kab))
                    Кабел_Тръба_PVC = Кабел_Тръба_PVC + dyl
                Case "HDP"
                    tryba = "HDPE гофрирана тръба за подземно полагане ø"
                    tryba = tryba + Mid(kab, 17, Len(kab))
                    Кабел_Тръба_HDPE = Кабел_Тръба_HDPE + dyl
                Case "PE "
                    tryba = "полиетиленова гофрирана тръба ø"
                    tryba = tryba + Mid(kab, 15, Len(kab))
                    Кабел_Тръба_PE = Кабел_Тръба_PE + dyl
                Case "мет"
                    tryba = "метална гофрирана тръба ø"
                    tryba = tryba + Mid(kab, 16, Len(kab))
                    Кабел_Тръба_MET = Кабел_Тръба_MET + dyl
                Case "нег"
                    tryba = "гофрирана матална тръба с PVC покритие ø"
                    tryba = tryba + Mid(kab, 22, Len(kab))
                    Кабел_Тръба_НЕГОР = Кабел_Тръба_НЕГОР + dyl
            End Select
            If tryba = "" Then Continue For
            If dyl > 0 Then
                With wsKol_Smetka
                    .Cells(index, 2).Value = "Доставка и полагане на " + Trim(tryba)
                    .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                    .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                    .Cells(index, 3).Value = "m"
                    .Cells(index, 4).Value = dyl
                End With
                index += 1
            End If
        Next
        '
        ' Изчислява кабелни канали
        '
        For i = 3 To 200
            If Trim(wsLines.Cells(i, 2).Value) = "" Then Exit For
            ProgressBar_Extrat.Value = i
            If wsLines.Cells(i, 1).Value <> Вид_Инсталация Then Continue For
            Dim tryba As String = ""
            Dim kab As String = wsLines.Cells(i, 2).value
            Dim dyl As String = wsLines.Cells(i, 6).Value

            If Mid(wsLines.Cells(i, 2).Value, 8, 4) = "каб." Then
                tryba = Mid(kab, 16, Len(kab))
                With wsKol_Smetka
                    .Cells(index, 2).Value = "Доставка и монтаж на кабелен канал " + Trim(tryba)
                    .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                    .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                    .Cells(index, 3).Value = "m"
                    .Cells(index, 4).Value = dyl
                End With
                index += 1
                Кабел_Канал = Кабел_Канал + dyl
            End If
        Next
        '
        ' Изчислява носещо въже
        '
        For i = 3 To 200
            If Trim(wsLines.Cells(i, 2).Value) = "" Then Exit For
            ProgressBar_Extrat.Value = i
            If wsLines.Cells(i, 1).Value <> Вид_Инсталация Then Continue For
            Dim tryba As String = ""
            Dim kab As String = wsLines.Cells(i, 2).value
            Dim dyl As String = wsLines.Cells(i, 6).Value

            If wsLines.Cells(i, 2).Value.ToString().Contains("въже") Then
                tryba = tryba + Mid(kab, 22, Len(kab))
                With wsKol_Smetka
                    .Cells(index, 2).Value = "Доставка и монтаж на носещо въже" + Trim(tryba)
                    .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                    .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                    .Cells(index, 3).Value = "m"
                    .Cells(index, 4).Value = dyl
                End With
                Кабел_Въже = Кабел_Въже + dyl
                index += 1
            End If
        Next
        '
        ' Изчислява ПКОМ скоби
        '
        For i = 3 To 200
            If Trim(wsLines.Cells(i, 2).Value) = "" Then Exit For
            ProgressBar_Extrat.Value = i
            If wsLines.Cells(i, 1).Value <> Вид_Инсталация Then Continue For
            Dim tryba As String = ""
            Dim kab As String = wsLines.Cells(i, 2).value
            Dim dyl As String = wsLines.Cells(i, 6).Value

            If wsLines.Cells(i, 2).Value.ToString().Contains("ПКОМ") Then
                tryba = tryba + Mid(kab, 22, Len(kab))
                With wsKol_Smetka
                    .Cells(index, 2).Value = "Доставка и монтаж на ПКОМ скоби" + Trim(tryba)
                    .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                    .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                    .Cells(index, 3).Value = "бр."
                    .Cells(index, 4).Value = dyl / 0.5
                End With
                index += 1
                Кабел_ПКОМ = Кабел_ПКОМ + dyl
            End If
        Next
        '
        ' Изчислява кабелна скара
        '
        For i = 3 To 200
            If Trim(wsLines.Cells(i, 2).Value) = "" Then Exit For
            ProgressBar_Extrat.Value = i
            If wsLines.Cells(i, 1).Value <> Вид_Инсталация Then Continue For
            If wsLines.Cells(i, 2).Value.ToString().Contains("скара") Then
                Кабел_Скара = Кабел_Скара + wsLines.Cells(i, 6).Value
            End If
        Next
        '
        ' Изчислява кабел по метална конструкция
        '
        For i = 3 To 200
            If Trim(wsLines.Cells(i, 2).Value) = "" Then Exit For
            ProgressBar_Extrat.Value = i
            If wsLines.Cells(i, 1).Value <> Вид_Инсталация Then Continue For
            If wsLines.Cells(i, 2).Value = "по конструкция" Then
                Кабел_Конструкция = Кабел_Конструкция + wsLines.Cells(i, 6).Value
            End If
        Next
        '
        ' Изчислява кабел по крепежни елементи
        '
        For i = 3 To 200
            If Trim(wsLines.Cells(i, 2).Value) = "" Then Exit For
            ProgressBar_Extrat.Value = i
            If wsLines.Cells(i, 1).Value <> Вид_Инсталация Then Continue For
            If wsLines.Cells(i, 2).Value = "крепежни елементи" Then
                Кабел_Креп_елем = Кабел_Креп_елем + wsLines.Cells(i, 6).Value
            End If
        Next
        '
        ' Изчислява кабел ВЪЗДУШНО
        '
        For i = 3 To 200
            If Trim(wsLines.Cells(i, 2).Value) = "" Then Exit For
            ProgressBar_Extrat.Value = i
            If wsLines.Cells(i, 1).Value <> Вид_Инсталация Then Continue For
            If wsLines.Cells(i, 2).Value = "изт. въздушно" Then
                Кабел_Въздушно = Кабел_Въздушно + wsLines.Cells(i, 6).Value
            End If
        Next
        '
        ' Изтегляне на кабели в монтирани тръби
        '
        Dim kab_tryb As Double = Кабел_Общо -
                                 Кабел_Скара -
                                 Кабел_ПКОМ -
                                 Кабел_Въже -
                                 Кабел_Канал -
                                 Кабел_Конструкция -
                                 Кабел_Креп_елем -
                                 Кабел_Изкоп -
                                 Кабел_Въздушно

        If kab_tryb > 0 Then
            With wsKol_Smetka
                .Cells(index, 2).Value = "Изтегляне на кабели в монтирани тръби"
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "m"
                .Cells(index, 4).Value = kab_tryb
            End With
            index += 1
        End If
        '
        ' Полагане на кабели ВЪЗДУШНО
        '
        If Кабел_Въздушно > 0 Then
            With wsKol_Smetka
                .Cells(index, 2).Value = "Полагане на кабел въздушно"
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "m"
                .Cells(index, 4).Value = Кабел_Въздушно
            End With
            index += 1
        End If
        '
        ' Направа улей в тухлен зид
        '
        If (Кабел_Тръба_PVC + Кабел_Тръба_PE) > 0 Then
            With wsKol_Smetka
                .Cells(index, 2).Value = "Направа улей в тухлен зид"
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "m"
                .Cells(index, 4).Value = Кабел_Тръба_PVC + Кабел_Тръба_PE
            End With
            index += 1
        End If
        '
        ' Полагане на кабели в кабелни канали
        '
        If Кабел_Канал > 0 Then
            With wsKol_Smetka
                .Cells(index, 2).Value = "Полагане на кабели в кабелни канали"
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "m"
                .Cells(index, 4).Value = Кабел_Канал
            End With
            index += 1
        End If
        '
        ' Полагане на кабели по монтирано въже
        '
        If Кабел_Въже > 0 Then
            With wsKol_Smetka
                .Cells(index, 2).Value = "Полагане на кабели по монтирано въже"
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "m"
                .Cells(index, 4).Value = Кабел_Въже
            End With
            index += 1
        End If
        '
        ' Полагане на кабели по ПКОМ скоби
        '
        If Кабел_ПКОМ > 0 Then
            With wsKol_Smetka
                .Cells(index, 2).Value = "Полагане на кабели по ПКОМ скоби"
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "m"
                .Cells(index, 4).Value = Кабел_ПКОМ
            End With
            index += 1
        End If
        '
        ' Полагане на кабели по кабелна скара
        '
        If Кабел_Скара > 0 Then
            With wsKol_Smetka
                .Cells(index, 2).Value = "Полагане на кабели по кабелна скара"
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "m"
                .Cells(index, 4).Value = Кабел_Скара
            End With
            index += 1
        End If
        '
        ' Полагане на кабели по метална конструкция
        '
        If Кабел_Конструкция > 0 Then
            With wsKol_Smetka
                .Cells(index, 2).Value = "Полагане на кабели открито по метална конструкция"
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "m"
                .Cells(index, 4).Value = Кабел_Конструкция
            End With
            index += 1
        End If
        '
        ' Полагане на кабели по крепежни елементи за мълниезащита
        '
        If Кабел_Креп_елем > 0 Then
            With wsKol_Smetka
                .Cells(index, 2).Value = "Доставка и монтаж на комплект крепежни елементи"
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "бр."
                .Cells(index, 4).Value = Кабел_Креп_елем
            End With
            index += 1
            With wsKol_Smetka
                .Cells(index, 2).Value = "Полагане на проводник по монтирани крепежни елементи"
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "m"
                .Cells(index, 4).Value = Кабел_Креп_елем
            End With
            index += 1
        End If
        Return index ' Връща използваните редове
    End Function
    Private Sub Button_Изчисти_ПИЦ_Click(sender As Object, e As EventArgs) Handles Button_Изчисти_ПИЦ.Click
        clearPIC(wsPIC)
    End Sub
    Private Sub clearPIC(ws As excel.Worksheet)
        With ws
            With .Range("A:AK")
                .Clear()
                .Value = ""
            End With
            With .Range("A:AK")
                .Font.Name = "Cambria"
                .Font.Size = 12
                .WrapText = vbTrue
            End With
            .Columns("A").ColumnWidth = 6
            .Columns("B").ColumnWidth = 15
            .Columns("C").ColumnWidth = 40
            .Columns("D").ColumnWidth = 25
            .Columns("E").ColumnWidth = 8
            .Columns("F").ColumnWidth = 11
            .Columns("G").ColumnWidth = 2
            .Columns("H").ColumnWidth = 21
            .Columns("I").ColumnWidth = 21
            .Columns("J").ColumnWidth = 21
            .Columns("K:AG").ColumnWidth = 30
            .Cells(1, 1).Value = "Брой"
            .Cells(1, 2).Value = "Име"
            .Cells(1, 3).Value = "Visibility"
            .Cells(1, 8).Value = "Качване Двужилен"
            .Cells(1, 9).Value = "Качване Трижилен"
        End With
        If IsNothing(wsLines) Then Exit Sub
        With wsLines
            .Cells(2, 39).Value = 0
            .Cells(2, 40).Value = 0
            .Cells(2, 41).Value = 0
            .Cells(2, 42).Value = 0
        End With
    End Sub
    Private Sub clearТабла(ws As excel.Worksheet)
        With ws
            With .Range("A:AK")
                .Clear()
                .Value = ""
            End With
            With .Range("A:AK")
                .Font.Name = "Cambria"
                .Font.Size = 12
                .WrapText = vbTrue
            End With
            .Columns("A").ColumnWidth = 12
            .Columns("B").ColumnWidth = 33
            .Columns("C").ColumnWidth = 8
            .Columns("D").ColumnWidth = 27
            .Columns("E").ColumnWidth = 15
            .Columns("F").ColumnWidth = 22
            .Columns("G").ColumnWidth = 31
            .Columns("H").ColumnWidth = 25
            .Columns("I").ColumnWidth = 8
            .Columns("J").ColumnWidth = 8
            .Columns("K").ColumnWidth = 35
            .Columns("L").ColumnWidth = 85
            .Columns("M:Z").ColumnWidth = 15
            .Cells(1, 1).Value = "Табло"
            .Cells(1, 2).Value = "Вид"
            .Cells(1, 3).Value = "Брой"
            .Cells(1, 4).Value = "_1"
            .Cells(1, 5).Value = "_2"
            .Cells(1, 6).Value = "_3"
            .Cells(1, 7).Value = "_4"
            .Cells(1, 8).Value = "_5"
            .Cells(1, 9).Value = "_6"
            .Cells(1, 10).Value = "_7"
            .Cells(1, 11).Value = "Име на блок"
            .Cells(1, 12).Value = "Дълго име"
            .Cells(1, 13).Value = "+MX"
            .Cells(1, 15).Value = "Брой модули"
            .Cells(1, 16).Value = "Брой връзки до 2,5"
            .Cells(1, 17).Value = "Брой връзки до 16"
            .Cells(1, 18).Value = "Брой връзки над 16"
            .Cells(1, 19).Value = "Брой ДЗТ"
            .Cells(1, 20).Value = "Монтаж"

        End With
    End Sub
    Private Sub Button_Вземи_ПИЦ_ДАТЧИЦИ_Click(sender As Object, e As EventArgs) Handles Button_Вземи_ПИЦ_ДАТЧИЦИ.Click
        If IsNothing(excel_Workbook) Then
            MsgBox("Първо да беше отворил файла?")
            Exit Sub
        End If
        Me.Visible = vbFalse
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim edt As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database
        Dim SelectedSet = cu.GetObjects("INSERT", "Изберете блок")
        Dim arrBlock(500) As strKontakt
        Dim Качване(100) As strКачване

        If SelectedSet Is Nothing Then
            MsgBox("Няма маркиран нито един блок.")
            Me.Visible = vbTrue
            Exit Sub
        End If
        Dim blkRecId As ObjectId = ObjectId.Null
        Dim index As Integer = 0
        Dim index_Качване As Integer = 0
        Dim index_Row As Integer = 0
        Dim broj_konturi As Integer = 0
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                For Each sObj As SelectedObject In SelectedSet
                    blkRecId = sObj.ObjectId

                    Dim acBlkRef As BlockReference = DirectCast(acTrans.GetObject(blkRecId, OpenMode.ForRead), BlockReference)
                    Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                    Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection
                    Dim Visibility As String = ""

                    Dim strKOTA_1 As String = ""
                    Dim strKOTA_2 As String = ""
                    Dim strТРЪБА_1 As String = ""
                    Dim strТРЪБА_2 As String = ""
                    Dim strKabel_d_0 As String = ""
                    Dim strKabel_d_1 As String = ""
                    Dim strKabel_d_2 As String = ""
                    Dim strKabel_d_3 As String = ""
                    Dim strKabel_d_4 As String = ""
                    Dim strKabel_d_5 As String = ""
                    Dim strKabel_d_6 As String = ""
                    Dim strKabel_d_7 As String = ""
                    Dim strKabel_d_8 As String = ""
                    Dim strKabel_d_9 As String = ""
                    Dim strKabel_d_10 As String = ""
                    Dim strKabel_g_0 As String = ""
                    Dim strKabel_g_1 As String = ""
                    Dim strKabel_g_2 As String = ""
                    Dim strKabel_g_3 As String = ""
                    Dim strKabel_g_4 As String = ""
                    Dim strKabel_g_5 As String = ""
                    Dim strKabel_g_6 As String = ""
                    Dim strKabel_g_7 As String = ""
                    Dim strKabel_g_8 As String = ""
                    Dim strKabel_g_9 As String = ""
                    Dim strKabel_g_10 As String = ""

                    For Each prop As DynamicBlockReferenceProperty In props
                        If prop.PropertyName = "Visibility" Then Visibility = prop.Value
                    Next

                    Dim blName As String = (CType(acBlkRef.DynamicBlockTableRecord.GetObject(OpenMode.ForRead), BlockTableRecord)).Name

                    Dim iVisib As Integer = -1
                    Select Case blName
                        Case "Датчик_ПАБ"
                            iVisib = Array.FindIndex(arrBlock, Function(f) f.blName = blName And
                                                                           f.blVisibility = Visibility)
                            For Each objID As ObjectId In attCol
                                Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                                Dim acAttRef As AttributeReference = dbObj
                                If acAttRef.Tag = "ZN" Then
                                    Dim sss As Integer = Val(acAttRef.TextString)
                                    broj_konturi = Math.Max(broj_konturi, sss)
                                End If
                            Next
                        Case "Кабел"
                            Continue For
                        Case "Качване"
                            For Each objID As ObjectId In attCol
                                Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                                Dim acAttRef As AttributeReference = dbObj

                                If acAttRef.Tag = "KOTA_1" Then strKOTA_1 = acAttRef.TextString
                                If acAttRef.Tag = "KOTA_2" Then strKOTA_2 = acAttRef.TextString
                                If acAttRef.Tag = "ТРЪБА_1" Then strТРЪБА_1 = acAttRef.TextString
                                If acAttRef.Tag = "ТРЪБА_2" Then strТРЪБА_2 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_0" Then strKabel_d_0 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_1" Then strKabel_d_1 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_2" Then strKabel_d_2 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_3" Then strKabel_d_3 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_4" Then strKabel_d_4 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_5" Then strKabel_d_5 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_6" Then strKabel_d_6 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_7" Then strKabel_d_7 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_8" Then strKabel_d_8 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_9" Then strKabel_d_9 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_10" Then strKabel_d_10 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_0" Then strKabel_g_0 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_1" Then strKabel_g_1 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_2" Then strKabel_g_2 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_3" Then strKabel_g_3 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_4" Then strKabel_g_4 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_5" Then strKabel_g_5 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_6" Then strKabel_g_6 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_7" Then strKabel_g_7 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_8" Then strKabel_g_8 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_9" Then strKabel_g_9 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_10" Then strKabel_g_10 = acAttRef.TextString
                            Next
                        Case Else
                            Continue For
                    End Select
                    If iVisib = -1 Then
                        arrBlock(index).count = 1
                        arrBlock(index).blName = blName
                        arrBlock(index).blVisibility = Visibility
                        If blName = "Качване" Then
                            Качване(index_Качване).KOTA_1 = strKOTA_1
                            Качване(index_Качване).KOTA_2 = strKOTA_2
                            Качване(index_Качване).ТРЪБА_1 = strТРЪБА_1
                            Качване(index_Качване).ТРЪБА_2 = strТРЪБА_2
                            Качване(index_Качване).Kabel_d_0 = strKabel_d_0
                            Качване(index_Качване).Kabel_d_1 = strKabel_d_1
                            Качване(index_Качване).Kabel_d_2 = strKabel_d_2
                            Качване(index_Качване).Kabel_d_3 = strKabel_d_3
                            Качване(index_Качване).Kabel_d_6 = strKabel_d_6
                            Качване(index_Качване).Kabel_d_7 = strKabel_d_7
                            Качване(index_Качване).Kabel_d_4 = strKabel_d_4
                            Качване(index_Качване).Kabel_d_5 = strKabel_d_5
                            Качване(index_Качване).Kabel_d_8 = strKabel_d_8
                            Качване(index_Качване).Kabel_d_9 = strKabel_d_9
                            Качване(index_Качване).Kabel_d_10 = strKabel_d_10
                            Качване(index_Качване).Kabel_g_0 = strKabel_g_0
                            Качване(index_Качване).Kabel_g_1 = strKabel_g_1
                            Качване(index_Качване).Kabel_g_2 = strKabel_g_2
                            Качване(index_Качване).Kabel_g_3 = strKabel_g_3
                            Качване(index_Качване).Kabel_g_4 = strKabel_g_4
                            Качване(index_Качване).Kabel_g_5 = strKabel_g_5
                            Качване(index_Качване).Kabel_g_6 = strKabel_g_6
                            Качване(index_Качване).Kabel_g_7 = strKabel_g_7
                            Качване(index_Качване).Kabel_g_8 = strKabel_g_8
                            Качване(index_Качване).Kabel_g_9 = strKabel_g_9
                            Качване(index_Качване).Kabel_g_10 = strKabel_g_10
                            index_Качване += 1
                        End If
                        index += 1
                    Else
                        arrBlock(iVisib).count = arrBlock(iVisib).count + 1
                    End If
                Next

                index_Качване = 0

                Call clearPIC(wsPIC)

                wsPIC.Cells(1, 4).Value = "ПИЦ-контури"
                wsPIC.Cells(1, 5).Value = broj_konturi.ToString

                index_Row = 2
                Dim dylvina1 As Double = 0
                Dim dylvina2 As Double = 0
                Dim br_linii As Double = 0
                Dim Kabel_2x05 As Double = 0
                Dim Kabel_3x05 As Double = 0

                For Each iarrBlock In arrBlock
                    If iarrBlock.count = 0 Then Exit For
                    With wsPIC
                        .Cells(index_Row, 1) = iarrBlock.count
                        .Cells(index_Row, 2) = iarrBlock.blName
                        .Cells(index_Row, 3) = iarrBlock.blVisibility
                        If iarrBlock.blName = "Качване" Then
                            .Cells(index_Row, 11) = Качване(index_Качване).KOTA_1
                            .Cells(index_Row, 12) = Качване(index_Качване).ТРЪБА_1
                            .Cells(index_Row, 13) = Качване(index_Качване).Kabel_d_0
                            .Cells(index_Row, 14) = Качване(index_Качване).Kabel_d_1
                            .Cells(index_Row, 15) = Качване(index_Качване).Kabel_d_2
                            .Cells(index_Row, 16) = Качване(index_Качване).Kabel_d_3
                            .Cells(index_Row, 17) = Качване(index_Качване).Kabel_d_6
                            .Cells(index_Row, 18) = Качване(index_Качване).Kabel_d_5
                            .Cells(index_Row, 19) = Качване(index_Качване).Kabel_d_8
                            .Cells(index_Row, 20) = Качване(index_Качване).Kabel_d_7
                            .Cells(index_Row, 21) = Качване(index_Качване).Kabel_d_4
                            .Cells(index_Row, 22) = Качване(index_Качване).Kabel_d_9
                            .Cells(index_Row, 23) = Качване(index_Качване).Kabel_d_10
                            .Cells(index_Row, 24) = Качване(index_Качване).KOTA_2
                            .Cells(index_Row, 25) = Качване(index_Качване).ТРЪБА_2
                            .Cells(index_Row, 26) = Качване(index_Качване).Kabel_g_0
                            .Cells(index_Row, 27) = Качване(index_Качване).Kabel_g_1
                            .Cells(index_Row, 28) = Качване(index_Качване).Kabel_g_2
                            .Cells(index_Row, 29) = Качване(index_Качване).Kabel_g_3
                            .Cells(index_Row, 30) = Качване(index_Качване).Kabel_g_4
                            .Cells(index_Row, 31) = Качване(index_Качване).Kabel_g_5
                            .Cells(index_Row, 32) = Качване(index_Качване).Kabel_g_6
                            .Cells(index_Row, 33) = Качване(index_Качване).Kabel_g_7
                            .Cells(index_Row, 34) = Качване(index_Качване).Kabel_g_8
                            .Cells(index_Row, 35) = Качване(index_Качване).Kabel_g_9
                            .Cells(index_Row, 36) = Качване(index_Качване).Kabel_g_10
                            index_Качване += 1

                            dylvina1 = Val(.Cells(index_Row, 11).value)
                            dylvina2 = Val(.Cells(index_Row, 24).value)
                            For i = 11 To 36
                                br_linii = InStr(.Cells(index_Row, i).value, "л.")
                                If br_linii = 0 Then Continue For
                                br_linii = Mid(.Cells(index_Row, i).value, 1, br_linii - 1)
                                If InStr(.Cells(index_Row, i).value, "FS 2x0,") > 1 Then
                                    .Cells(index_Row, 8).value =
                                        IIf(i < 24, br_linii * dylvina1, br_linii * dylvina2) +
                                        .Cells(index_Row, 8).value
                                    Kabel_2x05 = Kabel_2x05 + .Cells(index_Row, 8).value
                                End If
                                If InStr(.Cells(index_Row, i).value, "FS 3x0,") > 1 Then
                                    .Cells(index_Row, 9).value =
                                        IIf(i < 24, br_linii * dylvina1, br_linii * dylvina2) +
                                        .Cells(index_Row, 9).value
                                    Kabel_3x05 = Kabel_3x05 + .Cells(index_Row, 9).value
                                End If
                            Next
                        End If
                        If iarrBlock.blName = "Датчик_ПАБ" Then
                            Select Case iarrBlock.blVisibility
                                Case "Ръчен пожароизвестител адресируем",
                                     "Изпълнително устройство",
                                     "Ръчен пожароизвестител конвенционален"
                                    wsLines.Cells(2, 41).Value = wsLines.Cells(2, 41).Value +
                                                                iarrBlock.count
                                Case "ПАБ - Сирена и Звук",
                                     "ПАБ - Лампа"
                                    wsLines.Cells(2, 40).Value = wsLines.Cells(2, 40).Value +
                                                                iarrBlock.count
                                Case Else
                                    wsLines.Cells(2, 39).Value = wsLines.Cells(2, 39).Value +
                                                                iarrBlock.count
                            End Select
                        End If
                    End With
                    index_Row += 1
                Next

                wsLines.Cells(2, 42).Value = Kabel_2x05 / 2
                wsLines.Cells(2, 43).Value = Kabel_3x05 / 2
                With wsPIC
                    .Cells.Sort(Key1:= .Range("B2"),
                                Order1:=excel.XlSortOrder.xlAscending,
                                Header:=excel.XlYesNoGuess.xlYes,
                                OrderCustom:=1, MatchCase:=False,
                                Orientation:=excel.Constants.xlTopToBottom,
                                DataOption1:=excel.XlSortDataOption.xlSortTextAsNumbers,
                                Key2:= .Range("D2"),
                                Order2:=excel.XlSortOrder.xlAscending,
                                DataOption2:=excel.XlSortDataOption.xlSortTextAsNumbers
                                )
                End With
                acTrans.Commit()
            Catch ex As Exception
                MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
                acTrans.Abort()
            End Try
        End Using
        Me.Visible = vbTrue
    End Sub
    Private Sub Button_Генерирай_ПИЦ_Click(sender As Object, e As EventArgs) Handles Button_Генерирай_ПИЦ.Click
        Dim index As Integer = 0
        Dim index_internet As Integer = 0
        Dim index_Kabel As Integer = 0
        Dim i As Integer
        ProgressBar_Extrat.Maximum = ProgressBar_Maximum
        For i = 6 To 1000
            If wsKol_Smetka.Cells(i, 2).Value = "Раздел" Or wsKol_Smetka.Cells(i, 2).Value = "" Then
                index = i
                Exit For
            End If
            ProgressBar_Extrat.Value = i
        Next
        For i = 2 To 1000
            If wsPIC.Cells(i, 2).Value = "" Then
                Exit For
            End If
        Next
        index_internet = i
        ProgressBar_Extrat.Maximum = i
        Dim Text_Dostawka As String = ""
        Dim broj_elementi As Integer = 0
        Call Excel_Kol_smetka_Razdel(wsKol_Smetka, "СЛАБОТОКОВИ ИНСТАЛАЦИИ", "A" & Trim(index.ToString), "D" & Trim(index.ToString))
        index += 1
        Call Excel_Kol_smetka_Razdel(wsKol_Smetka, "ПОЖАРОИЗВЕСТИТЕЛНА ИНСТАЛАЦИЯ", "B" & Trim(index.ToString), "D" & Trim(index.ToString))
        Dim Kabel(1000) As strKabel
        index += 1
        Text_Dostawka = "Доставка и монтаж на пожароизвестителна централа; "
        Text_Dostawka = Text_Dostawka & IIf(RadioButton_Адресируема.Checked, "адресируема;", "конвенционална;")
        Text_Dostawka = Text_Dostawka & " Брой "
        Text_Dostawka = Text_Dostawka & IIf(RadioButton_Адресируема.Checked, "контури", "зони")
        Text_Dostawka = Text_Dostawka & ": " & wsPIC.Cells(1, 5).Value
        With wsKol_Smetka
            .Cells(index, 2).Value = Text_Dostawka
            .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
            .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
            .Cells(index, 3).Value = "бр."
            .Cells(index, 4).Value = 1
        End With
        index += 1
        Dim osnowi_konw As Integer = 0
        Dim osnowi_adres As Integer = 0

        For i = 2 To 10000
            If Trim(wsPIC.Cells(i, 2).Value) = "" Then Exit For
            ProgressBar_Extrat.Value = i
            Text_Dostawka = ""
            broj_elementi = wsPIC.Cells(i, 1).Value

            Select Case wsPIC.Cells(i, 2).Value
                Case "Датчик_ПАБ"
                    Select Case wsPIC.Cells(i, 3).Value
                        Case "Ръчен пожароизвестител конвенционален"
                            Text_Dostawka = "ръчен пожароизвестител - конвенционален"
                        Case "ПАБ - Термичен конвенционален комбиниран"
                            Text_Dostawka = "пожароизвестител конвенционален; топлинен; комбиниран топлинен диференциален и оптично-димен"                    '
                            osnowi_konw = osnowi_konw + broj_elementi
                        Case "ПАБ - Термичен конвенционален диференциален"
                            Text_Dostawka = "пожароизвестител конвенционален; топлинен; диференциален"
                            osnowi_konw = osnowi_konw + broj_elementi
                        Case "ПАБ - Термичен конвенционален"
                            Text_Dostawka = "пожароизвестител конвенционален; топлинен; максимален"
                            osnowi_konw = osnowi_konw + broj_elementi
                        Case "ПАБ - Пламъков конвенционален"
                            Text_Dostawka = "пожароизвестител конвенционален; оптичен; пламъков"
                            osnowi_konw = osnowi_konw + broj_elementi
                        Case "ПАБ - Димооптичен конвенционален"
                            Text_Dostawka = "пожароизвестител конвенционален; оптично-димен; самокомпенсация на замърсяването"
                            osnowi_konw = osnowi_konw + broj_elementi
                        Case "ПАБ - Сирена конвенционална"
                            Text_Dostawka = "конвенционална сирена"
                        Case "ПАБ - Лампа"
                            Text_Dostawka = "паралелен сигнализатор"
                        Case "ПАБ - Сирена и Звук"
                            Text_Dostawka = "светлинно-звуков сигнализатор; външен монтаж"
                        Case "Ръчен пожароизвестител адресируем"
                            Text_Dostawka = "ръчен пожароизвестител - адресируем"
                        Case "Изпълнително устройство"
                            Text_Dostawka = "адресируемо входно - изходно устройство; 1-вход; 1-изход"
                        Case "ПАБ - Сирена адресируема"
                            Text_Dostawka = "адресируема сирена"
                        Case "ПАБ - Термичен адресируем с адаптер-7120"
                            Text_Dostawka = "пожароизвестител адресируем; топлинен; диференциален; вграден адаптер"
                            osnowi_adres = osnowi_adres + broj_elementi
                        Case "ПАБ - Термичен адресируем комбиниран"
                            Text_Dostawka = "пожароизвестител адресируем; топлинен; комбиниран топлинен диференциален и оптично-димен"
                            osnowi_adres = osnowi_adres + broj_elementi
                        Case "ПАБ - Термичен адресируем диференциален"
                            Text_Dostawka = "пожароизвестител адресируем; топлинен; диференциален"
                            osnowi_adres = osnowi_adres + broj_elementi
                        Case "ПАБ - Термичен адресируем - 7101"
                            Text_Dostawka = "пожароизвестител адресируем; топлинен; максимален"
                            osnowi_adres = osnowi_adres + broj_elementi
                        Case "ПАБ - Димооптичен адресируем"
                            Text_Dostawka = "пожароизвестител адресируем; оптично-димен; самокомпенсация на замърсяването"
                            osnowi_adres = osnowi_adres + broj_elementi
                        Case "Линеен оптично димен приемник"
                            Text_Dostawka = "линеен оптично димен пожароизвестител; комплект"
                        Case "Изолатор"
                        Case "Зона", "Линеен оптично димен излъчвател"
                        Case Else
                            Text_Dostawka = "НЕИЗВЕСТЕН ЕЛЕМЕНТ --> " & wsPIC.Cells(i, 2).Value & " - " & wsPIC.Cells(i, 3).Value
                    End Select
                    index_Kabel += 1
                Case "Качване"
                    Continue For
            End Select
            If Trim(Text_Dostawka) <> "" Then
                With wsKol_Smetka
                    .Cells(index, 2).Value = "Доставка и монтаж на " + Text_Dostawka
                    .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                    .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                    .Cells(index, 3).Value = "бр."
                    .Cells(index, 4).Value = broj_elementi
                End With
                index += 1
            End If
        Next

        index = Kol_Smetka_Kabeli(index, vbFalse, "ПИЦ")

        If osnowi_konw > 0 Then
            Text_Dostawka = "основа за пожароизвестител конвенционален"
            With wsKol_Smetka
                .Cells(index, 2).Value = "Доставка и монтаж на " + Text_Dostawka
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "бр."
                .Cells(index, 4).Value = osnowi_konw
            End With
            index += 1
        End If
        If osnowi_adres > 0 Then
            Text_Dostawka = "основа за пожароизвестител адресируем"
            With wsKol_Smetka
                .Cells(index, 2).Value = "Доставка и монтаж на " + Text_Dostawka
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "бр."
                .Cells(index, 4).Value = osnowi_adres
            End With
            index += 1
        End If
        Text_Dostawka = "Доставка на стъкла за ръчен пожароизвестител; комплект 10бр."
        With wsKol_Smetka
            .Cells(index, 2).Value = Text_Dostawka
            .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
            .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
            .Cells(index, 3).Value = "бр."
            .Cells(index, 4).Value = 1
        End With
        index += 1
        Text_Dostawka = "акумулаторна батерия за пожароизвестителна централа 12V, 7Ah"
        With wsKol_Smetka
            .Cells(index, 2).Value = "Доставка и монтаж на " + Text_Dostawka
            .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
            .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
            .Cells(index, 3).Value = "бр."
            .Cells(index, 4).Value = 2
        End With
        index += 1
        Text_Dostawka = "Комплексно изпитване на системата"
        With wsKol_Smetka
            .Cells(index, 2).Value = Text_Dostawka
            .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
            .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
            .Cells(index, 3).Value = "бр."
            .Cells(index, 4).Value = 1
        End With
        index += 1
        Text_Dostawka = "Обучение на персонал за работа със системата"
        With wsKol_Smetka
            .Cells(index, 2).Value = Text_Dostawka
            .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
            .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
            .Cells(index, 3).Value = "бр."
            .Cells(index, 4).Value = 1
        End With
        index += 1
        ProgressBar_Extrat.Value = ProgressBar_Extrat.Minimum
    End Sub
    Private Sub Button_Вземи_ДОМОФ_ДАТЧИЦИ_Click(sender As Object, e As EventArgs) Handles Button_Вземи_ДОМОФ_ДАТЧИЦИ.Click
        If IsNothing(excel_Workbook) Then
            MsgBox("Първо да беше отворил файла?")
            Exit Sub
        End If
        Me.Visible = vbFalse
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim edt As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database
        Dim SelectedSet = cu.GetObjects("INSERT", "Изберете блок")
        Dim arrBlock(500) As strKontakt
        Dim Качване(100) As strКачване

        If SelectedSet Is Nothing Then
            MsgBox("Няма маркиран нито един блок.")
            Me.Visible = vbTrue
            Exit Sub
        End If
        Dim blkRecId As ObjectId = ObjectId.Null
        Dim index As Integer = 0
        Dim index_Качване As Integer = 0
        Dim index_Row As Integer = 0
        Dim broj_konturi As Integer = 0
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                For Each sObj As SelectedObject In SelectedSet
                    blkRecId = sObj.ObjectId

                    Dim acBlkRef As BlockReference = DirectCast(acTrans.GetObject(blkRecId, OpenMode.ForRead), BlockReference)
                    Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                    Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection
                    Dim Visibility As String = ""

                    Dim strKOTA_1 As String = ""
                    Dim strKOTA_2 As String = ""
                    Dim strТРЪБА_1 As String = ""
                    Dim strТРЪБА_2 As String = ""
                    Dim strKabel_d_0 As String = ""
                    Dim strKabel_d_1 As String = ""
                    Dim strKabel_d_2 As String = ""
                    Dim strKabel_d_3 As String = ""
                    Dim strKabel_d_4 As String = ""
                    Dim strKabel_d_5 As String = ""
                    Dim strKabel_d_6 As String = ""
                    Dim strKabel_d_7 As String = ""
                    Dim strKabel_d_8 As String = ""
                    Dim strKabel_d_9 As String = ""
                    Dim strKabel_d_10 As String = ""
                    Dim strKabel_g_0 As String = ""
                    Dim strKabel_g_1 As String = ""
                    Dim strKabel_g_2 As String = ""
                    Dim strKabel_g_3 As String = ""
                    Dim strKabel_g_4 As String = ""
                    Dim strKabel_g_5 As String = ""
                    Dim strKabel_g_6 As String = ""
                    Dim strKabel_g_7 As String = ""
                    Dim strKabel_g_8 As String = ""
                    Dim strKabel_g_9 As String = ""
                    Dim strKabel_g_10 As String = ""

                    For Each prop As DynamicBlockReferenceProperty In props
                        If prop.PropertyName = "Visibility" Then Visibility = prop.Value
                    Next

                    Dim blName As String = (CType(acBlkRef.DynamicBlockTableRecord.GetObject(OpenMode.ForRead), BlockTableRecord)).Name

                    Dim iVisib As Integer = -1
                    Select Case blName
                        Case "Домофон"
                            iVisib = Array.FindIndex(arrBlock, Function(f) f.blName = blName And
                                                                           f.blVisibility = Visibility)
                        Case "Кабел"
                            Continue For
                        Case "Качване"
                            For Each objID As ObjectId In attCol
                                Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                                Dim acAttRef As AttributeReference = dbObj

                                If acAttRef.Tag = "KOTA_1" Then strKOTA_1 = acAttRef.TextString
                                If acAttRef.Tag = "KOTA_2" Then strKOTA_2 = acAttRef.TextString
                                If acAttRef.Tag = "ТРЪБА_1" Then strТРЪБА_1 = acAttRef.TextString
                                If acAttRef.Tag = "ТРЪБА_2" Then strТРЪБА_2 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_0" Then strKabel_d_0 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_1" Then strKabel_d_1 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_2" Then strKabel_d_2 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_3" Then strKabel_d_3 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_4" Then strKabel_d_4 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_5" Then strKabel_d_5 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_6" Then strKabel_d_6 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_7" Then strKabel_d_7 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_8" Then strKabel_d_8 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_9" Then strKabel_d_9 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_10" Then strKabel_d_10 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_0" Then strKabel_g_0 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_1" Then strKabel_g_1 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_2" Then strKabel_g_2 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_3" Then strKabel_g_3 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_4" Then strKabel_g_4 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_5" Then strKabel_g_5 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_6" Then strKabel_g_6 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_7" Then strKabel_g_7 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_8" Then strKabel_g_8 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_9" Then strKabel_g_9 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_10" Then strKabel_g_10 = acAttRef.TextString
                            Next
                        Case Else
                            Continue For
                    End Select
                    If iVisib = -1 Then
                        arrBlock(index).count = 1
                        arrBlock(index).blName = blName
                        arrBlock(index).blVisibility = Visibility
                        If blName = "Качване" Then
                            Качване(index_Качване).KOTA_1 = strKOTA_1
                            Качване(index_Качване).KOTA_2 = strKOTA_2
                            Качване(index_Качване).ТРЪБА_1 = strТРЪБА_1
                            Качване(index_Качване).ТРЪБА_2 = strТРЪБА_2
                            Качване(index_Качване).Kabel_d_0 = strKabel_d_0
                            Качване(index_Качване).Kabel_d_1 = strKabel_d_1
                            Качване(index_Качване).Kabel_d_2 = strKabel_d_2
                            Качване(index_Качване).Kabel_d_3 = strKabel_d_3
                            Качване(index_Качване).Kabel_d_6 = strKabel_d_6
                            Качване(index_Качване).Kabel_d_7 = strKabel_d_7
                            Качване(index_Качване).Kabel_d_4 = strKabel_d_4
                            Качване(index_Качване).Kabel_d_5 = strKabel_d_5
                            Качване(index_Качване).Kabel_d_8 = strKabel_d_8
                            Качване(index_Качване).Kabel_d_9 = strKabel_d_9
                            Качване(index_Качване).Kabel_d_10 = strKabel_d_10
                            Качване(index_Качване).Kabel_g_0 = strKabel_g_0
                            Качване(index_Качване).Kabel_g_1 = strKabel_g_1
                            Качване(index_Качване).Kabel_g_2 = strKabel_g_2
                            Качване(index_Качване).Kabel_g_3 = strKabel_g_3
                            Качване(index_Качване).Kabel_g_4 = strKabel_g_4
                            Качване(index_Качване).Kabel_g_5 = strKabel_g_5
                            Качване(index_Качване).Kabel_g_6 = strKabel_g_6
                            Качване(index_Качване).Kabel_g_7 = strKabel_g_7
                            Качване(index_Качване).Kabel_g_8 = strKabel_g_8
                            Качване(index_Качване).Kabel_g_9 = strKabel_g_9
                            Качване(index_Качване).Kabel_g_10 = strKabel_g_10
                            index_Качване += 1
                        End If
                        index += 1
                    Else
                        arrBlock(iVisib).count = arrBlock(iVisib).count + 1
                    End If
                Next

                index_Качване = 0

                Call clearDomofon(wsDOMOF)

                index_Row = 2
                Dim dylvina1 As Double = 0
                Dim dylvina2 As Double = 0
                Dim br_linii As Double = 0
                Dim Kabel_FTP As Double = 0

                For Each iarrBlock In arrBlock
                    If iarrBlock.count = 0 Then Exit For
                    With wsDOMOF
                        .Cells(index_Row, 1) = iarrBlock.count
                        .Cells(index_Row, 2) = iarrBlock.blName
                        .Cells(index_Row, 3) = iarrBlock.blVisibility
                        If iarrBlock.blName = "Качване" Then
                            .Cells(index_Row, 11) = Качване(index_Качване).KOTA_1
                            .Cells(index_Row, 12) = Качване(index_Качване).ТРЪБА_1
                            .Cells(index_Row, 13) = Качване(index_Качване).Kabel_d_0
                            .Cells(index_Row, 14) = Качване(index_Качване).Kabel_d_1
                            .Cells(index_Row, 15) = Качване(index_Качване).Kabel_d_2
                            .Cells(index_Row, 16) = Качване(index_Качване).Kabel_d_3
                            .Cells(index_Row, 17) = Качване(index_Качване).Kabel_d_6
                            .Cells(index_Row, 18) = Качване(index_Качване).Kabel_d_5
                            .Cells(index_Row, 19) = Качване(index_Качване).Kabel_d_8
                            .Cells(index_Row, 20) = Качване(index_Качване).Kabel_d_7
                            .Cells(index_Row, 21) = Качване(index_Качване).Kabel_d_4
                            .Cells(index_Row, 22) = Качване(index_Качване).Kabel_d_9
                            .Cells(index_Row, 23) = Качване(index_Качване).Kabel_d_10
                            .Cells(index_Row, 24) = Качване(index_Качване).KOTA_2
                            .Cells(index_Row, 25) = Качване(index_Качване).ТРЪБА_2
                            .Cells(index_Row, 26) = Качване(index_Качване).Kabel_g_0
                            .Cells(index_Row, 27) = Качване(index_Качване).Kabel_g_1
                            .Cells(index_Row, 28) = Качване(index_Качване).Kabel_g_2
                            .Cells(index_Row, 29) = Качване(index_Качване).Kabel_g_3
                            .Cells(index_Row, 30) = Качване(index_Качване).Kabel_g_4
                            .Cells(index_Row, 31) = Качване(index_Качване).Kabel_g_5
                            .Cells(index_Row, 32) = Качване(index_Качване).Kabel_g_6
                            .Cells(index_Row, 33) = Качване(index_Качване).Kabel_g_7
                            .Cells(index_Row, 34) = Качване(index_Качване).Kabel_g_8
                            .Cells(index_Row, 35) = Качване(index_Качване).Kabel_g_9
                            .Cells(index_Row, 36) = Качване(index_Качване).Kabel_g_10
                            index_Качване += 1

                            dylvina1 = Val(.Cells(index_Row, 11).value)
                            dylvina2 = Val(.Cells(index_Row, 24).value)
                            For i = 11 To 36
                                br_linii = InStr(.Cells(index_Row, i).value, "л.")
                                If br_linii = 0 Then Continue For
                                br_linii = Mid(.Cells(index_Row, i).value, 1, br_linii - 1)
                                If InStr(.Cells(index_Row, i).value, "FTP ") > 1 Then
                                    .Cells(index_Row, 8).value =
                                        IIf(i < 24, br_linii * dylvina1, br_linii * dylvina2) +
                                        .Cells(index_Row, 8).value
                                    Kabel_FTP = Kabel_FTP + .Cells(index_Row, 8).value
                                End If
                            Next
                        End If
                        If iarrBlock.blName = "Домофон" Then
                            wsLines.Cells(2, 45).Value = wsLines.Cells(2, 45).Value +
                                                                iarrBlock.count
                        End If
                    End With
                    index_Row += 1
                Next

                wsLines.Cells(2, 46).Value = Kabel_FTP / 2
                acTrans.Commit()
                With wsDOMOF
                    .Cells.Sort(Key1:= .Range("B2"),
                                Order1:=excel.XlSortOrder.xlAscending,
                                Header:=excel.XlYesNoGuess.xlYes,
                                OrderCustom:=1, MatchCase:=False,
                                Orientation:=excel.Constants.xlTopToBottom,
                                DataOption1:=excel.XlSortDataOption.xlSortTextAsNumbers,
                                Key2:= .Range("D2"),
                                Order2:=excel.XlSortOrder.xlAscending,
                                DataOption2:=excel.XlSortDataOption.xlSortTextAsNumbers
                                )
                End With

            Catch ex As Exception
                MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
                acTrans.Abort()
            End Try
        End Using
        Me.Visible = vbTrue
    End Sub
    Private Sub Button_Генерирай_ДОМОФ_Click(sender As Object, e As EventArgs) Handles Button_Генерирай_ДОМОФ.Click
        Dim index As Integer = 0
        Dim index_internet As Integer = 0
        Dim index_Kabel As Integer = 0
        Dim i As Integer
        ProgressBar_Extrat.Maximum = ProgressBar_Maximum
        For i = 6 To 1000
            If wsKol_Smetka.Cells(i, 2).Value = "Раздел" Or wsKol_Smetka.Cells(i, 2).Value = "" Then
                index = i
                Exit For
            End If
            ProgressBar_Extrat.Value = i
        Next
        For i = 2 To 1000
            If wsDOMOF.Cells(i, 2).Value = "" Then
                Exit For
            End If
        Next
        index_internet = i
        ProgressBar_Extrat.Maximum = i
        Dim Text_Dostawka As String = ""
        Dim broj_elementi As Integer = 0
        Call Excel_Kol_smetka_Razdel(wsKol_Smetka, "СЛАБОТОКОВИ ИНСТАЛАЦИИ", "A" & Trim(index.ToString), "D" & Trim(index.ToString))
        index += 1
        Call Excel_Kol_smetka_Razdel(wsKol_Smetka, "ДОМОФОННА ИНСТАЛАЦИЯ", "B" & Trim(index.ToString), "D" & Trim(index.ToString))
        Dim Kabel(1000) As strKabel
        index += 1

        For i = 2 To 10000
            If Trim(wsDOMOF.Cells(i, 2).Value) = "" Then Exit For
            ProgressBar_Extrat.Value = i
            Text_Dostawka = ""
            broj_elementi = wsDOMOF.Cells(i, 1).Value

            Select Case wsDOMOF.Cells(i, 2).Value
                Case "Домофон"
                    Select Case wsDOMOF.Cells(i, 3).Value
                        Case "Домофонна централа(контролер за достъп)", "Централа"
                            Text_Dostawka = "домофонна централа"
                        Case "Карта четец"
                            Text_Dostawka = "безконтактен RFID четец за чипове и карти"
                        Case "Звънец"
                            Text_Dostawka = "домофонна централа"
                        Case "Клавиатура"
                            Text_Dostawka = "цифров аудио домофон със клавиатура"
                        Case "Табло"
                            Text_Dostawka = "входно домофонно табло:" & vbCrLf &
                            "Брой бутони: "
                        Case "Бутон"
                            Text_Dostawka = "бутон за звънец"
                        Case "Брава"
                            Text_Dostawka = "заключващ механизъм за врата; комплект с крепежни елементи"
                        Case "Домофон"
                            Text_Dostawka = "цифров аудио домофон без клавиатура"
                        Case Else
                            Text_Dostawka = "НЕИЗВЕСТЕН ЕЛЕМЕНТ --> " &
                                            wsPIC.Cells(i, 2).Value & " - " &
                                            wsPIC.Cells(i, 3).Value
                    End Select
                    index_Kabel += 1
                Case "Качване"
                    Continue For
            End Select
            If Trim(Text_Dostawka) <> "" Then
                With wsKol_Smetka
                    .Cells(index, 2).Value = "Доставка и монтаж на " + Text_Dostawka
                    .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                    .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                    .Cells(index, 3).Value = "бр."
                    .Cells(index, 4).Value = broj_elementi
                End With
                index += 1
            End If
        Next

        index = Kol_Smetka_Kabeli(index, vbFalse, "ДОМОФ")

        Text_Dostawka = "Комплексно изпитване на системата"
        With wsKol_Smetka
            .Cells(index, 2).Value = Text_Dostawka
            .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
            .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
            .Cells(index, 3).Value = "бр."
            .Cells(index, 4).Value = 1
        End With
        index += 1

        Text_Dostawka = "Обучение на персонал за работа със системата"
        With wsKol_Smetka
            .Cells(index, 2).Value = Text_Dostawka
            .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
            .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
            .Cells(index, 3).Value = "бр."
            .Cells(index, 4).Value = 1
        End With
        index += 1

        ProgressBar_Extrat.Value = ProgressBar_Extrat.Minimum
    End Sub
    Private Sub Button_Вземи_АПАРАТИ_Click(sender As Object, e As EventArgs) Handles Button_Вземи_АПАРАТИ.Click
        If IsNothing(excel_Workbook) Then
            MsgBox("Първо да беше отворил файла?")
            Exit Sub
        End If
        Me.Visible = vbFalse
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim edt As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database
        Dim SelectedSet = cu.GetObjects("INSERT", "Изберете блок")
        If SelectedSet Is Nothing Then
            MsgBox("Няма маркиран нито един блок.")
            Me.Visible = vbTrue
            Exit Sub
        End If
        ProgressBar_Extrat.Maximum = SelectedSet.Count + 1
        ProgressBar_Extrat.Value = ProgressBar_Extrat.Minimum
        If SelectedSet Is Nothing Then
            MsgBox("Няма маркиран нито един блок.")
            Me.Visible = vbTrue
            Exit Sub
        End If
        Dim blkRecId As ObjectId = ObjectId.Null
        Dim index As Integer = 0
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                Dim arrBlock = cu.GetAparati(SelectedSet)
                acTrans.Commit()
                clearТабла(wsElboard)
                index = 2
                For Each iarrBlock In arrBlock
                    If iarrBlock.bl_Брой = 0 Then Exit For
                    wsElboard.Cells(index, 1).Value = iarrBlock.bl_Табло
                    wsElboard.Cells(index, 2).Value = iarrBlock.bl_SHORTNAME
                    wsElboard.Cells(index, 3).Value = iarrBlock.bl_Брой
                    wsElboard.Cells(index, 4).Value = iarrBlock.bl_1
                    wsElboard.Cells(index, 5).Value = iarrBlock.bl_2
                    wsElboard.Cells(index, 6).Value = iarrBlock.bl_3
                    wsElboard.Cells(index, 7).Value = iarrBlock.bl_4
                    wsElboard.Cells(index, 8).Value = iarrBlock.bl_5
                    wsElboard.Cells(index, 9).Value = iarrBlock.bl_6
                    wsElboard.Cells(index, 10).Value = iarrBlock.bl_7
                    wsElboard.Cells(index, 11).Value = iarrBlock.bl_ИмеБлок
                    wsElboard.Cells(index, 12).Value = iarrBlock.bl_LONGNAME
                    wsElboard.Cells(index, 13).Value = iarrBlock.bl_DESIGNATION
                    index += 1
                Next
                With wsElboard
                    .Cells.Sort(Key1:= .Range("A2"),
                                Order1:=excel.XlSortOrder.xlAscending,
                                Header:=excel.XlYesNoGuess.xlYes,
                                OrderCustom:=1, MatchCase:=False,
                                Orientation:=excel.Constants.xlTopToBottom,
                                DataOption1:=excel.XlSortDataOption.xlSortTextAsNumbers,
                                Key2:= .Range("B2"),
                                Order2:=excel.XlSortOrder.xlDescending,
                                DataOption2:=excel.XlSortDataOption.xlSortTextAsNumbers,
                                Key3:= .Range("E2"),
                                Order3:=excel.XlSortOrder.xlAscending,
                                DataOption3:=excel.XlSortDataOption.xlSortTextAsNumbers
                                )

                End With
            Catch ex As Exception
                MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
                acTrans.Abort()
            End Try
        End Using
        Me.Visible = vbTrue
        ProgressBar_Extrat.Value = ProgressBar_Extrat.Minimum
    End Sub
    Private Sub Button_Вземи_СКАРИ_Click(sender As Object, e As EventArgs) Handles Button_Вземи_СКАРИ.Click
        If IsNothing(excel_Workbook) Then
            MsgBox("Първо да беше отворил файла?")
            Exit Sub
        End If

        Me.Visible = vbFalse
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim edt As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database
        Dim SelectedSet = cu.GetObjects("INSERT", "Изберете блок")
        Dim arrBlock(500) As strСкара
        If SelectedSet Is Nothing Then
            MsgBox("Няма маркиран нито един блок.")
            Me.Visible = vbTrue
            Exit Sub
        End If
        Dim blkRecId As ObjectId = ObjectId.Null
        Dim index As Integer = 0

        Dim formu As String = ""
        Dim rang As String = ""
        Dim colum As String = ""

        arrBlock(index).bl_ИмеБлок = "Канал"
        arrBlock(index).bl_Ширина = 11
        arrBlock(index).bl_Височина = 11
        arrBlock(index).bl_Дължина = 0
        arrBlock(index).bl_Брой = 0
        arrBlock(index).bl_Visible = ""
        index += 1
        arrBlock(index).bl_ИмеБлок = "Скара стълба"
        arrBlock(index).bl_Ширина = 200
        arrBlock(index).bl_Височина = 60
        arrBlock(index).bl_Дължина = 0
        arrBlock(index).bl_Брой = 0
        arrBlock(index).bl_Visible = ""
        index += 1

        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                For Each sObj As SelectedObject In SelectedSet
                    blkRecId = sObj.ObjectId

                    Dim acBlkRef As BlockReference = DirectCast(acTrans.GetObject(blkRecId, OpenMode.ForRead), BlockReference)
                    Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                    Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection

                    Dim blName As String = (CType(acBlkRef.DynamicBlockTableRecord.GetObject(OpenMode.ForRead), BlockTableRecord)).Name

                    If Not (blName = "Скара" Or
                            blName = "Канал" Or
                            blName = "Скара_ъгъл") Then Continue For

                    Dim Ширина As Integer = 0
                    Dim Височина As Integer = 0
                    Dim Дължина As Double = 0
                    Dim Visibility As String = ""

                    For Each prop As DynamicBlockReferenceProperty In props
                        If prop.PropertyName = "Ширина" Then Ширина = prop.Value * 10
                        If prop.PropertyName = "Височина" Then Височина = prop.Value * 10
                        If prop.PropertyName = "Дължина" Then Дължина = prop.Value / 100
                        If prop.PropertyName = "Visibility" Then Visibility = prop.Value
                    Next

                    Dim iVisib As Integer = -1

                    iVisib = Array.FindIndex(arrBlock, Function(f) f.bl_ИмеБлок = blName And
                                                                   f.bl_Ширина = Ширина And
                                                                   f.bl_Височина = Височина And
                                                                   f.bl_Visible = Visibility)

                    If iVisib = -1 Then
                        arrBlock(index).bl_ИмеБлок = blName
                        arrBlock(index).bl_Ширина = Ширина
                        arrBlock(index).bl_Височина = Височина
                        arrBlock(index).bl_Дължина = Дължина
                        arrBlock(index).bl_Брой = 1
                        arrBlock(index).bl_Visible = Visibility
                        index += 1
                    Else
                        arrBlock(iVisib).bl_Дължина = arrBlock(iVisib).bl_Дължина + Дължина
                        arrBlock(iVisib).bl_Брой = arrBlock(iVisib).bl_Брой + 1
                    End If
                Next

                clearКбелнаскара(wsCableTrays)

                index = 4
                For Each iarrBlock In arrBlock
                    If IsNothing(iarrBlock.bl_ИмеБлок) Then Exit For

                    With wsCableTrays
                        .Cells(index, 1).value = iarrBlock.bl_ИмеБлок
                        .Cells(index, 2).value = iarrBlock.bl_Ширина
                        .Cells(index, 3).value = iarrBlock.bl_Височина
                        .Cells(index, 4).value = IIf(iarrBlock.bl_ИмеБлок = "Скара_ъгъл",
                                                     iarrBlock.bl_Visible,
                                                     iarrBlock.bl_Дължина)
                        .Cells(index, 5).value = iarrBlock.bl_Брой
                        Select Case iarrBlock.bl_ИмеБлок
                            Case "Скара_ъгъл"
                                .Cells(index, 2).value = iarrBlock.bl_Ширина * 2
                                .Cells(index, 4).value = iarrBlock.bl_Visible

                                formu = "=E" & index.ToString
                                rang = "H" & Trim((index).ToString)
                                .Range(rang).Formula = formu
                                If iarrBlock.bl_Visible = "Т-Вертикален" Then
                                    .Range("J2").Value = .Range("J2").Value + iarrBlock.bl_Брой
                                End If
                                If iarrBlock.bl_Visible = "90-Вертикален" Then
                                    .Range("J2").Value = .Range("J2").Value + iarrBlock.bl_Брой
                                End If
                            Case "Скара стълба"
                                formu = "=F" & index.ToString & "*$G$2"
                                rang = "G" & Trim((index).ToString)
                                .Range(rang).Formula = formu

                                formu = "=IF((D" & index.ToString & "+F" & index.ToString & ")>0,INT(G" & index.ToString & "/3+1)*3,0)"
                                rang = "H" & Trim((index).ToString)
                                .Range(rang).Formula = formu

                                formu = "=SUM(J" & index.ToString & ":K" & index.ToString & ")+D" & index.ToString
                                rang = "F" & Trim((index).ToString)
                                .Range(rang).Formula = formu

                                formu = "=$J$2*$J$3"
                                rang = "J" & Trim((index).ToString)
                                .Range(rang).Formula = formu

                                formu = "=$K$2*$K$3"
                                rang = "K" & Trim((index).ToString)
                                .Range(rang).Formula = formu
                            Case "Скара"
                                formu = "=F" & index.ToString & "*$G$2"
                                rang = "G" & Trim((index).ToString)
                                .Range(rang).Formula = formu

                                formu = "=INT(G" & index.ToString & "/3+1)*3"
                                rang = "H" & Trim((index).ToString)
                                .Range(rang).Formula = formu

                                formu = "=SUM(J" & index.ToString & ":K" & index.ToString & ")+D" & index.ToString
                                rang = "F" & Trim((index).ToString)
                                .Range(rang).Formula = formu

                            Case "Канал"

                                formu = "=F" & index.ToString & "*$G$2"
                                rang = "G" & Trim((index).ToString)
                                .Range(rang).Formula = formu

                                formu = "=SUM(M" & index.ToString & ":Q" & index.ToString & ")+D" & index.ToString
                                rang = "F" & Trim((index).ToString)
                                .Range(rang).Formula = formu

                                If iarrBlock.bl_Ширина = 11 And
                                    iarrBlock.bl_Височина = 11 Then

                                    formu = "=IF((D" & index.ToString &
                                            "+F" & index.ToString &
                                            ")>0,INT(G" & index.ToString &
                                            "/2+1)*2,0)"

                                    rang = "H" & Trim((index).ToString)
                                    .Range(rang).Formula = formu
                                Else
                                    formu = "=INT(G" & index.ToString & "/2+1)*2"
                                    rang = "H" & Trim((index).ToString)
                                    .Range(rang).Formula = formu
                                End If

                                If iarrBlock.bl_Ширина = 11 And
                                    iarrBlock.bl_Височина = 11 Then

                                    formu = "=$M$2*$M$3"
                                    rang = "M" & Trim((index).ToString)
                                    .Range(rang).Formula = formu

                                    formu = "=$N$2*$N$3"
                                    rang = "N" & Trim((index).ToString)
                                    .Range(rang).Formula = formu

                                    formu = "=$O$2*$O$3"
                                    rang = "O" & Trim((index).ToString)
                                    .Range(rang).Formula = formu

                                    formu = "=$P$2*$P$3"
                                    rang = "P" & Trim((index).ToString)
                                    .Range(rang).Formula = formu

                                    formu = "=$Q$2*$Q$3"
                                    rang = "Q" & Trim((index).ToString)
                                    .Range(rang).Formula = formu
                                End If
                        End Select
                        index += 1
                    End With
                Next
                acTrans.Commit()

                With wsCableTrays

                    formu = (H_Етаж).ToString
                    rang = "J3"
                    .Range(rang).Formula = formu

                    formu = (H_Етаж - 1.5).ToString
                    rang = "K3"
                    .Range(rang).Formula = formu

                    formu = (H_Етаж - H_Ключ).ToString
                    rang = "M3"
                    .Range(rang).Formula = formu

                    formu = (H_Етаж - H_Контакт).ToString
                    rang = "N3"
                    .Range(rang).Formula = formu

                    formu = (H_Етаж - 2.1).ToString
                    rang = "O3"
                    .Range(rang).Formula = formu

                    formu = (H_Етаж).ToString
                    rang = "P3"
                    .Range(rang).Formula = formu

                    formu = (H_Етаж).ToString
                    rang = "Q3"
                    .Range(rang).Formula = formu

                End With

            Catch ex As Exception
                MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
                acTrans.Abort()
            End Try
        End Using

        With wsCableTrays

            .Sort.SortFields.Clear()
            rang = "A4:A" + (index - 1).ToString
            .Sort.SortFields.Add(Key:= .Range(rang), Order:=excel.XlSortOrder.xlAscending, DataOption:=excel.XlSortDataOption.xlSortNormal)
            rang = "B4:B" + (index - 1).ToString
            .Sort.SortFields.Add(Key:= .Range(rang), Order:=excel.XlSortOrder.xlAscending, DataOption:=excel.XlSortDataOption.xlSortNormal)
            rang = "C4:C" + (index - 1).ToString
            .Sort.SortFields.Add(Key:= .Range(rang), Order:=excel.XlSortOrder.xlAscending, DataOption:=excel.XlSortDataOption.xlSortNormal)
            Dim exRange As excel.Range
            exRange = .Range("A4:Z" + (index - 1).ToString)
            With .Sort
                .SetRange(exRange)
                .Header = excel.XlYesNoGuess.xlYes
                .MatchCase = False
                .SortMethod = excel.XlSortMethod.xlPinYin
                .Orientation = excel.XlSortOrientation.xlSortColumns
                .Apply()
            End With
        End With

        Me.Visible = vbTrue
    End Sub
    Private Sub Button_Изчисти_СКАРИ_Click(sender As Object, e As EventArgs) Handles Button_Изчисти_СКАРИ.Click
        Call clearКбелнаскара(wsCableTrays)
    End Sub
    Private Sub Button_Изчисти_АПАРАТИ_Click(sender As Object, e As EventArgs) Handles Button_Изчисти_АПАРАТИ.Click
        clearТабла(wsElboard)
    End Sub
    Private Sub Button_Генератор_АПАРАТИ_Click(sender As Object, e As EventArgs) Handles Button_Генератор_АПАРАТИ.Click
        Dim index As Integer = 0
        Dim i As Integer
        ProgressBar_Extrat.Maximum = ProgressBar_Maximum
        For i = 6 To 10000
            If wsKol_Smetka.Cells(i, 2).Value = "Раздел" Or wsKol_Smetka.Cells(i, 2).Value = "" Then
                index = i
                Exit For
            End If
            ProgressBar_Extrat.Value = i
        Next
        For i = 2 To 10000
            If wsElboard.Cells(i, 2).Value = "" Then
                Exit For
            End If
        Next
        ProgressBar_Extrat.Maximum = i
        Call Excel_Kol_smetka_Razdel(wsKol_Smetka, "СИЛОВИ РАЗПРЕДЕЛИТЕЛНИ ТАБЛА", "A" & Trim(index.ToString), "D" & Trim(index.ToString))
        index += 1
        Dim Tablo As String = ""
        Dim Tablo_old As String = ""
        Dim Брой_табла As Integer = 0
        Dim Брой_Дефек As Integer = 0
        Dim Нишa_Таблo As Boolean = vbFalse
        Dim Брой_Връзки_2 As Integer
        Dim Брой_Връзки_16 As Integer
        Dim Брой_Връзки_63 As Integer
        For i = 2 To 10000
            Dim text As String = ""
            ProgressBar_Extrat.Value = i
            If wsElboard.Cells(i, 2).Value = "" Then Exit For
            If Tablo <> "" Then
                If Tablo <> wsElboard.Cells(i, 1).Value Then
                    With wsKol_Smetka
                        text = "Монтаж на ел. табло '" + Tablo + "', в т.ч.:"
                        .Cells(index, 2).Value = text
                        index += 1

                        text = IIf(Нишa_Таблo, " - Направа нишa за ел.таблo", " - Монтаж на табло окачено на стена")
                        .Cells(index, 2).Value = text
                        .Cells(index, 3).Value = "бр."
                        .Cells(index, 4).Value = 1
                        index += 1

                        If Брой_Връзки_2 > 0 Then
                            text = " - Свързване проводник към табло до 2,5мм²"
                            .Cells(index, 2).Value = text
                            .Cells(index, 3).Value = "бр."
                            .Cells(index, 4).Value = Брой_Връзки_2
                            index += 1
                        End If

                        If Брой_Връзки_16 > 0 Then
                            text = " - Свързване проводник към табло до 16мм²"
                            .Cells(index, 2).Value = text
                            .Cells(index, 3).Value = "бр."
                            .Cells(index, 4).Value = Брой_Връзки_16
                            index += 1
                        End If

                        If Брой_Връзки_63 > 0 Then
                            text = " - Свързване проводник към табло до 16мм²"
                            .Cells(index, 2).Value = text
                            .Cells(index, 3).Value = "бр."
                            .Cells(index, 4).Value = Брой_Връзки_63
                            index += 1
                        End If

                    End With
                    Tablo_old = Tablo
                    Брой_Връзки_2 = 0
                    Брой_Връзки_16 = 0
                    Брой_Връзки_63 = 0
                    text = ""
                End If
            End If
            Select Case wsElboard.Cells(i, 2).Value
                Case "Контролен", "iEM2155", "iEM3155"
                    text = text + " - Eлектромер; директен; "
                    text = text + IIf(Trim(wsElboard.Cells(i, 7).Value) = "230V",
                                           "1P+N; 230V", "3P+N; 230/400V")
                Case "iEM3250", "iEM3255"
                    text = text + " - Eлектромер; индиректен; "
                    text = text + IIf(Trim(wsElboard.Cells(i, 7).Value) = "230V",
                                           "1P+N; 230V", "3P+N; 230/400V")
                Case "Метален шкаф стоящ"
                    Tablo = wsElboard.Cells(i, 1).Value
                    text = "Доставка на ел. табло "
                    text = text + "'" + wsElboard.Cells(i, 1).Value + "' - "
                    text = text + "метален шкаф за стоящ монтаж;"
                    text = text + " Размери: Вис.-" + wsElboard.Cells(i, 4).Value.ToString + "; "
                    text = text + "Шир. -" + wsElboard.Cells(i, 5).Value.ToString + "; "
                    text = text + "Дъл. -" + wsElboard.Cells(i, 6).Value.ToString + "; "
                    text = text + IIf(Trim(wsElboard.Cells(i, 7).Value) <> "",
                                      "Врата: " + Trim(wsElboard.Cells(i, 7).Value) + ";",
                                      "")
                    Call Excel_Kol_smetka_Razdel(wsKol_Smetka, text, "B" & Trim(index.ToString), "D" & Trim(index.ToString))
                    wsElboard.Cells(i, 20).Value = "стоящ монтаж"
                    index += 1
                    text = "В т.ч. доставени и монтирани в таблото елементи:"
                    Брой_табла += 1
                    Нишa_Таблo = vbFalse
                Case "Метален шкаф"
                    Tablo = wsElboard.Cells(i, 1).Value
                    text = "Доставка на ел. табло "
                    text = text + "'" + wsElboard.Cells(i, 1).Value + "'-"
                    text = text + "метален шкаф;"
                    text = text + " Размери: Вис.-" + wsElboard.Cells(i, 4).Value.ToString + "; "
                    text = text + "Шир. -" + wsElboard.Cells(i, 5).Value.ToString + "; "
                    text = text + "Дъл. -" + wsElboard.Cells(i, 6).Value.ToString + "; "
                    text = text + IIf(Trim(wsElboard.Cells(i, 7).Value) <> "",
                                      "Врата: " + Trim(wsElboard.Cells(i, 7).Value) + ";",
                                      "")
                    Call Excel_Kol_smetka_Razdel(wsKol_Smetka, text, "B" & Trim(index.ToString), "D" & Trim(index.ToString))
                    index += 1
                    wsElboard.Cells(i, 20).Value = "метален шкаф"
                    text = "В т.ч. доставени и монтирани в таблото елементи:"
                    Брой_табла += 1
                    Нишa_Таблo = vbFalse
                Case "Изпъкнал монтаж", "Открит монтаж"
                    Tablo = wsElboard.Cells(i, 1).Value
                    text = "Доставка на ел. табло "
                    text = text + "'" + wsElboard.Cells(i, 1).Value + "'-"
                    text = text + "полиестерен шкаф;"
                    text = text + " Брой модули: " + wsElboard.Cells(i, 4).Value.ToString + "; "
                    If Trim(wsElboard.Cells(i, 6).Value) <> "" Then
                        text = text + "Врата: " + Trim(wsElboard.Cells(i, 6).Value)
                    Else
                        text = text + "Врата: " + Trim(wsElboard.Cells(i, 5).Value)
                    End If
                    Call Excel_Kol_smetka_Razdel(wsKol_Smetka, text, "B" & Trim(index.ToString), "D" & Trim(index.ToString))
                    index += 1
                    wsElboard.Cells(i, 20).Value = "Изпъкнал монтаж"
                    text = "В т.ч. доставени и монтирани в таблото елементи:"
                    Брой_табла += 1
                    Нишa_Таблo = vbFalse
                Case "Вграден монтаж"
                    Tablo = wsElboard.Cells(i, 1).Value
                    text = "Доставка на ел. табло "
                    text = text + "'" + wsElboard.Cells(i, 1).Value + "'-"
                    text = text + "полиестерен шкаф;"
                    text = text + " Брой модули: " + wsElboard.Cells(i, 4).Value.ToString + "; "
                    If Trim(wsElboard.Cells(i, 6).Value) <> "" Then
                        text = text + "Врата: " + Trim(wsElboard.Cells(i, 6).Value)
                    Else
                        text = text + "Врата: " + Trim(wsElboard.Cells(i, 5).Value)
                    End If
                    Call Excel_Kol_smetka_Razdel(wsKol_Smetka, text, "B" & Trim(index.ToString), "D" & Trim(index.ToString))
                    wsElboard.Cells(i, 20).Value = "Вграден монтаж"
                    index += 1
                    text = "В т.ч. доставени и монтирани в таблото елементи:"
                    Брой_табла += 1
                    Нишa_Таблo = vbTrue
                Case "Mini Kaedra", "Kaedra"
                    text = "Доставка на ел. табло "
                    text += "'" + wsElboard.Cells(i, 1).Value + "'-"
                    text += "полиестерен шкаф;"
                    text += " Брой модули: " + wsElboard.Cells(i, 5).Value.ToString + "; "
                    text += "Брой редове: "
                    If wsElboard.Range("F" & i).Value IsNot Nothing Then
                        text += wsElboard.Range("F" & i).Value.ToString
                    Else
                        Select Case wsElboard.Range("E" & i).Value.ToString
                            Case "3", "4", "6", "8", "12", "18"
                                text += "1"
                            Case "24"
                                text += "2"
                            Case "54"
                                text += "3"
                            Case "72"
                                text += "4"
                        End Select
                    End If
                    text += "; Степен на защита: IP65"
                    Call Excel_Kol_smetka_Razdel(wsKol_Smetka, text, "B" & Trim(index.ToString), "D" & Trim(index.ToString))
                    wsElboard.Cells(i, 20).Value = "Вграден монтаж"
                    index += 1
                    text = "В т.ч. доставени и монтирани в таблото елементи:"
                    Брой_табла += 1
                    Нишa_Таблo = vbTrue
                Case "Kaedra - щепселни съединения"
                Case "Kaedra"
                Case "E60", "Е60", "iK60", "iC60", "EZ9 MCB", "EZ9 MCB ",
                     "Е120", "E120", "EZCV250", "C120", "С120"
                    text = text + " - Автоматичен прекъсвач - "
                    text = text + wsElboard.Cells(i, 2).Value + " "
                    text = text + wsElboard.Cells(i, 4).Value + "; "
                    text = text + wsElboard.Cells(i, 5).Value + "; "
                    text = text + wsElboard.Cells(i, 6).Value +
                                  IIf(Trim(wsElboard.Cells(i, 7).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 7).Value +
                                  IIf(Trim(wsElboard.Cells(i, 8).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 8).Value +
                                  IIf(Trim(wsElboard.Cells(i, 13).Value) <> "", "; ", "")
                    text = text + IIf(Len(Trim(wsElboard.Cells(i, 13).Value)) > 0, wsElboard.Cells(i, 13).Value & "; ", "")
                    Dim Полюси As Integer = Val(wsElboard.Range("F" & i.ToString).Value) + 2
                    Select Case Val(wsElboard.Cells(i, 7).Value)
                        Case <= 20
                            wsElboard.Cells(i, 16).Value = Val(wsElboard.Cells(i, 3).Value) * Полюси
                        Case <= 63
                            wsElboard.Cells(i, 17).Value = Val(wsElboard.Cells(i, 3).Value) * Полюси
                        Case > 63
                            wsElboard.Cells(i, 18).Value = Val(wsElboard.Cells(i, 3).Value) * Полюси
                    End Select
                Case "NG125", "NG160", "NS", "NW"
                    text = text + IIf(wsElboard.Cells(i, 4).Value = "NA",
                                  " - Мощностен разединител - ",
                                  " - Автоматичен прекъсвач - ")
                    text = text + wsElboard.Cells(i, 2).Value + " "
                    text = text + IIf(Len(Trim(wsElboard.Cells(i, 13).Value)) > 0, wsElboard.Cells(i, 13).Value & "; ", "")
                    text = text + wsElboard.Cells(i, 5).Value + "; "
                    text = text + wsElboard.Cells(i, 6).Value +
                                  IIf(Trim(wsElboard.Cells(i, 7).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 7).Value +
                                  IIf(Trim(wsElboard.Cells(i, 8).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 8).Value +
                                  IIf(Trim(wsElboard.Cells(i, 13).Value) <> "", "; ", "")
                    text = text + IIf(Len(Trim(wsElboard.Cells(i, 13).Value)) > 0, wsElboard.Cells(i, 13).Value & "; ", "")
                    text = " " + text
                    Dim Полюси As Integer = Val(wsElboard.Range("F" & i.ToString).Value) + 2
                    Select Case Val(wsElboard.Cells(i, 7).Value)
                        Case <= 20
                            wsElboard.Cells(i, 16).Value = Val(wsElboard.Cells(i, 3).Value) * Полюси
                        Case <= 63
                            wsElboard.Cells(i, 17).Value = Val(wsElboard.Cells(i, 3).Value) * Полюси
                        Case > 63
                            wsElboard.Cells(i, 18).Value = Val(wsElboard.Cells(i, 3).Value) * Полюси
                    End Select
                Case "C60H-DC"
                    text = text + " - Автоматичен прекъсвач за постоянен ток - "
                    text = text + wsElboard.Cells(i, 2).Value + " "
                    text = text + wsElboard.Cells(i, 4).Value + "; "
                    text = text + wsElboard.Cells(i, 5).Value + "; "
                    text = text + wsElboard.Cells(i, 6).Value +
                                  IIf(Trim(wsElboard.Cells(i, 7).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 7).Value +
                                  IIf(Trim(wsElboard.Cells(i, 8).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 8).Value +
                                  IIf(Trim(wsElboard.Cells(i, 13).Value) <> "", "; ", "")
                    text = text + IIf(Len(Trim(wsElboard.Cells(i, 13).Value)) > 0, wsElboard.Cells(i, 13).Value & "; ", "")
                Case "NSX100", "NSX160", "NSX250", "NSX400", "NSX630"
                    text = text + " - Автоматичен прекъсвач - Compact - "
                    text = text + wsElboard.Cells(i, 2).Value + " "
                    text = text + wsElboard.Cells(i, 4).Value + "; "
                    text = text + wsElboard.Cells(i, 6).Value + "; "
                    text = text + wsElboard.Cells(i, 5).Value +
                                  IIf(Trim(wsElboard.Cells(i, 7).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 7).Value +
                                  IIf(Trim(wsElboard.Cells(i, 8).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 8).Value +
                                  IIf(Trim(wsElboard.Cells(i, 13).Value) <> "", "; ", "")
                    text = text + IIf(Len(Trim(wsElboard.Cells(i, 13).Value)) > 0, wsElboard.Cells(i, 13).Value & "; ", "")

                    Dim Полюси As Integer = Val(wsElboard.Cells(i, 6).Value) + 2
                    Select Case Val(wsElboard.Cells(i, 7).Value)
                        Case <= 20
                            wsElboard.Cells(i, 16).Value = Val(wsElboard.Cells(i, 3).Value) * Полюси
                        Case <= 63
                            wsElboard.Cells(i, 17).Value = Val(wsElboard.Cells(i, 3).Value) * Полюси
                        Case > 63
                            wsElboard.Cells(i, 18).Value = Val(wsElboard.Cells(i, 3).Value) * Полюси
                    End Select
                Case "NS800", "NS800", "NS630b", "NS1600", "NS1250", "NS1000"
                    text = text + " - Автоматичен прекъсвач - Compact - "
                    text = text + wsElboard.Cells(i, 2).Value + " "
                    text = text + wsElboard.Cells(i, 4).Value + "; "
                    text = text + Mid(Trim(wsElboard.Cells(i, 5).Value), 1, 2) + "; "
                    If Len(Trim(wsElboard.Cells(i, 5).Value)) > 2 Then
                        text = text + Trim(Mid(Trim(wsElboard.Cells(i, 5).Value), 3, Len(wsElboard.Cells(i, 5).Value))) + "; "
                    End If
                    text = text + wsElboard.Cells(i, 6).Value +
                                  IIf(Trim(wsElboard.Cells(i, 7).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 7).Value +
                                  IIf(Trim(wsElboard.Cells(i, 8).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 8).Value +
                                  IIf(Trim(wsElboard.Cells(i, 13).Value) <> "", "; ", "")
                    text = text + IIf(Len(Trim(wsElboard.Cells(i, 13).Value)) > 0, wsElboard.Cells(i, 13).Value & "; ", "")
                Case "NT06", "NT08", "NT10", "NT12", "NT16"
                    text = text + " - Автоматичен прекъсвач - "
                    text = text + wsElboard.Cells(i, 2).Value + " "
                    text = text + wsElboard.Cells(i, 4).Value + "; "
                    text = text + wsElboard.Cells(i, 5).Value + "; "
                    text = text + wsElboard.Cells(i, 6).Value +
                                  IIf(Trim(wsElboard.Cells(i, 7).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 7).Value +
                                  IIf(Trim(wsElboard.Cells(i, 8).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 8).Value +
                                  IIf(Trim(wsElboard.Cells(i, 13).Value) <> "", "; ", "")
                    text = text + IIf(Len(Trim(wsElboard.Cells(i, 13).Value)) > 0, wsElboard.Cells(i, 13).Value & "; ", "")
                Case "NW08", "NW10", "NW12", "NW16", "NW20", "NW25", "NW32",
                     "NW40b", "NW40", "NW50", "NW63", "NB600", "ZC100", "EZC250", "EZC400"
                    text = text + " - Автоматичен прекъсвач - "
                    text = text + wsElboard.Cells(i, 2).Value + " "
                    text = text + wsElboard.Cells(i, 4).Value + "; "
                    text = text + wsElboard.Cells(i, 5).Value + "; "
                    text = text + wsElboard.Cells(i, 6).Value +
                                  IIf(Trim(wsElboard.Cells(i, 7).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 7).Value +
                                  IIf(Trim(wsElboard.Cells(i, 8).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 8).Value +
                                  IIf(Trim(wsElboard.Cells(i, 13).Value) <> "", "; ", "")
                    text = text + IIf(Len(Trim(wsElboard.Cells(i, 13).Value)) > 0, wsElboard.Cells(i, 13).Value & "; ", "")
                Case "EZC400", "EZC250", "EZC100"
                    text = text + " - Автоматичен прекъсвач - EasyPact - "
                    text = text + wsElboard.Cells(i, 2).Value + " "
                    text = text + wsElboard.Cells(i, 4).Value + "; "
                    text = text + wsElboard.Cells(i, 6).Value + "; "
                    text = text + wsElboard.Cells(i, 5).Value +
                                  IIf(Trim(wsElboard.Cells(i, 7).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 7).Value +
                                  IIf(Trim(wsElboard.Cells(i, 8).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 8).Value +
                                  IIf(Trim(wsElboard.Cells(i, 13).Value) <> "", "; ", "")
                    text = text + IIf(Len(Trim(wsElboard.Cells(i, 13).Value)) > 0, wsElboard.Cells(i, 13).Value & "; ", "")
                Case "iSW", "IN", "INS", "INV"
                    text = text + " - Мощностен разединител -  "
                    text = text + wsElboard.Cells(i, 2).Value + " "
                    text = text + wsElboard.Cells(i, 6).Value + "; "
                    text = text + wsElboard.Cells(i, 5).Value +
                                  IIf(Trim(wsElboard.Cells(i, 7).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 7).Value +
                                  IIf(Trim(wsElboard.Cells(i, 8).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 8).Value +
                                  IIf(Trim(wsElboard.Cells(i, 13).Value) <> "", "; ", "")
                    text = text + IIf(Len(Trim(wsElboard.Cells(i, 13).Value)) > 0, wsElboard.Cells(i, 13).Value & "; ", "")
                    Dim Полюси As Integer = Val(wsElboard.Cells(i, 5).Value) + 2
                    Select Case Val(wsElboard.Cells(i, 6).Value)
                        Case <= 20
                            wsElboard.Cells(i, 16).Value = Val(wsElboard.Cells(i, 3).Value) * Полюси
                        Case <= 63
                            wsElboard.Cells(i, 17).Value = Val(wsElboard.Cells(i, 3).Value) * Полюси
                        Case > 63
                            wsElboard.Cells(i, 18).Value = Val(wsElboard.Cells(i, 3).Value) * Полюси
                    End Select
                Case "NT"
                    text = text + " - Мощностен разединител - "
                    text = text + wsElboard.Cells(i, 2).Value + " "
                    text = text + wsElboard.Cells(i, 5).Value + "; "
                    text = text + wsElboard.Cells(i, 6).Value +
                                  IIf(Trim(wsElboard.Cells(i, 7).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 7).Value +
                                  IIf(Trim(wsElboard.Cells(i, 8).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 8).Value +
                                  IIf(Trim(wsElboard.Cells(i, 13).Value) <> "", "; ", "")
                Case "DPN N Vigi", "DPNa Vigi", "DPNа Vigi", "EZ9 RCBO"
                    text = text + " - Автоматичен прекъсвач с вградена дефектно токова защита - "
                    text = text + wsElboard.Cells(i, 2).Value + "; "
                    text = text + wsElboard.Cells(i, 4).Value + "; "
                    text = text + wsElboard.Cells(i, 6).Value + "; "
                    text = text + wsElboard.Cells(i, 5).Value +
                                  IIf(Trim(wsElboard.Cells(i, 7).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 7).Value +
                                  IIf(Trim(wsElboard.Cells(i, 8).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 8).Value +
                                  IIf(Trim(wsElboard.Cells(i, 13).Value) <> "", "; ", "")
                    Dim Полюси As Integer = Val(wsElboard.Range("E" & i.ToString).Value) + 2
                    Select Case Val(wsElboard.Cells(i, 7).Value)
                        Case <= 20
                            wsElboard.Cells(i, 16).Value = Val(wsElboard.Cells(i, 3).Value) * Полюси
                        Case <= 63
                            wsElboard.Cells(i, 17).Value = Val(wsElboard.Cells(i, 3).Value) * Полюси
                        Case > 63
                            wsElboard.Cells(i, 18).Value = Val(wsElboard.Cells(i, 3).Value) * Полюси
                    End Select
                    wsElboard.Cells(i, 19).Value = wsElboard.Cells(i, 3).Value
                    Брой_Дефек = Брой_Дефек + wsElboard.Cells(i, 3).Value
                Case "ID Domae", "iID", "Vigi iC60", "iID К", "EZ9 RCCB"
                    text = text + " - Дефектнотокова защита - "
                    text = text + wsElboard.Cells(i, 2).Value + "; "
                    text = text + wsElboard.Cells(i, 4).Value + "; "
                    text = text + wsElboard.Cells(i, 6).Value + "; "
                    text = text + wsElboard.Cells(i, 5).Value +
                                  IIf(Trim(wsElboard.Cells(i, 7).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 7).Value +
                                  IIf(Trim(wsElboard.Cells(i, 8).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 8).Value +
                                  IIf(Trim(wsElboard.Cells(i, 13).Value) <> "", "; ", "")
                    wsElboard.Cells(i, 19).Value = wsElboard.Cells(i, 3).Value
                    Брой_Дефек = Брой_Дефек + wsElboard.Cells(i, 3).Value
                    Dim Полюси As Integer = Val(wsElboard.Range("E" & i.ToString).Value) + 2
                    Select Case Val(wsElboard.Cells(i, 7).Value)
                        Case <= 20
                            wsElboard.Cells(i, 16).Value = Val(wsElboard.Cells(i, 3).Value) * Полюси
                        Case <= 63
                            wsElboard.Cells(i, 17).Value = Val(wsElboard.Cells(i, 3).Value) * Полюси
                        Case > 63
                            wsElboard.Cells(i, 18).Value = Val(wsElboard.Cells(i, 3).Value) * Полюси
                    End Select
                Case "PRC", "PRI"
                    text = text + " - Катоден отводител за телефонна линия - "
                    text = text + wsElboard.Cells(i, 2).Value
                Case "iTL"
                    text = text + " - Импулсно реле - "
                    text = text + wsElboard.Cells(i, 2).Value + " "
                    text = text + wsElboard.Cells(i, 5).Value + "; "
                    text = text + wsElboard.Cells(i, 6).Value +
                                  IIf(Trim(wsElboard.Cells(i, 7).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 7).Value +
                                  IIf(Trim(wsElboard.Cells(i, 8).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 8).Value +
                                  IIf(Trim(wsElboard.Cells(i, 13).Value) <> "", "; ", "")
                    wsElboard.Cells(i, 16).Value = Val(wsElboard.Range("C" & i.ToString).Value) * 4
                Case "IHP", "IH"
                    text = text + " - Програмируемо времереле - "
                    text = text + wsElboard.Cells(i, 2).Value + " "
                    text = text + wsElboard.Cells(i, 7).Value +
                                  IIf(Trim(wsElboard.Cells(i, 8).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 8).Value +
                                  IIf(Trim(wsElboard.Cells(i, 13).Value) <> "", "; ", "")
                    wsElboard.Cells(i, 16).Value = Val(wsElboard.Range("C" & i.ToString).Value) * 4
                Case "MIN", "MINp"
                    text = text + " - Стълбищен автомат - "
                    text = text + wsElboard.Cells(i, 2).Value + " "
                    text = text + wsElboard.Cells(i, 5).Value + "; "
                    text = text + wsElboard.Cells(i, 6).Value +
                                  IIf(Trim(wsElboard.Cells(i, 7).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 7).Value +
                                  IIf(Trim(wsElboard.Cells(i, 8).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 8).Value +
                                  IIf(Trim(wsElboard.Cells(i, 13).Value) <> "", "; ", "")
                    wsElboard.Cells(i, 16).Value = Val(wsElboard.Range("C" & i.ToString).Value) * 4
                Case "IC 100к", "IC Astro", "IC100", "IC2000", "IC2000P+"
                    text = text + " - Фото реле - "
                    text = text + wsElboard.Cells(i, 2).Value + " "
                    text = text + wsElboard.Cells(i, 5).Value + "; "
                    text = text + wsElboard.Cells(i, 6).Value +
                                  IIf(Trim(wsElboard.Cells(i, 7).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 7).Value +
                                  IIf(Trim(wsElboard.Cells(i, 8).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 8).Value +
                                  IIf(Trim(wsElboard.Cells(i, 13).Value) <> "", "; ", "")
                    wsElboard.Cells(i, 16).Value = Val(wsElboard.Range("C" & i.ToString).Value) * 4
                Case "LC2D", "LC2K", "LC8K", "LP2K", "LP5K"
                    text = text + " - Реверсивен контактор - "
                    text = text + wsElboard.Cells(i, 2).Value + "; "
                    text = text + "Брой полюси: " + wsElboard.Cells(i, 4).Value + "; "
                    text = text + "Номинален ток: " + wsElboard.Cells(i, 6).Value + "; "
                    text = text + "Помощни контакти: " + wsElboard.Cells(i, 5).Value + "; "
                    text = text + "Управляващо напрежение: " + wsElboard.Cells(i, 7).Value + "; "
                    If Trim(wsElboard.Cells(i, 8).Value) <> "" Then
                        text = text + "Клема: " + wsElboard.Cells(i, 8).Value + "; "
                    End If
                Case "iCT"
                    text = text + " - Модулен контактор - "
                    text = text + wsElboard.Cells(i, 2).Value + "; "
                    text = text + "Брой полюси: " + wsElboard.Cells(i, 4).Value + "; "
                    text = text + "Номинален ток: " + wsElboard.Cells(i, 6).Value + "; "
                    text = text + "Помощни контакти: без; "
                    text = text + "Управляващо напрежение: " + wsElboard.Cells(i, 7).Value + "; "
                    If Trim(wsElboard.Cells(i, 8).Value) <> "" Then
                        text = text + "Клема: " + wsElboard.Cells(i, 8).Value + "; "
                    End If
                Case "LC1D", "LC1K", "LC7K", "LP1D", "LP1K", "LP4K"
                    text = text + " - Контактор - "
                    text = text + wsElboard.Cells(i, 2).Value + "; "
                    text = text + "Брой полюси: " + wsElboard.Cells(i, 4).Value + "; "
                    text = text + "Номинален ток: " + wsElboard.Cells(i, 6).Value + "; "
                    text = text + "Помощни контакти: " + wsElboard.Cells(i, 5).Value + "; "
                    text = text + "Управляващо напрежение: " + wsElboard.Cells(i, 7).Value + "; "
                    If Trim(wsElboard.Cells(i, 8).Value) <> "" Then
                        text = text + "Клема: " + wsElboard.Cells(i, 8).Value + "; "
                    End If
                Case "_Тип 2 iPRD", "_Тип 2 iPRD", "_Тип 2 iPF", "_Тип 1+2 PRF1",
                     "_Тип 1+2 PRD1", "_Тип 1 PRF1 Master", "_Тип 1 PRD1 Master"
                    text = text + " - Катоден отводител - "
                    text = text + "Тип: " + wsElboard.Cells(i, 4).Value + "; "
                    text = text + "Ток: " + wsElboard.Cells(i, 6).Value + "; "
                    text = text + "Напрежение: " + wsElboard.Cells(i, 7).Value + "; "
                    text = text + "Брой полюси: " + wsElboard.Cells(i, 5).Value + "; "
                Case "GZ1", "GV4P", "GV3-P", "GV2-ME"
                    text = text + " - Термомагнитен моторен прекъсвач - "
                    text = text + "Pдвиг(400V): " + wsElboard.Cells(i, 4).Value + "; "
                    text = text + "Брой полюси: " + wsElboard.Cells(i, 5).Value + "; "
                    text = text + "Обхват термична защита: " + wsElboard.Cells(i, 6).Value + "; "
                Case Else
                    text = "НЕИЗВЕСТЕН ЕЛЕМЕНТ --> " + wsElboard.Cells(i, 12).Value
            End Select
            wsKol_Smetka.Cells(index, 2).Value = text
            wsKol_Smetka.Cells(index, 3).Value = "бр."
            wsKol_Smetka.Cells(index, 4).Value = wsElboard.Cells(i, 3).Value
            index += 1
            Брой_Връзки_2 = Брой_Връзки_2 + wsElboard.Range("P" & i.ToString).Value
            Брой_Връзки_16 = Брой_Връзки_16 + wsElboard.Range("Q" & i.ToString).Value
            Брой_Връзки_63 = Брой_Връзки_63 + wsElboard.Range("R" & i.ToString).Value
        Next

        With wsKol_Smetka
            Text = "Монтаж на ел. табло '" + Tablo + "', в т.ч.:"
            .Cells(index, 2).Value = Text
            index += 1

            Text = IIf(Нишa_Таблo, " - Направа нишa за ел.таблo", " - Монтаж на табло окачено на стена")
            .Cells(index, 2).Value = Text
            .Cells(index, 3).Value = "бр."
            .Cells(index, 4).Value = 1
            index += 1

            If Брой_Връзки_2 > 0 Then
                Text = " - Свързване проводник към табло до 2,5мм²"
                .Cells(index, 2).Value = Text
                .Cells(index, 3).Value = "бр."
                .Cells(index, 4).Value = Брой_Връзки_2
                index += 1
            End If

            If Брой_Връзки_16 > 0 Then
                Text = " - Свързване проводник към табло до 16мм²"
                .Cells(index, 2).Value = Text
                .Cells(index, 3).Value = "бр."
                .Cells(index, 4).Value = Брой_Връзки_16
                index += 1
            End If

            If Брой_Връзки_63 > 0 Then
                Text = " - Свързване проводник към табло до 16мм²"
                .Cells(index, 2).Value = Text
                .Cells(index, 3).Value = "бр."
                .Cells(index, 4).Value = Брой_Връзки_63
                index += 1
            End If

        End With

        wsKoef.Cells(3, 5).Value = Брой_Дефек
        wsKoef.Cells(5, 5).Value = Брой_табла

        ProgressBar_Extrat.Value = ProgressBar_Extrat.Minimum
    End Sub
    Private Sub Button_силова_КУТИИ_Click(sender As Object, e As EventArgs) Handles Button_силова_КУТИИ.Click
        If IsNothing(excel_Workbook) Then
            MsgBox("Първо да беше отворил файла?")
            Exit Sub
        End If
        Me.Visible = vbFalse
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim edt As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database
        Dim SelectedSet = cu.GetObjects("LWPOLYLINE", "Изберете разклонителни кутии")
        Dim Pline_Count As Integer = SelectedSet.Count
        wsKoef.Cells(10, 2).Value = Pline_Count
        Me.Visible = vbTrue
    End Sub
    Private Sub Button_Генератор_Спецификация_Click(sender As Object, e As EventArgs) Handles Button_Генератор_Спецификация.Click
        Dim index As Integer = 0
        Dim i As Integer
        ProgressBar_Extrat.Maximum = ProgressBar_Maximum
        For i = 6 To 1000
            If wsSpecefikaciq.Cells(i, 2).Value = "Раздел" Or wsSpecefikaciq.Cells(i, 2).Value = "" Then
                index = i
                Exit For
            End If
            ProgressBar_Extrat.Value = i
        Next
        For i = 2 To 1000
            If wsElboard.Cells(i, 2).Value = "" Then
                Exit For
            End If
        Next
        ProgressBar_Extrat.Maximum = i
        Call Excel_Kol_smetka_Razdel(wsSpecefikaciq, "СИЛОВИ РАЗПЕДЕЛИТЕЛНИ ТАБЛА", "A" & Trim(index.ToString), "D" & Trim(index.ToString))
        index += 1
        Dim tablo As String = ""
        For i = 2 To 10000
            Dim text As String = ""
            ProgressBar_Extrat.Value = i
            If wsElboard.Cells(i, 2).Value = "" Then Exit For
            Select Case wsElboard.Cells(i, 2).Value
                Case "Метален шкаф стоящ"
                    text = "Eл. табло "
                    text = text + "'" + wsElboard.Cells(i, 1).Value + "'"
                    Call Excel_Kol_smetka_Razdel(wsSpecefikaciq, text, "B" & Trim(index.ToString), "D" & Trim(index.ToString))
                    index += 1
                    text = " - метален шкаф за стоящ монтаж" + vbCrLf
                    text = text + " - Съответствие с БДС EN 61439-1" + vbCrLf
                    text = text + " - Противовлажно и противопрашно уплътнено" + vbCrLf
                    text = text + " - Степен на защита: IP 66" + vbCrLf
                    text = text + " - Размери: Вис.-" + wsElboard.Cells(i, 4).Value.ToString + "; "
                    text = text + " - Шир. -" + wsElboard.Cells(i, 5).Value.ToString + "; "
                    text = text + " - Дъл. -" + wsElboard.Cells(i, 6).Value.ToString + vbCrLf
                    text = text + IIf(Trim(wsElboard.Cells(i, 7).Value) <> "",
                                      " - Врата: " + Trim(wsElboard.Cells(i, 7).Value),
                                      "")
                Case "Метален шкаф"
                    text = "Eл. табло "
                    text = text + "'" + wsElboard.Cells(i, 1).Value + "'"
                    Call Excel_Kol_smetka_Razdel(wsSpecefikaciq, text, "B" & Trim(index.ToString), "D" & Trim(index.ToString))
                    index += 1
                    text = " - метален шкаф" + vbCrLf
                    text = text + " - Съответствие с БДС EN 61439-1" + vbCrLf
                    text = text + " - Противовлажно и противопрашно уплътнено" + vbCrLf
                    text = text + " - Степен на защита: IP 66" + vbCrLf
                    text = text + " - Размери: Вис.-" + wsElboard.Cells(i, 4).
                    text = text + " - Размери: Вис.-" + wsElboard.Cells(i, 4).Value.ToString + "; "
                    text = text + " - Шир. -" + wsElboard.Cells(i, 5).Value.ToString + "; "
                    text = text + " - Дъл. -" + wsElboard.Cells(i, 6).Value.ToString + vbCrLf
                    text = text + IIf(Trim(wsElboard.Cells(i, 7).Value) <> "",
                                      " - Врата: " + Trim(wsElboard.Cells(i, 7).Value),
                                      "")
                Case "Изпъкнал монтаж"
                    tablo = wsElboard.Cells(i, 1).Value
                    text = "Доставка на ел. табло "
                    text = text + "'" + wsElboard.Cells(i, 1).Value + "'-"
                    text = text + "полиестерен шкаф;"
                    text = text + " Брой модули: " + wsElboard.Cells(i, 4).Value.ToString + "; "
                    If Trim(wsElboard.Cells(i, 6).Value) <> "" Then
                        text = text + "Врата: " + Trim(wsElboard.Cells(i, 6).Value)
                    Else
                        text = text + "Врата: " + Trim(wsElboard.Cells(i, 5).Value)
                    End If
                    text = text + "; в т.ч. доставени и монтирани елементи:"
                    Call Excel_Kol_smetka_Razdel(wsKol_Smetka, text, "B" & Trim(index.ToString), "D" & Trim(index.ToString))
                    index += 1
                    Continue For
                Case "Вграден монтаж"
                    tablo = wsElboard.Cells(i, 1).Value
                    text = "Доставка на ел. табло "
                    text = text + "'" + wsElboard.Cells(i, 1).Value + "'-"
                    text = text + "полиестерен шкаф;"
                    text = text + " Брой модули: " + wsElboard.Cells(i, 4).Value.ToString + "; "
                    If Trim(wsElboard.Cells(i, 6).Value) <> "" Then
                        text = text + "Врата: " + Trim(wsElboard.Cells(i, 6).Value)
                    Else
                        text = text + "Врата: " + Trim(wsElboard.Cells(i, 5).Value)
                    End If
                    text = text + "; в т.ч. доставени и монтирани елементи:"
                    Call Excel_Kol_smetka_Razdel(wsKol_Smetka, text, "B" & Trim(index.ToString), "D" & Trim(index.ToString))
                    index += 1
                    text = "В т.ч. доставени и монтирани елементи:"
                Case "Mini Kaedra"
                Case "Kaedra - щепселни съединения"
                Case "Kaedra"
                Case "Kaedra"
                    '
                    '
                    '
                Case "E60", "Е60", "iK60", "iC60", "EZ9 MCB ",
                     "Е120", "E120", "EZCV250", "C120", "С120"
                    text = text + " - Автоматичен прекъсвач - "
                    text = text + wsElboard.Cells(i, 2).Value + " "
                    text = text + wsElboard.Cells(i, 4).Value + "; "
                    text = text + wsElboard.Cells(i, 5).Value + "; "
                    text = text + wsElboard.Cells(i, 6).Value +
                                  IIf(Trim(wsElboard.Cells(i, 7).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 7).Value +
                                  IIf(Trim(wsElboard.Cells(i, 8).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 8).Value +
                                  IIf(Trim(wsElboard.Cells(i, 13).Value) <> "", "; ", "")
                    text = text + IIf(wsElboard.Cells(i, 13).Value = "+MX", "; +MX", "")
                Case "NG125", "NG160", "NS", "NW"
                    text = text + IIf(wsElboard.Cells(i, 4).Value = "NA",
                                  " - Мощностен разединител - ",
                                  " - Автоматичен прекъсвач - ")
                    text = text + wsElboard.Cells(i, 2).Value + " "
                    text = Trim(text) + IIf(wsElboard.Cells(i, 4).Value = "NA", "; ", " " + wsElboard.Cells(i, 4).Value + "; ")
                    text = text + wsElboard.Cells(i, 5).Value + "; "
                    text = text + wsElboard.Cells(i, 6).Value +
                                  IIf(Trim(wsElboard.Cells(i, 7).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 7).Value +
                                  IIf(Trim(wsElboard.Cells(i, 8).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 8).Value +
                                  IIf(Trim(wsElboard.Cells(i, 13).Value) <> "", "; ", "")
                    text = text + IIf(wsElboard.Cells(i, 13).Value = "+MX", "; +MX", "")
                    text = " " + text
                Case "C60H-DC"
                    text = text + " - Автоматичен прекъсвач за постоянен ток - "
                    text = text + wsElboard.Cells(i, 2).Value + " "
                    text = text + wsElboard.Cells(i, 4).Value + "; "
                    text = text + wsElboard.Cells(i, 5).Value + "; "
                    text = text + wsElboard.Cells(i, 6).Value +
                                  IIf(Trim(wsElboard.Cells(i, 7).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 7).Value +
                                  IIf(Trim(wsElboard.Cells(i, 8).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 8).Value +
                                  IIf(Trim(wsElboard.Cells(i, 13).Value) <> "", "; ", "")
                    text = text + IIf(wsElboard.Cells(i, 13).Value = "+MX", "; +MX", "")
                Case "NSX100", "NSX160", "NSX250", "NSX400", "NSX630"
                    text = text + " - Автоматичен прекъсвач - Compact - "
                    text = text + wsElboard.Cells(i, 2).Value + " "
                    text = text + wsElboard.Cells(i, 4).Value + "; "
                    text = text + wsElboard.Cells(i, 6).Value + "; "
                    text = text + wsElboard.Cells(i, 5).Value +
                                  IIf(Trim(wsElboard.Cells(i, 7).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 7).Value +
                                  IIf(Trim(wsElboard.Cells(i, 8).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 8).Value +
                                  IIf(Trim(wsElboard.Cells(i, 13).Value) <> "", "; ", "")
                    text = text + IIf(wsElboard.Cells(i, 13).Value = "+MX", "; +MX", "")
                Case "NS800", "NS800", "NS630b", "NS1600", "NS1250", "NS1000"
                    text = text + " - Автоматичен прекъсвач - Compact - "
                    text = text + wsElboard.Cells(i, 2).Value + " "
                    text = text + wsElboard.Cells(i, 4).Value + "; "
                    text = text + Mid(Trim(wsElboard.Cells(i, 5).Value), 1, 2) + "; "
                    If Len(Trim(wsElboard.Cells(i, 5).Value)) > 2 Then
                        text = text + Trim(Mid(Trim(wsElboard.Cells(i, 5).Value), 3, Len(wsElboard.Cells(i, 5).Value))) + "; "
                    End If
                    text = text + wsElboard.Cells(i, 6).Value +
                                  IIf(Trim(wsElboard.Cells(i, 7).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 7).Value +
                                  IIf(Trim(wsElboard.Cells(i, 8).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 8).Value +
                                  IIf(Trim(wsElboard.Cells(i, 13).Value) <> "", "; ", "")
                    text = text + IIf(wsElboard.Cells(i, 13).Value = "+MX", "; +MX", "")
                Case "NT06", "NT08", "NT10", "NT12", "NT16"
                    text = text + " - Автоматичен прекъсвач - "
                    text = text + wsElboard.Cells(i, 2).Value + " "
                    text = text + wsElboard.Cells(i, 4).Value + "; "
                    text = text + wsElboard.Cells(i, 5).Value + "; "
                    text = text + wsElboard.Cells(i, 6).Value +
                                  IIf(Trim(wsElboard.Cells(i, 7).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 7).Value +
                                  IIf(Trim(wsElboard.Cells(i, 8).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 8).Value +
                                  IIf(Trim(wsElboard.Cells(i, 13).Value) <> "", "; ", "")
                    text = text + IIf(wsElboard.Cells(i, 13).Value = "+MX", "; +MX", "")
                Case "NW08", "NW10", "NW12", "NW16", "NW20", "NW25", "NW32",
                     "NW40b", "NW40", "NW50", "NW63", "NB600", "ZC100", "EZC250", "EZC400"
                    text = text + " - Автоматичен прекъсвач - "
                    text = text + wsElboard.Cells(i, 2).Value + " "
                    text = text + wsElboard.Cells(i, 4).Value + "; "
                    text = text + wsElboard.Cells(i, 5).Value + "; "
                    text = text + wsElboard.Cells(i, 6).Value +
                                  IIf(Trim(wsElboard.Cells(i, 7).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 7).Value +
                                  IIf(Trim(wsElboard.Cells(i, 8).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 8).Value +
                                  IIf(Trim(wsElboard.Cells(i, 13).Value) <> "", "; ", "")
                    text = text + IIf(wsElboard.Cells(i, 13).Value = "+MX", "; +MX", "")
                Case "EZC400", "EZC250", "EZC100"
                    text = text + " - Автоматичен прекъсвач - EasyPact - "
                    text = text + wsElboard.Cells(i, 2).Value + " "
                    text = text + wsElboard.Cells(i, 4).Value + "; "
                    text = text + wsElboard.Cells(i, 6).Value + "; "
                    text = text + wsElboard.Cells(i, 5).Value +
                                  IIf(Trim(wsElboard.Cells(i, 7).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 7).Value +
                                  IIf(Trim(wsElboard.Cells(i, 8).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 8).Value +
                                  IIf(Trim(wsElboard.Cells(i, 13).Value) <> "", "; ", "")
                    text = text + IIf(wsElboard.Cells(i, 13).Value = "+MX", "; +MX", "")
                Case "iSW", "IN", "INS", "INV"
                    text = text + " - Мощностен разединител -  "
                    text = text + wsElboard.Cells(i, 2).Value + " "
                    text = text + wsElboard.Cells(i, 6).Value + "; "
                    text = text + wsElboard.Cells(i, 5).Value +
                                  IIf(Trim(wsElboard.Cells(i, 7).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 7).Value +
                                  IIf(Trim(wsElboard.Cells(i, 8).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 8).Value +
                                  IIf(Trim(wsElboard.Cells(i, 13).Value) <> "", "; ", "")
                    text = text + IIf(wsElboard.Cells(i, 13).Value = "+MX", "; +MX", "")
                Case "NT"
                    text = text + " - Мощностен разединител - "
                    text = text + wsElboard.Cells(i, 2).Value + " "
                    text = text + wsElboard.Cells(i, 5).Value + "; "
                    text = text + wsElboard.Cells(i, 6).Value +
                                  IIf(Trim(wsElboard.Cells(i, 7).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 7).Value +
                                  IIf(Trim(wsElboard.Cells(i, 8).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 8).Value +
                                  IIf(Trim(wsElboard.Cells(i, 13).Value) <> "", "; ", "")
                Case "DPN N Vigi", "DPNa Vigi", "DPNа Vigi", "EZ9 RCBO"
                    text = text + " - Автоматичен прекъсвач с вградена дефектно токова защита - "
                    text = text + wsElboard.Cells(i, 2).Value + "; "
                    text = text + wsElboard.Cells(i, 4).Value + "; "
                    text = text + wsElboard.Cells(i, 6).Value + "; "
                    text = text + wsElboard.Cells(i, 5).Value +
                                  IIf(Trim(wsElboard.Cells(i, 7).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 7).Value +
                                  IIf(Trim(wsElboard.Cells(i, 8).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 8).Value +
                                  IIf(Trim(wsElboard.Cells(i, 13).Value) <> "", "; ", "")
                Case "ID Domae", "iID", "Vigi iC60", "iID К", "EZ9 RCCB"
                    text = text + " - Дефектнотокова защита - "
                    text = text + wsElboard.Cells(i, 2).Value + "; "
                    text = text + wsElboard.Cells(i, 4).Value + "; "
                    text = text + wsElboard.Cells(i, 6).Value + "; "
                    text = text + wsElboard.Cells(i, 5).Value +
                                  IIf(Trim(wsElboard.Cells(i, 7).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 7).Value +
                                  IIf(Trim(wsElboard.Cells(i, 8).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 8).Value +
                                  IIf(Trim(wsElboard.Cells(i, 13).Value) <> "", "; ", "")
                    wsElboard.Cells(i, 19).Value = wsElboard.Cells(i, 3).Value
                Case "PRC", "PRI"
                    text = text + " - Катоден отводител за телефонна линия - "
                    text = text + wsElboard.Cells(i, 2).Value
                Case "iTL"
                    text = text + " - Импулсно реле - "
                    text = text + wsElboard.Cells(i, 2).Value + " "
                    text = text + wsElboard.Cells(i, 5).Value + "; "
                    text = text + wsElboard.Cells(i, 6).Value +
                                  IIf(Trim(wsElboard.Cells(i, 7).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 7).Value +
                                  IIf(Trim(wsElboard.Cells(i, 8).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 8).Value +
                                  IIf(Trim(wsElboard.Cells(i, 13).Value) <> "", "; ", "")
                Case "IHP", "IH"
                    text = text + " - Програмируемо времереле - "
                    text = text + wsElboard.Cells(i, 2).Value + " "
                    text = text + wsElboard.Cells(i, 7).Value +
                                  IIf(Trim(wsElboard.Cells(i, 8).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 8).Value +
                                  IIf(Trim(wsElboard.Cells(i, 13).Value) <> "", "; ", "")
                Case "MIN", "MINp"
                    text = text + " - Стълбищен автомат - "
                    text = text + wsElboard.Cells(i, 2).Value + " "
                    text = text + wsElboard.Cells(i, 5).Value + "; "
                    text = text + wsElboard.Cells(i, 6).Value +
                                  IIf(Trim(wsElboard.Cells(i, 7).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 7).Value +
                                  IIf(Trim(wsElboard.Cells(i, 8).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 8).Value +
                                  IIf(Trim(wsElboard.Cells(i, 13).Value) <> "", "; ", "")
                Case "IC 100к", "IC Astro", "IC100", "IC2000", "IC2000P+"
                    text = text + " - Фото реле - "
                    text = text + wsElboard.Cells(i, 2).Value + " "
                    text = text + wsElboard.Cells(i, 5).Value + "; "
                    text = text + wsElboard.Cells(i, 6).Value +
                                  IIf(Trim(wsElboard.Cells(i, 7).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 7).Value +
                                  IIf(Trim(wsElboard.Cells(i, 8).Value) <> "", "; ", "")
                    text = text + wsElboard.Cells(i, 8).Value +
                                  IIf(Trim(wsElboard.Cells(i, 13).Value) <> "", "; ", "")
                Case "LC2D", "LC2K", "LC8K", "LP2K", "LP5K"
                    text = text + " - Реверсивен контактор - "
                    text = text + wsElboard.Cells(i, 2).Value + "; "
                    text = text + "Брой полюси: " + wsElboard.Cells(i, 4).Value + "; "
                    text = text + "Номинален ток: " + wsElboard.Cells(i, 6).Value + "; "
                    text = text + "Помощни контакти: " + wsElboard.Cells(i, 5).Value + "; "
                    text = text + "Управляващо напрежение: " + wsElboard.Cells(i, 7).Value + "; "
                    If Trim(wsElboard.Cells(i, 8).Value) <> "" Then
                        text = text + "Клема: " + wsElboard.Cells(i, 8).Value + "; "
                    End If
                Case "iCT"
                    text = text + " - Модулен контактор - "
                    text = text + wsElboard.Cells(i, 2).Value + "; "
                    text = text + "Брой полюси: " + wsElboard.Cells(i, 4).Value + "; "
                    text = text + "Номинален ток: " + wsElboard.Cells(i, 6).Value + "; "
                    text = text + "Помощни контакти: без; "
                    text = text + "Управляващо напрежение: " + wsElboard.Cells(i, 7).Value + "; "
                    If Trim(wsElboard.Cells(i, 8).Value) <> "" Then
                        text = text + "Клема: " + wsElboard.Cells(i, 8).Value + "; "
                    End If
                Case "LC1D", "LC1K", "LC7K", "LP1D", "LP1K", "LP4K"
                    text = text + " - Контактор - "
                    text = text + wsElboard.Cells(i, 2).Value + "; "
                    text = text + "Брой полюси: " + wsElboard.Cells(i, 4).Value + "; "
                    text = text + "Номинален ток: " + wsElboard.Cells(i, 6).Value + "; "
                    text = text + "Помощни контакти: " + wsElboard.Cells(i, 5).Value + "; "
                    text = text + "Управляващо напрежение: " + wsElboard.Cells(i, 7).Value + "; "
                    If Trim(wsElboard.Cells(i, 8).Value) <> "" Then
                        text = text + "Клема: " + wsElboard.Cells(i, 8).Value + "; "
                    End If
                Case "_Тип 2 iPRD", "_Тип 2 iPRD", "_Тип 2 iPF", "_Тип 1+2 PRF1",
                     "_Тип 1+2 PRD1", "_Тип 1 PRF1 Master", "_Тип 1 PRD1 Master"
                    text = text + " - Катоден отводител - "
                    text = text + "Тип: " + wsElboard.Cells(i, 4).Value + "; "
                    text = text + "Ток: " + wsElboard.Cells(i, 6).Value + "; "
                    text = text + "Напрежение: " + wsElboard.Cells(i, 7).Value + "; "
                    text = text + "Брой полюси: " + wsElboard.Cells(i, 5).Value + "; "
                Case Else
                    text = "НЕИЗВЕСТЕН ЕЛЕМЕНТ --> " + wsElboard.Cells(i, 12).Value
            End Select
            wsSpecefikaciq.Cells(index, 2).Value = text
            wsSpecefikaciq.Cells(index, 3).Value = "бр."
            wsSpecefikaciq.Cells(index, 4).Value = wsElboard.Cells(i, 3).Value
            index += 1
        Next
        ProgressBar_Extrat.Value = ProgressBar_Extrat.Minimum
    End Sub
    Private Sub Button_Вземи_МЪЛНИЯ_Click(sender As Object, e As EventArgs) Handles Button_Вземи_МЪЛНИЯ.Click
        If IsNothing(excel_Workbook) Then
            MsgBox("Първо да беше отворил файла?")
            Exit Sub
        End If
        Me.Visible = vbFalse
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim edt As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database
        Dim SelectedSet = cu.GetObjects("INSERT", "Изберете блок")
        Dim arrBlock(500) As strKontakt
        If SelectedSet Is Nothing Then
            MsgBox("Няма маркиран нито един блок.")
            Me.Visible = vbTrue
            Exit Sub
        End If
        Dim blkRecId As ObjectId = ObjectId.Null
        Dim index As Integer = 0
        Me.Visible = vbTrue
        ProgressBar_Extrat.Maximum = SelectedSet.Count
        ProgressBar_Extrat.Value = ProgressBar_Extrat.Minimum
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                For Each sObj As SelectedObject In SelectedSet

                    ProgressBar_Extrat.Value += 1

                    blkRecId = sObj.ObjectId

                    Dim acBlkRef As BlockReference = DirectCast(acTrans.GetObject(blkRecId, OpenMode.ForRead), BlockReference)
                    Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                    Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection
                    Dim Visibility As String = ""

                    For Each prop As DynamicBlockReferenceProperty In props
                        If prop.PropertyName = "Visibility1" Then Visibility = prop.Value
                        If prop.PropertyName = "Visibility" Then Visibility = prop.Value
                        If prop.PropertyName = "Тип" Then Visibility = prop.Value
                    Next
                    Dim blName As String = (CType(acBlkRef.DynamicBlockTableRecord.GetObject(OpenMode.ForRead), BlockTableRecord)).Name
                    Dim iVisib As Integer = -1
                    Dim strWis As String = ""

                    Select Case blName
                        Case "Заземление"
                            For Each objID As ObjectId In attCol
                                Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                                Dim acAttRef As AttributeReference = dbObj
                                If acAttRef.Tag = "ВИС" Then strWis = acAttRef.TextString
                            Next
                            iVisib = Array.FindIndex(arrBlock, Function(f) f.blName = blName And
                                                                           f.blVisibility = Visibility 'And f.blWis = strWis
                                                                           )
                        Case Else
                            Continue For
                    End Select

                    If iVisib = -1 Then
                        arrBlock(index).blVisibility = Visibility
                        arrBlock(index).blName = blName
                        arrBlock(index).count = 1
                        Select Case blName
                            Case "Заземление"
                                arrBlock(index).blWis = strWis
                        End Select
                        index += 1
                    Else
                        arrBlock(iVisib).count = arrBlock(iVisib).count + 1
                    End If

                Next
                index = 2
                ProgressBar_Extrat.Value = ProgressBar_Extrat.Minimum
                For Each iarrBlock In arrBlock
                    If iarrBlock.count = 0 Then Exit For
                    ProgressBar_Extrat.Value += 1
                    ' #################################################################################################################################
                    Dim broj_elementi As Integer = iarrBlock.count
                    Select Case iarrBlock.blName
                        Case "Заземление"
                            wsLines.Cells(2, 50).Value = iarrBlock.count
                    End Select
                    ' #################################################################################################################################
                    index += 1
                Next
                acTrans.Commit()
            Catch ex As Exception
                MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
                acTrans.Abort()
            End Try
        End Using
        ProgressBar_Extrat.Value = ProgressBar_Extrat.Minimum
    End Sub
    Private Sub Button_МЪЛНИЯ_Приемник_Click(sender As Object, e As EventArgs) Handles Button_МЪЛНИЯ_Приемник.Click
        If IsNothing(excel_Workbook) Then
            MsgBox("Първо да беше отворил файла?")
            Exit Sub
        End If
        Me.Visible = vbFalse
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim edt As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database
        Dim SelectedSet = cu.GetObjects("INSERT", "Изберете блок")
        Dim arrBlock(500) As strKontakt
        Me.Visible = vbTrue
        If SelectedSet Is Nothing Then
            MsgBox("Няма маркиран нито един блок.")
            Me.Visible = vbTrue
            Exit Sub
        End If
        Dim blkRecId As ObjectId = ObjectId.Null

        Dim Hm As Double = 0
        Dim Rb As Double = 0
        Dim Tip As String = ""
        Dim Kategoroq As String = ""

        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                For Each sObj As SelectedObject In SelectedSet

                    ProgressBar_Extrat.Value += 1
                    blkRecId = sObj.ObjectId

                    Dim acBlkRef As BlockReference = DirectCast(acTrans.GetObject(blkRecId, OpenMode.ForRead), BlockReference)
                    Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                    Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection

                    For Each prop As DynamicBlockReferenceProperty In props
                        If prop.PropertyName = "Hm" Then Hm = prop.Value
                        If prop.PropertyName = "Rb" Then Rb = prop.Value
                        If prop.PropertyName = "Тип" Then Tip = prop.Value
                        If prop.PropertyName = "Категория" Then Kategoroq = prop.Value
                    Next
                Next
                With wsKoef
                    .Cells(12, 2).Value = Tip
                    .Cells(13, 2).Value = Kategoroq
                    .Cells(14, 2).Value = Hm / 100
                    .Cells(15, 2).Value = Rb / 100
                End With
                acTrans.Commit()
            Catch ex As Exception
                MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
                acTrans.Abort()
            End Try
        End Using
        ProgressBar_Extrat.Value = ProgressBar_Extrat.Minimum
    End Sub
    Private Sub Button_Генерирай_МЪЛНИЯ_Click(sender As Object, e As EventArgs) Handles Button_Генерирай_МЪЛНИЯ.Click
        Dim index As Integer = 0
        Dim i As Integer
        For i = 6 To 1000
            If wsKol_Smetka.Cells(i, 2).Value = "Раздел" Or wsKol_Smetka.Cells(i, 2).Value = "" Then
                index = i
                Exit For
            End If
        Next
        Call Excel_Kol_smetka_Razdel(wsKol_Smetka, "МЪЛНИЕЗАЩИТНА ИНСТАЛАЦИЯ", "A" & Trim(index.ToString), "D" & Trim(index.ToString))
        index += 1
        Dim Text As String = ""
        ' Проверка дали дадена клетка в Excel не е празна
        If wsKoef.Cells(14, 2).Value IsNot Nothing Then
            Text = "Доставка и монтаж на мълниеприемник с изпреварващо действие, ниво на защита: "
            Text = Text + wsKoef.Cells(13, 2).Value
            Text = Text + " при h(m) ="
            Text = Text + wsKoef.Cells(14, 2).Value.ToString
            Text = Text + ", защитен радиус Rз= "
            Text = Text + wsKoef.Cells(15, 2).Value.ToString
            Text = Text + "m"
            With wsKol_Smetka
                .Cells(index, 2).Value = Text
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "бр."
                .Cells(index, 4).Value = 1
            End With
            index += 1
            Text = "Доставка и монтаж на мълниеприемна мачта с дължина: H="
            Text = Text + (wsKoef.Cells(14, 2).Value + 1).ToString
            Text = Text + "m"
            With wsKol_Smetka
                .Cells(index, 2).Value = Text
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "бр."
                .Cells(index, 4).Value = 1
            End With
            index += 1
            Text = "Доставка и монтаж на крепежни елементи за укрепване за мълниеприемна мачта"
            With wsKol_Smetka
                .Cells(index, 2).Value = Text
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "копл."
                .Cells(index, 4).Value = 1
            End With          '
        Else
            ' Клетката е празна, пропуснете кодовия блок
        End If

        index += 1
        index = Kol_Smetka_Kabeli(index, vbFalse, "МЪЛНИЯ")
        Text = "Доставка и монтаж на контролна клема"
        With wsKol_Smetka
            .Cells(index, 2).Value = Text
            .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
            .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
            .Cells(index, 3).Value = "бр."
            .Cells(index, 4).Value = wsLines.Cells(2, 50).value
        End With
        index += 1
        Text = "Доставка и монтаж на поцинковани заземителени колoве 63х63х4х1500mm" +
               " + поцинкована шина 40х3х1500mm"
        With wsKol_Smetka
            .Cells(index, 2).Value = Text
            .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
            .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
            .Cells(index, 3).Value = "бр."
            .Cells(index, 4).Value = wsLines.Cells(2, 50).value * 2
        End With
        wsKoef.Cells(2, 5).Value = wsLines.Cells(2, 50).value
        index += 1
        Text = "Доставка на поцинкована шина 40х4mm"
        With wsKol_Smetka
            .Cells(index, 2).Value = Text
            .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
            .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
            .Cells(index, 3).Value = "m"
            .Cells(index, 4).Value = wsLines.Cells(2, 50).value * 3
        End With
        index += 1
        Text = "Направа на изкоп 0,80х0,40m, с обратно зариване и трамбоване"
        With wsKol_Smetka
            .Cells(index, 2).Value = Text
            .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
            .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
            .Cells(index, 3).Value = "m"
            .Cells(index, 4).Value = wsLines.Cells(2, 50).value * 3 + 1
        End With
        index += 1
        Text = "Полагане и свързване на поцинкована шина 40х4mm към заземители и контролна клема"
        With wsKol_Smetka
            .Cells(index, 2).Value = Text
            .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
            .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
            .Cells(index, 3).Value = "бр."
            .Cells(index, 4).Value = wsLines.Cells(2, 50).value * 2
        End With
        index += 1
        Text = "Направа на антикорозионно покритие в местата на заварката"
        With wsKol_Smetka
            .Cells(index, 2).Value = Text
            .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
            .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
            .Cells(index, 3).Value = "бр."
            .Cells(index, 4).Value = wsLines.Cells(2, 50).value * 2
        End With
        index += 1
        ProgressBar_Extrat.Value = ProgressBar_Extrat.Minimum
    End Sub
    Private Sub Button_Вземи_ВЪНШНО_Click(sender As Object, e As EventArgs) Handles Button_Вземи_ВЪНШНО.Click
        If IsNothing(excel_Workbook) Then
            MsgBox("Първо да беше отворил файла?")
            Exit Sub
        End If
        Me.Visible = vbFalse
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim edt As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database
        Dim SelectedSet = cu.GetObjects("INSERT", "Изберете блоковете в чертеж за външно захранване на сградата:")
        Dim arrBlock = cu.GET_Zazemlenie(SelectedSet)
        If SelectedSet Is Nothing Then
            MsgBox("Няма маркиран нито един блок.")
            Me.Visible = vbTrue
            Exit Sub
        End If
        Dim blkRecId As ObjectId = ObjectId.Null
        Me.Visible = vbTrue
        ProgressBar_Extrat.Maximum = SelectedSet.Count
        ProgressBar_Extrat.Value = ProgressBar_Extrat.Minimum
        With wsKoef
            .Cells(red_Външно + 0, 1).Value = "Табло същ.със заземител"
            .Cells(red_Външно + 1, 1).Value = "Табло същ.без заземител"
            .Cells(red_Външно + 2, 1).Value = "Табло електромерно стоящо"
            .Cells(red_Външно + 3, 1).Value = "Табло електромерно на стълб"
            .Cells(red_Външно + 4, 1).Value = "Табло електромерно на пилoн"
            .Cells(red_Външно + 5, 1).Value = "Табло ГРТ - СЪС Заземител"
            .Cells(red_Външно + 6, 1).Value = "Табло ГРТ - БЕЗ Заземител"
            .Cells(red_Външно + 7, 1).Value = "Стълб-съществуващ"
            .Cells(red_Външно + 8, 1).Value = "Стълб-нов НЦ 835"
            .Cells(red_Външно + 9, 1).Value = "Стълб-нов НЦ 590"
            .Cells(red_Външно + 10, 1).Value = "Стълб-нов НЦ 250"
            .Cells(red_Външно + 11, 1).Value = "Стълб-нов"
            .Cells(red_Външно + 12, 1).Value = "Сечение-траншея"
            .Cells(red_Външно + 13, 1).Value = "Сечение"
            .Cells(red_Външно + 14, 1).Value = "Репер"
            .Cells(red_Външно + 15, 1).Value = "Пилон-нов СЪС Заземител"
            .Cells(red_Външно + 16, 1).Value = "Пилон-нов БЕЗ Заземител"
            .Cells(red_Външно + 17, 1).Value = "Опъвач_Регулируем"
            .Cells(red_Външно + 18, 1).Value = "Опъвач_НЕрегулируем"
            .Cells(red_Външно + 19, 1).Value = "Надпокривна конзола"
            .Cells(red_Външно + 20, 1).Value = "Заземител-СЪС контролна клема"
            .Cells(red_Външно + 21, 1).Value = "Заземител-БЕЗ контролна клема"

            .Cells(red_Външно + 0, 4).Value = "Табло електромерно на СЪЩЕСТУВАЩ стълб"
            .Cells(red_Външно + 1, 4).Value = "Табло електромерно на СЪЩЕСТУВАЩ  пилoн"

            .Cells(red_Външно + 0, 5).Value = 0
            .Cells(red_Външно + 1, 5).Value = 0



            .Cells(red_Външно + 0, 2).Value = 0
            .Cells(red_Външно + 1, 2).Value = 0
            .Cells(red_Външно + 2, 2).Value = 0
            .Cells(red_Външно + 3, 2).Value = 0
            .Cells(red_Външно + 4, 2).Value = 0
            .Cells(red_Външно + 5, 2).Value = 0
            .Cells(red_Външно + 6, 2).Value = 0
            .Cells(red_Външно + 7, 2).Value = 0
            .Cells(red_Външно + 8, 2).Value = 0
            .Cells(red_Външно + 9, 2).Value = 0
            .Cells(red_Външно + 10, 2).Value = 0
            .Cells(red_Външно + 11, 2).Value = 0
            .Cells(red_Външно + 12, 2).Value = 0
            .Cells(red_Външно + 13, 2).Value = 0
            .Cells(red_Външно + 14, 2).Value = 0
            .Cells(red_Външно + 15, 2).Value = 0
            .Cells(red_Външно + 16, 2).Value = 0
            .Cells(red_Външно + 17, 2).Value = 0
            .Cells(red_Външно + 18, 2).Value = 0
            .Cells(red_Външно + 19, 2).Value = 0
            .Cells(red_Външно + 20, 2).Value = 0
            .Cells(red_Външно + 21, 2).Value = 0

        End With
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                ProgressBar_Extrat.Value = ProgressBar_Extrat.Minimum
                For Each iarrBlock In arrBlock
                    If iarrBlock.count = 0 Then Exit For
                    Select Case iarrBlock.blVisibility
                        Case "Табло същ.със заземител"
                            wsKoef.Cells(red_Външно + 0, 2).Value = iarrBlock.count
                        Case "Табло същ.без заземител"
                            wsKoef.Cells(red_Външно + 1, 2).Value = iarrBlock.count
                        Case "Табло електромерно стоящо"
                            wsKoef.Cells(red_Външно + 2, 2).Value = iarrBlock.count
                        Case "Табло електромерно на стълб"
                            wsKoef.Cells(red_Външно + 3, 2).Value = iarrBlock.count
                        Case "Табло електромерно на пилoн"
                            wsKoef.Cells(red_Външно + 4, 2).Value = iarrBlock.count
                        Case "Табло ГРТ - СЪС Заземител"
                            wsKoef.Cells(red_Външно + 5, 2).Value = iarrBlock.count
                        Case "Табло ГРТ - БЕЗ Заземител"
                            wsKoef.Cells(red_Външно + 6, 2).Value = iarrBlock.count
                        Case "Стълб-съществуващ"
                            wsKoef.Cells(red_Външно + 7, 2).Value = iarrBlock.count
                        Case "Стълб-нов НЦ 835"
                            wsKoef.Cells(red_Външно + 8, 2).Value = iarrBlock.count
                        Case "Стълб-нов НЦ 590"
                            wsKoef.Cells(red_Външно + 9, 2).Value = iarrBlock.count
                        Case "Стълб-нов НЦ 250"
                            wsKoef.Cells(red_Външно + 10, 2).Value = iarrBlock.count
                        Case "Стълб-нов"
                            wsKoef.Cells(red_Външно + 11, 2).Value = iarrBlock.count
                        Case "Сечение-траншея"
                            wsKoef.Cells(red_Външно + 12, 2).Value = iarrBlock.count
                        Case "Сечение"
                            wsKoef.Cells(red_Външно + 13, 2).Value = iarrBlock.count
                        Case "Репер"
                            wsKoef.Cells(red_Външно + 14, 2).Value = iarrBlock.count
                        Case "Пилон-нов СЪС Заземител"
                            wsKoef.Cells(red_Външно + 15, 2).Value = iarrBlock.count
                        Case "Пилон-нов БЕЗ Заземител"
                            wsKoef.Cells(red_Външно + 16, 2).Value = iarrBlock.count
                        Case "Опъвач_Регулируем"
                            wsKoef.Cells(red_Външно + 17, 2).Value = iarrBlock.count
                        Case "Опъвач_НЕрегулируем"
                            wsKoef.Cells(red_Външно + 18, 2).Value = iarrBlock.count
                        Case "Надпокривна конзола"
                            wsKoef.Cells(red_Външно + 19, 2).Value = iarrBlock.count
                        Case "Заземител-СЪС контролна клема"
                            wsKoef.Cells(red_Външно + 20, 2).Value = iarrBlock.count
                        Case "Заземител-БЕЗ контролна клема"
                            wsKoef.Cells(red_Външно + 21, 2).Value = iarrBlock.count


                        Case "Табло електромерно на СЪЩЕСТУВАЩ стълб"
                            wsKoef.Cells(red_Външно + 0, 5).Value = iarrBlock.count
                        Case "Табло електромерно на СЪЩЕСТУВАЩ  пилoн"
                            wsKoef.Cells(red_Външно + 1, 5).Value = iarrBlock.count

                    End Select
                Next
                acTrans.Commit()
            Catch ex As Exception
                MsgBox("Възникна грешка:   " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
                acTrans.Abort()
            End Try
        End Using
        ProgressBar_Extrat.Value = ProgressBar_Extrat.Minimum
    End Sub
    Private Sub Button_Генератор_ВЪНШНО_Click(sender As Object, e As EventArgs) Handles Button_Генератор_ВЪНШНО.Click
        Dim response = MsgBox("За изчисление на кабелите на които ще се прави разделка е необходимо да се ЗАПИШЕ броя на кабелите!" & vbCrLf & vbCrLf & "Ако няма кабели със сечение над 16mm² не е необходмо.", vbYesNo)
        If response = vbNo Then
            Exit Sub
        End If
        Dim index As Integer = 0
        Dim i As Integer
        ProgressBar_Extrat.Maximum = ProgressBar_Maximum
        For i = 6 To 1000
            If wsKol_Smetka.Cells(i, 2).Value = "Раздел" Or wsKol_Smetka.Cells(i, 2).Value = "" Then
                index = i
                Exit For
            End If
            ProgressBar_Extrat.Value = i
        Next
        For i = 2 To 1000
            If wsKontakti.Cells(i, 2).Value = "" Then
                Exit For
            End If
        Next
        Call Excel_Kol_smetka_Razdel(wsKol_Smetka, "ВЪНШНО ЗАХРАНВАНЕ", "A" & Trim(index.ToString), "D" & Trim(index.ToString))
        index += 1
        Dim Text_Dostawka As String = ""
        Dim Кабел_Общо As Double = 0
        '
        ' ИЗКОП
        '
        For i = 3 To 2000
            If Trim(wsLines.Cells(i, 2).Value) = "" Then Exit For
            ProgressBar_Extrat.Value = i
            If wsLines.Cells(i, 1).Value = "ТРАНШЕЯ-КАБЕЛ" Then
                Кабел_Общо = Кабел_Общо + wsLines.Cells(i, 6).Value
            End If
        Next
        If Кабел_Общо > 0 Then
            With wsKol_Smetka
                .Cells(index, 2).Value = "Трасиране на кабелна линия по равен терен с колчета"
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "m"
                .Cells(index, 4).Value = Кабел_Общо
            End With
            index += 1
            With wsKol_Smetka
                For i = 3 To 2000
                    If Trim(wsLines.Cells(i, 2).Value) = "" Then Exit For
                    If wsLines.Cells(i, 1).Value = "ТРАНШЕЯ-КАБЕЛ" Then
                        Dim dddd As String = wsLines.Cells(i, 2).Value
                        .Cells(index, 2).Value = "Направа на изкоп " + wsLines.Cells(i, 2).Value + ", с обратно зариване и трамбоване"
                        .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                        .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                        .Cells(index, 3).Value = "m"
                        .Cells(index, 4).Value = wsLines.Cells(i, 6).Value
                        index += 1
                    End If
                Next
            End With
            With wsKol_Smetka
                .Cells(index, 2).Value = "Направа на кабелна подложка с пясък/пресята пръст с дeбелина 10cm"
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "m"
                .Cells(index, 4).Value = Кабел_Общо
            End With
            index += 1
            With wsKol_Smetka
                .Cells(index, 2).Value = "Доставка и полагане на сигнална лента, ''Внимание електрически кабел''; Широчина 200 мм; жълта и дебелина не по-малка от 0,25 мм"
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "m"
                .Cells(index, 4).Value = Кабел_Общо
            End With
            index += 1
        End If
        '
        ' РЕПЕРИ
        '
        If wsKoef.Cells(29, 2).Value > 0 Then
            With wsKol_Smetka
                .Cells(index, 2).Value = "Направа репери за кабелна линия"
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "бр."
                .Cells(index, 4).Value = wsKoef.Cells(red_Външно + 14, 2).Value
            End With
            index += 1
        End If
        '
        ' Записва кабели в тръби
        '

        index = Kol_Smetka_Kabeli(index, vbTrue, "ВЪНШНО")
        '
        'Изправяне на ПИЛОН 
        '
        If (wsKoef.Cells(23, 2).Value + wsKoef.Cells(24, 2).Value) > 0 Then
            Text_Dostawka = "Доставка и изправяне на пилон H=7,5m;"
            With wsKol_Smetka
                .Cells(index, 2).Value = Text_Dostawka
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "бр."
                .Cells(index, 4).Value = wsKoef.Cells(red_Външно + 15, 2).Value +   ' Пилон-нов СЪС Заземител
                                         wsKoef.Cells(red_Външно + 16, 2).Value     ' Пилон-нов БЕЗ Заземител
            End With
            index += 1
        End If
        '
        'Изправяне на СТЪЛБ
        '
        If (wsKoef.Cells(25, 2).Value) > 0 Then
            Text_Dostawka = "Доставка и изправяне на СБС H=9,5m;"
            With wsKol_Smetka
                .Cells(index, 2).Value = Text_Dostawka
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "бр."
                .Cells(index, 4).Value = wsKoef.Cells(red_Външно + 8, 2).Value +    ' Стълб-нов НЦ 835
                                         wsKoef.Cells(red_Външно + 9, 2).Value +    ' Стълб-нов НЦ 590
                                         wsKoef.Cells(red_Външно + 10, 2).Value +   ' Стълб-нов НЦ 250
                                         wsKoef.Cells(red_Външно + 11, 2).Value     ' Стълб-нов
            End With
            index += 1
        End If
        '
        ' КРЕПЕЖНИ ЕЛЕМЕНТИ ЗА УСУКАН КАБЕЛ
        '
        Dim krep As Integer = wsKoef.Cells(red_Външно + 17, 2).Value + ' "Опъвач_Регулируем"
                              wsKoef.Cells(red_Външно + 18, 2).Value + ' "Опъвач_НЕрегулируем"
                              wsKoef.Cells(red_Външно + 19, 2).Value   ' "Надпокривна конзола"

        Dim stylb = wsKoef.Cells(red_Външно + 3, 2).Value +  ' "Табло електромерно на стълб"
                    wsKoef.Cells(red_Външно + 4, 2).Value +  ' "Табло електромерно на пилoн"
                    wsKoef.Cells(red_Външно + 7, 2).Value +  '  "Стълб-съществуващ"
                    wsKoef.Cells(red_Външно + 8, 2).Value +  '  "Стълб-нов НЦ 835"
                    wsKoef.Cells(red_Външно + 9, 2).Value +  '  "Стълб-нов НЦ 590"
                    wsKoef.Cells(red_Външно + 10, 2).Value + '  "Стълб-нов НЦ 250"
                    wsKoef.Cells(red_Външно + 11, 2).Value + '  "Стълб-нов"
                    wsKoef.Cells(red_Външно + 15, 2).Value + '  "Пилон-нов СЪС Заземител"
                    wsKoef.Cells(red_Външно + 0, 5).Value +  '  "Табло електромерно на СЪЩЕСТУВАЩ стълб"
                    wsKoef.Cells(red_Външно + 1, 5).Value ' Табло електромерно на СЪЩЕСТУВАЩ  пилoн





        If krep > 0 Or stylb > 0 Then
            ' проверка брой крепежни елементи в чертежа
            If krep = 0 Then
                MsgBox("Някой трябваше да се сети да сложи ОПЪВАЧИ!!!")
            End If
            If stylb = 0 Then
                MsgBox("Те'з ОПЪВАЧИ де жа ги сложиш????")
            End If
            If krep Mod 2 = 0 Then
                krep = krep / 2
            Else
                krep = 0
                MsgBox("Броя на Опъвачите в чертежа не се дели на 2!!!")
            End If
            If krep <> stylb Then
                MsgBox("Броя на Опъвачите в чертежа не съответства на броя на стълбовете!!!")
            End If
            If krep <> 0 Then
                Text_Dostawka = "комплект крепежни елементи за усукан кабел;"
                Text_Dostawka += " в т.ч. опъвач - нерегулируем;"
                Text_Dostawka += " опъвач - регулируем;"
                Text_Dostawka += IIf(wsKoef.Cells(red_Външно + 19, 2).Value = 0, " кука с дюбел Ф12мм;", "")
            End If
            Text_Dostawka += " пресови маншони"
            With wsKol_Smetka
                .Cells(index, 2).Value = "Доставка и монтаж на " + Text_Dostawka
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "бр."
                .Cells(index, 4).Value = krep
            End With
            index += 1
            Text_Dostawka = "Полагане на кабел тип Al/R по стълб;"
            Text_Dostawka += " в т.ч. крепежни елементи;"
            With wsKol_Smetka
                .Cells(index, 2).Value = Text_Dostawka
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "m"
                .Cells(index, 4).Value = 6
            End With
            index += 1
        End If
        '
        ' Надпокривна конзола
        '
        If wsKoef.Cells(red_Външно + 19, 2).Value > 0 Then
            Text = "Доставка и монтаж на надпокривна конзола с дължина Н=2м, със заварена скоба за захващане на опъвача."
            With wsKol_Smetka
                .Cells(index, 2).Value = Text
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "бр."
                .Cells(index, 4).Value = wsKoef.Cells(red_Външно + 19, 2).Value   ' "Надпокривна конзола"
            End With
            index += 1
        End If
        '
        ' ЗАЗЕМИТЕЛИ
        '
        Dim Broj_zazemiteli = wsKoef.Cells(red_Външно + 5, 2).Value * 2 +   ' "Табло ГРТ - СЪС Заземител"
                              wsKoef.Cells(red_Външно + 15, 2).Value +      ' "Пилон-нов СЪС Заземител"
                              wsKoef.Cells(red_Външно + 20, 2).Value * 2 +  ' "Заземител-СЪС контролна клема"
                              wsKoef.Cells(red_Външно + 21, 2).Value * 2    ' "Заземител-БЕЗ контролна клема"

        Dim Broj_Kontrol_Klemi = wsKoef.Cells(red_Външно + 20, 2).Value +   ' "Заземител-СЪС контролна клема"
                                 wsKoef.Cells(red_Външно + 5, 2).Value      ' "Табло ГРТ - СЪС Заземител"

        Dim Broj_izmerwaniq = wsKoef.Cells(red_Външно + 0, 2).Value +       ' "Табло същ.със заземител"
                              wsKoef.Cells(red_Външно + 5, 2).Value +       ' "Табло ГРТ - СЪС Заземител"
                              wsKoef.Cells(red_Външно + 15, 2).Value +      ' "Пилон-нов СЪС Заземител"
                              wsKoef.Cells(red_Външно + 15, 2).Value +      ' "Пилон-нов СЪС Заземител"
                              wsKoef.Cells(red_Външно + 20, 2).Value +      ' "Заземител-СЪС контролна клема"
                              wsKoef.Cells(red_Външно + 21, 2).Value        ' "Заземител-БЕЗ контролна клема"

        wsKoef.Cells(1, 5).Value = Broj_izmerwaniq
        If wsKoef.Cells(1, 5).Value > 0 Then
            Text = "Доставка и монтаж на поцинковани заземителени колoве 63х63х5х1500mm"
            Text += " с поцинкована шина 40х4х1500mm"
            With wsKol_Smetka
                .Cells(index, 2).Value = Text
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "бр."
                .Cells(index, 4).Value = Broj_zazemiteli
            End With
            index += 1
            Text = "Доставка на поцинкована шина 40х4mm"
            With wsKol_Smetka
                .Cells(index, 2).Value = Text
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "m"
                .Cells(index, 4).Value = Broj_zazemiteli * 3
            End With
            index += 1
            If Broj_Kontrol_Klemi > 0 Then
                Text = "Доставка и монтаж на контролна клема"
                With wsKol_Smetka
                    .Cells(index, 2).Value = Text
                    .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                    .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                    .Cells(index, 3).Value = "бр."
                    .Cells(index, 4).Value = Broj_Kontrol_Klemi
                End With
                index += 1
                Text = "Полагане и свързване на поцинкована шина 40х4mm към заземители и контролна клема"
                With wsKol_Smetka
                    .Cells(index, 2).Value = Text
                    .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                    .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                    .Cells(index, 3).Value = "бр."
                    .Cells(index, 4).Value = Broj_Kontrol_Klemi
                End With
                index += 1
            End If
            Text = "Направа на антикорозионно покритие в местата на заварката"
            With wsKol_Smetka
                .Cells(index, 2).Value = Text
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "бр."
                .Cells(index, 4).Value = Broj_zazemiteli
            End With
            index += 1
        End If
        '
        ' СТРОИТЕЛНИ ОТПАДЪЦИ
        '
        If Кабел_Общо > 0 Then
            Dim Отпадък As Double = Кабел_Общо * 0.5 * 0.025
            With wsKol_Smetka
                .Cells(index, 2).Value = "Механично натоварване на строителни отпадъци и излишна земна пръст"
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "m³"
                .Cells(index, 4).Value = Отпадък * 4 / 5
                .Range("D" & Trim(index.ToString)).NumberFormat = "#,##0.00"
            End With
            index += 1
            With wsKol_Smetka
                .Cells(index, 2).Value = "Ръчно натоварване на строителни отпадъци и излишна земна пръст"
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "m³"
                .Cells(index, 4).Value = Отпадък * 1 / 5
                .Range("D" & Trim(index.ToString)).NumberFormat = "#,##0.00"
            End With
            index += 1
            With wsKol_Smetka
                .Cells(index, 2).Value = "Почистване и извозване на строителни отпадъци и излишна земна пръст"
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "m³"
                .Cells(index, 4).Value = Отпадък
                .Range("D" & Trim(index.ToString)).NumberFormat = "#,##0.00"
            End With
            index += 1
        End If
        ProgressBar_Extrat.Value = ProgressBar_Extrat.Minimum
    End Sub
    Private Sub Button_Протоколи_Click(sender As Object, e As EventArgs) Handles Button_Протоколи.Click
        Dim index As Integer = 0
        Dim i As Integer
        Dim Text_Dostawka As String = ""
        For i = 6 To 1000
            If wsKol_Smetka.Cells(i, 2).Value = "Раздел" Or wsKol_Smetka.Cells(i, 2).Value = "" Then
                index = i
                Exit For
            End If
        Next
        Call Excel_Kol_smetka_Razdel(wsKol_Smetka, "ПУСКОВО-НАЛАДЪЧНИ РАБОТИ", "A" & Trim(index.ToString), "D" & Trim(index.ToString))
        index += 1
        Dim Text As String =
                "Сертификат за контрол в т.ч. протоколи за: "
        Call Excel_Kol_smetka_Razdel(wsKol_Smetka, Text, "B" & Trim(index.ToString), "D" & Trim(index.ToString))
        index += 1
        If wsKoef.Cells(1, 5).Value > 0 Then
            Text_Dostawka = "Контрол на съпротивление на заземителна уредба"
            With wsKol_Smetka
                .Cells(index, 2).Value = Text_Dostawka
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "бр. т."
                .Cells(index, 4).Value = wsKoef.Cells(1, 5).Value
            End With
            index += 1
        End If
        If wsKoef.Cells(2, 5).Value > 0 Then
            Text_Dostawka = "Контрол на съпротивление на мълниезащитна заземителна уредба"
            With wsKol_Smetka
                .Cells(index, 2).Value = Text_Dostawka
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "бр. т."
                .Cells(index, 4).Value = wsKoef.Cells(2, 5).Value
            End With
            index += 1
        End If
        If wsKoef.Cells(3, 5).Value > 0 Then
            Text_Dostawka = "Контрол на функционална годност на защитни прекъсвачи за токове с нулева последователност (RCD) в схеми TN и TT"
            With wsKol_Smetka
                .Cells(index, 2).Value = Text_Dostawka
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "бр. т."
                .Cells(index, 4).Value = wsKoef.Cells(3, 5).Value
            End With
            index += 1
        End If
        If wsKoef.Cells(4, 5).Value > 0 Then
            Text_Dostawka = "Контрол на импедансът Zs на контур 'фаза-защитен проводник'"
            With wsKol_Smetka
                .Cells(index, 2).Value = Text_Dostawka
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "бр. т."
                .Cells(index, 4).Value = wsKoef.Cells(4, 5).Value
            End With
            index += 1
        End If
        If wsKoef.Cells(5, 5).Value > 0 Then
            Text_Dostawka = "Контрол изолационно съпротивление на захранващи кабели"
            With wsKol_Smetka
                .Cells(index, 2).Value = Text_Dostawka
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "бр. т."
                .Cells(index, 4).Value = wsKoef.Cells(5, 5).Value
            End With
            index += 1
        Else
            With wsLines
                ' Намиране на последния използван ред в колона BH
                Dim lastRow As Integer = .Cells(.Rows.Count, "BH").End(excel.XlDirection.xlUp).Row

                ' Инициализиране на сумата
                Dim sum As Double = 0

                ' Цикъл през всички клетки в колона BH до последния използван ред
                For i = 1 To lastRow
                    Dim cellValue As Object = .Cells(i, "BH").Value
                    If Not IsNothing(cellValue) AndAlso IsNumeric(cellValue) Then
                        sum += CDbl(cellValue)
                    End If
                Next
                wsKoef.Cells(5, 5).Value = sum
            End With
            Text_Dostawka = "Контрол изолационно съпротивление на захранващи кабели"
            With wsKol_Smetka
                .Cells(index, 2).Value = Text_Dostawka
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "бр. т."
                .Cells(index, 4).Value = wsKoef.Cells(5, 5).Value
            End With
            index += 1
        End If
    End Sub
    Private Sub Button_Генерирай_СКАРИ_Click(sender As Object, e As EventArgs) Handles Button_Генерирай_СКАРИ.Click
        Dim index As Integer = 0
        Dim i As Integer
        ProgressBar_Extrat.Maximum = ProgressBar_Maximum
        For i = 6 To 1000
            If wsKol_Smetka.Cells(i, 2).Value = "Раздел" Or wsKol_Smetka.Cells(i, 2).Value = "" Then
                index = i
                Exit For
            End If
            ProgressBar_Extrat.Value = i
        Next
        For i = 2 To 1000
            If wsKontakti.Cells(i, 2).Value = "" Then
                Exit For
            End If
        Next
        ProgressBar_Extrat.Maximum = ProgressBar_Maximum

        Dim Text_Dostawka As String = ""
        Dim broj_elementi As Integer = 0
        Dim Kanal As Integer = 0
        Dim skara As Integer = 0
        Dim ygli As Integer = 0
        Dim Planka_35 As Integer = 0
        Dim Planka_60 As Integer = 0
        Dim Planka_85 As Integer = 0
        Dim Planka_110 As Integer = 0

        Dim Konzola_50 As Integer = 0
        Dim Konzola_100 As Integer = 0
        Dim Konzola_150 As Integer = 0
        Dim Konzola_200 As Integer = 0
        Dim Konzola_300 As Integer = 0
        Dim Konzola_400 As Integer = 0
        Dim Konzola_500 As Integer = 0
        Dim Konzola_600 As Integer = 0

        Dim Konzola_V_50 As Integer = 0
        Dim Konzola_V_100 As Integer = 0
        Dim Konzola_V_150 As Integer = 0
        Dim Konzola_V_200 As Integer = 0
        Dim Konzola_V_300 As Integer = 0
        Dim Konzola_V_400 As Integer = 0
        Dim Konzola_V_500 As Integer = 0
        Dim Konzola_V_600 As Integer = 0

        Dim Bolt_М6X12 As Integer = 0
        Dim Anker_М10X10 As Integer = 0

        Dim brKonzoli As Double = 1.5

        Call Excel_Kol_smetka_Razdel(wsKol_Smetka, "КАБЕЛНИ КАНАЛИ И КАБЕЛНИ СКАРИ", "A" & Trim(index.ToString), "D" & Trim(index.ToString))
        index += 1
        Dim range As String = ""
        For i = 4 To 10000
            If Trim(wsCableTrays.Cells(i, 1).Value) = "" Then Exit For
            ProgressBar_Extrat.Value = i
            Text_Dostawka = ""
            range = ""
            broj_elementi = wsCableTrays.Range("H" + i.ToString).Value.ToString
            If broj_elementi = 0 Then Continue For
            Dim m_br As Boolean = True ' Определя дали да се пише бр. или m в колона 3
            ' True - m
            ' Face - бр
            Select Case wsCableTrays.Cells(i, 1).Value
                Case "Канал"
                    Text_Dostawka = "на PVC кабелен канал с капак " +
                                wsCableTrays.Range("B" + i.ToString).Value.ToString +
                                "x" +
                                wsCableTrays.Range("C" + i.ToString).Value.ToString +
                                "mm"
                    Kanal = Kanal + broj_elementi
                Case "Скара"
                    Text_Dostawka = "на перфорирана, кабелна скара " +
                                    wsCableTrays.Range("B" + i.ToString).Value.ToString +
                                    "x" +
                                    wsCableTrays.Range("C" + i.ToString).Value.ToString +
                                    "mm, комоплект с капак за кабелна скара"
                    skara = skara + broj_elementi
                    Select Case wsCableTrays.Range("C" + i.ToString).Value
                        Case = 35
                            Planka_35 = Planka_35 + 2 * broj_elementi / 3
                        Case = 60
                            Planka_60 = Planka_60 + 2 * broj_elementi / 3
                        Case = 85
                            Planka_85 = Planka_85 + 2 * broj_elementi / 3
                        Case = 110
                            Planka_110 = Planka_110 + 2 * broj_elementi / 3
                    End Select
                    Select Case wsCableTrays.Range("B" + i.ToString).Value
                        Case = 50
                            Konzola_50 = Konzola_50 + broj_elementi / brKonzoli
                        Case = 100
                            Konzola_100 = Konzola_100 + broj_elementi / brKonzoli
                        Case = 150
                            Konzola_150 = Konzola_150 + broj_elementi / brKonzoli
                        Case = 200
                            Konzola_200 = Konzola_200 + broj_elementi / brKonzoli
                        Case = 300
                            Konzola_300 = Konzola_300 + broj_elementi / brKonzoli
                        Case = 400
                            Konzola_400 = Konzola_400 + broj_elementi / brKonzoli
                        Case = 500
                            Konzola_500 = Konzola_500 + broj_elementi / brKonzoli
                        Case = 600
                            Konzola_600 = Konzola_600 + broj_elementi / brKonzoli
                    End Select
                Case "Скара_ъгъл"
                    Select Case wsCableTrays.Range("D" + i.ToString).Value
                        Case "Т-Вертикален"
                            Text_Dostawka = "на вертикален Т-преход за каб. скара "
                            Select Case wsCableTrays.Range("C" + i.ToString).Value
                                Case = 35
                                    Planka_35 = Planka_35 + broj_elementi * 6
                                Case = 60
                                    Planka_60 = Planka_60 + broj_elementi * 6
                                Case = 85
                                    Planka_85 = Planka_85 + broj_elementi * 6
                                Case = 110
                                    Planka_110 = Planka_110 + broj_elementi * 6
                            End Select
                        Case "90-Вертикален"
                            Text_Dostawka = "на вертикален ъгъл за каб. скара "
                            Select Case wsCableTrays.Range("C" + i.ToString).Value
                                Case = 35
                                    Planka_35 = Planka_35 + broj_elementi * 4
                                Case = 60
                                    Planka_60 = Planka_60 + broj_elementi * 4
                                Case = 85
                                    Planka_85 = Planka_85 + broj_elementi * 4
                                Case = 110
                                    Planka_110 = Planka_110 + broj_elementi * 4
                            End Select
                        Case "Кръст"
                            Text_Dostawka = "на кръстат ъгъл за каб. скара "
                            Select Case wsCableTrays.Range("C" + i.ToString).Value
                                Case = 35
                                    Planka_35 = Planka_35 + broj_elementi * 8
                                Case = 60
                                    Planka_60 = Planka_60 + broj_elementi * 8
                                Case = 85
                                    Planka_85 = Planka_85 + broj_elementi * 8
                                Case = 110
                                    Planka_110 = Planka_110 + broj_elementi * 8
                            End Select
                        Case "Т-Хоризонтален"
                            Text_Dostawka = "на хоризонтален, Т-преход за каб. скара "
                            Select Case wsCableTrays.Range("C" + i.ToString).Value
                                Case = 35
                                    Planka_35 = Planka_35 + broj_elementi * 6
                                Case = 60
                                    Planka_60 = Planka_60 + broj_elementi * 6
                                Case = 85
                                    Planka_85 = Planka_85 + broj_elementi * 6
                                Case = 110
                                    Planka_110 = Planka_110 + broj_elementi * 6
                            End Select
                        Case "90-Хоризонтален"
                            Text_Dostawka = "на хоризонтален ъгъл 90° за каб. скара "
                            Select Case wsCableTrays.Range("C" + i.ToString).Value
                                Case = 35
                                    Planka_35 = Planka_35 + broj_elementi * 2
                                Case = 60
                                    Planka_60 = Planka_60 + broj_elementi * 2
                                Case = 85
                                    Planka_85 = Planka_85 + broj_elementi * 2
                                Case = 110
                                    Planka_110 = Planka_110 + broj_elementi * 2
                            End Select
                    End Select
                    Text_Dostawka = Text_Dostawka +
                                    wsCableTrays.Range("B" + i.ToString).Value.ToString +
                                    "x" +
                                    wsCableTrays.Range("C" + i.ToString).Value.ToString +
                                    " mm"
                    skara = skara + broj_elementi
                    m_br = False
                Case "Скара стълба"
                    Text_Dostawka = "на кабелна стълба " +
                                    wsCableTrays.Range("B" + i.ToString).Value.ToString +
                                    "x" +
                                    wsCableTrays.Range("C" + i.ToString).Value.ToString +
                                    "mm"
                    skara = skara + broj_elementi
                    Select Case wsCableTrays.Range("C" + i.ToString).Value
                        Case = 35
                            Planka_35 = Planka_35 + 2 * broj_elementi / 3
                        Case = 60
                            Planka_60 = Planka_60 + 2 * broj_elementi / 3
                        Case = 85
                            Planka_85 = Planka_85 + 2 * broj_elementi / 3
                        Case = 110
                            Planka_110 = Planka_110 + 2 * broj_elementi / 3
                    End Select
                    Select Case wsCableTrays.Range("B" + i.ToString).Value
                        Case = 50
                            Konzola_V_50 = Konzola_V_50 + broj_elementi / brKonzoli
                        Case = 100
                            Konzola_V_100 = Konzola_V_100 + broj_elementi / brKonzoli
                        Case = 150
                            Konzola_V_150 = Konzola_V_150 + broj_elementi / brKonzoli
                        Case = 200
                            Konzola_V_200 = Konzola_V_200 + broj_elementi / brKonzoli
                        Case = 300
                            Konzola_V_300 = Konzola_V_300 + broj_elementi / brKonzoli
                        Case = 400
                            Konzola_V_400 = Konzola_V_400 + broj_elementi / brKonzoli
                        Case = 500
                            Konzola_V_500 = Konzola_V_500 + broj_elementi / brKonzoli
                        Case = 600
                            Konzola_V_600 = Konzola_V_600 + broj_elementi / brKonzoli
                    End Select
                Case Else
                    Text_Dostawka = "НЕИЗВЕСТЕН ЕЛЕМЕНТ --> " & wsKontakti.Cells(i, 2).Value & " - " & wsKontakti.Cells(i, 4).Value
            End Select
            If Trim(Text_Dostawka) <> "" Then
                With wsKol_Smetka
                    .Cells(index, 2).Value = "Доставка и монтаж " + Text_Dostawka
                    .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                    .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                    .Cells(index, 3).Value = If(m_br, "m", "бр.")
                    .Cells(index, 4).Value = broj_elementi
                End With
                index += 1
            End If
        Next
        If Kanal > 0 Then
            With wsKol_Smetka
                .Cells(index, 2).Value = "Доставка на комплект крепежни елементи за PVC кабелен канал в т.ч.:"
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "бр."
                .Cells(index, 4).Value = 1
            End With
            index += 1
            With wsKol_Smetka
                .Cells(index, 2).Value = " - Дюбел PVC 4х40 mm " +
                                         (Kanal * 3).ToString +
                                         " броя"
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
            End With
            index += 1
            With wsKol_Smetka
                .Cells(index, 2).Value = " - Винт 40х4 mm " +
                                         (Kanal * 3).ToString +
                                         " броя"
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
            End With
            index += 1
        End If
        If skara > 0 Then
            With wsKol_Smetka
                .Cells(index, 2).Value = "Доставка на комплект крепежни елементи за кабелна скара в т.ч.:"
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "бр."
                .Cells(index, 4).Value = 1
            End With
            index += 1
            '
            ' Изчислява болтове
            '

            ' За конзоли хоризонтални
            Anker_М10X10 = Anker_М10X10 + 2 * (Konzola_50 + Konzola_100 + Konzola_150 +
                                               Konzola_200 + Konzola_300 + Konzola_400 +
                                               Konzola_500 + Konzola_600)
            ' За конзоли вертикални
            Anker_М10X10 = Anker_М10X10 + 2 * (Konzola_V_50 + Konzola_V_100 + Konzola_V_150 +
                                               Konzola_V_200 + Konzola_V_300 + Konzola_V_400 +
                                               Konzola_V_500 + Konzola_V_600)
            'За кабелни стълби
            '
            ' Изчислява болтове
            '

            'Захващане скара към конзола
            Bolt_М6X12 = Bolt_М6X12 + 2 * (Konzola_50 + Konzola_100 + Konzola_150 + Konzola_200 + Konzola_300)
            Bolt_М6X12 = Bolt_М6X12 + 3 * (Konzola_400 + Konzola_500 + Konzola_600)
            'Захващане скара към планка за вертикален монтаж 
            Bolt_М6X12 = Bolt_М6X12 + 2 * (Konzola_V_50 + Konzola_V_100 + Konzola_V_150 + Konzola_V_200 + Konzola_V_300)
            Bolt_М6X12 = Bolt_М6X12 + 3 * (Konzola_V_400 + Konzola_V_500 + Konzola_V_600)
            'Захващане скара с хоризонтална планка 
            Bolt_М6X12 = Bolt_М6X12 + 4 * (Planka_35 + Planka_60 + Planka_85 + Planka_110)

            With wsKol_Smetka
                If Bolt_М6X12 > 0 Then
                    .Cells(index, 2).Value = " - Комплект болт, гайка и шайба осигурителна М6х12 mm - " +
                                         Bolt_М6X12.ToString +
                                         " бр."
                    .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                    .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                    index += 1
                End If
                If Anker_М10X10 > 0 Then
                    .Cells(index, 2).Value = " - Анкерен болт М10X10 mm - " +
                                             Anker_М10X10.ToString +
                                             " бр."
                    .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                    .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                    index += 1
                End If
                If Planka_35 > 0 Then
                    .Cells(index, 2).Value = " - Планка съединителна H=35 mm - " +
                                         Planka_35.ToString +
                                         " бр."
                    .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                    .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                    index += 1
                End If
                If Planka_60 > 0 Then
                    .Cells(index, 2).Value = " - Планка съединителна H=60 mm - " +
                                         Planka_60.ToString +
                                         " бр."
                    .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                    .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                    index += 1
                End If
                If Planka_85 > 0 Then
                    .Cells(index, 2).Value = " - Планка съединителна H=85 mm - " +
                                         Planka_85.ToString +
                                         " бр."
                    .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                    .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                    index += 1
                End If
                If Planka_110 > 0 Then
                    .Cells(index, 2).Value = " - Планка съединителна H=110 mm - " +
                                         Planka_110.ToString +
                                         " бр."
                    .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                    .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                    index += 1
                End If
                If Konzola_50 > 0 Then
                    .Cells(index, 2).Value = " - Конзола за стена за кабелна скара 50 mm - " +
                                         Konzola_50.ToString +
                                         " бр."
                    .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                    .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                    index += 1
                End If
                If Konzola_100 > 0 Then
                    .Cells(index, 2).Value = " - Конзола за стена за кабелна скара 100 mm - " +
                                         Konzola_100.ToString +
                                         " бр."
                    .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                    .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                    index += 1
                End If
                If Konzola_150 > 0 Then
                    .Cells(index, 2).Value = " - Конзола за стена за кабелна скара 150 mm - " +
                                         Konzola_150.ToString +
                                         " бр."
                    .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                    .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                    index += 1
                End If
                If Konzola_200 > 0 Then
                    .Cells(index, 2).Value = " - Конзола за стена за кабелна скара 200 mm - " +
                                         Konzola_200.ToString +
                                         " бр."
                    .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                    .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                    index += 1
                End If
                If Konzola_300 > 0 Then
                    .Cells(index, 2).Value = " - Конзола за стена за кабелна скара 300 mm - " +
                                         Konzola_300.ToString +
                                         " бр."
                    .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                    .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                    index += 1
                End If
                If Konzola_400 > 0 Then
                    .Cells(index, 2).Value = " - Конзола за стена за кабелна скара 400 mm - " +
                                         Konzola_400.ToString +
                                         " бр."
                    .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                    .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                    index += 1
                End If
                If Konzola_500 > 0 Then
                    .Cells(index, 2).Value = " - Конзола за стена за кабелна скара 500 mm - " +
                                         Konzola_500.ToString +
                                         " бр."
                    .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                    .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                    index += 1
                End If
                If Konzola_600 > 0 Then
                    .Cells(index, 2).Value = " - Конзола за стена за кабелна скара 600 mm - " +
                                         Konzola_600.ToString +
                                         " бр."
                    .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                    .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                    index += 1
                End If
                If Konzola_V_50 > 0 Then
                    .Cells(index, 2).Value = " - Планка за вертикален монтаж на кабелна скара 50 mm - " +
                                         Konzola_V_50.ToString +
                                         " бр."
                    .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                    .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                    index += 1
                End If
                If Konzola_V_100 > 0 Then
                    .Cells(index, 2).Value = " - Планка за вертикален монтаж на кабелна скара 100 mm - " +
                                         Konzola_V_100.ToString +
                                         " бр."
                    .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                    .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                    index += 1
                End If
                If Konzola_V_150 > 0 Then
                    .Cells(index, 2).Value = " - Планка за вертикален монтаж на кабелна скара 150 mm - " +
                                         Konzola_V_150.ToString +
                                         " бр."
                    .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                    .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                    index += 1
                End If
                If Konzola_V_200 > 0 Then
                    .Cells(index, 2).Value = " - Планка за вертикален монтаж на кабелна скара 200 mm - " +
                                         Konzola_V_200.ToString +
                                         " бр."
                    .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                    .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                    index += 1
                End If
                If Konzola_V_300 > 0 Then
                    .Cells(index, 2).Value = " - Планка за вертикален монтаж на кабелна скара 300 mm - " +
                                         Konzola_V_300.ToString +
                                         " бр."
                    .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                    .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                    index += 1
                End If
                If Konzola_V_400 > 0 Then
                    .Cells(index, 2).Value = " - Планка за вертикален монтаж на кабелна скара 400 mm - " +
                                         Konzola_V_400.ToString +
                                         " бр."
                    .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                    .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                    index += 1
                End If
                If Konzola_V_500 > 0 Then
                    .Cells(index, 2).Value = " - Планка за вертикален монтаж на кабелна скара 500 mm - " +
                                         Konzola_V_500.ToString +
                                         " бр."
                    .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                    .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                    index += 1
                End If
                If Konzola_V_600 > 0 Then
                    .Cells(index, 2).Value = " - Планка за вертикален монтаж на кабелна скара 600 mm - " +
                                         Konzola_V_500.ToString +
                                         " бр."
                    .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                    .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                    index += 1
                End If
            End With
        End If

        ProgressBar_Extrat.Value = ProgressBar_Extrat.Minimum
    End Sub
    Private Sub Button_Вземи_ВИДЕО_КАМЕРИ_Click(sender As Object, e As EventArgs) Handles Button_Вземи_ВИДЕО_КАМЕРИ.Click
        If IsNothing(excel_Workbook) Then
            MsgBox("Първо да беше отворил файла?")
            Exit Sub
        End If
        Me.Visible = vbFalse
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim edt As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database
        Dim SelectedSet = cu.GetObjects("INSERT", "Изберете блок")
        Dim arrBlock(500) As strKontakt
        Dim Качване(100) As strКачване

        If SelectedSet Is Nothing Then
            MsgBox("Няма маркиран нито един блок.")
            Me.Visible = vbTrue
            Exit Sub
        End If
        Dim blkRecId As ObjectId = ObjectId.Null
        Dim index As Integer = 0
        Dim index_Качване As Integer = 0
        Dim index_Row As Integer = 0
        Dim broj_konturi As Integer = 0
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                For Each sObj As SelectedObject In SelectedSet
                    blkRecId = sObj.ObjectId

                    Dim acBlkRef As BlockReference = DirectCast(acTrans.GetObject(blkRecId, OpenMode.ForRead), BlockReference)
                    Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                    Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection
                    Dim Visibility As String = ""

                    Dim strKOTA_1 As String = ""
                    Dim strKOTA_2 As String = ""
                    Dim strТРЪБА_1 As String = ""
                    Dim strТРЪБА_2 As String = ""
                    Dim strKabel_d_0 As String = ""
                    Dim strKabel_d_1 As String = ""
                    Dim strKabel_d_2 As String = ""
                    Dim strKabel_d_3 As String = ""
                    Dim strKabel_d_4 As String = ""
                    Dim strKabel_d_5 As String = ""
                    Dim strKabel_d_6 As String = ""
                    Dim strKabel_d_7 As String = ""
                    Dim strKabel_d_8 As String = ""
                    Dim strKabel_d_9 As String = ""
                    Dim strKabel_d_10 As String = ""
                    Dim strKabel_g_0 As String = ""
                    Dim strKabel_g_1 As String = ""
                    Dim strKabel_g_2 As String = ""
                    Dim strKabel_g_3 As String = ""
                    Dim strKabel_g_4 As String = ""
                    Dim strKabel_g_5 As String = ""
                    Dim strKabel_g_6 As String = ""
                    Dim strKabel_g_7 As String = ""
                    Dim strKabel_g_8 As String = ""
                    Dim strKabel_g_9 As String = ""
                    Dim strKabel_g_10 As String = ""

                    For Each prop As DynamicBlockReferenceProperty In props
                        If prop.PropertyName = "Visibility" Then Visibility = prop.Value
                    Next

                    Dim blName As String = (CType(acBlkRef.DynamicBlockTableRecord.GetObject(OpenMode.ForRead), BlockTableRecord)).Name

                    Dim iVisib As Integer = -1
                    Select Case blName
                        Case "Камери"
                            iVisib = Array.FindIndex(arrBlock, Function(f) f.blName = blName And
                                                                           f.blVisibility = Visibility)
                        Case "Кабел"
                            Continue For
                        Case "Качване"
                            For Each objID As ObjectId In attCol
                                Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                                Dim acAttRef As AttributeReference = dbObj

                                If acAttRef.Tag = "KOTA_1" Then strKOTA_1 = acAttRef.TextString
                                If acAttRef.Tag = "KOTA_2" Then strKOTA_2 = acAttRef.TextString
                                If acAttRef.Tag = "ТРЪБА_1" Then strТРЪБА_1 = acAttRef.TextString
                                If acAttRef.Tag = "ТРЪБА_2" Then strТРЪБА_2 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_0" Then strKabel_d_0 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_1" Then strKabel_d_1 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_2" Then strKabel_d_2 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_3" Then strKabel_d_3 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_4" Then strKabel_d_4 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_5" Then strKabel_d_5 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_6" Then strKabel_d_6 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_7" Then strKabel_d_7 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_8" Then strKabel_d_8 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_9" Then strKabel_d_9 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_10" Then strKabel_d_10 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_0" Then strKabel_g_0 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_1" Then strKabel_g_1 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_2" Then strKabel_g_2 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_3" Then strKabel_g_3 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_4" Then strKabel_g_4 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_5" Then strKabel_g_5 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_6" Then strKabel_g_6 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_7" Then strKabel_g_7 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_8" Then strKabel_g_8 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_9" Then strKabel_g_9 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_10" Then strKabel_g_10 = acAttRef.TextString
                            Next
                        Case Else
                            Continue For
                    End Select
                    If iVisib = -1 Then
                        arrBlock(index).count = 1
                        arrBlock(index).blName = blName
                        arrBlock(index).blVisibility = Visibility
                        If blName = "Качване" Then
                            Качване(index_Качване).KOTA_1 = strKOTA_1
                            Качване(index_Качване).KOTA_2 = strKOTA_2
                            Качване(index_Качване).ТРЪБА_1 = strТРЪБА_1
                            Качване(index_Качване).ТРЪБА_2 = strТРЪБА_2
                            Качване(index_Качване).Kabel_d_0 = strKabel_d_0
                            Качване(index_Качване).Kabel_d_1 = strKabel_d_1
                            Качване(index_Качване).Kabel_d_2 = strKabel_d_2
                            Качване(index_Качване).Kabel_d_3 = strKabel_d_3
                            Качване(index_Качване).Kabel_d_6 = strKabel_d_6
                            Качване(index_Качване).Kabel_d_7 = strKabel_d_7
                            Качване(index_Качване).Kabel_d_4 = strKabel_d_4
                            Качване(index_Качване).Kabel_d_5 = strKabel_d_5
                            Качване(index_Качване).Kabel_d_8 = strKabel_d_8
                            Качване(index_Качване).Kabel_d_9 = strKabel_d_9
                            Качване(index_Качване).Kabel_d_10 = strKabel_d_10
                            Качване(index_Качване).Kabel_g_0 = strKabel_g_0
                            Качване(index_Качване).Kabel_g_1 = strKabel_g_1
                            Качване(index_Качване).Kabel_g_2 = strKabel_g_2
                            Качване(index_Качване).Kabel_g_3 = strKabel_g_3
                            Качване(index_Качване).Kabel_g_4 = strKabel_g_4
                            Качване(index_Качване).Kabel_g_5 = strKabel_g_5
                            Качване(index_Качване).Kabel_g_6 = strKabel_g_6
                            Качване(index_Качване).Kabel_g_7 = strKabel_g_7
                            Качване(index_Качване).Kabel_g_8 = strKabel_g_8
                            Качване(index_Качване).Kabel_g_9 = strKabel_g_9
                            Качване(index_Качване).Kabel_g_10 = strKabel_g_10
                            index_Качване += 1
                        End If
                        index += 1
                    Else
                        arrBlock(iVisib).count = arrBlock(iVisib).count + 1
                    End If
                Next

                index_Качване = 0

                Call clearDomofon(wsVIDEO)

                index_Row = 2
                Dim dylvina1 As Double = 0
                Dim dylvina2 As Double = 0
                Dim br_linii As Double = 0
                Dim Kabel_FTP As Double = 0

                For Each iarrBlock In arrBlock
                    If iarrBlock.count = 0 Then Exit For
                    With wsVIDEO
                        .Cells(index_Row, 1) = iarrBlock.count
                        .Cells(index_Row, 2) = iarrBlock.blName
                        .Cells(index_Row, 3) = iarrBlock.blVisibility
                        If iarrBlock.blName = "Качване" Then
                            .Cells(index_Row, 11) = Качване(index_Качване).KOTA_1
                            .Cells(index_Row, 12) = Качване(index_Качване).ТРЪБА_1
                            .Cells(index_Row, 13) = Качване(index_Качване).Kabel_d_0
                            .Cells(index_Row, 14) = Качване(index_Качване).Kabel_d_1
                            .Cells(index_Row, 15) = Качване(index_Качване).Kabel_d_2
                            .Cells(index_Row, 16) = Качване(index_Качване).Kabel_d_3
                            .Cells(index_Row, 17) = Качване(index_Качване).Kabel_d_6
                            .Cells(index_Row, 18) = Качване(index_Качване).Kabel_d_5
                            .Cells(index_Row, 19) = Качване(index_Качване).Kabel_d_8
                            .Cells(index_Row, 20) = Качване(index_Качване).Kabel_d_7
                            .Cells(index_Row, 21) = Качване(index_Качване).Kabel_d_4
                            .Cells(index_Row, 22) = Качване(index_Качване).Kabel_d_9
                            .Cells(index_Row, 23) = Качване(index_Качване).Kabel_d_10
                            .Cells(index_Row, 24) = Качване(index_Качване).KOTA_2
                            .Cells(index_Row, 25) = Качване(index_Качване).ТРЪБА_2
                            .Cells(index_Row, 26) = Качване(index_Качване).Kabel_g_0
                            .Cells(index_Row, 27) = Качване(index_Качване).Kabel_g_1
                            .Cells(index_Row, 28) = Качване(index_Качване).Kabel_g_2
                            .Cells(index_Row, 29) = Качване(index_Качване).Kabel_g_3
                            .Cells(index_Row, 30) = Качване(index_Качване).Kabel_g_4
                            .Cells(index_Row, 31) = Качване(index_Качване).Kabel_g_5
                            .Cells(index_Row, 32) = Качване(index_Качване).Kabel_g_6
                            .Cells(index_Row, 33) = Качване(index_Качване).Kabel_g_7
                            .Cells(index_Row, 34) = Качване(index_Качване).Kabel_g_8
                            .Cells(index_Row, 35) = Качване(index_Качване).Kabel_g_9
                            .Cells(index_Row, 36) = Качване(index_Качване).Kabel_g_10
                            index_Качване += 1

                            dylvina1 = Val(.Cells(index_Row, 11).value)
                            dylvina2 = Val(.Cells(index_Row, 24).value)
                            For i = 11 To 36
                                br_linii = InStr(.Cells(index_Row, i).value, "л.")
                                If br_linii = 0 Then Continue For
                                br_linii = Mid(.Cells(index_Row, i).value, 1, br_linii - 1)
                                If InStr(.Cells(index_Row, i).value, "FTP ") > 1 Then
                                    .Cells(index_Row, 8).value =
                                        IIf(i < 24, br_linii * dylvina1, br_linii * dylvina2) +
                                        .Cells(index_Row, 8).value
                                    Kabel_FTP = Kabel_FTP + .Cells(index_Row, 8).value
                                End If
                            Next
                        End If
                        If iarrBlock.blName = "Камери" Then
                            wsLines.Range("AZ2").Value = wsLines.Range("AZ2").Value +
                                                                iarrBlock.count
                        End If
                    End With
                    index_Row += 1
                Next

                wsLines.Cells(2, 46).Value = Kabel_FTP / 2
                acTrans.Commit()
                With wsVIDEO
                    .Cells.Sort(Key1:= .Range("B2"),
                                Order1:=excel.XlSortOrder.xlAscending,
                                Header:=excel.XlYesNoGuess.xlYes,
                                OrderCustom:=1, MatchCase:=False,
                                Orientation:=excel.Constants.xlTopToBottom,
                                DataOption1:=excel.XlSortDataOption.xlSortTextAsNumbers,
                                Key2:= .Range("D2"),
                                Order2:=excel.XlSortOrder.xlAscending,
                                DataOption2:=excel.XlSortDataOption.xlSortTextAsNumbers
                                )
                End With
            Catch ex As Exception
                MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
                acTrans.Abort()
            End Try
        End Using
        Me.Visible = vbTrue
    End Sub
    Private Sub Button_Изчисти_ВИДЕО_Click(sender As Object, e As EventArgs) Handles Button_Изчисти_ВИДЕО.Click
        clearDomofon(wsVIDEO)
    End Sub
    Private Sub Button_Изчисти_ДОМОФ_Click(sender As Object, e As EventArgs) Handles Button_Изчисти_ДОМОФ.Click
        clearDomofon(wsDOMOF)
    End Sub
    Private Sub Button_Генерирай_ВИДЕО_Click(sender As Object, e As EventArgs) Handles Button_Генерирай_ВИДЕО.Click
        Dim index As Integer = 0
        Dim index_internet As Integer = 0
        Dim index_Kabel As Integer = 0
        Dim i As Integer
        ProgressBar_Extrat.Maximum = ProgressBar_Maximum
        For i = 6 To 1000
            If wsKol_Smetka.Cells(i, 2).Value = "Раздел" Or wsKol_Smetka.Cells(i, 2).Value = "" Then
                index = i
                Exit For
            End If
            ProgressBar_Extrat.Value = i
        Next
        For i = 2 To 1000
            If wsVIDEO.Cells(i, 2).Value = "" Then
                Exit For
            End If
        Next
        index_internet = i
        Dim Text_Dostawka As String = ""
        Dim broj_elementi As Integer = 0
        Call Excel_Kol_smetka_Razdel(wsKol_Smetka, "СЛАБОТОКОВИ ИНСТАЛАЦИИ", "A" & Trim(index.ToString), "D" & Trim(index.ToString))
        index += 1
        Call Excel_Kol_smetka_Razdel(wsKol_Smetka, "ИНСТАЛАЦИЯ ВИДЕОНАБЮДЕНИЕ", "B" & Trim(index.ToString), "D" & Trim(index.ToString))
        Dim Kabel(1000) As strKabel
        index += 1

        For i = 2 To 10000
            If Trim(wsVIDEO.Cells(i, 2).Value) = "" Then Exit For
            ProgressBar_Extrat.Value = i
            Text_Dostawka = ""
            broj_elementi = wsVIDEO.Cells(i, 1).Value

            Select Case wsVIDEO.Cells(i, 2).Value
                Case "Камери"
                    Select Case wsVIDEO.Cells(i, 3).Value
                        Case "Зона"
                        Case "Камерa 360°"
                            Text_Dostawka = "куполна IP видеокамера камера 5 мегапиксе-ла:" & vbCrLf &
                            "- Резолюция: 5 мегапиксела;" & vbCrLf &
                            "- Захранване с PoE инжектор или суич: PoE 802.3af;" & vbCrLf &
                            "- Материал на корпуса: ПВЦ купол и ПВЦ основа;" & vbCrLf &
                            "- Комплект с основа."
                        Case "Насочен камера-IP66"
                            Text_Dostawka = "корпусна IP видеокамера камера 5 мегапиксе-ла:" & vbCrLf &
                            "- Резолюция: 5 мегапиксела;" & vbCrLf &
                            "- Захранване с PoE инжектор или суич: PoE 802.3af;" & vbCrLf &
                            "- Степен на вандалоустойчивост: IK10;" & vbCrLf &
                            "- Показател за водоустойчивост: IP67;" & vbCrLf &
                            "- Работна температура: -30~+60°C;" & vbCrLf &
                            "- Материал на корпуса: Метален корпус и метална стойка;" & vbCrLf &
                            "- Комплект със стойка."
                        Case "Насочен камера-20"
                            Text_Dostawka = "корпусна IP видеокамера камера 5 мегапиксе-ла:" & vbCrLf &
                            "- Резолюция: 5 мегапиксела;" & vbCrLf &
                            "- Захранване с PoE инжектор или суич: PoE 802.3af;" & vbCrLf &
                            "- Материал на корпуса: корпусна със стойка;" & vbCrLf &
                            "- Комплект със стойка."
                        Case Else
                            Text_Dostawka = "НЕИЗВЕСТЕН ЕЛЕМЕНТ --> " &
                                            wsPIC.Cells(i, 2).Value & " - " &
                                            wsPIC.Cells(i, 3).Value
                    End Select
                    index_Kabel += 1
                Case "Качване"
                    Continue For
            End Select
            If Trim(Text_Dostawka) <> "" Then
                With wsKol_Smetka
                    .Cells(index, 2).Value = "Доставка и монтаж на " + Text_Dostawka
                    .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                    .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                    .Cells(index, 3).Value = "бр."
                    .Cells(index, 4).Value = broj_elementi
                End With
                index += 1
            End If
        Next

        index = Kol_Smetka_Kabeli(index, vbFalse, "ВИДЕО")

        Text_Dostawka = "Комплексно изпитване на системата"
        With wsKol_Smetka
            .Cells(index, 2).Value = Text_Dostawka
            .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
            .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
            .Cells(index, 3).Value = "бр."
            .Cells(index, 4).Value = 1
        End With
        index += 1

        Text_Dostawka = "Обучение на персонал за работа със системата"
        With wsKol_Smetka
            .Cells(index, 2).Value = Text_Dostawka
            .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
            .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
            .Cells(index, 3).Value = "бр."
            .Cells(index, 4).Value = 1
        End With
        index += 1
        ProgressBar_Extrat.Value = ProgressBar_Extrat.Minimum
    End Sub
    Private Sub Button_Вземи_СОТ_ДАТЧИЦИ_Click(sender As Object, e As EventArgs) Handles Button_Вземи_СОТ_ДАТЧИЦИ.Click
        If IsNothing(excel_Workbook) Then
            MsgBox("Първо да беше отворил файла?")
            Exit Sub
        End If
        Me.Visible = vbFalse
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim edt As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database
        Dim SelectedSet = cu.GetObjects("INSERT", "Изберете блок")
        Dim arrBlock(500) As strKontakt
        Dim Качване(100) As strКачване

        If SelectedSet Is Nothing Then
            MsgBox("Няма маркиран нито един блок.")
            Me.Visible = vbTrue
            Exit Sub
        End If
        Dim blkRecId As ObjectId = ObjectId.Null
        Dim index As Integer = 0
        Dim index_Качване As Integer = 0
        Dim index_Row As Integer = 0
        Dim broj_konturi As Integer = 0
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                For Each sObj As SelectedObject In SelectedSet
                    blkRecId = sObj.ObjectId

                    Dim acBlkRef As BlockReference = DirectCast(acTrans.GetObject(blkRecId, OpenMode.ForRead), BlockReference)
                    Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                    Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection
                    Dim Visibility As String = ""

                    Dim strKOTA_1 As String = ""
                    Dim strKOTA_2 As String = ""
                    Dim strТРЪБА_1 As String = ""
                    Dim strТРЪБА_2 As String = ""
                    Dim strKabel_d_0 As String = ""
                    Dim strKabel_d_1 As String = ""
                    Dim strKabel_d_2 As String = ""
                    Dim strKabel_d_3 As String = ""
                    Dim strKabel_d_4 As String = ""
                    Dim strKabel_d_5 As String = ""
                    Dim strKabel_d_6 As String = ""
                    Dim strKabel_d_7 As String = ""
                    Dim strKabel_d_8 As String = ""
                    Dim strKabel_d_9 As String = ""
                    Dim strKabel_d_10 As String = ""
                    Dim strKabel_g_0 As String = ""
                    Dim strKabel_g_1 As String = ""
                    Dim strKabel_g_2 As String = ""
                    Dim strKabel_g_3 As String = ""
                    Dim strKabel_g_4 As String = ""
                    Dim strKabel_g_5 As String = ""
                    Dim strKabel_g_6 As String = ""
                    Dim strKabel_g_7 As String = ""
                    Dim strKabel_g_8 As String = ""
                    Dim strKabel_g_9 As String = ""
                    Dim strKabel_g_10 As String = ""

                    For Each prop As DynamicBlockReferenceProperty In props
                        If prop.PropertyName = "Visibility" Then Visibility = prop.Value
                    Next

                    Dim blName As String = (CType(acBlkRef.DynamicBlockTableRecord.GetObject(OpenMode.ForRead), BlockTableRecord)).Name

                    Dim iVisib As Integer = -1
                    Select Case blName
                        Case "СОТ"
                            iVisib = Array.FindIndex(arrBlock, Function(f) f.blName = blName And
                                                                           f.blVisibility = Visibility)
                        Case "Кабел"
                            Continue For
                        Case "Качване"
                            For Each objID As ObjectId In attCol
                                Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                                Dim acAttRef As AttributeReference = dbObj

                                If acAttRef.Tag = "KOTA_1" Then strKOTA_1 = acAttRef.TextString
                                If acAttRef.Tag = "KOTA_2" Then strKOTA_2 = acAttRef.TextString
                                If acAttRef.Tag = "ТРЪБА_1" Then strТРЪБА_1 = acAttRef.TextString
                                If acAttRef.Tag = "ТРЪБА_2" Then strТРЪБА_2 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_0" Then strKabel_d_0 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_1" Then strKabel_d_1 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_2" Then strKabel_d_2 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_3" Then strKabel_d_3 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_4" Then strKabel_d_4 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_5" Then strKabel_d_5 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_6" Then strKabel_d_6 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_7" Then strKabel_d_7 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_8" Then strKabel_d_8 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_9" Then strKabel_d_9 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_d_10" Then strKabel_d_10 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_0" Then strKabel_g_0 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_1" Then strKabel_g_1 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_2" Then strKabel_g_2 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_3" Then strKabel_g_3 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_4" Then strKabel_g_4 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_5" Then strKabel_g_5 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_6" Then strKabel_g_6 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_7" Then strKabel_g_7 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_8" Then strKabel_g_8 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_9" Then strKabel_g_9 = acAttRef.TextString
                                If acAttRef.Tag = "Kabel_g_10" Then strKabel_g_10 = acAttRef.TextString
                            Next
                        Case Else
                            Continue For
                    End Select
                    If iVisib = -1 Then
                        arrBlock(index).count = 1
                        arrBlock(index).blName = blName
                        arrBlock(index).blVisibility = Visibility
                        If blName = "Качване" Then
                            Качване(index_Качване).KOTA_1 = strKOTA_1
                            Качване(index_Качване).KOTA_2 = strKOTA_2
                            Качване(index_Качване).ТРЪБА_1 = strТРЪБА_1
                            Качване(index_Качване).ТРЪБА_2 = strТРЪБА_2
                            Качване(index_Качване).Kabel_d_0 = strKabel_d_0
                            Качване(index_Качване).Kabel_d_1 = strKabel_d_1
                            Качване(index_Качване).Kabel_d_2 = strKabel_d_2
                            Качване(index_Качване).Kabel_d_3 = strKabel_d_3
                            Качване(index_Качване).Kabel_d_6 = strKabel_d_6
                            Качване(index_Качване).Kabel_d_7 = strKabel_d_7
                            Качване(index_Качване).Kabel_d_4 = strKabel_d_4
                            Качване(index_Качване).Kabel_d_5 = strKabel_d_5
                            Качване(index_Качване).Kabel_d_8 = strKabel_d_8
                            Качване(index_Качване).Kabel_d_9 = strKabel_d_9
                            Качване(index_Качване).Kabel_d_10 = strKabel_d_10
                            Качване(index_Качване).Kabel_g_0 = strKabel_g_0
                            Качване(index_Качване).Kabel_g_1 = strKabel_g_1
                            Качване(index_Качване).Kabel_g_2 = strKabel_g_2
                            Качване(index_Качване).Kabel_g_3 = strKabel_g_3
                            Качване(index_Качване).Kabel_g_4 = strKabel_g_4
                            Качване(index_Качване).Kabel_g_5 = strKabel_g_5
                            Качване(index_Качване).Kabel_g_6 = strKabel_g_6
                            Качване(index_Качване).Kabel_g_7 = strKabel_g_7
                            Качване(index_Качване).Kabel_g_8 = strKabel_g_8
                            Качване(index_Качване).Kabel_g_9 = strKabel_g_9
                            Качване(index_Качване).Kabel_g_10 = strKabel_g_10
                            index_Качване += 1
                        End If
                        index += 1
                    Else
                        arrBlock(iVisib).count = arrBlock(iVisib).count + 1
                    End If
                Next

                index_Качване = 0

                Call clearDomofon(wsSOT)

                index_Row = 2
                Dim dylvina1 As Double = 0
                Dim dylvina2 As Double = 0
                Dim br_linii As Double = 0
                Dim Kabel_FTP As Double = 0

                For Each iarrBlock In arrBlock
                    If iarrBlock.count = 0 Then Exit For
                    With wsSOT
                        .Cells(index_Row, 1) = iarrBlock.count
                        .Cells(index_Row, 2) = iarrBlock.blName
                        .Cells(index_Row, 3) = iarrBlock.blVisibility
                        If iarrBlock.blName = "СОТ" And iarrBlock.blVisibility = "Зона" Then
                            Continue For
                        End If
                        If iarrBlock.blName = "Качване" Then
                            .Cells(index_Row, 11) = Качване(index_Качване).KOTA_1
                            .Cells(index_Row, 12) = Качване(index_Качване).ТРЪБА_1
                            .Cells(index_Row, 13) = Качване(index_Качване).Kabel_d_0
                            .Cells(index_Row, 14) = Качване(index_Качване).Kabel_d_1
                            .Cells(index_Row, 15) = Качване(index_Качване).Kabel_d_2
                            .Cells(index_Row, 16) = Качване(index_Качване).Kabel_d_3
                            .Cells(index_Row, 17) = Качване(index_Качване).Kabel_d_6
                            .Cells(index_Row, 18) = Качване(index_Качване).Kabel_d_5
                            .Cells(index_Row, 19) = Качване(index_Качване).Kabel_d_8
                            .Cells(index_Row, 20) = Качване(index_Качване).Kabel_d_7
                            .Cells(index_Row, 21) = Качване(index_Качване).Kabel_d_4
                            .Cells(index_Row, 22) = Качване(index_Качване).Kabel_d_9
                            .Cells(index_Row, 23) = Качване(index_Качване).Kabel_d_10
                            .Cells(index_Row, 24) = Качване(index_Качване).KOTA_2
                            .Cells(index_Row, 25) = Качване(index_Качване).ТРЪБА_2
                            .Cells(index_Row, 26) = Качване(index_Качване).Kabel_g_0
                            .Cells(index_Row, 27) = Качване(index_Качване).Kabel_g_1
                            .Cells(index_Row, 28) = Качване(index_Качване).Kabel_g_2
                            .Cells(index_Row, 29) = Качване(index_Качване).Kabel_g_3
                            .Cells(index_Row, 30) = Качване(index_Качване).Kabel_g_4
                            .Cells(index_Row, 31) = Качване(index_Качване).Kabel_g_5
                            .Cells(index_Row, 32) = Качване(index_Качване).Kabel_g_6
                            .Cells(index_Row, 33) = Качване(index_Качване).Kabel_g_7
                            .Cells(index_Row, 34) = Качване(index_Качване).Kabel_g_8
                            .Cells(index_Row, 35) = Качване(index_Качване).Kabel_g_9
                            .Cells(index_Row, 36) = Качване(index_Качване).Kabel_g_10
                            index_Качване += 1

                            dylvina1 = Val(.Cells(index_Row, 11).value)
                            dylvina2 = Val(.Cells(index_Row, 24).value)
                            For i = 11 To 36
                                br_linii = InStr(.Cells(index_Row, i).value, "л.")
                                If br_linii = 0 Then Continue For
                                br_linii = Mid(.Cells(index_Row, i).value, 1, br_linii - 1)
                                If InStr(.Cells(index_Row, i).value, "CAB/") > 1 Then
                                    .Cells(index_Row, 8).value =
                                        IIf(i < 24, br_linii * dylvina1, br_linii * dylvina2) +
                                        .Cells(index_Row, 8).value
                                    Kabel_FTP = Kabel_FTP + .Cells(index_Row, 8).value
                                End If
                            Next
                        End If
                        If iarrBlock.blName = "СОТ" And
                           iarrBlock.blVisibility <> "Клавиатура" Then
                            wsLines.Range("BC2").Value = wsLines.Range("BC2").Value +
                                                         iarrBlock.count
                        End If
                        If iarrBlock.blName = "СОТ" And iarrBlock.blVisibility = "Клавиатура" Then
                            wsLines.Range("BD2").Value = wsLines.Range("BD2").Value +
                                                         iarrBlock.count
                        End If
                    End With
                    index_Row += 1
                Next

                wsLines.Range("BE2").Value = Kabel_FTP / 2
                acTrans.Commit()
                With wsSOT
                    .Cells.Sort(Key1:= .Range("B2"),
                                Order1:=excel.XlSortOrder.xlAscending,
                                Header:=excel.XlYesNoGuess.xlYes,
                                OrderCustom:=1, MatchCase:=False,
                                Orientation:=excel.Constants.xlTopToBottom,
                                DataOption1:=excel.XlSortDataOption.xlSortTextAsNumbers,
                                Key2:= .Range("D2"),
                                Order2:=excel.XlSortOrder.xlAscending,
                                DataOption2:=excel.XlSortDataOption.xlSortTextAsNumbers
                                )
                End With
            Catch ex As Exception
                MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
                acTrans.Abort()
            End Try
        End Using
        Me.Visible = vbTrue
    End Sub
    Private Sub Button_Генерирай_СОТ_Click(sender As Object, e As EventArgs) Handles Button_Генерирай_СОТ.Click
        Dim index As Integer = 0
        Dim index_internet As Integer = 0
        Dim index_Kabel As Integer = 0
        Dim i As Integer
        ProgressBar_Extrat.Maximum = ProgressBar_Maximum
        For i = 6 To 1000
            If wsKol_Smetka.Cells(i, 2).Value = "Раздел" Or wsKol_Smetka.Cells(i, 2).Value = "" Then
                index = i
                Exit For
            End If
            ProgressBar_Extrat.Value = i
        Next
        For i = 2 To 1000
            If wsSOT.Cells(i, 2).Value = "" Then
                Exit For
            End If
        Next
        index_internet = i
        ProgressBar_Extrat.Maximum = i
        Dim Text_Dostawka As String = ""
        Dim broj_elementi As Integer = 0
        Call Excel_Kol_smetka_Razdel(wsKol_Smetka, "СЛАБОТОКОВИ ИНСТАЛАЦИИ", "A" & Trim(index.ToString), "D" & Trim(index.ToString))
        index += 1
        Call Excel_Kol_smetka_Razdel(wsKol_Smetka, "СОТ ИНСТАЛАЦИЯ", "B" & Trim(index.ToString), "D" & Trim(index.ToString))
        Dim Kabel(1000) As strKabel
        index += 1
        With wsKol_Smetka
            .Cells(index, 2).Value = "Доставка и монтаж на контролен панел за алармени системи - СОТ централа:" & vbCrLf &
                 "- 8 зонови входа (16 с дублиране на зони) разширяеми до 192;" & vbCrLf &
                 "- Поддържа до 8 разделения на алармената система;" & vbCrLf &
                 "- Поддържа до 254 разширителни модула;" & vbCrLf &
                 "- Поддържа до 999 потребителски кода;" & vbCrLf &
                 "- Поддържа до 999 дистанционни управления (при RTX3);" & vbCrLf &
                 "- Памет за 2048 събития;" & vbCrLf &
                 "- 5 програмируеми(PGM) изхода на платката."
            .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
            .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
            .Cells(index, 3).Value = "бр."
            .Cells(index, 4).Value = 1
        End With
        index += 1
        Dim akum As Integer = 1
        For i = 2 To 10000
            If Trim(wsSOT.Cells(i, 2).Value) = "" Then Exit For
            ProgressBar_Extrat.Value = i
            Text_Dostawka = ""
            broj_elementi = wsSOT.Cells(i, 1).Value
            Select Case wsSOT.Cells(i, 2).Value
                Case "СОТ"
                    Select Case wsSOT.Cells(i, 3).Value
                        Case "СОТ_МУК"
                            Text_Dostawka = "магнитен контакт"
                        Case "СОТ_Насочен"
                            Text_Dostawka = "детектор за движение" & vbCrLf &
                                            "- Ключ за защита от външна намеса;" & vbCrLf &
                                            "- Ударно и температурно защитен корпус;" & vbCrLf &
                                            "- Без регистриране на движението на домашни животни до 40 кг;" & vbCrLf &
                                            "- Обхват 11м х 11м ъгъл на наблюдение."
                        Case "СОТ_Сирена"
                            Text_Dostawka = "сирена за външен монтаж:" & vbCrLf &
                                "- Сила на звука dB(A) на 1m: > 120 dB;" & vbCrLf &
                                "- с жълта лампа;" & vbCrLf &
                                "- тампер;" & vbCrLf &
                                "- Защита срещу пробиване, пяна или прегряване;" & vbCrLf &
                                "- Работна честота: 1400 ~ 1600 Hz;" & vbCrLf &
                                "- Максимално време за работа на сирената след активация: 15 минути;" & vbCrLf &
                                "- Автономна работа при покой до 60 часа след прекъсване на връзката с алармената система;" & vbCrLf &
                                "- Степен на защита: IP44 / IK10."
                        Case "Клавиатура"
                            Text_Dostawka = "чувствителна на допир клавиатура:" & vbCrLf &
                                         "- Визуализация на зоните в аларма;" & vbCrLf &
                                         "- 32-символен син LCD екран;" & vbCrLf &
                                         "- 1 адресируема зона и 1 PGM изход;" & vbCrLf &
                                         "- 3 паник аларми, задействани с едно докосване на клавиш."
                        Case "Паник бутон"
                            Text_Dostawka = "паник бутон за алармени системи:" & vbCrLf &
                                            "- Механичен;" & vbCrLf &
                                            "- Конфигурация (отворен/затворен контакт): По избор."
                        Case "Приемник паник бутон"
                            Text_Dostawka = "безжичен паник бутон/дистанционно:" & vbCrLf &
                                            "- Обхват: 30m."
                            With wsKol_Smetka
                                .Cells(index, 2).Value = "Доставка и монтаж на " + Text_Dostawka
                                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                                .Cells(index, 3).Value = "бр."
                                .Cells(index, 4).Value = broj_elementi
                            End With
                            index += 1
                            Text_Dostawka = "приемник паник бутон:" & vbCrLf &
                                            "- Обхват: 30m;" & vbCrLf &
                                            "- Максимален брой дистанционни: 32;" & vbCrLf &
                                            "- Следене на честотния спектър против заглушаване;" & vbCrLf &
                                            "- Следене състоянието и изпращане на сигнали за слаба батерия, тампер;" & vbCrLf &
                                            "- Визуализация на силата на сигнала."
                        Case "Разпрефелител-32"
                            Text_Dostawka = "32 зонов разширителен модул за алармени системи:" & vbCrLf &
                                            "- Брой зони: 32;" & vbCrLf &
                                            "- Вградено 2A захранване;" & vbCrLf &
                                            "- Монтира се на DIN шина."
                            akum = akum + broj_elementi
                        Case "Разпределител-16"
                            Text_Dostawka = "16 зонов разширителен модул за алармени системи:" & vbCrLf &
                                            "- Брой зони: 16;" & vbCrLf &
                                            "- Вградено 2A захранване;" & vbCrLf &
                                            "- Монтира се на DIN шина."
                            akum = akum + broj_elementi
                        Case "Разпределител-8"
                            Text_Dostawka = "8 зонов разширителен модул за алармени системи:" & vbCrLf &
                                            "- Брой зони: 8;" & vbCrLf &
                                            "- Вградено 2A захранване;" & vbCrLf &
                                            "- Монтира се на DIN шина."
                            akum = akum + broj_elementi
                        Case "СОТ_360"
                            Text_Dostawka = "360° цифров таванен датчик за движение:" & vbCrLf &
                                            "- 360° ъгъл на наблюдение" & vbCrLf &
                                            "- Обхват 7м на височина 2.4м" & vbCrLf &
                                            "- Обхват 11м на височина 3.7м"
                        Case "Вибрационен"
                            Text_Dostawka = "сеизмичен, вибрационeн детектор за трезори" & vbCrLf &
                                            "- 4 регулируеми нива на чувствителността;" & vbCrLf &
                                            "- Тампер ключ за защита от външна намеса;" & vbCrLf &
                                            "- Обхват на действие 2,5 м;" & vbCrLf &
                                            "- Температура на работа: -20°C до +50°C (-40°F до +122°F)."
                        Case "Датчик каса"
                            Text_Dostawka = "вибрационен детектор за каса" & vbCrLf &
                                         "- Брой на ударите, необходими за детектиране на алармен сигнал – регулира се чрез тример в зависимост от конкретните условия:" & vbCrLf &
                                         "  • при силни удари – алармата се задейства при по-малък брой удари" & vbCrLf &
                                         "  • при слаби удари – алармата се задейства при по-голям брой удари" & vbCrLf &
                                         "- Радиус на охраняваната повърхност :" & vbCrLf &
                                         "  • при панели, дървени, талашитни конструкции – до 1,80 m" & vbCrLf &
                                         "  • метални шкафове – до 1,50m" & vbCrLf &
                                         "  • дебелостенни бронирани каси и сейфове – до 1,0 m"
                        Case "СОТ_Звуков"
                            Text_Dostawka = "акустичен цифров датчик за стъкло:" & vbCrLf &
                                            "- 7 цифрови честотни филтъра;" & vbCrLf &
                                            "- Цифров усилвател на нивото и оценка на колебанията на честотата;" & vbCrLf &
                                            "- Настройка на чувствителността: до 9 м при висока и до 4.5 м при ниска чувствителност;" & vbCrLf &
                                            "- Ключ за защита от външна намеса."
                        Case "Микровълнов"
                            Text_Dostawka = "комбиниран ултразвуков-микровълнов детектор на движение:" & vbCrLf &
                                "- Охранявана зона: до 9х7 m;" & vbCrLf &
                                "- Скорост на детекция: 0,2 ÷ 3 m/sec."
                        Case Else
                            Text_Dostawka = "НЕИЗВЕСТЕН ЕЛЕМЕНТ --> " &
                                            wsPIC.Cells(i, 2).Value & " - " &
                                            wsPIC.Cells(i, 3).Value
                    End Select
                    index_Kabel += 1
                Case "Качване"
                    Continue For
            End Select
            If Trim(Text_Dostawka) <> "" Then
                With wsKol_Smetka
                    .Cells(index, 2).Value = "Доставка и монтаж на " + Text_Dostawka
                    .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                    .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                    .Cells(index, 3).Value = "бр."
                    .Cells(index, 4).Value = broj_elementi
                End With
                index += 1
            End If
        Next
        index = Kol_Smetka_Kabeli(index, vbFalse, "СОТ")
        Text_Dostawka = "акумолатор 7Аh" & vbCrLf &
                        "- Работно напрежение: 12V;" & vbCrLf &
                        "- Капацитет: 7Ah."
        With wsKol_Smetka
            .Cells(index, 2).Value = "Доставка и монтаж на " + Text_Dostawka
            .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
            .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
            .Cells(index, 3).Value = "бр."
            .Cells(index, 4).Value = akum
        End With
        index += 1
        Text_Dostawka = "захранваща платка за алармени системи:" & vbCrLf &
                        "- 13.8Vdc, 1.75A импулсно захранване" & vbCrLf &
                        "- Електронна защита с автоматично възстановяване;" & vbCrLf &
                        "- Автоматично прехвърляне на резервната батерия;" & vbCrLf &
                        "- Конектор за втора резервна батерия (опция);" & vbCrLf &
                        "- Вход за тест на акумулаторната батерия."
        With wsKol_Smetka
            .Cells(index, 2).Value = "Доставка и монтаж на " + Text_Dostawka
            .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
            .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
            .Cells(index, 3).Value = "бр."
            .Cells(index, 4).Value = akum
        End With
        index += 1
        Text_Dostawka = "мрежов трансформатор с предпазител:" & vbCrLf &
                        "- Мощност:45 VA;" & vbCrLf &
                        "- Захранващо напрежение: 220V;" & vbCrLf &
                        "- Вторично напрежение: 16,5V."
        With wsKol_Smetka
            .Cells(index, 2).Value = "Доставка и монтаж на " + Text_Dostawka
            .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
            .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
            .Cells(index, 3).Value = "бр."
            .Cells(index, 4).Value = akum
        End With
        index += 1
        With wsKol_Smetka
            .Cells(index, 2).Value = "Доставка и монтаж на метална кутия монтаж на контролни панели за алармени системи:" & vbCrLf &
                                    "- Вграден тампер ключ"
            .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
            .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
            .Cells(index, 3).Value = "бр."
            .Cells(index, 4).Value = akum
        End With
        index += 1
        Text_Dostawka = "Комплексно изпитване на системата"
        With wsKol_Smetka
            .Cells(index, 2).Value = Text_Dostawka
            .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
            .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
            .Cells(index, 3).Value = "бр."
            .Cells(index, 4).Value = 1
        End With
        index += 1
        Text_Dostawka = "Обучение на персонал за работа със системата"
        With wsKol_Smetka
            .Cells(index, 2).Value = Text_Dostawka
            .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
            .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
            .Cells(index, 3).Value = "бр."
            .Cells(index, 4).Value = 1
        End With
        index += 1
        ProgressBar_Extrat.Value = ProgressBar_Extrat.Minimum
    End Sub
    Private Sub Button_силова_КАЧВАНИЯ_Click(sender As Object, e As EventArgs)
        If IsNothing(excel_Workbook) Then
            MsgBox("Първо да беше отворил файла?")
            Exit Sub
        End If
        Me.Visible = vbFalse
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim edt As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database
        Dim SelectedSet = cu.GetObjects("INSERT", "Изберете ВСИЧКИ качвания:")


        Me.Visible = vbTrue
    End Sub
    Private Sub Button_Траншея_ВЪНШНО_Click(sender As Object, e As EventArgs) Handles Button_Траншея_ВЪНШНО.Click,
                                                                                      Button_Траншея_ФОТОВОЛТАИЦИ.Click
        Me.Visible = vbFalse
        Dim cu As CommonUtil = New CommonUtil()
        Dim ss = cu.GetObjects("LINE", "Изберете Линия")
        Dim i, index As Integer
        Dim Инсталация As String = "ТРАНШЕЯ"
        Me.Visible = vbTrue
        If ss Is Nothing Then
            MsgBox("Няма маркиран линия в слой 'EL'.")
            Exit Sub
        End If
        Dim Kabel_fec As Double = 1
        If sender.name = "Button_Траншея_ФОТОВОЛТАИЦИ" Then
            Kabel_fec = 10
        End If
        Dim Kabel(1000, 2) As String
        Kabel = cu.GET_LINE_TYPE_KABEL(Kabel, ss, vbFalse)
        For index = 2 To 1000
            If wsLines.Cells(index, 1).Value = "" Then Exit For
        Next
        Dim formu As String = ""
        Dim rang As String = ""
        Dim colum As String = ""
        Dim kabelSort(500) As strKabel

        For i = 0 To UBound(Kabel)
            If Kabel(i, 0) = Nothing Then Exit For
            With wsLines
                .Cells(index, 1) = Инсталация + "-КАБЕЛ"
                .Cells(index, 2) = Kabel(i, 0) + "x" + Kabel(i, 1) + "cm"
                .Cells(index, 3) = Kabel(i, 0)
                .Cells(index, 4) = Kabel(i, 1)
                .Cells(index, 5) = Kabel(i, 2) / 100 / Kabel_fec
                .Cells(index, 6) = Int(Kabel(i, 2) / 1000 + 1) * 10 / Kabel_fec
            End With
            index += 1
        Next
        '
        '  Сортира EXCEL листа
        ' 
        With wsLines
            .Cells.Sort(Key1:= .Range("A2"),
                        Order1:=excel.XlSortOrder.xlAscending,
                        Header:=excel.XlYesNoGuess.xlYes,
                        OrderCustom:=1, MatchCase:=False,
                        Orientation:=excel.Constants.xlTopToBottom,
                        DataOption1:=excel.XlSortDataOption.xlSortTextAsNumbers,
                        Key2:= .Range("B2"),
                        Order2:=excel.XlSortOrder.xlAscending,
                        DataOption2:=excel.XlSortDataOption.xlSortTextAsNumbers
                        )
        End With

    End Sub
    Private Sub Button_Вземи_ФОТОВОЛТАИЦИ_Click(sender As Object, e As EventArgs) Handles Button_Вземи_ФОТОВОЛТАИЦИ.Click

        fullName = Application.DocumentManager.MdiActiveDocument.Name
        filePath = Mid(fullName, 1, InStrRev(fullName, "\"))
        fileName = Mid(fullName, InStrRev(fullName, "\") + 1, Len(fullName) - 6)
        '
        OpenFileDialog1.InitialDirectory = filePath
        OpenFileDialog1.Filter = "Excel files (*.xls or *.xlsx)|*.xls;*.xlsx"
        OpenFileDialog1.FileName = "ФВЦ.xlsx"
        '
        If OpenFileDialog1.ShowDialog <> Windows.Forms.DialogResult.OK Then
            Exit Sub
        End If
        '
        nameExcel = OpenFileDialog1.FileName
        '
        ' Проверява дали EXCEL е отворен
        '
        Dim stream As FileStream = Nothing
        Try
            stream = File.Open(nameExcel, FileMode.Open, FileAccess.ReadWrite, FileShare.None)
            stream.Close()
        Catch ex As Exception
            MsgBox("Отворен е файл с име : " + Chr(13) + Chr(13) +
                   nameExcel + Chr(13) + Chr(13) +
                   "Моля затворете го преди да продължите!")
            Exit Sub
        End Try
        '
        'Get all currently running process Ids for Excel applications
        '
        NoMyExcelProcesses = Process.GetProcessesByName("Excel")
        '
        Dim objExcel_FEC As excel.Application = New excel.Application()
        Dim excel_Workbook_FEC As excel.Workbook = objExcel_FEC.Workbooks.Open(nameExcel)
        Dim wsFEC_Tabloca As excel.Worksheet = excel_Workbook_FEC.Worksheets("Таблица")
        Dim wsFEC_ФЕЦ As excel.Worksheet = excel_Workbook_FEC.Worksheets("ФЕЦ")
        Dim Verifi_FEC As Boolean = True
        Dim Kонектор As Integer = 0
        Dim Групи As Integer = 0
        '
        For i = 1 To 6
            If Verifi_FEC And wsFEC_Tabloca.Cells(i, 6).Value <> "OK" Then
                Dim response = MsgBox(wsFEC_Tabloca.Cells(i, 1).Value + "-> not OK" + vbCrLf + vbCrLf + "Да продължа ли?", vbYesNo)
                If response = vbNo Then
                    Exit Sub
                End If
            End If
        Next
        '
        For i = 2 To 11
            wsKoef.Cells(red_Фотоволтаици + 0, i).Value = wsFEC_Tabloca.Cells(11 + i, 5).Value ' "Брой панели"
            wsKoef.Cells(red_Фотоволтаици + 1, i).Value = wsFEC_Tabloca.Cells(11 + i, 1).Value ' "Тип панели"
            wsKoef.Cells(red_Фотоволтаици + 2, i).Value = wsFEC_Tabloca.Cells(25 + i, 5).Value ' "Брой инвертори"
            wsKoef.Cells(red_Фотоволтаици + 3, i).Value = wsFEC_Tabloca.Cells(25 + i, 1).Value ' "Тип инвертори"
        Next
        '
        With wsFEC_ФЕЦ
            For j As Integer = 0 To 9
                If Val(.Cells(5, 3 + j * 5).Value) = 0 Then
                    Continue For
                Else
                    Групи = Val(.Cells(5, 3 + j * 5).Value)
                End If
                For i = 37 To 59
                    Kонектор += IIf(Val(.Cells(i, 3 + j * 5).Value) > 0, Групи * 2, 0)
                Next
            Next
        End With
        '
        wsKoef.Cells(red_Фотоволтаици + 7, 2).Value = Kонектор
        excel_Workbook_FEC.Close()
        excel_Workbook_FEC = Nothing
        '
    End Sub
    Private Sub Button_Генератор_ФОТОВОЛТАИЦИ_Click(sender As Object, e As EventArgs) Handles Button_Генератор_ФОТОВОЛТАИЦИ.Click
        Dim index As Integer = 0
        Dim i As Integer
        ProgressBar_Extrat.Maximum = ProgressBar_Maximum
        For i = 6 To 1000
            If wsKol_Smetka.Cells(i, 2).Value = "Раздел" Or wsKol_Smetka.Cells(i, 2).Value = "" Then
                index = i
                Exit For
            End If
            ProgressBar_Extrat.Value = i
        Next
        For i = 2 To 1000
            If wsKontakti.Cells(i, 2).Value = "" Then
                Exit For
            End If
        Next
        Call Excel_Kol_smetka_Razdel(wsKol_Smetka, "ФОТОВОЛТАИЧНА ЦЕНТРАЛА", "A" & Trim(index.ToString), "D" & Trim(index.ToString))
        index += 1
        Dim Text_Dostawka As String = ""
        Dim Кабел_Общо As Double = 0
        For i = 2 To 11
            If Val(wsKoef.Cells(red_Фотоволтаици, i).Value) = 0 Then Continue For
            With wsKol_Smetka
                .Cells(index, 2).Value = "Доставка и монтаж на монокристален соларен панел тип " +
                    wsKoef.Cells(red_Фотоволтаици + 1, i).Value
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "бр."
                .Cells(index, 4).Value = wsKoef.Cells(red_Фотоволтаици, i).Value
                index += 1
            End With
        Next
        With wsKol_Smetka
            .Cells(index, 2).Value = "Доставка и монтаж фиксираща планка за соларен панел - /външна единична/ с болт и квадратна гайка М8 мм"
            .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
            .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
            .Cells(index, 3).Value = "бр."
            .Cells(index, 4).Value = "###"
            index += 1
        End With
        With wsKol_Smetka
            .Cells(index, 2).Value = "Доставка и монтаж фиксираща планка за соларен панел - /вътрешна двойна/ с болт и квадратна гайка М8 мм"
            .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
            .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
            .Cells(index, 3).Value = "бр."
            .Cells(index, 4).Value = "###"
            index += 1
        End With
        With wsKol_Smetka
            .Cells(index, 2).Value = "Доставка и монтаж соларен комплект конектор МС пожарозащитен клас UL 94-V0Q, IP-67 за кабел Cu 1х6мм²"
            .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
            .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
            .Cells(index, 3).Value = "бр."
            .Cells(index, 4).Value = wsKoef.Cells(red_Фотоволтаици + 7, 2).Value
            index += 1
        End With
        For i = 2 To 11
            If Val(wsKoef.Cells(red_Фотоволтаици + 2, i).Value) = 0 Then Continue For
            With wsKol_Smetka
                .Cells(index, 2).Value = "Доставка и монтаж на инвертор трифазен многострингов " +
                    wsKoef.Cells(red_Фотоволтаици + 3, i).Value
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "бр."
                .Cells(index, 4).Value = wsKoef.Cells(red_Фотоволтаици + 2, i).Value
                index += 1
            End With
            Dim brAkum As Integer = 0
            Dim strbat As String = ""
            Dim strKletka As String = ""
            Dim strMoshnost As String = ""
            Dim strKapacitet As String = ""

            With wsKoef
                brAkum = .Cells(red_Фотоволтаици + 10, 2).Value
                strbat = .Cells(red_Фотоволтаици + 10, 3).Value
                strKletka = .Cells(red_Фотоволтаици + 10, 4).Value
                strMoshnost = .Cells(red_Фотоволтаици + 10, 5).Value
                strKapacitet = .Cells(red_Фотоволтаици + 10, 6).Value
            End With
            With wsKol_Smetka
                .Cells(index, 2).Value = "Доставка и монтаж на батерия " + strbat +
                                         ", клетка " + strKletka + vbCrLf +
                                         "- Мощност: " + strMoshnost + " kW;" +
                                         vbCrLf +
                                         "- Капацитет: " + strKapacitet + " kWh."
                .Cells(index, 2).VerticalAlignment = excel.XlVAlign.xlVAlignTop
                .Cells(index, 2).HorizontalAlignment = excel.XlHAlign.xlHAlignLeft
                .Cells(index, 3).Value = "бр."
                .Cells(index, 4).Value = brAkum
                index += 1
            End With
        Next
        index = Kol_Smetka_Kabeli(index, vbTrue, "ФОТОВОЛТАИЦИ")

        Button_Генератор_ВЪНШНО_Click(sender, e)
    End Sub
    Private Sub Button_Записка_Click(sender As Object, e As EventArgs) Handles Button_Записка.Click
        Dim Zapiska As New Zapiska()
        Zapiska.New_zapiska()
    End Sub
    Private Sub Button_Вземи_БАТЕРИИ_Click(sender As Object, e As EventArgs) Handles Button_Вземи_БАТЕРИИ.Click
        If IsNothing(excel_Workbook) Then
            MsgBox("Първо да беше отворил файла?")
            Exit Sub
        End If
        Me.Visible = vbFalse
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim edt As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database
        Dim SelectedSet = cu.GetObjects("INSERT", "Изберете блок")
        Dim arrBlock(500) As strСкара
        If SelectedSet Is Nothing Then
            MsgBox("Няма маркиран нито един блок.")
            Me.Visible = vbTrue
            Exit Sub
        End If
        Dim blkRecId As ObjectId = ObjectId.Null
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                Dim brAkum As Integer = 0
                Dim strbat As String = ""
                Dim strKletka As String = ""
                Dim strMoshnost As String = ""
                Dim strKapacitet As String = ""
                For Each sObj As SelectedObject In SelectedSet
                    blkRecId = sObj.ObjectId
                    Dim acBlkRef As BlockReference = DirectCast(acTrans.GetObject(blkRecId, OpenMode.ForRead), BlockReference)
                    Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                    Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection

                    Dim blName As String = (CType(acBlkRef.DynamicBlockTableRecord.GetObject(OpenMode.ForRead), BlockTableRecord)).Name

                    If Not blName = "Батерия" Then Continue For
                    For Each objID As ObjectId In attCol
                        Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                        Dim acAttRef As AttributeReference = dbObj
                        ' Съхранение на стойностите на атрибутите в съответните променливи
                        If acAttRef.Tag = "БАТЕРИЯ_ВИД" Then strbat = acAttRef.TextString
                        If acAttRef.Tag = "БАТЕРИЯ_КЛЕТКА" Then strKletka = acAttRef.TextString
                        If acAttRef.Tag = "МОЩНОСТ" Then strMoshnost = acAttRef.TextString
                        If acAttRef.Tag = "КАПАЦИТЕТ" Then strKapacitet = acAttRef.TextString
                    Next
                    brAkum += 1
                Next
                With wsKoef
                    .Cells(red_Фотоволтаици + 10, 1) = "БАТЕРИЯ"
                    .Cells(red_Фотоволтаици + 10, 2) = brAkum
                    .Cells(red_Фотоволтаици + 10, 3) = strbat
                    .Cells(red_Фотоволтаици + 10, 4) = strKletka
                    .Cells(red_Фотоволтаици + 10, 5) = strMoshnost
                    .Cells(red_Фотоволтаици + 10, 6) = strKapacitet
                End With

            Catch ex As Exception
                MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
                acTrans.Abort()
            End Try
        End Using
    End Sub
End Class