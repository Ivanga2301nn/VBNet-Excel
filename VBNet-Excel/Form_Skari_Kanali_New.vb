Imports System.ComponentModel
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.IO
Imports System.Net
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Public Class Form_Skari_Kanali_New
    Dim PI As Double = 3.1415926535897931
    Dim cu As CommonUtil = New CommonUtil()
    Dim Line_Selected As SelectionSet
    Dim Slabo_Yes As Boolean = vbFalse
    Dim Solar_Yes As Boolean = vbFalse
    Dim Silow_Yes As Boolean = vbFalse
    Dim summaKabeli As Double = 0
    Dim skara As Double = 0
    Dim Шир As Double = 0
    Dim Вис As Double = 0
    Dim razdelitel As Integer = 0
    Structure Скара
        Dim Ширина As String
        Dim Височина As String
        Dim Площ As Double
        Dim Процент As Double
    End Structure
    Structure strLine
        Dim Layer As String
        Dim Linetype As String
        Dim count As Double
        Dim Diam As Double
        Dim Se4enie As Double
        Dim Start_Point As Point3d
        Dim End_Point As Point3d
        Dim Lenght As Double
        Dim Angle As Double
    End Structure
    Public TrayCatalog As New List(Of Скара)
    Public DuctCatalog As New List(Of Скара)
    Dim Kabel(200) As strLine
    Private Sub DataGridView_Кабели_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) _
    Handles DataGridView_Кабели.DataError
        e.Cancel = True  ' Игнорирай грешката
    End Sub
    Private Sub Skari_Kanali_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Line_Selected = cu.GetObjects("LINE", "Изберете линии за кабеланата скара/канал:")
        If Line_Selected Is Nothing Then
            MsgBox("НЕ Е маркиранa линия в слой 'EL'.")
            Exit Sub
        End If
        InitializeCatalog_Скари()
        InitializeCatalog_Канали()

        CreateGradientProgressBars()

        GroupBox_Размери_Скари.Visible = False
        GroupBox_Размери_Скари.Dock = DockStyle.Fill
        Set_array_Kabel()
        Label_Площ.Text = "Площ: " + summaKabeli.ToString + " mm"
    End Sub
    Private Sub Set_array_Kabel()
        If Line_Selected Is Nothing Then Exit Sub
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Using trans As Transaction = acDoc.TransactionManager.StartTransaction()
            Try
                Dim Index As Integer
                For Each sObj As SelectedObject In Line_Selected
                    Dim line As Line = TryCast(trans.GetObject(sObj.ObjectId, OpenMode.ForRead), Line)
                    Dim iVisib As Integer = -1
                    iVisib = Array.FindIndex(Kabel, Function(f) f.Layer = line.Layer)
                    If iVisib = -1 Then
                        Kabel(Index).Layer = line.Layer
                        Kabel(Index).Linetype = line.Linetype
                        Kabel(Index).Lenght = line.Length
                        Kabel(Index).Start_Point = line.StartPoint
                        Kabel(Index).End_Point = line.EndPoint
                        Kabel(Index).Angle = line.Angle
                        Kabel(Index).count = 1
                        Index += 1
                    Else
                        Kabel(iVisib).count = Kabel(iVisib).count + 1
                    End If
                Next
                trans.Commit()
            Catch ex As Exception
                MsgBox("Възникна грешка:  " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
                trans.Abort()
            End Try
        End Using
        summaKabeli = 0
        For br = 0 To UBound(Kabel)
            If Kabel(br).Layer = Nothing Then Exit For
            Kabel(br).Diam = cu.GET_line_Diamet(Kabel(br).Layer)
            Kabel(br).Se4enie = Kabel(br).Diam * Kabel(br).Diam * Kabel(br).count
            summaKabeli = summaKabeli + Kabel(br).Se4enie
            Select Case Kabel(br).Layer
                Case "EL_JC_2x05_sig", "EL_JC_2x05_zahr", "EL_JC_2x15_А", "EL_JC_2x15_Б",
                         "EL_JC_2x05", "EL_JC_2x10", "EL_JC_2x75", "EL_JC_3x05", "EL_JC_3x10",
                         "EL_JC_3x75", "EL_JC_4x05", "EL_JC_4x10", "EL_JC_4x75", "EL_SOT", "EL_Tel",
                         "EL_TV", "EL_UTP", "EL_Video", "EL_Video_FTP", "EL_Video_RG59CU", "EL_HDMI", "EL_6BPL2x1_0"
                    Slabo_Yes = vbTrue
                    DataGridView_Кабели.Rows.Add(New String() {"Слаботоков", "1", "1,5", Kabel(br).count, Kabel(br).Diam, Kabel(br).Se4enie})
                Case "EL_стринг",
                     "EL_стринг1", "EL_стринг2", "EL_стринг3", "EL_стринг4",
                     "EL_стринг5", "EL_стринг6", "EL_стринг7", "EL_стринг8",
                     "EL_стринг9", "EL_стринг10", "EL_стринг11", "EL_стринг12",
                     "EL_стринг13", "EL_стринг14", "EL_стринг15", "EL_стринг16",
                     "EL_стринг17", "EL_стринг18", "EL_стринг19", "EL_стринг20"

                    DataGridView_Кабели.Rows.Add(New String() {"Соларен", "1", "4", Kabel(br).count, Kabel(br).Diam, Kabel(br).Se4enie})
                    Solar_Yes = vbTrue
                Case "ELEKTRO", "EL_AlMgSi Ф8мм", "EL__DIM", "EL__KOTI", "EL__ORAZ", "EL_Канали", "EL_Скари", "EL_ТАБЛА"
                    Continue For
                Case Else
                    If Mid(Kabel(br).Layer, 4, 5) = "NHXCH" Then Continue For
                    If Mid(Kabel(br).Layer, 4, 2) = "UK" Then Continue For
                    If Mid(Kabel(br).Layer, 4, 2) = "PB" Then
                        DataGridView_Кабели.Rows.Add(New String() {"СВТ/САВТ", "1",
                                                     "16", Kabel(br).count,
                                                     Kabel(br).Diam, Kabel(br).Se4enie})
                        Continue For
                    End If

                    Dim brvi As String = Mid(Kabel(br).Layer, 4, 1) +
                                             IIf(InStr("+", Kabel(br).Layer), "+", "")
                    Dim sev As String = Mid(Kabel(br).Layer, 6, Len(Kabel(br).Layer))
                    sev = Replace(sev, "_", ",")


                    DataGridView_Кабели.Rows.Add(New String() {"СВТ/САВТ", brvi,
                                                     sev, Kabel(br).count,
                                                     Kabel(br).Diam, Kabel(br).Se4enie})
                    Silow_Yes = vbTrue
            End Select
        Next
        If summaKabeli > 66000 Then
            MsgBox("АЕЦ НЯМА ДА СМЯТАМЕ! АКО ВЕСЕ ПАК Е АЕЦ РАЗДЕЛИ ТРАСЕТО НА НЯКОЛКО СКАРИ!!!")
            Me.Close()
            Exit Sub
        End If
    End Sub
    Private Async Sub RadioButton_Скара_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton_Скара.CheckedChanged
        If RadioButton_Скара.Checked Then
            GroupBox_Размери_Скари.Text = "Размер на кабелната скара, mm"
            GroupBox_Размери_Скари.Visible = True
            GroupBox3.Visible = True
            Izbor_Element(15, TrayCatalog)
            CreateGrid("Скара")
            ' 2. Вземи процента от ComboBox
            Dim fillPercent As Double = 40
            If ComboBox_Процент_Запълване.SelectedItem IsNot Nothing Then
                Double.TryParse(ComboBox_Процент_Запълване.SelectedItem.ToString(), fillPercent)
            End If
            ' 3. Изчакай анимацията да завърши
            Await UpdateProgressBarsAnimated(TrayCatalog, 250, fillPercent)
            Label_Skara.Text = "Скара [ШхВ]"
        End If
    End Sub
    Private Async Sub RadioButton_Канал_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton_Канал.CheckedChanged
        If RadioButton_Канал.Checked Then
            GroupBox_Размери_Скари.Visible = False
            GroupBox_Размери_Скари.Text = "Размер на кабелните канали, mm"
            GroupBox3.Visible = False
            Izbor_Element(0, DuctCatalog)
            CreateGrid("Канал")
            ' 2. Вземи процента от ComboBox
            Dim fillPercent As Double = 40
            If ComboBox_Процент_Запълване.SelectedItem IsNot Nothing Then
                Double.TryParse(ComboBox_Процент_Запълване.SelectedItem.ToString(), fillPercent)
            End If
            ' 3. Изчакай анимацията да завърши
            Await UpdateProgressBarsAnimated(DuctCatalog, 250, fillPercent)
            Label_Skara.Text = "Канал [ШхВ]"
            GroupBox_Размери_Скари.Visible = True
        End If
    End Sub
    Private Sub RadioButton_Тръба_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton_Тръба.CheckedChanged
        GroupBox_Размери_Скари.Visible = False
        GroupBox3.Visible = True
        Label_Skara.Text = "Тръба [ø]"
    End Sub
    Private Async Sub NumericUpDown_Razdelitel_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown_Razdelitel.ValueChanged
        razdelitel = NumericUpDown_Razdelitel.Value
        Izbor_Element(15, TrayCatalog)

        ' 2. Вземи процента от ComboBox
        Dim fillPercent As Double = 40
        If ComboBox_Процент_Запълване.SelectedItem IsNot Nothing Then
            Double.TryParse(ComboBox_Процент_Запълване.SelectedItem.ToString(), fillPercent)
        End If
        ' 3. Изчакай анимацията да завърши
        Await UpdateProgressBarsAnimated(TrayCatalog, 250, fillPercent)
    End Sub
    Private Sub DataGridView_Кабели_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_Кабели.CellValueChanged
        Set_array_Kabel()
        Label_Площ.Text = "Площ: " + summaKabeli.ToString + " mm²"
    End Sub
    Private Async Sub ComboBox_Процент_Запълване_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_Процент_Запълване.SelectedIndexChanged
        ' 2. Вземи процента от ComboBox
        Dim fillPercent As Double = 40
        If ComboBox_Процент_Запълване.SelectedItem IsNot Nothing Then
            Double.TryParse(ComboBox_Процент_Запълване.SelectedItem.ToString(), fillPercent)
        End If

        If RadioButton_Канал.Checked = True Then
            Izbor_Element(0, DuctCatalog)
            Await UpdateProgressBarsAnimated(DuctCatalog, 250, fillPercent)
        Else
            Izbor_Element(15, TrayCatalog)
            Await UpdateProgressBarsAnimated(TrayCatalog, 250, fillPercent)
        End If
    End Sub
    ''' <summary>
    ''' Избира най-подходящата кабелна скара от каталога (TrayCatalog)
    ''' според:
    '''    • сумарната площ на кабелите
    '''    • избрания процент на запълване
    '''    • броя на разделителите
    '''
    ''' Логика на изчисление:
    '''
    ''' 1) Взема от интерфейса:
    '''    - Процент запълване (ComboBox_Процент_Запълване)
    '''    - Брой разделители (NumericUpDown_Razdelitel)
    '''
    ''' 2) Изчислява необходимата минимална площ:
    '''        requiredArea = totalCableArea / fillFactor
    '''
    ''' 3) За всяка скара от каталога (сортирана по площ, възходящо):
    '''    - Изчислява полезната ширина:
    '''        effectiveWidth = Ширина - (брой разделители × 15mm)
    '''
    '''    - Ако остава положителна ширина:
    '''        effectiveArea = effectiveWidth × Височина
    '''
    '''    - Проверява дали effectiveArea >= requiredArea
    '''
    ''' 4) Връща първата (най-малката) подходяща скара.
    '''
    ''' 5) Автоматично попълва TextBox_Кабелна_Скара
    '''    във формат "Ширина x Височина".
    '''
    ''' Забележка:
    ''' Сепараторът намалява полезната ширина с 15 mm за всеки брой.
    ''' </summary>
    '''
    ''' <returns>
    ''' Обект от тип Скара, отговарящ на изчислените условия.
    ''' Връща Nothing, ако няма подходящ размер.
    ''' </returns>
    Function Izbor_Element(sepWidth As Double, ByRef Catalog As List(Of Скара)) As Скара
        ' 1. Подготовка на данните
        Dim selectedPercent As String = ComboBox_Процент_Запълване.SelectedItem.ToString()
        Dim fillFactor As Double = CDbl(selectedPercent) / 100
        Dim numDividers As Integer = CInt(NumericUpDown_Razdelitel.Value)
        Dim separatorWidth As Double = sepWidth
        Dim totalCableArea As Double = summaKabeli
        Dim foundTray As Скара = Nothing
        ' Сортираме, за да намерим най-малкия подходящ модел за TextBox-а
        Dim sortedCatalog = Catalog.OrderBy(Function(x) x.Площ).ToList()
        ' 2. Минаваме през абсолютно всички елементи в каталога
        For i As Integer = 0 To sortedCatalog.Count - 1
            Dim currentTray = sortedCatalog(i)
            Dim w As Double = CDbl(currentTray.Ширина)
            Dim h As Double = CDbl(currentTray.Височина)
            Dim effectiveWidth As Double = w - (numDividers * separatorWidth)

            If effectiveWidth <= 0 Then
                ' Трейът е твърде малък за толкова разделители - не може да се използва
                currentTray.Процент = 100  ' Маркирай го като "препълнен"
                sortedCatalog(i) = currentTray
                Continue For  ' Премини към следващия трей
            End If

            ' Изчисляваме реалния процент (математически)
            Dim effectiveArea As Double = effectiveWidth * h
            ' Тук записваме реалната стойност (може да е 20%, може да е 500%)
            currentTray.Процент = (totalCableArea / effectiveArea) * 100
            ' Актуализираме обекта в списъка (тъй като е Structure, трябва да го върнем обратно)
            sortedCatalog(i) = currentTray
            ' 3. Логика за избор на оптимален модел (за TextBox)
            ' Търсим първия, чийто Процент е по-малък или равен на избрания в ComboBox-а
            If foundTray.Ширина = Nothing AndAlso currentTray.Процент <= (fillFactor * 100) Then
                foundTray = currentTray
            End If
        Next
        ' 4. Визуализация в TextBox
        If foundTray.Ширина <> Nothing Then
            TextBox_Кабелна_Скара.Text = foundTray.Ширина & "x" & foundTray.Височина
        Else
            TextBox_Кабелна_Скара.Text = "=Няма="
        End If
        ' Връщаме целия списък обратно в оригиналния Catalog, за да ползваме изчислените проценти
        Catalog = sortedCatalog

        Return foundTray
    End Function
    ''' <summary>
    ''' Инициализира каталога със скари (TrayCatalog),
    ''' като зарежда предварително дефинирани комбинации
    ''' от ширини и височини.
    '''
    ''' Процедурата:
    ''' 1) Изчиства съществуващия каталог.
    '''
    ''' 2) Дефинира допустимите ширини за всяка височина:
    '''    • Височина 35 mm  →  50, 100, 150, 200, 300, 400, 500, 600
    '''    • Височина 60 mm  →  50, 100, 150, 200, 300, 400, 500, 600
    '''    • Височина 85 mm  →  100, 150, 200, 300, 400, 500, 600
    '''    • Височина 110 mm →  200, 300, 400, 500, 600
    '''
    ''' 3) За всяка комбинация ширина/височина
    '''    извиква AddTrayToCatalog(), която добавя
    '''    съответния елемент в каталога.
    '''
    ''' Целта е да се гарантира, че системата работи
    ''' само с допустими фабрични размери на кабелни скари.
    ''' </summary>
    Public Sub InitializeCatalog_Скари()
        TrayCatalog.Clear()
        Dim widths35() As Integer = {50, 100, 150, 200, 300, 400, 500, 600}
        Dim widths60() As Integer = {50, 100, 150, 200, 300, 400, 500, 600}
        Dim widths85() As Integer = {100, 150, 200, 300, 400, 500, 600}
        Dim widths110() As Integer = {200, 300, 400, 500, 600}
        ' Пълнене за височина 35
        For Each w In widths35
            AddToCatalog(w, 35, TrayCatalog)
        Next
        ' Пълнене за височина 60
        For Each w In widths60
            AddToCatalog(w, 60, TrayCatalog)
        Next
        ' Пълнене за височина 85
        For Each w In widths85
            AddToCatalog(w, 85, TrayCatalog)
        Next
        ' Пълнене за височина 110/100
        For Each w In widths110
            AddToCatalog(w, 110, TrayCatalog)
        Next
    End Sub
    ''' <summary>
    ''' Помощна функция за създаване на обект Скара и добавяне в списъка
    ''' </summary>
    Private Sub AddToCatalog(w As Integer, h As Integer, ByRef Catalog As List(Of Скара))
        Dim новЕлемент As New Скара
        новЕлемент.Ширина = w.ToString()
        новЕлемент.Височина = h.ToString()
        новЕлемент.Площ = CDbl(w * h)
        Catalog.Add(новЕлемент)
    End Sub
    Public Sub InitializeCatalog_Канали()
        DuctCatalog.Clear()
        Dim widths16() As Integer = {16, 25, 40}
        Dim widths20() As Integer = {20, 25, 40, 60, 80}
        Dim widths25() As Integer = {25, 40, 80}
        Dim widths40() As Integer = {25, 40, 60, 80, 100, 120}
        Dim widths60() As Integer = {40, 60, 80, 100, 140}
        Dim widths80() As Integer = {40, 60, 80}

        For Each w In widths16
            AddToCatalog(w, 16, DuctCatalog)
        Next
        For Each w In widths20
            AddToCatalog(w, 20, DuctCatalog)
        Next
        For Each w In widths25
            AddToCatalog(w, 25, DuctCatalog)
        Next
        For Each w In widths40
            AddToCatalog(w, 40, DuctCatalog)
        Next
        For Each w In widths60
            AddToCatalog(w, 60, DuctCatalog)
        Next
        For Each w In widths80
            AddToCatalog(w, 80, DuctCatalog)
        Next
    End Sub
    Private Sub CreateGradientProgressBars()
        ' Почисти старите контроли
        TableLayoutPanel.Controls.Clear()

        ' Брой редове и колони (адаптирай според твоята таблица)
        Dim rows As Integer = 9
        Dim cols As Integer = 9

        ' Настрой TableLayoutPanel
        TableLayoutPanel.RowCount = rows
        TableLayoutPanel.ColumnCount = cols

        ' Изчисти и добави стилове за редовете
        TableLayoutPanel.RowStyles.Clear()
        For i As Integer = 1 To rows - 1
            TableLayoutPanel.RowStyles.Add(New RowStyle(SizeType.Percent, 100 / rows))
        Next

        ' Изчисти и добави стилове за колоните
        TableLayoutPanel.ColumnStyles.Clear()
        For i As Integer = 1 To cols - 1
            TableLayoutPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100 / cols))
        Next

        ' Създай и добави GradientProgressBar за всяка клетка
        For row As Integer = 1 To rows - 2
            For col As Integer = 1 To cols - 2
                Dim gp As New GradientProgressBar()
                gp.Name = $"gp_{row}_{col}"
                gp.Dock = DockStyle.Fill      ' Попълва цялата клетка
                gp.Margin = New Padding(2)    ' Малко разстояние между тях
                gp.Minimum = 0
                gp.Maximum = 100
                gp.Value = 0
                gp.ShowText = True

                ' Различни цветове по колони (по желание)
                Select Case col
                    Case 0 ' 35 мм
                        gp.StartColor = Color.FromArgb(0, 192, 0)
                        gp.EndColor = Color.FromArgb(0, 128, 255)
                    Case 1 ' 60 мм
                        gp.StartColor = Color.FromArgb(255, 205, 86)
                        gp.EndColor = Color.FromArgb(255, 159, 64)
                    Case 2 ' 85 мм
                        gp.StartColor = Color.FromArgb(54, 162, 235)
                        gp.EndColor = Color.FromArgb(0, 192, 0)
                    Case 3 ' 110 мм
                        gp.StartColor = Color.FromArgb(255, 99, 132)
                        gp.EndColor = Color.FromArgb(255, 159, 64)
                End Select

                TableLayoutPanel.Controls.Add(gp, col, row)
            Next
        Next
    End Sub
    Private Sub CreateGrid(tip As String)
        ' Изчистваме старите контроли
        TableLayoutPanel.Controls.Clear()
        TableLayoutPanel.RowStyles.Clear()
        TableLayoutPanel.ColumnStyles.Clear()

        ' Дефиниции за Скари и Канали
        Dim widths() As Integer      ' ⬅️ ХОРИЗОНТАЛНО (колони)
        Dim heights() As Integer     ' ⬅️ ВЕРТИКАЛНО (редове)
        Dim rows As Integer
        Dim cols As Integer

        If tip = "Скара" Then
            ' Кабелни скари: 8 ширини × 4 височини
            widths = {50, 100, 150, 200, 300, 400, 500, 600}
            heights = {35, 60, 85, 110}
            rows = widths.Length + 1   ' 8 + 1 header = 9
            cols = heights.Length + 1  ' 4 + 1 header = 5
        ElseIf tip = "Канал" Then
            ' Кабелни канали: 9 ширини × 9 височини
            widths = {25, 40, 60, 80, 100, 120, 140}
            heights = {20, 25, 40, 60, 80, 100, 140}
            rows = widths.Length + 1   ' 9 + 1 header = 10
            cols = heights.Length + 1  ' 9 + 1 header = 10
        End If

        ' Настройваме TableLayoutPanel
        TableLayoutPanel.RowCount = rows
        TableLayoutPanel.ColumnCount = cols

        ' Добавяме стилове за редовете
        TableLayoutPanel.RowStyles.Add(New RowStyle(SizeType.Absolute, 40)) ' Header ред
        For i As Integer = 0 To widths.Length - 1
            TableLayoutPanel.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0 / widths.Length))
        Next

        ' Добавяме стилове за колоните
        TableLayoutPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 60)) ' Header колона
        For i As Integer = 0 To heights.Length - 1
            TableLayoutPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0 / heights.Length))
        Next

        ' === Ред 0: Височини (Headers) ===
        For col As Integer = 0 To heights.Length - 1
            Dim headerLabel As New Label()
            headerLabel.Name = $"Label_1_{col + 1}"
            headerLabel.Text = heights(col).ToString()
            headerLabel.TextAlign = ContentAlignment.MiddleCenter
            headerLabel.Font = New Drawing.Font("Segoe UI", 10, FontStyle.Bold)
            headerLabel.Dock = DockStyle.Fill
            headerLabel.BackColor = Color.FromArgb(200, 220, 255)
            headerLabel.BorderStyle = BorderStyle.FixedSingle
            TableLayoutPanel.Controls.Add(headerLabel, col + 1, 0)
        Next

        ' === Колона 0: Ширини (Headers) + ProgressBars ===
        For row As Integer = 0 To widths.Length - 1
            ' Header за ширината
            Dim widthLabel As New Label()
            widthLabel.Name = $"Label_0_{row + 1}"
            widthLabel.Text = widths(row).ToString()
            widthLabel.TextAlign = ContentAlignment.MiddleCenter
            widthLabel.Font = New Drawing.Font("Segoe UI", 9, FontStyle.Bold)
            widthLabel.Dock = DockStyle.Fill
            widthLabel.BackColor = Color.FromArgb(200, 220, 255)
            widthLabel.BorderStyle = BorderStyle.FixedSingle
            TableLayoutPanel.Controls.Add(widthLabel, 0, row + 1)

            ' ProgressBars за всяка височина
            For col As Integer = 0 To heights.Length - 1
                Dim gp As New GradientProgressBar()
                gp.Name = $"gp_{widths(row)}_{heights(col)}"
                gp.Dock = DockStyle.Fill
                gp.Margin = New Padding(2)
                gp.Minimum = 0
                gp.Maximum = 100
                gp.Value = 0
                gp.ShowText = True
                gp.Font = New Drawing.Font("Segoe UI", 12.0F, FontStyle.Bold)

                ' ⬇️ ТОВА ЛИ ЛИПСВА? ⬇️
                gp.TrayWidth = widths(row).ToString()
                gp.TrayHeight = heights(col).ToString()

                ' ⬇️ ДОБАВИ ТОВА ⬇️
                AddHandler gp.ProgressBarClicked, AddressOf Me.TrayProgressBar_Click
                TableLayoutPanel.Controls.Add(gp, col + 1, row + 1)
            Next
        Next
    End Sub
    Private Sub UpdateProgressBars(ByRef Catalog As List(Of Скара))
        For row As Integer = 1 To TableLayoutPanel.RowCount - 1
            For col As Integer = 1 To TableLayoutPanel.ColumnCount - 1
                Dim ctrl = TableLayoutPanel.GetControlFromPosition(col, row)

                If TypeOf ctrl Is GradientProgressBar Then
                    Dim gp = DirectCast(ctrl, GradientProgressBar)

                    ' Извличаме размерите от името на контрола (gp_50_35)
                    Dim parts() As String = gp.Name.Split("_"c)
                    If parts.Length >= 3 Then
                        Dim width As String = parts(1)
                        Dim height As String = parts(2)

                        ' Търсим съответния елемент в каталога
                        Dim tray = Catalog.FirstOrDefault(Function(t) _
                            t.Ширина = width AndAlso t.Височина = height)

                        If tray.Ширина IsNot Nothing Then
                            ' Задаваме стойността
                            Dim percent As Double = tray.Процент
                            gp.Value = CInt(Math.Min(100, Math.Max(0, percent)))
                            gp.ShowText = True

                            ' ⬇️ ДОБАВИ ТОВА ⬇️
                            gp.TrayWidth = width
                            gp.TrayHeight = height

                            ' Оцветяване според запълването
                            If percent < 50 Then
                                gp.StartColor = Color.FromArgb(0, 192, 0)   ' Зелено
                                gp.EndColor = Color.FromArgb(0, 128, 255)   ' Синьо
                            ElseIf percent < 80 Then
                                gp.StartColor = Color.FromArgb(255, 205, 86) ' Жълто
                                gp.EndColor = Color.FromArgb(255, 159, 64)   ' Оранжево
                            Else
                                gp.StartColor = Color.FromArgb(255, 99, 132) ' Червено
                                gp.EndColor = Color.FromArgb(255, 159, 64)   ' Оранжево
                            End If
                        Else
                            ' Няма съвпадение в каталога - празен прогрес бар
                            gp.Value = 0
                            gp.ShowText = False
                        End If
                    End If
                End If
            Next
        Next
        TableLayoutPanel.Refresh()
    End Sub
    Private Sub TrayProgressBar_Click(sender As Object, e As TrayClickEventArgs)
        ' ⬇️ 1. ОПРЕДЕЛИ ТИПА (Скара или Канал) ⬇️
        Dim tip As String = ""
        Dim layer As String = ""
        If RadioButton_Скара.Checked Then
            tip = "Скара"
            layer = "EL_Скари"
        ElseIf RadioButton_Канал.Checked Then
            tip = "Канал"
            layer = "EL_Канали"
        Else
            MessageBox.Show("Моля, избери начин на полагане!", "Грешка",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        ' ⬇️ 2. ВЗЕМИ AutoCAD DOCUMENT ⬇️
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        ' ⬇️ 3. СКРИЙ ФОРМАТА (AutoCAD става активен) ⬇️
        Me.Visible = False
        ' ⬇️ 4. ИЗБОР НА ТОЧКА ОТ ПОТРЕБИТЕЛЯ ⬇️
        Dim pPtOpts As New PromptPointOptions("")
        pPtOpts.Message = vbLf & $"Изберете точка на вмъкване на {tip}: "
        Dim pPtRes As PromptPointResult = acDoc.Editor.GetPoint(pPtOpts)
        ' ⬇️ 5. ПОКАЖИ ФОРМАТА ОБРАТНО ⬇️
        Me.Visible = True
        If pPtRes.Status <> PromptStatus.OK Then
            Exit Sub
        End If
        Dim InsertPoint As Point3d = pPtRes.Value
        ' ⬇️ 6. ЗАКЛЮЧИ ДОКУМЕНТА И ВМЪКНИ БЛОКА ⬇️
        Using docLock As DocumentLock = acDoc.LockDocument()
            Dim blkRecId As ObjectId = cu.InsertBlock(tip, InsertPoint, layer, New Scale3d(1, 1, 1))

            If blkRecId.IsNull Then
                MessageBox.Show($"Блокът '{tip}' не съществува в чертежа!", "Грешка",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            ' ⬇️ 7. ЗАДАЙ ДИНАМИЧНИТЕ ПАРАМЕТРИ ⬇️
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Try
                    Dim acBlkRef As BlockReference = DirectCast(acTrans.GetObject(blkRecId, OpenMode.ForWrite), BlockReference)
                    Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection
                    Dim attCol = acBlkRef.AttributeCollection
                    ' ⬇️ Задай "ШИРИНА" и "ВИСОЧИНА" ⬇️
                    For Each prop As DynamicBlockReferenceProperty In props
                        If prop.PropertyName = "Ширина" Then prop.Value = CDbl(e.Width / 10)
                        If prop.PropertyName = "Височина" Then prop.Value = CDbl(e.Height / 10)
                        If prop.PropertyName = "Размер" Then prop.Value = $"{e.Width}x{e.Height}"
                    Next

                    For Each objID As ObjectId In attCol
                        Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForWrite)
                        Dim acAttRef As AttributeReference = dbObj
                        Dim ddd As String = ""
                        If acAttRef.Tag = "ШИРИНА" Then acAttRef.TextString = CDbl(e.Width)
                        If acAttRef.Tag = "ВИСОЧИНА" Then acAttRef.TextString = CDbl(e.Height)
                    Next
                    acTrans.Commit()
                Catch ex As Exception
                    acTrans.Abort()
                    MessageBox.Show($"Грешка: {ex.Message}", "Грешка", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

            End Using
        End Using
    End Sub
    ''' <summary>
    ''' Анимира всички GradientProgressBar контроли в TableLayoutPanel
    ''' според процента на запълване на съответните скари от подадения каталог.
    '''
    ''' Процедурата изпълнява:
    '''
    ''' 1) Нулиране на всички ProgressBar контроли
    '''    - Value = 0
    '''    - ShowText = False
    '''
    ''' 2) Изчисляване на крайни стойности
    '''    - За всяка скари в каталога намира съответния ProgressBar
    '''      по име (формат: име_ширина_височина)
    '''    - Определя крайна стойност (0–100%)
    '''    - Определя цветове според процента спрямо избрания праг:
    '''
    '''        🟡 Под оптимума  → има резерв
    '''        🟢 В оптимален диапазон
    '''        🔴 Над оптимума → препълнена скара
    '''
    '''    Диапазонът се определя чрез:
    '''        fillPercent ± 10%
    '''
    ''' 3) Плавна анимация
    '''    - Пълни баровете постепенно за зададената продължителност
    '''    - Показва текста при последната стъпка
    '''
    ''' 4) Финализация
    '''    - Задава точните крайни стойности
    '''    - Прилага изчислените градиентни цветове
    '''    - Освежава TableLayoutPanel
    '''
    ''' Цел:
    ''' Да визуализира състоянието на запълване на всяка възможна скара
    ''' с ясна цветова индикация за натоварване.
    ''' </summary>
    '''
    ''' <param name="Catalog">
    ''' Списък от обекти тип Скара, съдържащи ширина, височина и процент запълване.
    ''' </param>
    '''
    ''' <param name="durationMs">
    ''' Продължителност на анимацията в милисекунди.
    ''' </param>
    '''
    ''' <param name="fillPercent">
    ''' Целеви процент на оптимално запълване.
    ''' Използва се за определяне на цветови диапазон (по подразбиране 40%).
    ''' </param>
    '''
    ''' <returns>
    ''' Асинхронна Task операция.
    ''' </returns>
    Private Async Function UpdateProgressBarsAnimated(Catalog As List(Of Скара),
                                                  durationMs As Integer,
                                                  Optional fillPercent As Double = 40) As Task
        ' 1. Първо нулирай всички барове
        For row As Integer = 1 To TableLayoutPanel.RowCount - 1
            For col As Integer = 1 To TableLayoutPanel.ColumnCount - 1
                Dim ctrl = TableLayoutPanel.GetControlFromPosition(col, row)
                If TypeOf ctrl Is GradientProgressBar Then
                    DirectCast(ctrl, GradientProgressBar).Value = 0
                    DirectCast(ctrl, GradientProgressBar).ShowText = False
                End If
            Next
        Next
        ' 2. Изчисли крайните стойности
        Dim targetValues As New Dictionary(Of String, Integer)
        For row As Integer = 1 To TableLayoutPanel.RowCount - 1
            For col As Integer = 1 To TableLayoutPanel.ColumnCount - 1
                Dim ctrl = TableLayoutPanel.GetControlFromPosition(col, row)
                If TypeOf ctrl Is GradientProgressBar Then
                    Dim gp = DirectCast(ctrl, GradientProgressBar)
                    Dim parts() As String = gp.Name.Split("_"c)

                    If parts.Length >= 3 Then
                        Dim width As String = parts(1)
                        Dim height As String = parts(2)

                        Dim tray = Catalog.FirstOrDefault(Function(t) _
                                                              t.Ширина = width AndAlso t.Височина = height)
                        If tray.Ширина IsNot Nothing Then

                            Dim targetValue As Integer = CInt(Math.Min(100, Math.Max(0, tray.Процент)))
                            targetValues(gp.Name) = targetValue
                            Dim Procent As Double = 10
                            ' Прагове от ComboBox-а
                            Dim targetPercent As Double = fillPercent
                            Dim lowerBound As Double = Math.Max(0, targetPercent - Procent)   ' Не по-малко от 0
                            Dim upperBound As Double = Math.Min(100, targetPercent + Procent) ' Не повече от 100

                            Dim startColor As Color
                            Dim endColor As Color
                            Select Case tray.Процент
                                Case Is < lowerBound
                                    ' 🟡 ЖЪЛТО - под оптимума (имаш капацитет)
                                    startColor = Color.FromArgb(255, 255, 200)  ' Много бледо жълто
                                    endColor = Color.FromArgb(200, 150, 0)       ' Тъмно жълто/златисто
                                Case Is <= upperBound
                                    ' 🟢 ЗЕЛЕНО - оптимално
                                    startColor = Color.FromArgb(200, 255, 200)  ' Бледо зелено
                                    endColor = Color.FromArgb(0, 120, 0)         ' Тъмно зелено
                                Case Else
                                    ' 🔴 ЧЕРВЕНО - над оптимума (препълнено)
                                    startColor = Color.FromArgb(255, 200, 200)  ' Бледо червено
                                    endColor = Color.FromArgb(150, 0, 0)        ' Тъмно червено
                            End Select
                            gp.StartColor = startColor
                            gp.EndColor = endColor
                        End If
                    End If
                End If
            Next
        Next
        ' 3. Анимация - плавно пълнене
        Dim steps As Integer = 10
        Dim delayPerStep As Integer = durationMs \ steps
        For iStep As Integer = 1 To steps  ' 
            Dim currentPercent As Double = iStep / steps
            For Each kvp In targetValues
                Dim gpName As String = kvp.Key
                Dim targetValue As Integer = kvp.Value
                Dim animatedValue As Integer = CInt(targetValue * currentPercent)
                ' Намери контрола по име
                For row As Integer = 1 To TableLayoutPanel.RowCount - 1
                    For col As Integer = 1 To TableLayoutPanel.ColumnCount - 1
                        Dim ctrl = TableLayoutPanel.GetControlFromPosition(col, row)
                        If TypeOf ctrl Is GradientProgressBar Then
                            Dim gp = DirectCast(ctrl, GradientProgressBar)
                            If gp.Name = gpName Then
                                gp.Value = animatedValue
                                ' Показвай текста само на последната стъпка
                                If iStep = steps Then
                                    gp.ShowText = True
                                End If
                            End If
                        End If
                    Next
                Next
            Next
            ' Изчакай преди следващата стъпка
            Await Task.Delay(delayPerStep)
        Next
        ' 5. Финализирай (само стойностите и текста, цветовете вече са зададени)
        For row As Integer = 1 To TableLayoutPanel.RowCount - 1
            For col As Integer = 1 To TableLayoutPanel.ColumnCount - 1
                Dim ctrl = TableLayoutPanel.GetControlFromPosition(col, row)
                If TypeOf ctrl Is GradientProgressBar Then
                    Dim gp = DirectCast(ctrl, GradientProgressBar)
                    If targetValues.ContainsKey(gp.Name) Then
                        gp.Value = targetValues(gp.Name)
                        gp.ShowText = True
                    End If
                End If
            Next
        Next
        TableLayoutPanel.Refresh()
    End Function
End Class
Public Class TrayClickEventArgs
    Inherits EventArgs

    Public Property Width As String
    Public Property Height As String
    Public Property Percent As Double

    Public Sub New(w As String, h As String, p As Double)
        Me.Width = w
        Me.Height = h
        Me.Percent = p
    End Sub
End Class
Public Class GradientProgressBar
    Inherits Control
    Public Event ProgressBarClicked As EventHandler(Of TrayClickEventArgs)
    Private _value As Integer = 0
    Private _maximum As Integer = 100
    Private _minimum As Integer = 0
    Private _startColor As Color = Color.FromArgb(230, 230, 230)  ' Светло сиво
    Private _endColor As Color = Color.FromArgb(200, 200, 200)    ' Тъмно сиво
    Private _showText As Boolean = True
    Private _borderColor As Color = Color.Gray
    Private _width As String = ""
    Private _height As String = ""
    <Category("Data")>
    Public Property TrayWidth As String
        Get
            Return _width
        End Get
        Set(v As String)
            _width = v
        End Set
    End Property
    <Category("Data")>
    Public Property TrayHeight As String
        Get
            Return _height
        End Get
        Set(v As String)
            _height = v
        End Set
    End Property
    <Category("Behavior")>
    Public Property Value As Integer
        Get
            Return _value
        End Get
        Set(v As Integer)
            _value = Math.Max(_minimum, Math.Min(_maximum, v))
            Me.Invalidate()
        End Set
    End Property
    <Category("Behavior")>
    Public Property Maximum As Integer
        Get
            Return _maximum
        End Get
        Set(v As Integer)
            _maximum = Math.Max(1, v)
            Me.Invalidate()
        End Set
    End Property
    <Category("Behavior")>
    Public Property Minimum As Integer
        Get
            Return _minimum
        End Get
        Set(v As Integer)
            _minimum = v
            Me.Invalidate()
        End Set
    End Property
    <Category("Appearance")>
    Public Property StartColor As System.Drawing.Color
        Get
            Return _startColor
        End Get
        Set(v As System.Drawing.Color)
            _startColor = v
            Me.Invalidate()
        End Set
    End Property
    <Category("Appearance")>
    Public Property EndColor As System.Drawing.Color
        Get
            Return _endColor
        End Get
        Set(v As System.Drawing.Color)
            _endColor = v
            Me.Invalidate()
        End Set
    End Property
    <Category("Appearance")>
    Public Property ShowText As Boolean
        Get
            Return _showText
        End Get
        Set(v As Boolean)
            _showText = v
            Me.Invalidate()
        End Set
    End Property
    <Category("Appearance")>
    Public Property BorderColor As System.Drawing.Color
        Get
            Return _borderColor
        End Get
        Set(v As System.Drawing.Color)
            _borderColor = v
            Me.Invalidate()
        End Set
    End Property
    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)

        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias

        Dim percent As Single = CSng(_value - _minimum) / CSng(_maximum - _minimum)
        Dim progressWidth As Integer = CInt(Me.Width * percent)

        ' Фон - с пълни namespaces
        Dim backBrush As SolidBrush = New SolidBrush(System.Drawing.Color.FromArgb(240, 240, 240))
        e.Graphics.FillRectangle(backBrush, 0, 0, Me.Width, Me.Height)
        backBrush.Dispose()

        ' Градиентен прогрес
        If progressWidth > 0 Then
            Dim gradientBrush As New System.Drawing.Drawing2D.LinearGradientBrush(
                New Rectangle(0, 0, progressWidth, Me.Height),
                _startColor, _endColor,
                LinearGradientMode.Horizontal)

            Dim path As New GraphicsPath()
            Dim radius As Integer = 3

            If progressWidth > radius * 2 Then
                path.AddArc(0, 0, radius * 2, radius * 2, 180, 90)
                path.AddLine(radius, 0, progressWidth - radius, 0)
                path.AddArc(progressWidth - radius * 2, 0, radius * 2, radius * 2, 270, 90)
                path.AddLine(progressWidth, radius, progressWidth, Me.Height - radius)
                path.AddArc(progressWidth - radius * 2, Me.Height - radius * 2, radius * 2, radius * 2, 0, 90)
                path.AddLine(progressWidth - radius, Me.Height, radius, Me.Height)
                path.AddArc(0, Me.Height - radius * 2, radius * 2, radius * 2, 90, 90)
                path.CloseFigure()
            Else
                path.AddRectangle(New System.Drawing.Rectangle(0, 0, progressWidth, Me.Height))
            End If

            e.Graphics.FillPath(gradientBrush, path)

            Dim pen As New System.Drawing.Pen(_borderColor, 1)
            e.Graphics.DrawPath(pen, path)
            pen.Dispose()
            gradientBrush.Dispose()
            path.Dispose()
        End If

        ' Текст - НАПЪЛНО КВАЛИФИЦИРАН
        If _showText AndAlso Me.Height > 15 Then
            Dim text As String = String.Format("{0}%", CInt(percent * 100))

            Using font As New System.Drawing.Font("Segoe UI", 12.0F, System.Drawing.FontStyle.Bold)
                Dim textSize As System.Drawing.SizeF = e.Graphics.MeasureString(text, font)
                Dim textX As Single = (Me.Width - textSize.Width) / 2
                Dim textY As Single = (Me.Height - textSize.Height) / 2

                '' Shadow за по-добра четимост
                'Using shadowBrush As New System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(100, System.Drawing.Color.Black))
                '    e.Graphics.DrawString(text, font, shadowBrush, textX + 1, textY + 1)
                'End Using

                ' ⬇️ ТУК Е ПРОМЯНАТА ⬇️
                Using textBrush As New System.Drawing.SolidBrush(System.Drawing.Color.Black)  ' Бяло → Черно
                    e.Graphics.DrawString(text, font, textBrush, textX, textY)
                End Using
            End Using
        End If
        ' Външен border
        Dim borderPen As New System.Drawing.Pen(System.Drawing.Color.Gray, 1)
        e.Graphics.DrawRectangle(borderPen, 0, 0, Me.Width - 1, Me.Height - 1)
        borderPen.Dispose()
    End Sub
    Protected Overrides Sub OnResize(e As System.EventArgs)
        MyBase.OnResize(e)
        Me.SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer, True)
    End Sub
    Protected Overrides Sub OnClick(e As EventArgs)
        MyBase.OnClick(e)

        ' Вдигни събитието с данните
        RaiseEvent ProgressBarClicked(Me, New TrayClickEventArgs(_width, _height, CSng(_value)))
    End Sub
    Protected Overrides Sub OnMouseClick(e As MouseEventArgs)
        MyBase.OnMouseClick(e)
        Me.Invalidate()  ' Опционално: визуален反馈 при клик
    End Sub
End Class