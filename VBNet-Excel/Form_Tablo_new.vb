Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Internal.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Imports Autodesk.AutoCAD.PlottingServices
Imports System.Collections.Generic
Imports Autodesk.AutoCAD.DatabaseServices.Filters
Imports System.Drawing
'Imports System.IO
'Imports System.Windows.Forms

Public Class Form_Tablo_new
    Dim PI As Double = 3.1415926535897931
    Dim cu As CommonUtil = New CommonUtil()
    Dim brTabla As Integer = 20
    Dim brTokKryg As Integer = 100
    Dim brKonsumator As Integer = 150
    Dim brKonsum As Integer = 10000
    'Dim form_AS_tablo As New Form_Tablo()
    Dim appNameKonso As String = "EWG_KONSO"
    Dim appNameTablo As String = "EWG_TABLO"
    Dim za6t As Integer = 14
    Dim defkt As Integer = 8

    Dim arrKonsum(brKonsum) As ObjectId
    Dim arrTablo(brTabla) As strTablo
    'Dim tablo As New List(Of strTablo)
    Public Structure strTokow
        Dim CountKonsumator As Integer  ' Брой консуматри в токовия кръг
        Dim Tablo As String             ' Табло към което е включен токовия кръг
        Dim ТоковКръг As String         ' Номер на токов кръг
        Dim brLamp As Integer           ' Брой лампи
        Dim brKontakt As Integer        ' Брой контакти
        Dim Мощност As Double           ' Мощност на токов кръг - в kW
        Dim Kabebel_Se4enie As String   ' Сечение на кабела
        Dim faza As String              ' Фаза
        Dim konsuator1 As String
        Dim konsuator2 As String
        'Изчислителни полета
        '
        '
        Dim BrojPol As String           ' Брой на полюсите
        ' Ток на токовия кръг
        ' За трифазен консуматор
        ' .Мощност * 1.2 / (0.38 * Math.Sqrt(3) * 0.9)
        '
        ' За монофазен консуматор
        ' .Мощност * 1.2 / (0.22 * 0.9)
        '
        Dim Tok As Double               ' Ток на токовия кръг
        ' Полета за защита
        '
        '
        Dim BlockName As String         ' Име на блок който се вмъква
        Dim Designation As String
        Dim ShortName As String         ' Кратко име - вид на апарата
        Dim Type As String              ' 
        Dim NumberPoles As String       ' Брой на модулите / 
        Dim RatedCurrent As String      ' Номинален ток
        Dim Curve As String             ' Крива
        Dim Current As String
        Dim Control As String
        Dim Sensitivity As String       ' Изключвателна възможност
        Dim Protection As String
        '
        ' Консуатори
        '
        Dim Konsumator() As strKonsumator
    End Structure
    Public Structure strTablo
        Dim countTablo As Integer       ' Брой на таблата
        Dim Name As String              ' Име на таблото
        Dim prevTablo As String         ' Име на предходното табло
        Dim countTokKryg As Integer     ' Брой токови кръгове ?????
        Dim Tokowkryg() As strTokow     ' Токов кръг
        Dim TabloType As String         ' тип на таблото
        ' Ще трябва да се добавят полета за рамери, тип и други за таблото!!!!
        '
        '
    End Structure
    Public Structure strKonsumator
        Dim Name As String              ' Име на блока
        Dim ID_Block As ObjectId        ' Блок на елемента
        Dim ТоковКръг As String         ' Токов кръг към който е свързан
        Dim strМОЩНОСТ As String        ' Мощност от блока
        Dim doubМОЩНОСТ As Double       ' Изчислена мощност
        Dim ТАБЛО As String             ' Табло към което е включен токовия кръг
        Dim Pewdn As String             ' Предназначение 
        Dim PEWDN1 As String            ' Предназначение 
        Dim Dylvina_Led As Double       ' Дължина на LED лента
        Dim Visibility As String        ' 
    End Structure
    Private Sub Form_Tablo_new_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Height = 1000
        Me.Width = 1600
        DataGridView.Visible = False
    End Sub
    Private Sub GetObjects(index As Integer)
        Me.Visible = False
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim edt As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database
        Dim SelectedSet = cu.GetObjects("INSERT", "Изберете блок")
        Me.Visible = True
        If SelectedSet Is Nothing Then
            MsgBox("НЕ Е маркиран нито един блок.")
            Exit Sub
        End If
        Dim blkRecId As ObjectId = ObjectId.Null
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                Dim acBlkTbl As BlockTable
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                ToolStripProgressBar1.Maximum = SelectedSet.Count
                ToolStripProgressBar1.Value = 0
                For Each sObj As SelectedObject In SelectedSet
                    blkRecId = sObj.ObjectId
                    Dim Yes As Boolean = False
                    For ind_arr = 0 To index
                        If arrKonsum(ind_arr) = sObj.ObjectId Then
                            Yes = True
                            Exit For
                        End If
                    Next
                    If Yes Then Continue For
                    arrKonsum(index) = sObj.ObjectId
                    index += 1
                    ToolStripProgressBar1.Value += 1
                Next
                acTrans.Commit()
            Catch ex As Exception
                MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
                acTrans.Abort()
            End Try
        End Using
    End Sub
    Private Sub NewToolStripButton_Click(sender As Object, e As EventArgs) Handles NewToolStripButton.Click
        'tablo.Clear()
        ReDim arrKonsum(brKonsum)

        GetObjects(0)

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim edt As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database

        Dim brTablo As Integer = 0
        ToolStripProgressBar1.Minimum = 0
        ToolStripProgressBar1.Value = 0
        Dim blkRecId As ObjectId = ObjectId.Null

        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                Dim acBlkTbl As BlockTable
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim index As Integer = 0

                ToolStripProgressBar1.Maximum = arrKonsum.Count

                Dim index_Tablo As Integer = 0

                For Each sObj As ObjectId In arrKonsum

                    Dim _strTablo As New strTablo
                    Dim _strTokow As New strTokow
                    Dim _strKonsumator As New strKonsumator

                    _strKonsumator.ID_Block = sObj

                    ToolStripProgressBar1.Value += 1
                    If sObj.IsNull Then Exit For
                    Dim acBlkRef As BlockReference = DirectCast(acTrans.GetObject(sObj, OpenMode.ForRead), BlockReference)
                    Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                    Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection

                    _strKonsumator.ТоковКръг = ""
                    _strKonsumator.strМОЩНОСТ = ""
                    _strKonsumator.doubМОЩНОСТ = 0.0
                    _strKonsumator.ТАБЛО = ""
                    _strKonsumator.Pewdn = ""
                    _strKonsumator.PEWDN1 = ""
                    _strKonsumator.Dylvina_Led = 0.0

                    For Each objID As ObjectId In attCol
                        Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                        Dim acAttRef As AttributeReference = dbObj
                        If acAttRef.Tag = "МОЩНОСТ" Then _strKonsumator.strМОЩНОСТ = acAttRef.TextString
                        If acAttRef.Tag = "LED" Then _strKonsumator.strМОЩНОСТ = acAttRef.TextString
                        If acAttRef.Tag = "КРЪГ" Then _strKonsumator.ТоковКръг = acAttRef.TextString
                        If acAttRef.Tag = "ТАБЛО" Then _strKonsumator.ТАБЛО = acAttRef.TextString
                        If acAttRef.Tag = "Pewdn" Then _strKonsumator.Pewdn = acAttRef.TextString
                        If acAttRef.Tag = "PEWDN1" Then _strKonsumator.PEWDN1 = acAttRef.TextString
                    Next

                    If _strKonsumator.strМОЩНОСТ = "" Then Continue For

                    Dim Visibility As String = ""
                    For Each prop As DynamicBlockReferenceProperty In props
                        If prop.PropertyName = "Visibility1" Then _strKonsumator.Visibility = prop.Value
                        If prop.PropertyName = "Visibility" Then _strKonsumator.Visibility = prop.Value
                        If prop.PropertyName = "Дължина" Then _strKonsumator.Dylvina_Led = prop.Value
                    Next

                    _strKonsumator.Name = (CType(acBlkRef.DynamicBlockTableRecord.GetObject(OpenMode.ForRead), BlockTableRecord)).Name

                    If _strKonsumator.Dylvina_Led <> 0 Then
                        _strKonsumator.doubМОЩНОСТ = Val(_strKonsumator.Dylvina_Led) * 14 / 100
                    Else
                        Dim brМОЩНОСТ, moМОЩНОСТ As String
                        Dim poz As Integer = Math.Max(InStr(_strKonsumator.strМОЩНОСТ, "х"), InStr(_strKonsumator.strМОЩНОСТ, "x"))
                        If poz > 0 Then
                            brМОЩНОСТ = Mid(_strKonsumator.strМОЩНОСТ, 1, poz - 1)
                            moМОЩНОСТ = Mid(_strKonsumator.strМОЩНОСТ, poz + 1, Len(_strKonsumator.strМОЩНОСТ))
                            _strKonsumator.doubМОЩНОСТ = Val(brМОЩНОСТ) * Val(moМОЩНОСТ)
                        Else
                            _strKonsumator.doubМОЩНОСТ = Val(_strKonsumator.strМОЩНОСТ)
                        End If
                    End If

                    _strKonsumator.doubМОЩНОСТ = _strKonsumator.doubМОЩНОСТ / 1000

                    If _strKonsumator.ТАБЛО = "" Or _strKonsumator.ТАБЛО = "Табло" Then _strKonsumator.ТАБЛО = "Гл.Р.Т."

                    _strTokow.Tablo = _strKonsumator.ТАБЛО

                    Dim iTablo = Array.FindIndex(arrTablo, Function(f) f.Name = _strKonsumator.ТАБЛО)

                    If iTablo = -1 Then
                        arrTablo(brTablo).Name = _strKonsumator.ТАБЛО
                        iTablo = brTablo
                        ReDim arrTablo(brTablo).Tokowkryg(brTokKryg)
                        arrTablo(brTablo).countTokKryg = 0
                        brTablo += 1
                    End If

                    Dim iKryg As Integer = Array.FindIndex(arrTablo(iTablo).Tokowkryg, Function(f) f.ТоковКръг = _strKonsumator.ТоковКръг)

                    If iKryg = -1 Then
                        ReDim arrTablo(iTablo).Tokowkryg(arrTablo(iTablo).countTokKryg).Konsumator(brKonsumator)
                        iKryg = arrTablo(iTablo).countTokKryg
                        arrTablo(iTablo).countTokKryg += 1
                    End If

                    With arrTablo(iTablo).Tokowkryg(iKryg)
                        .ТоковКръг = _strKonsumator.ТоковКръг
                        .konsuator1 = _strKonsumator.PEWDN1
                        .konsuator2 = _strKonsumator.Pewdn
                        With .Konsumator(.CountKonsumator)
                            .ID_Block = _strKonsumator.ID_Block
                            .Name = _strKonsumator.Name                     ' Име на блока
                            .ID_Block = _strKonsumator.ID_Block             ' Блок на елемента
                            .ТоковКръг = _strKonsumator.ТоковКръг           ' Токов кръг към който е свързан
                            .strМОЩНОСТ = _strKonsumator.strМОЩНОСТ         ' Мощност от блока
                            .doubМОЩНОСТ = _strKonsumator.doubМОЩНОСТ       ' Изчислена мощност
                            .ТАБЛО = _strKonsumator.ТАБЛО                   ' Табло към което е включен токовия кръг
                            .Pewdn = _strKonsumator.Pewdn                   ' Предназначение 
                            .PEWDN1 = _strKonsumator.PEWDN1                 ' Предназначение 
                            .Dylvina_Led = _strKonsumator.Dylvina_Led       ' Дължина на LED лента
                            .Visibility = _strKonsumator.Visibility
                        End With
                        .CountKonsumator += 1
                    End With
                Next
                arrTablo.Count
                acTrans.Commit()
            Catch ex As Exception
                MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
                acTrans.Abort()
            End Try
        End Using
        ToolStripProgressBar1.Value = 0
    End Sub
    'Private Sub Form_Tablo_new_Closed(sender As Object, e As EventArgs) Handles Me.Closed
    '    TreeView.Nodes.Clear()
    '    Me.Close()
    'End Sub


    'Private Sub insDataGrid(name As String, colArray As strTablo)
    '    If colArray.countTokKryg = 0 Then Exit Sub

    '    Dim dagrid As System.Windows.Forms.DataGridView = New Windows.Forms.DataGridView
    '    Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
    '    Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
    '    Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
    '    With DataGridViewCellStyle1
    '        .Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
    '        .BackColor = System.Drawing.SystemColors.ControlDark
    '        .Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
    '        .ForeColor = System.Drawing.SystemColors.WindowText
    '        .SelectionBackColor = System.Drawing.SystemColors.Highlight
    '        .SelectionForeColor = System.Drawing.SystemColors.HighlightText
    '        .WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
    '    End With
    '    With DataGridViewCellStyle2
    '        .BackColor = System.Drawing.Color.Silver
    '        .ForeColor = System.Drawing.Color.Black
    '        .Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
    '        .Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
    '        .SelectionBackColor = System.Drawing.SystemColors.Highlight
    '        .SelectionForeColor = System.Drawing.SystemColors.HighlightText
    '        .WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
    '    End With
    '    With DataGridViewCellStyle3
    '        .Format = "N2"
    '        .NullValue = Nothing
    '    End With

    '    DataGridView.Rows.Clear()

    '    With DataGridView
    '        '.Name = name
    '        .BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
    '        .ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
    '        .RowHeadersDefaultCellStyle = DataGridViewCellStyle2
    '        .Size = New System.Drawing.Size(432, 450)
    '        .Dock = System.Windows.Forms.DockStyle.Fill
    '        .ColumnCount = 2 + colArray.countTokKryg + 1
    '        .ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
    '        With .Columns(0)
    '            .Width = 100
    '            .HeaderText = "Параметър"
    '            .Name = "Параметър"
    '            .Frozen = vbTrue
    '            .ReadOnly = vbTrue
    '            .SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
    '            .DefaultCellStyle = DataGridViewCellStyle3
    '        End With
    '        With .Columns(1)
    '            .DefaultCellStyle = DataGridViewCellStyle3
    '            .Width = 40
    '            .HeaderText = "Дим."
    '            .Name = "Дим."
    '            .Frozen = vbTrue
    '            .ReadOnly = vbTrue
    '            .SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
    '        End With

    '        With .Rows
    '            .Add({"Автоматичен прекъсвач", " "})
    '            .Add({"Изчислен ток", "А"})
    '            .Add({"Тип на апарата", " "})
    '            .Add({"Номинален ток", "А"})
    '            .Add({"Изкл. възможност", " "})
    '            .Add({"Крива", " "})
    '            .Add({"Брой полюси", "бр."})
    '            .Add({"-----------", "---"})
    '            .Add({"ДТЗ", " "})
    '            .Add({"Вид на апарата", " "})
    '            .Add({"Клас на апарата", " "})
    '            .Add({"Номинален ток", "А"})
    '            .Add({"Изкл. възможност", "mA"})
    '            .Add({"Брой полюси", "бр."})
    '            .Add({"-----------", "---"})
    '            .Add({"Брой лампи", "бр."})
    '            .Add({"Брой контакти", "бр."})
    '            .Add({"Инст. мощност", "kW"})
    '            .Add({"Тип кабел", "---"})
    '            .Add({"Сечение", "---"})
    '            .Add({"Фаза", "---"})
    '            .Add({"Консуматор", "---"})
    '            .Add({"", "---"})
    '        End With

    '        Dim Мощност As Double = 0.0
    '        Dim brLamp As Integer = 0
    '        Dim brKontakt As Integer = 0
    '        Dim brCol As Integer = 0
    '        Dim ИзлазКонтакти As Integer = 0
    '        Dim ИзлазТок As Double = 0
    '        Dim ИзлазТрифази As Boolean = False
    '        Dim ТаблоТрифази As Boolean = False

    '        Dim j As Integer = 0
    '        For i = 0 To colArray.countTokKryg
    '            With .Columns(2 + brCol)
    '                .Width = 100
    '                .HeaderText = colArray.Tokowkryg(i).ТоковКръг
    '                .Name = colArray.Tokowkryg(i).ТоковКръг
    '                .SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
    '            End With
    '            If colArray.Tokowkryg(i).faza <> "L1" And colArray.Tokowkryg(i).faza <> Nothing Then
    '                ТаблоТрифази = True
    '            End If
    '            .Rows(1).Cells(2 + brCol).Value = colArray.Tokowkryg(i).Tok.ToString("0.00")
    '            .Rows(2).Cells(2 + brCol).Value = "EZ9 MCB"
    '            .Rows(3).Cells(2 + brCol).Value = colArray.Tokowkryg(i).RatedCurrent & "A"
    '            .Rows(4).Cells(2 + brCol).Value = colArray.Tokowkryg(i).Sensitivity
    '            .Rows(5).Cells(2 + brCol).Value = colArray.Tokowkryg(i).Curve
    '            .Rows(6).Cells(2 + brCol).Value = colArray.Tokowkryg(i).NumberPoles

    '            If colArray.Tokowkryg(i).konsuator1 = "Контакти" Then
    '                ИзлазКонтакти += 1
    '                If colArray.Tokowkryg(i).faza <> "L1" And colArray.Tokowkryg(i).faza <> Nothing Then
    '                    ИзлазТок += colArray.Tokowkryg(i).Мощност * 1.2 / (0.38 * Math.Sqrt(3) * 0.9)
    '                Else
    '                    ИзлазТок += colArray.Tokowkryg(i).Мощност * 1.2 / (0.22 * 0.9)
    '                End If
    '                If colArray.Tokowkryg(i).faza <> "L1" And colArray.Tokowkryg(i).faza <> Nothing Then
    '                    ИзлазТрифази = True
    '                End If
    '                If ИзлазКонтакти = 3 Then
    '                    .Rows(defkt + 1).Cells(2 + brCol).Value = "EZ9 RCCB"
    '                    .Rows(defkt + 2).Cells(2 + brCol).Value = "AC"
    '                    Select Case ИзлазТок
    '                        Case < 25
    '                            .Rows(defkt + 3).Cells(2 + brCol).Value = "25А"
    '                        Case < 40
    '                            .Rows(defkt + 3).Cells(2 + brCol).Value = "40А"
    '                        Case < 63
    '                            .Rows(defkt + 3).Cells(2 + brCol).Value = "63А"
    '                        Case Else
    '                            .Rows(defkt + 3).Cells(2 + brCol).Value = "#####"
    '                    End Select
    '                    .Rows(defkt + 4).Cells(2 + brCol).Value = "30mA"
    '                    .Rows(defkt + 5).Cells(2 + brCol).Value = IIf(ИзлазТрифази, "4p", "2p")
    '                    ИзлазТрифази = False
    '                    ИзлазКонтакти = 0
    '                    ИзлазТок = 0
    '                End If
    '            End If
    '            If colArray.Tokowkryg(i).konsuator1 = "Бойлер" Then
    '                .Rows(defkt + 1).Cells(2 + brCol).Value = "EZ9 RCBO"
    '                .Rows(defkt + 2).Cells(2 + brCol).Value = "AC"
    '                .Rows(defkt + 3).Cells(2 + brCol).Value = "25А"
    '                .Rows(defkt + 4).Cells(2 + brCol).Value = "30mA"
    '                .Rows(defkt + 5).Cells(2 + brCol).Value = "2p"
    '            End If
    '            .Rows(za6t + 1).Cells(2 + brCol).Value = colArray.Tokowkryg(i).brLamp
    '            .Rows(za6t + 2).Cells(2 + brCol).Value = colArray.Tokowkryg(i).brKontakt
    '            .Rows(za6t + 3).Cells(2 + brCol).Value = colArray.Tokowkryg(i).Мощност.ToString("0.000")
    '            .Rows(za6t + 4).Cells(2 + brCol).Value = "СВТ"
    '            .Rows(za6t + 5).Cells(2 + brCol).Value = colArray.Tokowkryg(i).Kabebel_Se4enie
    '            .Rows(za6t + 6).Cells(2 + brCol).Value = colArray.Tokowkryg(i).faza
    '            .Rows(za6t + 7).Cells(2 + brCol).Value = colArray.Tokowkryg(i).konsuator1
    '            .Rows(za6t + 8).Cells(2 + brCol).Value = colArray.Tokowkryg(i).konsuator2
    '            brLamp += colArray.Tokowkryg(i).brLamp
    '            brKontakt += colArray.Tokowkryg(i).brKontakt
    '            Мощност += colArray.Tokowkryg(i).Мощност
    '            brCol += 1
    '        Next
    '        With .Columns(1 + brCol)
    '            .DefaultCellStyle = DataGridViewCellStyle3
    '            .Width = 75
    '            .HeaderText = "ОБЩО"
    '            .Name = "ОБЩО"
    '            .SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
    '        End With

    '        .Rows(za6t + 1).Cells(.ColumnCount - 1).Value = brLamp
    '        .Rows(za6t + 2).Cells(.ColumnCount - 1).Value = brKontakt
    '        .Rows(za6t + 3).Cells(.ColumnCount - 1).Value = Мощност
    '        .Rows(za6t + 4).Cells(1 + brCol).Value = "СВТ"

    '        If ТаблоТрифази Then
    '            .Rows(1).Cells(.ColumnCount - 1).Value = Мощност * 1.2 / (0.38 * Math.Sqrt(3) * 0.9)
    '        Else
    '            .Rows(1).Cells(.ColumnCount - 1).Value = Мощност * 1.2 / (0.22 * 0.9)
    '        End If
    '        .Rows(2).Cells(.ColumnCount - 1).Value = "iSW"
    '        Dim Общо_Сечение As String = ""
    '        Dim Общо_Ном_Ток As String = ""
    '        Select Case .Rows(1).Cells(.ColumnCount - 1).Value
    '            Case < 6        ' АП - 6А
    '                Общо_Сечение = IIf(ТаблоТрифази, "5x", "3x") & "1,5"
    '                Общо_Ном_Ток = "6"
    '            Case < 10       ' АП - 10А
    '                Общо_Сечение = IIf(ТаблоТрифази, "5x", "3x") & "1,5"
    '                Общо_Ном_Ток = "10"
    '            Case < 16       ' АП - 16А
    '                Общо_Сечение = IIf(ТаблоТрифази, "5x", "3x") & "2,5"
    '                Общо_Ном_Ток = "16"
    '            Case < 20       ' АП - 20А
    '                Общо_Сечение = IIf(ТаблоТрифази, "5x", "3x") & "2,5"
    '                Общо_Ном_Ток = "20"
    '            Case < 25       ' АП - 25А
    '                Общо_Сечение = IIf(ТаблоТрифази, "5x", "3x") & "4,0"
    '                Общо_Ном_Ток = "25"
    '            Case < 32       ' АП - 32А
    '                Общо_Сечение = IIf(ТаблоТрифази, "5x", "3x") & "6,0"
    '                Общо_Ном_Ток = "32"
    '            Case < 40       ' АП - 40А
    '                Общо_Сечение = IIf(ТаблоТрифази, "5x", "3x") & "10,0"
    '                Общо_Ном_Ток = "40"
    '            Case < 50       ' АП - 50А
    '                Общо_Сечение = IIf(ТаблоТрифази, "5x", "3x") & "10,0"
    '                Общо_Ном_Ток = "50"
    '            Case < 63       ' АП - 63А
    '                Общо_Сечение = IIf(ТаблоТрифази, "5x", "3x") & "16,0"
    '                Общо_Ном_Ток = "63"
    '            Case < 80       ' АП - 80А
    '                Общо_Сечение = IIf(ТаблоТрифази, "5x", "3x") & "25,0"
    '                Общо_Ном_Ток = "80"
    '            Case < 100      ' АП - 100А
    '                Общо_Сечение = IIf(ТаблоТрифази, "5x", "3x") & "35,0"
    '                Общо_Ном_Ток = "100"
    '            Case < 125      ' АП - 125А
    '                Общо_Сечение = IIf(ТаблоТрифази, "5x", "3x") & "50,0"
    '                Общо_Ном_Ток = "125"
    '            Case Else
    '                Общо_Сечение = "######"
    '        End Select

    '        .Rows(3).Cells(.ColumnCount - 1).Value = Общо_Ном_Ток & "A"
    '        .Rows(4).Cells(.ColumnCount - 1).Value = "-"
    '        .Rows(5).Cells(.ColumnCount - 1).Value = "-"
    '        .Rows(6).Cells(.ColumnCount - 1).Value = IIf(ТаблоТрифази, "3p", "1p")

    '        .Rows(za6t + 5).Cells(1 + brCol).Value = Общо_Сечение
    '        .Rows(za6t + 6).Cells(1 + brCol).Value = IIf(ТаблоТрифази, "L1,L2,L3", "L1")
    '        .Rows(za6t + 7).Cells(1 + brCol).Value = "---"
    '        .Rows(za6t + 8).Cells(1 + brCol).Value = "---"

    '        If ИзлазКонтакти = 1 Then
    '            For brCol = .ColumnCount To 2 Step -1
    '                If colArray.Tokowkryg(brCol).konsuator1 = "Контакти" Then
    '                    If .Rows(defkt + 1).Cells(brCol - 1).Value = "EZ9 RCCB" Then

    '                    Else
    '                        .Rows(defkt + 1).Cells(2 + brCol - 1).Value = ""
    '                        .Rows(defkt + 2).Cells(2 + brCol - 1).Value = ""
    '                        .Rows(defkt + 3).Cells(2 + brCol - 1).Value = ""
    '                        .Rows(defkt + 4).Cells(2 + brCol - 1).Value = ""
    '                        .Rows(defkt + 5).Cells(2 + brCol - 1).Value = ""
    '                        .Rows(defkt + 1).Cells(2 + brCol).Value = "EZ9 RCCB"
    '                        .Rows(defkt + 2).Cells(2 + brCol).Value = "AC"
    '                        Select Case ИзлазТок
    '                            Case < 25
    '                                .Rows(defkt + 3).Cells(2 + brCol).Value = "25А"
    '                            Case < 40
    '                                .Rows(defkt + 3).Cells(2 + brCol).Value = "40А"
    '                            Case < 63
    '                                .Rows(defkt + 3).Cells(2 + brCol).Value = "63А"
    '                        End Select
    '                        .Rows(defkt + 4).Cells(2 + brCol).Value = "30mA"
    '                        .Rows(defkt + 5).Cells(2 + brCol).Value = IIf(ИзлазТрифази, "4p", "2p")
    '                        Exit For
    '                    End If
    '                End If
    '            Next
    '        End If

    '        If ИзлазКонтакти = 2 Then
    '            For brCol = .ColumnCount To 2 Step -1
    '                If colArray.Tokowkryg(brCol).konsuator1 = "Контакти" Then
    '                    .Rows(defkt + 1).Cells(2 + brCol).Value = "EZ9 RCCB"
    '                    .Rows(defkt + 2).Cells(2 + brCol).Value = "AC"
    '                    Select Case ИзлазТок
    '                        Case < 25
    '                            .Rows(defkt + 3).Cells(2 + brCol).Value = "25А"
    '                        Case < 40
    '                            .Rows(defkt + 3).Cells(2 + brCol).Value = "40А"
    '                        Case < 63
    '                            .Rows(defkt + 3).Cells(2 + brCol).Value = "63А"
    '                    End Select
    '                    .Rows(defkt + 4).Cells(2 + brCol).Value = "30mA"
    '                    .Rows(defkt + 5).Cells(2 + brCol).Value = IIf(ИзлазТрифази, "4p", "2p")
    '                    Exit For
    '                End If
    '            Next
    '        End If

    '    End With
    'End Sub
    Private Sub TreeView_AfterSelect(sender As Object, e As Windows.Forms.TreeViewEventArgs) Handles TreeView.AfterSelect
        DataGridView.Visible = True
        'insDataGrid("DataGridView", arrTablo(TreeView.SelectedNode.Index))
    End Sub
    Private Sub TreeView_ADD()
        TreeView.BeginUpdate()
        TreeView.Nodes.Clear()
        For i As Integer = 0 To arrTablo.Count - 1
            If arrTablo(i).Name = Nothing Then Exit For
            TreeView.Nodes.Add(arrTablo(i).Name)
            For j As Integer = 0 To arrTablo(i).Tokowkryg.Count - 1
                If arrTablo(i).Tokowkryg(j).ТоковКръг = Nothing Then Exit For
                TreeView.Nodes(i).Nodes.Add(arrTablo(i).Tokowkryg(j).ТоковКръг)
                For k As Integer = 0 To arrTablo(i).Tokowkryg(j).Konsumator.Count - 1
                    If arrTablo(i).Tokowkryg(j).Konsumator(k).Name = Nothing Then Exit For
                    TreeView.Nodes(i).Nodes(j).Nodes.Add(
                         arrTablo(i).Tokowkryg(j).Konsumator(k).Name & " | " &
                         arrTablo(i).Tokowkryg(j).Konsumator(k).Visibility & " | " &
                         arrTablo(i).Tokowkryg(j).Konsumator(k).strМОЩНОСТ & " | " &
                         arrTablo(i).Tokowkryg(j).Konsumator(k).doubМОЩНОСТ
                        )
                Next
            Next
        Next
        TreeView.EndUpdate()
    End Sub
    Private Sub Button_Print_Click(sender As Object, e As EventArgs) Handles Button_Print.Click
        '' Get the current database and start the Transaction Manager
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ptBasePointRes As PromptPointResult
        Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
        '' Prompt for the start point

        pPtOpts.Message = vbLf & "Изберете долен ляв ъгъл на таблото: "
        ptBasePointRes = acDoc.Editor.GetPoint(pPtOpts)

        ' Exit if the user presses ESC or cancels the command
        Dim brColums As Double = DataGridView.Columns.Count
        If ptBasePointRes.Status = PromptStatus.Cancel Then Exit Sub
        Dim ptBasePoint As Point3d = ptBasePointRes.Value

        Dim TabloName As String = TreeView.SelectedNode.Text
        Dim widthColom As Double = 120
        Dim heightRow As Double = 25
        Dim widthText As Double = 140
        Dim widthTextDim As Double = 40
        Dim lengthProw As Double = 90
        Dim lengthProwBlock As Double = 0
        Dim padingText As Double = 3
        Dim widthTablo As Double = 410
        Dim heightText As Double = 12
        Dim Y_Шина As Double = 620

        Dim blkRecId_D As ObjectId = ObjectId.Null
        Dim blkRecId_L As ObjectId = ObjectId.Null
        Dim index_D As Integer = 0

        Dim Faza_Tablo As Boolean = False

        Try

            Dim arrPoint(15, 1) As Point3d
            Dim prX As Double = ptBasePoint.X + widthText + widthTextDim + (brColums - 2) * widthColom
            Dim prY As Double = ptBasePoint.X + widthText + widthTextDim + (brColums - 2) * widthColom

            arrPoint(0, 0) = New Point3d(ptBasePoint.X, ptBasePoint.Y, 0)
            arrPoint(0, 1) = New Point3d(prX, ptBasePoint.Y, 0)
            arrPoint(1, 0) = New Point3d(ptBasePoint.X, ptBasePoint.Y + 3 * heightRow, 0)
            arrPoint(1, 1) = New Point3d(prX, ptBasePoint.Y + 3 * heightRow, 0)
            arrPoint(2, 0) = New Point3d(ptBasePoint.X, ptBasePoint.Y + 4 * heightRow, 0)
            arrPoint(2, 1) = New Point3d(prX, ptBasePoint.Y + 4 * heightRow, 0)
            arrPoint(3, 0) = New Point3d(ptBasePoint.X, ptBasePoint.Y + 5 * heightRow, 0)
            arrPoint(3, 1) = New Point3d(prX, ptBasePoint.Y + 5 * heightRow, 0)
            arrPoint(4, 0) = New Point3d(ptBasePoint.X, ptBasePoint.Y + 6 * heightRow, 0)
            arrPoint(4, 1) = New Point3d(prX, ptBasePoint.Y + 6 * heightRow, 0)
            arrPoint(5, 0) = New Point3d(ptBasePoint.X, ptBasePoint.Y + 7 * heightRow, 0)
            arrPoint(5, 1) = New Point3d(prX, ptBasePoint.Y + 7 * heightRow, 0)
            arrPoint(6, 0) = New Point3d(ptBasePoint.X, ptBasePoint.Y + 8 * heightRow, 0)
            arrPoint(6, 1) = New Point3d(prX, ptBasePoint.Y + 8 * heightRow, 0)
            arrPoint(7, 0) = New Point3d(ptBasePoint.X, ptBasePoint.Y + 9 * heightRow, 0)
            arrPoint(7, 1) = New Point3d(prX, ptBasePoint.Y + 9 * heightRow, 0)
            arrPoint(8, 0) = New Point3d(ptBasePoint.X, ptBasePoint.Y + 10 * heightRow, 0)
            arrPoint(8, 1) = New Point3d(prX, ptBasePoint.Y + 10 * heightRow, 0)

            arrPoint(9, 0) = New Point3d(ptBasePoint.X, ptBasePoint.Y, 0)
            arrPoint(9, 1) = New Point3d(ptBasePoint.X, ptBasePoint.Y + 10 * heightRow, 0)
            arrPoint(10, 0) = New Point3d(ptBasePoint.X + widthText, ptBasePoint.Y, 0)
            arrPoint(10, 1) = New Point3d(ptBasePoint.X + widthText, ptBasePoint.Y + 10 * heightRow, 0)
            arrPoint(11, 0) = New Point3d(ptBasePoint.X + widthText + widthTextDim, ptBasePoint.Y, 0)
            arrPoint(11, 1) = New Point3d(ptBasePoint.X + widthText + widthTextDim,
                                          ptBasePoint.Y + 10 * heightRow, 0)

            arrPoint(12, 0) = New Point3d(ptBasePoint.X + widthText,
                                          ptBasePoint.Y + 10 * heightRow + lengthProw, 0)
            arrPoint(12, 1) = New Point3d(ptBasePoint.X + widthText,
                                          ptBasePoint.Y + 10 * heightRow + lengthProw + widthTablo, 0)
            arrPoint(13, 0) = New Point3d(ptBasePoint.X + widthText,
                                          ptBasePoint.Y + 10 * heightRow + lengthProw, 0)
            arrPoint(13, 1) = New Point3d(prX, ptBasePoint.Y + 10 * heightRow + lengthProw, 0)
            arrPoint(14, 0) = New Point3d(prX, ptBasePoint.Y + 10 * heightRow + lengthProw, 0)
            arrPoint(14, 1) = New Point3d(prX, ptBasePoint.Y + 10 * heightRow + lengthProw + widthTablo, 0)
            arrPoint(15, 0) = New Point3d(ptBasePoint.X + widthText,
                                          ptBasePoint.Y + 10 * heightRow + lengthProw + widthTablo, 0)
            arrPoint(15, 1) = New Point3d(prX,
                                          ptBasePoint.Y + 10 * heightRow + lengthProw + widthTablo, 0)

            Dim X As Double = 0

            For index As Integer = 0 To UBound(arrPoint)
                If index < 12 Then
                    cu.DrowLine(arrPoint(index, 0), arrPoint(index, 1), "EL_ТАБЛА", LineWeight.ByLayer, "ByLayer")
                Else
                    cu.DrowLine(arrPoint(index, 0), arrPoint(index, 1), "EL_ТАБЛА", LineWeight.ByLayer, "CENTER")
                End If

            Next
            For index As Integer = 1 To brColums - 2
                X = ptBasePoint.X + widthText + widthTextDim + index * widthColom
                cu.DrowLine(New Point3d(X, ptBasePoint.Y, 0),
                                  New Point3d(X, ptBasePoint.Y + 10 * heightRow, 0),
                                  "EL_ТАБЛА",
                                  LineWeight.ByLayer,
                                  "ByLayer")
            Next

            prX = ptBasePoint.X + padingText
            prY = ptBasePoint.Y + (heightRow - heightText) / 2

            cu.InsertText("Токов кръг", New Point3d(prX, prY + 9 * heightRow, 0),
                                 "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("Брой лампи", New Point3d(prX, prY + 8 * heightRow, 0),
                                 "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("Брой контакти", New Point3d(prX, prY + 7 * heightRow, 0),
                                 "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("Инстал. мощност", New Point3d(prX, prY + 6 * heightRow, 0),
                                 "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("Тип кабел", New Point3d(prX, prY + 5 * heightRow, 0),
                                 "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("Сечение кабел", New Point3d(prX, prY + 4 * heightRow, 0),
                                 "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("Фаза", New Point3d(prX, prY + 3 * heightRow, 0),
                                 "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("Консуматор", New Point3d(prX, prY + 2 * heightRow, 0),
                                 "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)

            prX = prX + widthText
            prY = ptBasePoint.Y + (heightRow - heightText) / 2

            cu.InsertText("№", New Point3d(prX, prY + 9 * heightRow, 0),
                                 "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("бр.", New Point3d(prX, prY + 8 * heightRow, 0),
                                 "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("бр.", New Point3d(prX, prY + 7 * heightRow, 0),
                                 "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("kW", New Point3d(prX, prY + 6 * heightRow, 0),
                                 "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("---", New Point3d(prX, prY + 5 * heightRow, 0),
                                 "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("mm²", New Point3d(prX, prY + 4 * heightRow, 0),
                                 "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("---", New Point3d(prX, prY + 3 * heightRow, 0),
                                 "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("---", New Point3d(prX, prY + 2 * heightRow, 0),
                                 "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            Dim krygkontakt As Integer = 0
            Dim Broj_N As Integer = 1
            For index As Integer = 2 To brColums - 1
                Dim ТоковКръг As String = DataGridView.Columns(index).HeaderText.ToString
                Dim brLap As String = IIf(DataGridView.Rows(za6t + 1).Cells(index).Value = 0,
                                          "----",
                                          DataGridView.Rows(za6t + 1).Cells(index).Value.ToString)
                Dim brKontakt As String = IIf(DataGridView.Rows(za6t + 2).Cells(index).Value = 0,
                                          "----",
                                          DataGridView.Rows(za6t + 2).Cells(index).Value.ToString)
                Dim Мощност As String = DataGridView.Rows(za6t + 3).Cells(index).Value.ToString
                Dim typeKabel As String = DataGridView.Rows(za6t + 4).Cells(index).Value.ToString
                Dim sechKabel As String = DataGridView.Rows(za6t + 5).Cells(index).Value.ToString
                Dim Faza As String = DataGridView.Rows(za6t + 6).Cells(index).Value.ToString
                Dim konsuator1 As String = DataGridView.Rows(za6t + 7).Cells(index).Value
                Dim konsuator2 As String = DataGridView.Rows(za6t + 8).Cells(index).Value

                X = ptBasePoint.X + widthText + widthTextDim + (index - 2) * widthColom + widthColom / 2
                cu.InsertText(ТоковКръг,
                              New Point3d(X + padingText,
                                          ptBasePoint.Y + 9 * heightRow + heightRow / 2, 0),
                              "EL__DIM", heightText, TextHorizontalMode.TextMid, TextVerticalMode.TextBase)
                cu.InsertText(brLap,
                              New Point3d(X + padingText,
                                          ptBasePoint.Y + 8 * heightRow + heightRow / 2, 0),
                              "EL__DIM", heightText, TextHorizontalMode.TextMid, TextVerticalMode.TextBase)
                cu.InsertText(brKontakt,
                              New Point3d(X + padingText,
                                          ptBasePoint.Y + 7 * heightRow + heightRow / 2, 0),
                              "EL__DIM", heightText, TextHorizontalMode.TextMid, TextVerticalMode.TextBase)
                cu.InsertText(Мощност,
                              New Point3d(X + padingText,
                                          ptBasePoint.Y + 6 * heightRow + heightRow / 2, 0),
                              "EL__DIM", heightText, TextHorizontalMode.TextMid, TextVerticalMode.TextBase)
                cu.InsertText(typeKabel,
                              New Point3d(X + padingText,
                                          ptBasePoint.Y + 5 * heightRow + heightRow / 2, 0),
                              "EL__DIM", heightText, TextHorizontalMode.TextMid, TextVerticalMode.TextBase)
                cu.InsertText(sechKabel,
                              New Point3d(X + padingText,
                                          ptBasePoint.Y + 4 * heightRow + heightRow / 2, 0),
                              "EL__DIM", heightText, TextHorizontalMode.TextMid, TextVerticalMode.TextBase)
                cu.InsertText(Faza,
                              New Point3d(X + padingText,
                                          ptBasePoint.Y + 3 * heightRow + heightRow / 2, 0),
                              "EL__DIM", heightText, TextHorizontalMode.TextMid, TextVerticalMode.TextBase)
                cu.InsertText(konsuator1,
                              New Point3d(X - widthColom / 2 + padingText,
                                          ptBasePoint.Y + 2 * heightRow + (heightRow - heightText) / 2, 0),
                              "EL__DIM", 12, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
                cu.InsertText(konsuator2,
                              New Point3d(X - widthColom / 2 + padingText,
                                          ptBasePoint.Y + 1 * heightRow + (heightRow - heightText) / 2, 0),
                              "EL__DIM", 12, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
                If Faza <> "L1" And Not Faza_Tablo And Faza <> Nothing Then
                    Faza_Tablo = True
                End If
                Dim doc As Document = Application.DocumentManager.MdiActiveDocument
                Dim blkRecId As ObjectId = ObjectId.Null

                If index < brColums - 1 Then
                    Select Case konsuator1
                        Case "Бойлер"
                            blkRecId = cu.InsertBlock("s_dpnn_vigi_circ_break",
                                           New Point3d(X, ptBasePoint.Y + Y_Шина, 0),
                                           "EL_ТАБЛА",
                                           New Scale3d(5, 5, 5)
                                           )
                            lengthProwBlock = 132.5
                        Case "Контакти"
                            blkRecId = cu.InsertBlock("s_c60_circ_break",
                                           New Point3d(X, ptBasePoint.Y + Y_Шина - 117.5, 0),
                                           "EL_ТАБЛА",
                                           New Scale3d(5, 5, 5)
                                           )
                            krygkontakt += 1

                            index_D = index

                            If krygkontakt = 3 Then
                                X = ptBasePoint.X + widthText + widthTextDim + (index - 3) * widthColom + widthColom / 2

                                blkRecId_D = cu.InsertBlock("s_id_res_circ_break",
                                               New Point3d(X, ptBasePoint.Y + Y_Шина, 0),
                                               "EL_ТАБЛА",
                                               New Scale3d(5, 5, 5)
                                               )

                                blkRecId_L = cu.DrowLine(New Point3d(X - widthColom - widthColom / 4, ptBasePoint.Y + Y_Шина - 117.5, 0),
                                                         New Point3d(X + widthColom + widthColom / 4, ptBasePoint.Y + Y_Шина - 117.5, 0),
                                                         "EL_ТАБЛА",
                                                         LineWeight.LineWeight070,
                                                         "ByLayer")

                                cu.InsertText(Faza & ",N" & Broj_N.ToString & ",PE",
                                              New Point3d(X - widthColom - widthColom / 4, ptBasePoint.Y + Y_Шина - 117.5 + 6, 0),
                                              "EL__DIM",
                                              heightText,
                                              TextHorizontalMode.TextLeft,
                                              TextVerticalMode.TextBase)
                                Broj_N += 1

                            End If
                            lengthProwBlock = 27.5
                        Case Else
                            blkRecId = cu.InsertBlock("s_c60_circ_break",
                                           New Point3d(X, ptBasePoint.Y + Y_Шина, 0),
                                           "EL_ТАБЛА",
                                           New Scale3d(5, 5, 5)
                                           )
                            lengthProwBlock = 145
                    End Select
                    Using trans As Transaction = doc.TransactionManager.StartTransaction()
                        Dim acBlkTbl As BlockTable = trans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                        Dim acBlkRef As BlockReference =
                        DirectCast(trans.GetObject(blkRecId, OpenMode.ForWrite), BlockReference)

                        Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                        For Each objID As ObjectId In attCol
                            Dim dbObj As DBObject = trans.GetObject(objID, OpenMode.ForWrite)
                            Dim acAttRef As AttributeReference = dbObj
                            If DataGridView.Rows(9).Cells(index).Value = "EZ9 RCBO" Then
                                If acAttRef.Tag = "1" Then acAttRef.TextString = DataGridView.Rows(10).Cells(index).Value ' АЦ
                                If acAttRef.Tag = "2" Then acAttRef.TextString = DataGridView.Rows(13).Cells(index).Value ' 2п
                                If acAttRef.Tag = "3" Then acAttRef.TextString = "C"
                                If acAttRef.Tag = "4" Then acAttRef.TextString = "25A"
                                If acAttRef.Tag = "5" Then acAttRef.TextString = DataGridView.Rows(12).Cells(index).Value ' 30мА
                                If acAttRef.Tag = "SHORTNAME" Then acAttRef.TextString = DataGridView.Rows(9).Cells(index).Value
                            Else
                                If acAttRef.Tag = "1" Then acAttRef.TextString = ""
                                If acAttRef.Tag = "2" Then acAttRef.TextString = DataGridView.Rows(5).Cells(index).Value
                                If acAttRef.Tag = "3" Then acAttRef.TextString = DataGridView.Rows(6).Cells(index).Value
                                If acAttRef.Tag = "4" Then acAttRef.TextString = DataGridView.Rows(3).Cells(index).Value
                                If acAttRef.Tag = "SHORTNAME" Then acAttRef.TextString = DataGridView.Rows(2).Cells(index).Value
                            End If

                            If acAttRef.Tag = "DESIGNATION" Then acAttRef.TextString = ""
                            If acAttRef.Tag = "REFNB" Then acAttRef.TextString = TabloName

                        Next
                        trans.Commit()
                    End Using
                    If krygkontakt = 3 Then
                        Using trans As Transaction = doc.TransactionManager.StartTransaction()
                            Dim acBlkTbl As BlockTable = trans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                            Dim acBlkRef As BlockReference =
                                DirectCast(trans.GetObject(blkRecId_D, OpenMode.ForWrite), BlockReference)

                            Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                            For Each objID As ObjectId In attCol
                                Dim dbObj As DBObject = trans.GetObject(objID, OpenMode.ForWrite)
                                Dim acAttRef As AttributeReference = dbObj
                                If acAttRef.Tag = "DESIGNATION" Then acAttRef.TextString = ""
                                If acAttRef.Tag = "1" Then acAttRef.TextString = DataGridView.Rows(10).Cells(index).Value ' АЦ
                                If acAttRef.Tag = "2" Then acAttRef.TextString = DataGridView.Rows(13).Cells(index).Value ' 2п
                                If acAttRef.Tag = "3" Then acAttRef.TextString = DataGridView.Rows(11).Cells(index).Value
                                If acAttRef.Tag = "4" Then acAttRef.TextString = "Мигновена"
                                If acAttRef.Tag = "5" Then acAttRef.TextString = DataGridView.Rows(12).Cells(index).Value ' 30мА
                                If acAttRef.Tag = "REFNB" Then acAttRef.TextString = TabloName
                                If acAttRef.Tag = "SHORTNAME" Then acAttRef.TextString = DataGridView.Rows(9).Cells(index).Value
                            Next
                            trans.Commit()
                        End Using
                        krygkontakt = 0
                    End If
                    X = ptBasePoint.X + widthText + widthTextDim + (index - 1) * widthColom
                    cu.DrowLine(New Point3d(X - widthColom / 2, ptBasePoint.Y + 10 * heightRow, 0),
                                New Point3d(X - widthColom / 2, ptBasePoint.Y + 10 * heightRow + lengthProw + lengthProwBlock, 0),
                                "EL_ТАБЛА",
                                LineWeight.ByLayer,
                                "ByLayer")
                End If
            Next

            Select Case krygkontakt
                Case 1
                    blkRecId_D = blkRecId_D
                    blkRecId_L = blkRecId_L
                    index_D = index_D
                Case 2
                    blkRecId_D = blkRecId_D
                    blkRecId_L = blkRecId_L
                    index_D = index_D
            End Select


            'blkRecId_L = cu.DrowLine(New Point3d(X - widthColom - widthColom / 4, ptBasePoint.Y + Y_Шина - 117.5, 0),
            '                            New Point3d(X + widthColom + widthColom / 4, ptBasePoint.Y + Y_Шина - 117.5, 0),
            '                            "EL_ТАБЛА",
            '                            LineWeight.LineWeight070,
            '                            "ByLayer")

            X = ptBasePoint.X + widthText + widthTextDim

            cu.InsertText(TabloName,
                          New Point3d(X + (brColums - 3) * widthColom,
                                      ptBasePoint.Y + Y_Шина + 95,
                                      0),
                          "EL__DIM", heightText + 5, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)

            cu.InsertText(IIf(Faza_Tablo, "L1,L2,L3,N,PE", "L,N,PE"),
                          New Point3d(X, ptBasePoint.Y + Y_Шина + 2 * padingText, 0),
                          "EL__DIM",
                          heightText,
                          TextHorizontalMode.TextLeft,
                          TextVerticalMode.TextBase)

            cu.DrowLine(New Point3d(X, ptBasePoint.Y + Y_Шина, 0),
                        New Point3d(X + (brColums - 3) * widthColom, ptBasePoint.Y + Y_Шина, 0),
                        "EL_ТАБЛА",
                        LineWeight.LineWeight070,
                        "ByLayer")

            cu.DrowLine(New Point3d(X + (brColums - 3) * widthColom / 2, ptBasePoint.Y + Y_Шина + 95, 0),
                            New Point3d(X + (brColums - 3) * widthColom / 2, ptBasePoint.Y + Y_Шина + 220, 0),
                            "EL_ТАБЛА",
                            LineWeight.ByLayer,
                            "ByLayer")

            '
            ' Поставя знак за заземление
            '
            If TabloName = "Гл.Р.Т." Then

                cu.DrowLine(New Point3d(X, ptBasePoint.Y + Y_Шина, 0),
                        New Point3d(X - widthColom, ptBasePoint.Y + Y_Шина, 0),
                        "EL_ТАБЛА",
                        LineWeight.ByLayer,
                        "ByLayer")

                cu.InsertText("R<30Ω",
                          New Point3d(X - widthColom, ptBasePoint.Y + Y_Шина + 2 * padingText, 0),
                          "EL__DIM",
                          heightText,
                          TextHorizontalMode.TextLeft,
                          TextVerticalMode.TextBase)

                Dim blkRecId = cu.InsertBlock("Заземление",
                               New Point3d(X - widthColom, ptBasePoint.Y + Y_Шина, 0),
                               "EL_ТАБЛА",
                               New Scale3d(1, 1, 1)
                               )

                Dim doc As Document = Application.DocumentManager.MdiActiveDocument
                Using trans As Transaction = doc.TransactionManager.StartTransaction()
                    Dim acBlkTbl As BlockTable = trans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                    Dim acBlkRef As BlockReference =
                        DirectCast(trans.GetObject(blkRecId, OpenMode.ForWrite), BlockReference)

                    Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection

                    For Each prop As DynamicBlockReferenceProperty In props
                        'This Is where you change states based on input
                        If prop.PropertyName = "Visibility" Then prop.Value = "Без ревизионна кутия"
                        If prop.PropertyName = "Position1 X" Then prop.Value = -15.0
                        If prop.PropertyName = "Position1 Y" Then prop.Value = -80.0
                        If prop.PropertyName = "Angle1" Then prop.Value = PI
                    Next
                    Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                    For Each objID As ObjectId In attCol
                        Dim dbObj As DBObject = trans.GetObject(objID, OpenMode.ForWrite)
                        Dim acAttRef As AttributeReference = dbObj
                        If acAttRef.Tag = "ТАБЛО" Then acAttRef.TextString = "2к"
                    Next

                    trans.Commit()
                End Using


                cu.InsertText("PE",
                          New Point3d(X - widthColom + 3 * padingText, ptBasePoint.Y + Y_Шина - heightText - padingText, 0),
                          "EL__DIM",
                          heightText,
                          TextHorizontalMode.TextLeft,
                          TextVerticalMode.TextBase)
            End If
            '
            ' Поставя товаров прекъсвач 
            '
            Dim doc_ As Document = Application.DocumentManager.MdiActiveDocument
            Dim blkRecId_ As ObjectId = ObjectId.Null
            blkRecId_ = cu.InsertBlock("s_i_ng_switch_disconn",
                           New Point3d(ptBasePoint.X + widthText + widthTextDim +
                           (brColums - 3) * widthColom / 2,
                                       ptBasePoint.Y + Y_Шина + 95,
                                       0),
                           "EL_ТАБЛА",
                           New Scale3d(5, 5, 5)
                           )

            Using trans As Transaction = doc_.TransactionManager.StartTransaction()
                Dim acBlkTbl As BlockTable = trans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim acBlkRef As BlockReference =
                                DirectCast(trans.GetObject(blkRecId_, OpenMode.ForWrite), BlockReference)

                Dim Index As Integer = DataGridView.Columns.Count - 1
                Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                For Each objID As ObjectId In attCol
                    Dim dbObj As DBObject = trans.GetObject(objID, OpenMode.ForWrite)
                    Dim acAttRef As AttributeReference = dbObj
                    If acAttRef.Tag = "DESIGNATION" Then acAttRef.TextString = ""
                    If acAttRef.Tag = "1" Then acAttRef.TextString = ""
                    If acAttRef.Tag = "2" Then acAttRef.TextString = DataGridView.Rows(6).Cells(Index).Value ' 2п
                    If acAttRef.Tag = "3" Then acAttRef.TextString = DataGridView.Rows(3).Cells(Index).Value
                    If acAttRef.Tag = "4" Then acAttRef.TextString = ""
                    If acAttRef.Tag = "5" Then acAttRef.TextString = ""
                    If acAttRef.Tag = "REFNB" Then acAttRef.TextString = TabloName
                    If acAttRef.Tag = "SHORTNAME" Then acAttRef.TextString = DataGridView.Rows(2).Cells(Index).Value
                Next
                trans.Commit()
            End Using

            Dim Zabelevka As String = "1. Таблото да се изпълни в съответствие с изискванията на БДС EN 61439-1."
            Zabelevka += vbCrLf & "2. Aпаратурата и тоководящите части да бъдат монтирани зад защитни капаци. "
            Zabelevka += vbCrLf & "3. Достъпа до палците и ръкохватките на комутационните апарати се осигурява посредством отвори в защитните капаци."
            Zabelevka += vbCrLf & "4. Апаратурата е избрана по каталога на SCHNEIDER ELECTRIC."
            Zabelevka += vbCrLf & "5. Изборът на автоматичните прекъсвачи е съобразен с токовете на к.с., спазени са изискванията за селективност."
            Zabelevka += vbCrLf & "6. При замяна типа на апаратурата да се преизчисли схемата."
            cu.InsertMText("ЗАБЕЛЕЖКИ:",
                                     New Point3d(ptBasePoint.X,
                                                 ptBasePoint.Y - 20, 0),
                                     "EL__DIM", 10)
            cu.InsertMText(Zabelevka,
                                     New Point3d(ptBasePoint.X + 30,
                                                 ptBasePoint.Y - 20 - heightRow, 0),
                                     "EL__DIM", 10)

        Catch ex As Exception
            MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
        End Try

    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TreeView_ADD()
    End Sub
    Private Sub GroupBox3_Enter(sender As Object, e As EventArgs) Handles GroupBox3.Enter

    End Sub
End Class
Public Class Tablo_new
    Dim form_AS As New Form_Tablo_new()
    <CommandMethod("Tablo_new")>
    Public Sub Tablo_new()
        form_AS.Show()
    End Sub
End Class