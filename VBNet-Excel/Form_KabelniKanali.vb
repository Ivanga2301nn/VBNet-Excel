Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports System.Net
Imports System.IO

Public Class Form_KabelniKanali
    '
    'C:\Users\I\source\repos\VBNet-Excel\VBNet-Excel\bin\x64\Debug
    '
    Dim PI As Double = 3.1415926535897931
    Dim cu As CommonUtil = New CommonUtil()
    Dim Line_Selected As SelectionSet
    Dim Slabo_Yes As Boolean = vbFalse
    Dim Solar_Yes As Boolean = vbFalse
    Dim Silow_Yes As Boolean = vbFalse
    Dim summaKabeli As Double = 0
    Dim skara As Double = 0
    Dim arrSkari(10, 4) As Integer
    Dim arrKanal(10, 4) As Integer
    Dim Шир As Double = 0
    Dim Вис As Double = 0
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
    Dim Kabel(200) As strLine
    Private Sub Form_KabelniKanali_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Line_Selected = cu.GetObjects("LINE", "Изберете линии за кабеланата скара/канал:")
        If Line_Selected Is Nothing Then
            MsgBox("НЕ Е маркиранa линия в слой 'EL'.")
            Exit Sub
        End If

        arrSkari(1, 0) = 50
        arrSkari(1, 1) = 35
        arrSkari(1, 2) = 60
        arrSkari(1, 3) = 0
        arrSkari(1, 4) = 0

        arrSkari(2, 0) = 100
        arrSkari(2, 1) = 35
        arrSkari(2, 2) = 60
        arrSkari(2, 3) = 85
        arrSkari(2, 4) = 0

        arrSkari(3, 0) = 150
        arrSkari(3, 1) = 35
        arrSkari(3, 2) = 60
        arrSkari(3, 3) = 85
        arrSkari(3, 4) = 0

        arrSkari(4, 0) = 200
        arrSkari(4, 1) = 35
        arrSkari(4, 2) = 60
        arrSkari(4, 3) = 85
        arrSkari(4, 4) = 110

        arrSkari(5, 0) = 300
        arrSkari(5, 1) = 35
        arrSkari(5, 2) = 60
        arrSkari(5, 3) = 85
        arrSkari(5, 4) = 110

        arrSkari(6, 0) = 400
        arrSkari(6, 1) = 60
        arrSkari(6, 2) = 85
        arrSkari(6, 3) = 110
        arrSkari(6, 4) = 0

        arrSkari(7, 0) = 500
        arrSkari(7, 1) = 60
        arrSkari(7, 2) = 85
        arrSkari(7, 3) = 110
        arrSkari(7, 4) = 0

        arrSkari(8, 0) = 600
        arrSkari(8, 1) = 60
        arrSkari(8, 2) = 85
        arrSkari(8, 3) = 110
        arrSkari(8, 4) = 0

        arrSkari(9, 0) = 0
        arrSkari(9, 1) = 0
        arrSkari(9, 2) = 0
        arrSkari(9, 3) = 0
        arrSkari(9, 4) = 0

        arrSkari(10, 0) = 0
        arrSkari(10, 1) = 0
        arrSkari(10, 2) = 0
        arrSkari(10, 3) = 0
        arrSkari(10, 4) = 0

        '########################################################################

        arrKanal(1, 0) = 12
        arrKanal(1, 1) = 12
        arrKanal(1, 2) = 0
        arrKanal(1, 3) = 0
        arrKanal(1, 4) = 0

        arrKanal(2, 0) = 16
        arrKanal(2, 1) = 16
        arrKanal(2, 2) = 0
        arrKanal(2, 3) = 0
        arrKanal(2, 4) = 0

        arrKanal(3, 0) = 20
        arrKanal(3, 1) = 20
        arrKanal(3, 2) = 0
        arrKanal(3, 3) = 0
        arrKanal(3, 4) = 0

        arrKanal(4, 0) = 25
        arrKanal(4, 1) = 20
        arrKanal(4, 2) = 25
        arrKanal(4, 3) = 0
        arrKanal(4, 4) = 0

        arrKanal(5, 0) = 30
        arrKanal(5, 1) = 25
        arrKanal(5, 2) = 0
        arrKanal(5, 3) = 0
        arrKanal(5, 4) = 0

        arrKanal(6, 0) = 40
        arrKanal(6, 1) = 20
        arrKanal(6, 2) = 25
        arrKanal(6, 3) = 40
        arrKanal(6, 4) = 0

        arrKanal(7, 0) = 60
        arrKanal(7, 1) = 20
        arrKanal(7, 2) = 40
        arrKanal(7, 3) = 60
        arrKanal(7, 4) = 0

        arrKanal(8, 0) = 80
        arrKanal(8, 1) = 20
        arrKanal(8, 2) = 25
        arrKanal(8, 3) = 40
        arrKanal(8, 4) = 0

        arrKanal(9, 0) = 100
        arrKanal(9, 1) = 40
        arrKanal(9, 2) = 60
        arrKanal(9, 3) = 0
        arrKanal(9, 4) = 0

        arrKanal(10, 0) = 140
        arrKanal(10, 1) = 60
        arrKanal(10, 2) = 70
        arrKanal(10, 3) = 0
        arrKanal(10, 4) = 0

        Call Set_array_Kabel()
    End Sub
    Private Sub Izbor_Skara()
        Label_Сечение_кабели.Text = summaKabeli
        Dim Процент As Double = Val(ComboBox_Процент_Запълване.SelectedItem)
        For i As Integer = 0 To 10
            For j As Integer = 1 To 4
                Шир = arrSkari(i, 0)
                Вис = arrSkari(i, j)
                If summaKabeli < (Шир * Вис * Процент / 100) Then
                    Изчислява_Skara()
                    Exit Sub
                End If
            Next
        Next
    End Sub
    Private Sub Изчислява_Skara()
        TextBox_Кабелна_Скара.Text = Шир & "х" & Вис
        Dim procent As Integer = Int(summaKabeli / (Шир * Вис) * 100)
        Label_Процент_Запълване.Text = procent
        ProgressBar_Procent.Value = IIf(procent > 100, 100, procent)
    End Sub
    Private Sub RadioButton_Скара_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton_Скара.CheckedChanged
        GroupBox_Избор.Text = "Избор на кабелна скара"
        Label4.Text = "Скара [ШхВ],mm"
        clear_GroupBox_Избор()
        Call Izbor_Skara()
    End Sub
    Private Sub RadioButton_Канал_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton_Канал.CheckedChanged
        GroupBox_Избор.Text = "Избор на кабелен канал"
        Label4.Text = "Канал [ШхВ],mm"
        clear_GroupBox_Избор()
        Call Izbor_Skara()
    End Sub
    Private Sub RadioButton_Тръба_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton_Тръба.CheckedChanged
        GroupBox_Избор.Text = "Избор на тръба"
        clear_GroupBox_Избор()
    End Sub
    Private Sub clear_GroupBox_Избор()
        GroupBox_Избор.Visible = True
        clearButton_Ш()
    End Sub
    Private Sub clearButton_Ш()
        With Button_Ш_1
            .BackColor = System.Drawing.SystemColors.ControlLight
            .Text = IIf(RadioButton_Скара.Checked, arrSkari(1, 0), arrKanal(1, 0))
        End With
        With Button_Ш_2
            .BackColor = System.Drawing.SystemColors.ControlLight
            .Text = IIf(RadioButton_Скара.Checked, arrSkari(2, 0), arrKanal(2, 0))
        End With
        With Button_Ш_3
            .BackColor = System.Drawing.SystemColors.ControlLight
            .Text = IIf(RadioButton_Скара.Checked, arrSkari(3, 0), arrKanal(3, 0))
        End With
        With Button_Ш_4
            .BackColor = System.Drawing.SystemColors.ControlLight
            .Text = IIf(RadioButton_Скара.Checked, arrSkari(4, 0), arrKanal(4, 0))
        End With
        With Button_Ш_5
            .BackColor = System.Drawing.SystemColors.ControlLight
            .Text = IIf(RadioButton_Скара.Checked, arrSkari(5, 0), arrKanal(5, 0))
        End With
        With Button_Ш_6
            .BackColor = System.Drawing.SystemColors.ControlLight
            .Text = IIf(RadioButton_Скара.Checked, arrSkari(6, 0), arrKanal(6, 0))
        End With
        With Button_Ш_7
            .BackColor = System.Drawing.SystemColors.ControlLight
            .Text = IIf(RadioButton_Скара.Checked, arrSkari(7, 0), arrKanal(7, 0))
        End With
        With Button_Ш_8
            .BackColor = System.Drawing.SystemColors.ControlLight
            .Text = IIf(RadioButton_Скара.Checked, arrSkari(8, 0), arrKanal(8, 0))
        End With
        With Button_Ш_9
            .BackColor = System.Drawing.SystemColors.ControlLight
            .Text = IIf(RadioButton_Скара.Checked, arrSkari(9, 0), arrKanal(9, 0))
        End With
        With Button_Ш_10
            .BackColor = System.Drawing.SystemColors.ControlLight
            .Text = IIf(RadioButton_Скара.Checked, arrSkari(10, 0), arrKanal(10, 0))
        End With
        With Button_В_1
            .BackColor = System.Drawing.SystemColors.ControlLight
            .Text = IIf(RadioButton_Скара.Checked, 50, 20)
        End With
        With Button_В_2
            .BackColor = System.Drawing.SystemColors.ControlLight
            .Text = IIf(RadioButton_Скара.Checked, 60, 25)
        End With
        With Button_В_3
            .BackColor = System.Drawing.SystemColors.ControlLight
            .Text = IIf(RadioButton_Скара.Checked, 85, 40)
        End With
        With Button_В_4
            .BackColor = System.Drawing.SystemColors.ControlLight
            .Text = IIf(RadioButton_Скара.Checked, 100, 60)
        End With
    End Sub
    Private Sub Button_Ш_1_Click(sender As Object, e As EventArgs) Handles Button_Ш_1.Click, Button_Ш_2.Click, Button_Ш_3.Click,
                                                                           Button_Ш_4.Click, Button_Ш_5.Click, Button_Ш_6.Click,
                                                                           Button_Ш_7.Click, Button_Ш_8.Click, Button_Ш_9.Click,
                                                                           Button_Ш_10.Click

        clearButton_Ш()
        With sender
            .BackColor = System.Drawing.SystemColors.Info
        End With
        Button_В_1.Enabled = vbTrue
        Button_В_2.Enabled = vbTrue
        Button_В_3.Enabled = vbTrue
        Button_В_4.Enabled = vbTrue
        Select Case sender.name
            Case "Button_Ш_1"
                Button_В_1.Text = IIf(RadioButton_Скара.Checked, arrSkari(1, 1), arrKanal(1, 1))
                Button_В_2.Text = IIf(RadioButton_Скара.Checked, arrSkari(1, 2), arrKanal(1, 2))
                Button_В_3.Text = IIf(RadioButton_Скара.Checked, arrSkari(1, 3), arrKanal(1, 3))
                Button_В_4.Text = IIf(RadioButton_Скара.Checked, arrSkari(1, 4), arrKanal(1, 4))
            Case "Button_Ш_2"
                Button_В_1.Text = IIf(RadioButton_Скара.Checked, arrSkari(2, 1), arrKanal(2, 1))
                Button_В_2.Text = IIf(RadioButton_Скара.Checked, arrSkari(2, 2), arrKanal(2, 2))
                Button_В_3.Text = IIf(RadioButton_Скара.Checked, arrSkari(2, 3), arrKanal(2, 3))
                Button_В_4.Text = IIf(RadioButton_Скара.Checked, arrSkari(2, 4), arrKanal(2, 4))
            Case "Button_Ш_3"
                Button_В_1.Text = IIf(RadioButton_Скара.Checked, arrSkari(3, 1), arrKanal(3, 1))
                Button_В_2.Text = IIf(RadioButton_Скара.Checked, arrSkari(3, 2), arrKanal(3, 2))
                Button_В_3.Text = IIf(RadioButton_Скара.Checked, arrSkari(3, 3), arrKanal(3, 3))
                Button_В_4.Text = IIf(RadioButton_Скара.Checked, arrSkari(3, 4), arrKanal(3, 4))
            Case "Button_Ш_4"
                Button_В_1.Text = IIf(RadioButton_Скара.Checked, arrSkari(4, 1), arrKanal(4, 1))
                Button_В_2.Text = IIf(RadioButton_Скара.Checked, arrSkari(4, 2), arrKanal(4, 2))
                Button_В_3.Text = IIf(RadioButton_Скара.Checked, arrSkari(4, 3), arrKanal(4, 3))
                Button_В_4.Text = IIf(RadioButton_Скара.Checked, arrSkari(4, 4), arrKanal(4, 4))
            Case "Button_Ш_5"
                Button_В_1.Text = IIf(RadioButton_Скара.Checked, arrSkari(5, 1), arrKanal(4, 1))
                Button_В_2.Text = IIf(RadioButton_Скара.Checked, arrSkari(5, 2), arrKanal(5, 2))
                Button_В_3.Text = IIf(RadioButton_Скара.Checked, arrSkari(5, 3), arrKanal(5, 3))
                Button_В_4.Text = IIf(RadioButton_Скара.Checked, arrSkari(5, 4), arrKanal(5, 4))
            Case "Button_Ш_6"
                Button_В_1.Text = IIf(RadioButton_Скара.Checked, arrSkari(6, 1), arrKanal(6, 1))
                Button_В_2.Text = IIf(RadioButton_Скара.Checked, arrSkari(6, 2), arrKanal(6, 2))
                Button_В_3.Text = IIf(RadioButton_Скара.Checked, arrSkari(6, 3), arrKanal(6, 3))
                Button_В_4.Text = IIf(RadioButton_Скара.Checked, arrSkari(6, 4), arrKanal(6, 4))
            Case "Button_Ш_7"
                Button_В_1.Text = IIf(RadioButton_Скара.Checked, arrSkari(7, 1), arrKanal(7, 1))
                Button_В_2.Text = IIf(RadioButton_Скара.Checked, arrSkari(7, 2), arrKanal(7, 2))
                Button_В_3.Text = IIf(RadioButton_Скара.Checked, arrSkari(7, 3), arrKanal(7, 3))
                Button_В_4.Text = IIf(RadioButton_Скара.Checked, arrSkari(7, 4), arrKanal(7, 4))
            Case "Button_Ш_8"
                Button_В_1.Text = IIf(RadioButton_Скара.Checked, arrSkari(8, 1), arrKanal(8, 1))
                Button_В_2.Text = IIf(RadioButton_Скара.Checked, arrSkari(8, 2), arrKanal(8, 2))
                Button_В_3.Text = IIf(RadioButton_Скара.Checked, arrSkari(8, 3), arrKanal(8, 3))
                Button_В_4.Text = IIf(RadioButton_Скара.Checked, arrSkari(8, 4), arrKanal(8, 4))
            Case "Button_Ш_9"
                Button_В_1.Text = IIf(RadioButton_Скара.Checked, arrSkari(9, 1), arrKanal(9, 1))
                Button_В_2.Text = IIf(RadioButton_Скара.Checked, arrSkari(9, 2), arrKanal(9, 2))
                Button_В_3.Text = IIf(RadioButton_Скара.Checked, arrSkari(9, 3), arrKanal(9, 3))
                Button_В_4.Text = IIf(RadioButton_Скара.Checked, arrSkari(9, 4), arrKanal(9, 4))
            Case "Button_Ш_10"
                Button_В_1.Text = IIf(RadioButton_Скара.Checked, arrSkari(10, 1), arrKanal(10, 1))
                Button_В_2.Text = IIf(RadioButton_Скара.Checked, arrSkari(10, 2), arrKanal(10, 2))
                Button_В_3.Text = IIf(RadioButton_Скара.Checked, arrSkari(10, 3), arrKanal(10, 3))
                Button_В_4.Text = IIf(RadioButton_Скара.Checked, arrSkari(10, 4), arrKanal(10, 4))
        End Select
        If Button_В_2.Text = 0 Then
            Button_В_2.Visible = vbFalse
        Else
            Button_В_2.Visible = vbTrue
        End If
        If Button_В_3.Text = 0 Then
            Button_В_3.Visible = vbFalse
        Else
            Button_В_3.Visible = vbTrue
        End If
        If Button_В_4.Text = 0 Then
            Button_В_4.Visible = vbFalse
        Else
            Button_В_4.Visible = vbTrue
        End If
        Шир = sender.text
        Call Изчислява_Skara()
    End Sub
    Private Sub Button_В_1_Click(sender As Object, e As EventArgs) Handles Button_В_1.Click, Button_В_2.Click, Button_В_3.Click,
                                                                           Button_В_4.Click
        With Button_В_1
            .BackColor = System.Drawing.SystemColors.ControlLight
        End With
        With Button_В_2
            .BackColor = System.Drawing.SystemColors.ControlLight
        End With
        With Button_В_3
            .BackColor = System.Drawing.SystemColors.ControlLight
        End With
        With Button_В_4
            .BackColor = System.Drawing.SystemColors.ControlLight
        End With
        With sender
            .BackColor = System.Drawing.SystemColors.Info
        End With
        Вис = sender.text
        Call Изчислява_Skara()
    End Sub
    Private Sub ComboBox_Процент_Запълване_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_Процент_Запълване.SelectedIndexChanged
        Izbor_Skara()
    End Sub
    Private Sub DataGridView_Кабели_CellEndEdit(sender As Object, e As Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView_Кабели.CellEndEdit
        Dim rows As Integer = e.RowIndex
        Dim cells As Integer = e.ColumnIndex

        Dim n1 As Boolean = Not IsNothing(DataGridView_Кабели.Rows(rows).Cells(0).Value)
        Dim n2 As Boolean = Not IsNothing(DataGridView_Кабели.Rows(rows).Cells(1).Value)
        Dim n3 As Boolean = Not IsNothing(DataGridView_Кабели.Rows(rows).Cells(2).Value)
        Dim n4 As Boolean = IIf(Val(DataGridView_Кабели.Rows(rows).Cells(3).Value) = 0, vbFalse, vbTrue)
        Dim kab As String = "EL_"
        If n1 And n2 And n3 And n4 Then
            Select Case DataGridView_Кабели.Rows(rows).Cells(0).Value.ToString
                Case "СВТ/САВТ"
                    If DataGridView_Кабели.Rows(rows).Cells(2).Value = "2,5" Then
                        kab = kab & DataGridView_Кабели.Rows(rows).Cells(1).Value.ToString &
                                    "x" &
                                    "2_5"
                        Exit Select
                    End If
                    If DataGridView_Кабели.Rows(rows).Cells(2).Value = "1,5" Then
                        kab = kab & DataGridView_Кабели.Rows(rows).Cells(1).Value.ToString &
                                    "x" &
                                    "1_5"
                        Exit Select
                    End If
                    If DataGridView_Кабели.Rows(rows).Cells(1).Value.ToString <> "3+" Then
                        kab = kab &
                                        DataGridView_Кабели.Rows(rows).Cells(1).Value.ToString &
                                        "x" &
                                        DataGridView_Кабели.Rows(rows).Cells(2).Value.ToString
                    Else
                        Select Case DataGridView_Кабели.Rows(rows).Cells(2).Value.ToString
                            Case "25"
                                kab = kab + "3x25+16"
                            Case "35"
                                kab = kab + "3х35+16"
                            Case "50"
                                kab = kab + "3х50+25"
                            Case "70"
                                kab = kab + "3х70+35"
                            Case "95"
                                kab = kab + "3x95+50"
                            Case "120"
                                kab = kab + "3х120+70"
                            Case "150"
                                kab = kab + "3x150+70"
                            Case "185"
                                kab = kab + "3x185+95"
                            Case "240"
                                kab = kab + "3x240+120"
                        End Select
                    End If
                Case "Слаботоков"
                    kab = kab & "UTP"
                Case "Соларен"
                    kab = kab & "стринг1"
            End Select
            Dim sec As Double = cu.GET_line_Diamet(kab)
            sec = IIf(sec = -1, 0, sec)
            Dim Plo As Double = sec * sec * Val(DataGridView_Кабели.Rows(rows).Cells(3).Value)
            DataGridView_Кабели.Rows(rows).Cells(4).Value = sec
            DataGridView_Кабели.Rows(rows).Cells(5).Value = Plo
            summaKabeli = 0
            For Each row As Windows.Forms.DataGridViewRow In DataGridView_Кабели.Rows
                summaKabeli = summaKabeli + row.Cells(5).Value
            Next
        End If
        Label_Сечение_кабели.Text = summaKabeli
    End Sub
    Private Sub Label_Процент_Запълване_TextChanged(sender As Object, e As EventArgs) Handles Label_Процент_Запълване.TextChanged
        Dim procent As Integer = Val(Label_Процент_Запълване.Text)
        Dim R As Integer = 0
        Dim G As Integer = 0
        Dim B As Integer = 0
        Select Case procent
            Case < 10
                R = 0
                G = 255
                B = 0
            Case < 20
                R = 50
                G = 255
                B = 0
            Case < 30
                R = 100
                G = 255
                B = 0
            Case < 40
                R = 150
                G = 255
                B = 0
            Case < 50
                R = 200
                G = 255
                B = 0
            Case < 60
                R = 255
                G = 255
                B = 0
            Case < 70
                R = 255
                G = 200
                B = 0
            Case < 80
                R = 255
                G = 150
                B = 0
            Case < 90
                R = 255
                G = 100
                B = 0
            Case < 100
                R = 255
                G = 50
                B = 0
            Case > 100
                R = 255
                G = 0
                B = 0
        End Select
        Label_Процент_Запълване.BackColor = System.Drawing.Color.FromArgb(R, G, B)
    End Sub
    Private Sub Set_array_Kabel()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
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
        GroupBox_Избор.Visible = vbFalse
        Button_В_1.Enabled = vbFalse
        Button_В_2.Enabled = vbFalse
        Button_В_3.Enabled = vbFalse
        Button_В_4.Enabled = vbFalse
    End Sub
    Private Sub OpenToolStripButton_Click(sender As Object, e As EventArgs) Handles OpenToolStripButton.Click
        ReDim Kabel(200)
        Line_Selected = cu.GetObjects("LINE", "Изберете линии за кабеланата скара/канал:")
        If Line_Selected Is Nothing Then
            MsgBox("НЕ Е маркиранa линия в слой 'EL'.")
            Exit Sub
        End If
        Call Set_array_Kabel()
    End Sub
End Class

