Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports System.Net
Imports System.IO
Public Class Form_Skari_Kanali_New
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
    Private Sub Skari_Kanali_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Me.Height = 1000
        'Me.Width = 1600

        Line_Selected = cu.GetObjects("LINE", "Изберете линии за кабеланата скара/канал:")
        If Line_Selected Is Nothing Then
            MsgBox("Нама маркиранa линия в слой 'EL'.")
            Exit Sub
        End If

        GroupBox_Размери_Скари.Visible = False
        GroupBox_Размери_Скари.Dock = Windows.Forms.DockStyle.Fill



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
    End Sub
    Private Sub RadioButton_Скара_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton_Скара.CheckedChanged
        GroupBox_Размери_Скари.Visible = True
    End Sub
    Private Sub RadioButton_Канал_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton_Канал.CheckedChanged
        GroupBox_Размери_Скари.Visible = False
    End Sub
    Private Sub RadioButton_Тръба_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton_Тръба.CheckedChanged
        GroupBox_Размери_Скари.Visible = False
    End Sub

End Class