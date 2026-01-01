Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports System.IO
Imports System.Net
Imports System.Security.Cryptography
Imports System.Diagnostics.Eventing.Reader
Imports System.Drawing.Drawing2D
Imports Autodesk.AutoCAD.DatabaseServices.Filters
Imports Autodesk.AutoCAD.GraphicsInterface

Imports excel = Microsoft.Office.Interop.Excel

Imports ACSMCOMPONENTS24Lib
Imports Microsoft.Office.Interop

Public Class Obekti
    Dim cu As CommonUtil = New CommonUtil()
    Private Sub Obekti_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim name_file As String = Application.DocumentManager.MdiActiveDocument.Name
        Dim File_Path As String = Path.GetDirectoryName(name_file)
        Dim File_name As String = Path.GetFileName(name_file)
        Dim File_DST As String = File_Path + "\" + "SheetSet.dst"
        Dim Set_Name As String = "БЛОК"
        Dim Set_Desc As String = "Създадено от Бат Генчо"
        With DataGridView
            With .Rows
                .Add({"Наименование на ОБЕКТА", ""})
                .Add({"Местоположение на ОБЕКТА", ""})
                .Add({"ВЪЗЛОЖИТЕЛ на проекта", ""})
                .Add({"СОСТВЕНИК на обекта", ""})
                .Add({"ФАЗА на проекта", ""})
                .Add({"ДАТА на проекта", ""})
                .Add({"Част АРХИТЕКТУРА", ""})
                .Add({"Част КОНСТРУКЦИИ", ""})
                .Add({"Част ТЕХНОЛОГИЯ", ""})
                .Add({"Част ВиК", ""})
                .Add({"Част ОВ", ""})
                .Add({"Част ГЕОДЕЗИЯ", ""})
                .Add({"Част ВП", ""})
                .Add({"Част ЕЕФ", ""})
                .Add({"Част ПБ", ""})
                .Add({"Част ПБЗ", ""})
                .Add({"Част ПУСО", ""})
                .Add({"Проектант", "инж. М.Тонкова-Генчева"})
                .Add({"Място на диска", File_Path})
            End With
        End With
    End Sub
    Private Sub DataGridView_CellContentClick(sender As Object, e As Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView.CellContentClick
        Me.Visible = False
        Dim sss As String
        sss = DataGridView.Rows(e.RowIndex).Cells(0).Value
        DataGridView.Rows(e.RowIndex).Cells(1).Value = cu.GetObjects_TEXT("Изберете " & sss)
        Me.Visible = True
    End Sub

    Private Sub butExit_Click(sender As Object, e As EventArgs) Handles butExit.Click
        Me.Close()
    End Sub
End Class

Public Class Set_SheetSet
    Dim Form_Obekti As New Obekti
    <CommandMethod("Set_SheetSet")>
    Public Sub Set_SheetSet()
        Form_Obekti.Show()
    End Sub
End Class


