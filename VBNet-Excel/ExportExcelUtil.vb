
Imports Autodesk.AutoCAD.Runtime


Public Class ShowExcelUtilForm

    <CommandMethod("ShowExcelUtilForm")>
    Public Sub ShowExcelUtilForm()
        Dim form_AS_tablo As New Form_ExcelUtilForm()
        form_AS_tablo.Show()
    End Sub
End Class