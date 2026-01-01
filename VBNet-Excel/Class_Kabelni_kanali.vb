Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports System.Net
Imports excel = Microsoft.Office.Interop.Excel
Imports System.IO

Public Class Class_Kabelni_kanali

    <CommandMethod("Skari_Kanali")>
    Public Sub Skari_Kanali_()
        Dim Form_Skari_Kanali As New Form_KabelniKanali()
        Form_Skari_Kanali.Show()
    End Sub
End Class
