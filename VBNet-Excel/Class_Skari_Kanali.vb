Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Internal.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Imports Autodesk.AutoCAD.PlottingServices
Imports System.Collections.Generic

Public Class Class_Skari_Kanali
    <CommandMethod("Skari_Kanali_New")>
    Public Sub Skari_Kanali_new()
        Dim Form_Skari_Kanali_new As New Form_Skari_Kanali_New()
        Form_Skari_Kanali_new.Show()
    End Sub
End Class
