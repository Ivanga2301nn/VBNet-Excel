Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Internal.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Imports Autodesk.AutoCAD.PlottingServices
Imports System.Collections.Generic

Public Class Class_Skari_Kanali
    Dim PI As Double = 3.1415926535897931
    Dim cu As CommonUtil = New CommonUtil()
    Dim brTabla As Integer = 20
    Dim brTokKryg As Integer = 50
    Dim form_AS_tablo As New Form_Tablo()
    Dim appNameKonso As String = "EWG_KONSO"
    Dim appNameTablo As String = "EWG_TABLO"
    <CommandMethod("Skari_Kanali_New")>
    Public Sub Skari_Kanali_new()
        Dim Form_Skari_Kanali_new As New Form_Skari_Kanali_New()
        Form_Skari_Kanali_new.Show()
    End Sub
End Class
