Imports System.Collections.Generic
Imports System.Diagnostics.Eventing.Reader
Imports System.Drawing.Drawing2D
Imports System.IO
Imports System.Net
Imports System.Security.Cryptography
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.TrackBar
Imports System.Windows.Input
Imports ACSMCOMPONENTS24Lib
Imports Autodesk.AutoCAD
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.DatabaseServices.Filters
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.GraphicsInterface
Imports Autodesk.AutoCAD.Internal
Imports Autodesk.AutoCAD.Runtime
Imports AXDBLib
Imports Microsoft.Office.Interop.Excel
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application
Imports excel = Microsoft.Office.Interop.Excel

Public Class SheetPrinter
    <CommandMethod("SheetPrinter")>
    Sub RunSheetPrinter()
        '1. Свържи се с текущия AutoCAD документ
        Dim name_file As String = Application.DocumentManager.MdiActiveDocument.Name    ' Взимаме името на текущо отворения документ в AutoCAD
        Dim File_Path As String = Path.GetDirectoryName(name_file)                      ' Взимаме пътя до директорията, в която се намира текущия файл
        Dim File_name As String = Path.GetFileName(name_file)                           ' Взимаме само името на файла от пълния път



        '2. Създай екземпляр на Form_Print
        '3. Извикай LoadCurrentSheetSet() -> List<Sheet>
        '4. Попълни TreeView на формата с листовете
        '5. Покажи формата (ShowDialog или Show)
    End Sub
End Class
