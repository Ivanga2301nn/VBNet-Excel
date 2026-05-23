Imports System.Collections.Generic
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Linq
Imports System.Security.Cryptography
Imports System.Security.Cryptography.X509Certificates
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Autodesk.AutoCAD
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.DatabaseServices.Filters
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.GraphicsSystem
Imports Autodesk.AutoCAD.Internal.DatabaseServices
Imports Autodesk.AutoCAD.PlottingServices
Imports Autodesk.AutoCAD.Runtime
Imports AXDBLib
Imports iTextSharp.text.pdf
Imports Microsoft.Office.Interop.Word
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Org.BouncyCastle.Asn1.Cmp
Imports Org.BouncyCastle.Math.EC.ECCurve
Imports Button = System.Windows.Forms.Button
Imports Font = System.Drawing.Font
#Region "📂 ИНДЕКС: Отделени класове и файлове (Версия 3)"
' 1. Form_Tablo_new_strTokow | Form_Tablo_new_strTokow.vb
'    → Отговорност: Основни структури от данни на проекта. Съдържа класовете strTokow 
'      (токов кръг) и strKonsumator (консуматор), изчистени и подготвени за Data Binding и клониране.
'
' 2. Form_Tablo_CatalogManager | Form_Tablo_CatalogManager.vb
'    → Отговорност: Централна техническа библиотека (Каталози). Съдържа класовете BreakerCatalog 
'      (MCB, MCCB, ACB прекъсвачи) и CableCatalog (кабели, сечения, монтаж). Предстои добавяне на още каталози.
'
' 3. Form_Tablo_new_BatchAddCircuits | Form_Tablo_new_BatchAddCircuits.vb
'    → Отговорност: Обработка на масово добавяне на кръгове (strTokow) резерва и съществуващи.
'      Създава се като отделна форма с фокус върху UX за тази конкретна задача.
'
' 4. Form_Tablo_new_AutoCadInserter | Form_Tablo_new_AutoCadInserter.vb
'    → Отговорност: AutoCAD чертане на табла, шини, кръгове и анотации
'
' 🌟 [🔥 НА ДНЕВЕН РЕД] 🌟
' 5. Form_Tablo_new_ProjectPathResolver | Form_Tablo_new_ProjectPathResolver.vb
'    → Отговорност: Файлова логика (Save/Load), пътища, BuildingName, 
'      сериализация и безопасна обработка на ListTokow (ProcessAndRepairList).
'      Служи като централизирана система за управление на JSON файловете на проекта.
'
' 6. TreeViewManager | Form_Tablo_new_TreeViewManager.vb
'    → Отговорност: Йерархия (Сграда→Табло→TK→Консуматори), Drag&Drop (засега само табла), 
'      синхронизация с UI и ListTokow
'
' 7. [ПЛАНИРАНО] LoadCalculator | Form_Tablo_new_LoadCalculator.vb
'    → Отговорност: Изчисления на токове, избор на ДТЗ/прекъсвачи, фазов баланс
#End Region

' ============================================================
' 1. КОМАНДА ЗА СТАРТИРАНЕ (Трябва да е извън класа на формата)
' ============================================================
Public Module AcadCommands
    <CommandMethod("Tablo_new")>
    Public Sub StartTabloForm()
        Dim frm As New Form_Tablo_new()
        frm.ShowDialog()
    End Sub
End Module
Public Class Form_Tablo_new
    Private Sub Form_Tablo_new_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Height = 950
        Me.Width = 1600
    End Sub

End Class
