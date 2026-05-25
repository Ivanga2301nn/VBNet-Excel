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
Imports Autodesk.AutoCAD.Internal
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
' Form_Tablo_new_strTokow | Form_Tablo_new_strTokow.vb
'    → Отговорност: Основни структури от данни на проекта. Съдържа класовете strTokow 
'      (токов кръг) и strKonsumator (консуматор), изчистени и подготвени за Data Binding и клониране.
'
' Form_Tablo_CatalogManager | Form_Tablo_CatalogManager.vb
'    → Отговорност: Централна техническа библиотека (Каталози). Съдържа класовете BreakerCatalog 
'      (MCB, MCCB, ACB прекъсвачи) и CableCatalog (кабели, сечения, монтаж). Предстои добавяне на още каталози.
'
' Form_Tablo_new_BatchAddCircuits | Form_Tablo_new_BatchAddCircuits.vb
'    → Отговорност: Обработка на масово добавяне на кръгове (strTokow) резерва и съществуващи.
'      Създава се като отделна форма с фокус върху UX за тази конкретна задача.
'
' Form_Tablo_new_AutoCadInserter | Form_Tablo_new_AutoCadInserter.vb
'    → Отговорност:Този VB.NET клас съдържа част от логиката на модул за
'    автоматизация в среда на AutoCAD,
'    чиято основна цел е автоматично генериране и изчертаване на еднолинейни схеми
'    на електрически табла по предварително подадени данни
'    (структуриран списък от токови кръгове List(Of strTokow)).
'
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
'
#End Region

' ============================================================
' 1. КОМАНДА ЗА СТАРТИРАНЕ (Трябва да е извън класа на формата)
' ============================================================
Public Module AcadCommands
    <CommandMethod("Tablo_new", CommandFlags.UsePickSet)>
    Public Sub StartTabloForm()
        ' Извикване на метода и записване на резултата в списък
        Dim extractedConsumers As List(Of strKonsumator) =
            Form_Tablo_new_ConsumerExtractor.ExtractSelectedConsumers()
        ' 2. Подаваме извлечените консуматори и получаваме списък
        Dim extractedTokowList As List(Of strTokow) =
            Form_Tablo_new_ConsumerExtractor.CreateTokowList(extractedConsumers)
        ' 3. Подаваме И ДВАТА списъка на формата
        Dim frm As New Form_Tablo_new(extractedConsumers, extractedTokowList)
        frm.ShowDialog()
    End Sub
End Module
Public Class Form_Tablo_new
#Region "КЛАС: MotorProtectionCatalog (Моторни защити - GV)"
    'Тази структура описва един електрически консуматор (товар),
    ' извлечен от блок в AutoCAD чертеж.
    Private ListKonsumator As New List(Of strKonsumator)
    ' Списък за токовите кръгове
    Dim ListTokow As New List(Of strTokow)
#End Region
    Public Sub New(ByVal consumersList As List(Of strKonsumator), ByVal extractedTokowList As List(Of strTokow))
        ' Този ред е задължителен за инициализиране на контролите по формата (дизайнера)
        InitializeComponent()
        ' Записваме подадения списък в нашата променлива, за да го ползваме навсякъде във формата
        ListKonsumator = consumersList
        ListTokow = extractedTokowList
    End Sub
    Private Sub Form_Tablo_new_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Height = 950
        Me.Width = 1600
        ' Извикваме фабриката за класове
        InitializeProjectComponents()
    End Sub
    Private Sub InitializeProjectComponents()
        ' Инициализация на компонентите на проекта
        ' Каталози (техническа библиотека)
        ' КЛАС: MotorProtectionCatalog (Моторни защити - GV)
        Dim catalogManager As New MotorProtectionCatalog()
        ' КЛАС: CableCatalog (Кабели)
        Dim catalogCable As New CableCatalog()
        ' КЛАС: BreakerCatalog (Прекъсвачи)
        Dim catalogBreaker As New BreakerCatalog()
        ' Йерархия и дървовидна структура (TreeView)
        Dim treeViewManager As New Form_Tablo_new_TreeViewManager(TreeView_Табло, ListTokow)
        ' Масово добавяне на кръгове (Batch Add Circuits)
        'Dim batchAddCircuits As New Form_BatchAddCircuits(ListTokow)
        ' Автоматично генериране в AutoCAD (AutoCAD Inserter)
        'Dim autoCadInserter As New Form_Tablo_new_AutoCadInserter(ListTokow)
        ' Файлова логика и управление на проекти (Project Path Resolver)
        'Dim projectPathResolver As New Form_Tablo_new_ProjectPathResolver()
    End Sub
End Class


