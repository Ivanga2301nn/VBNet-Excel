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
Public Module AppSettings
    Public Property CurrentManufacturer As String = "Schneider"
End Module
Public Class Form_Tablo_new
    ' --- Данни за извлечените от AutoCAD консуматори и токови кръгове ---
    Private ListKonsumator As New List(Of strKonsumator)
    Private ListTokow As New List(Of strTokow)

    ' --- КАТАЛОЗИ (Глобални за формата, за да живеят през цялото време) ---
    Private _motorCatalog As MotorProtectionCatalog
    Private _cableCatalog As CableCatalog
    Private _breakerCatalog As BreakerCatalog
    Private _disconnectorCatalog As DisconnectorCatalog
    Private _rcdCatalog As RCDCatalog

    ' Създаваме инстанция на изчислителния двигател
    Private _calculationEngine As ElectricalCalculationEngine
    Private _boardStructureManager As BoardStructureManager

    ' --- ГЛОБАЛНОТО СЪСТОЯНИЕ ЗА МАРКАТА ---
    ' Полето е Shared
    Private _currentManufacturer As String = AppSettings.CurrentManufacturer

    ' СВОЙСТВОТО СЪЩО ТРЯБВА ДА Е SHARED:
    Public Property CurrentManufacturer As String
        Get
            Return _currentManufacturer
        End Get
        Set(value As String)
            If _currentManufacturer <> value Then
                _currentManufacturer = value
                ' Внимание: Методът OnManufacturerChanged също ще трябва да стане Shared!
                OnManufacturerChanged()
            End If
        End Set
    End Property
    ' --- МЕНИДЖЪРИ НА ИНТЕРФЕЙСА ---
    Private _treeViewManager As Form_Tablo_new_TreeViewManager

    ''' <summary>
    ''' Конструктор на формата - приема данните от AutoCAD
    ''' </summary>
    Public Sub New(ByVal consumersList As List(Of strKonsumator), ByVal extractedTokowList As List(Of strTokow))
        InitializeComponent()
        ' 1. ПЪРВО създаваме каталозите в паметта, за да са готови
        InitializeProjectComponents()
        ' Записваме подадените списъци
        ListKonsumator = consumersList
        ListTokow = extractedTokowList
    End Sub

    Private Sub Form_Tablo_new_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Height = 950
        Me.Width = 1600
        ' Извикваме фабриката за класове
        InitializeProjectComponents()
        ' каталозите вече ще съществуват в паметта и няма да има NullReference!)
        FillManufacturerCombo()
        ' Подаваме списъка за изчисление
        _calculationEngine.ExecuteCalculations(ListTokow)
        _boardStructureManager.SortListTokow(ListTokow)

    End Sub
    ''' <summary>
    ''' Инициализация на компонентите на проекта
    ''' </summary>
    Private Sub InitializeProjectComponents()
        ' Инициализираме обектите на ниво клас.
        ' Когато се изпълни "New()", в самите каталози автоматично ще се извика техния LoadCatalog()

        _motorCatalog = New MotorProtectionCatalog()
        _cableCatalog = New CableCatalog()
        _breakerCatalog = New BreakerCatalog()
        _disconnectorCatalog = New DisconnectorCatalog()
        _rcdCatalog = New RCDCatalog()

        _calculationEngine = New ElectricalCalculationEngine(_breakerCatalog,
                                                             _motorCatalog,
                                                             _disconnectorCatalog,
                                                             _cableCatalog,
                                                             _rcdCatalog
                                                             )
        _boardStructureManager = New BoardStructureManager()

        ' Йерархия и дървовидна структура (TreeView)
        _treeViewManager = New Form_Tablo_new_TreeViewManager(TreeView_Табло, ListTokow)

        ' Масово добавяне на кръгове (Batch Add Circuits)
        'Dim batchAddCircuits As New Form_BatchAddCircuits(ListTokow)

        ' Автоматично генериране в AutoCAD (AutoCAD Inserter)
        'Dim autoCadInserter As New Form_Tablo_new_AutoCadInserter(ListTokow)

        ' Файлова логика и управление на проекти (Project Path Resolver)
        'Dim projectPathResolver As New Form_Tablo_new_ProjectPathResolver()
    End Sub
    ' --- СВОЙСТВА (Properties) за достъп от външни изчислителни класове ---
    ' Ако утре направиш друг клас, който ще смята, той ще иска достъп до тези каталози през формата:
    Public ReadOnly Property CableCatalog As CableCatalog
        Get
            Return _cableCatalog
        End Get
    End Property
    Public ReadOnly Property BreakerCatalog As BreakerCatalog
        Get
            Return _breakerCatalog
        End Get
    End Property
    Public ReadOnly Property MotorCatalog As MotorProtectionCatalog
        Get
            Return _motorCatalog
        End Get
    End Property
    Public ReadOnly Property DisconnectorCatalog As DisconnectorCatalog
        Get
            Return _disconnectorCatalog
        End Get
    End Property
    Public ReadOnly Property RCDCatalog As RCDCatalog
        Get
            Return _rcdCatalog
        End Get
    End Property
    ''' <summary>
    ''' Пълни ToolStripComboBox-а за производител с данните от каталога за прекъсвачи.
    ''' </summary>
    Private Sub FillManufacturerCombo()
        ' Достъпваме директно вътрешния ComboBox чрез .ComboBox
        TscboManufacturer.ComboBox.Items.Clear()
        ' Вземаме марките от каталога на прекъсвачите
        For Each brand As String In _breakerCatalog.Brand_For_combo
            TscboManufacturer.ComboBox.Items.Add(brand)
        Next
        ' Избираме първата марка по подразбиране
        If TscboManufacturer.ComboBox.Items.Count > 0 Then
            TscboManufacturer.ComboBox.SelectedIndex = 0
        End If
    End Sub
    ''' <summary>
    ''' Метод, който се извиква при смяна на марката. 
    ''' Тук ще добавим логика за обновяване на UI и данни.
    '''
    ''' Този метод се грижи за всичко, когато марката се смени глобално
    ''' </summary>
    Private Sub OnManufacturerChanged()
        ' 1. Казваме на DataGridView1 да си пренареди Combo клетките за прекъсвачи
        ' 2. Казваме на ДТЗ (RCD) частта да се филтрира по новата марка
        ' 3. Казваме на Товаровите прекъсвачи (Разединителите) да превключат

        UpdateGridCombosForNewManufacturer(_currentManufacturer)
    End Sub
    ' --- СЪБИТИЕТО НА TOOLSTRIPCOMBOBOX-А СТАВА СУПЕР КРАТКО ---
    Private Sub TscboManufacturer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TscboManufacturer.SelectedIndexChanged
        Dim selectedBrand As String = TscboManufacturer.ComboBox.SelectedItem?.ToString()
        If Not String.IsNullOrEmpty(selectedBrand) Then
            ' Просто променяме глобалното свойство! То само ще свърши останалото.
            CurrentManufacturer = selectedBrand
        End If
    End Sub
    ''' <summary>
    ''' Метод за обновяване на Combo клетките в DataGridView1 при смяна на марката.
    ''' Тук ще се филтрират прекъсвачите от каталога.
    ''' </summary>  
    Private Sub UpdateGridCombosForNewManufacturer(brandName As String)
        ' 🛡️ ЗАЩИТА: Ако каталозите още не са готови, излез кротко!
        If _breakerCatalog Is Nothing Then Exit Sub
        ' Филтрираме списъците за комбо кутиите според избраната марка
        _breakerCatalog.FilterComboLists(brandName)
    End Sub
End Class

