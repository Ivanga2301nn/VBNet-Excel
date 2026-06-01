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
Imports Autodesk.AutoCAD.ComponentModel
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
        ' 1. Извличаме консуматорите
        Dim extractedConsumers As List(Of strKonsumator) =
            ConsumerExtractor.ExtractSelectedConsumers()
        ' 2. Генерираме токовите кръгове и ги НАЛИВАМЕ ДИРЕКТНО в глобалния източник
        AppSettings.ListTokow = ConsumerExtractor.CreateTokowList(extractedConsumers)
        ' 3. Отваряме формата "чиста" - тя сама ще си вземе данните от AppSettings
        Dim frm As New Form_Tablo_new(extractedConsumers)
        frm.ShowDialog()
    End Sub
End Module
Public Module AppSettings
    Public Property CurrentManufacturer As String = "Schneider"
    Public Const ROOT_NODE_TEXT As String = "Гл.Р.Т."

    ' ЦЕНТРАЛНИЯТ ИЗТОЧНИК НА ИСТИНАТА ЗА ЦЕЛИЯ ПРОЕКТ:    Public Property ListTokow As New List(Of clsTokow)
    Public Property ListTokow As New List(Of clsTokow)

    ' Флаг, който спира събитията на Grid-а, докато трае зареждането на данни
    Public Property IsGridLoading As Boolean = False
End Module
Public Class Form_Tablo_new
    ' --- Данни за извлечените от AutoCAD консуматори и токови кръгове ---
    Private ListKonsumator As New List(Of strKonsumator)

    ' --- КАТАЛОЗИ (Глобални за формата, за да живеят през цялото време) ---
    Private _motorCatalog As MotorProtectionCatalog
    Private _cableCatalog As CableCatalog
    Private _breakerCatalog As BreakerCatalog
    Private _disconnectorCatalog As DisconnectorCatalog
    Private _rcdCatalog As RCDCatalog

    ' Създаваме инстанция на електрическите класове, която ще се грижи за всички изчисления в проекта
    Private _calculationEngine As ElectricalCalculationEngine
    Private _boardStructureManager As BoardStructureManager
    Private _panelBalanceManager As PanelBalanceManager

    Private _DataGridViewManager As DataGridViewManager
    Private _gridChangeManager As DataGridViewChangeManager

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
    Public WithEvents _treeViewManager As TreeViewManager
    ''' <summary>
    ''' Конструктор на формата - приема данните от AutoCAD
    ''' </summary>
    Public Sub New(ByVal consumersList As List(Of strKonsumator))
        ' Записваме подадените списъци
        ListKonsumator = consumersList
        InitializeComponent()
        ' 1. ПЪРВО създаваме каталозите в паметта, за да са готови
        InitializeProjectComponents()
    End Sub
    Private Sub Form_Tablo_new_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Height = 950
        Me.Width = 1600
        ' Извикваме фабриката за класове
        'InitializeProjectComponents()
        ' каталозите вече ще съществуват в паметта и няма да има NullReference!)
        FillManufacturerCombo()
        ' Подаваме списъка за изчисление
        _calculationEngine.ExecuteCalculations()
        _boardStructureManager.SortListTokow()
        _boardStructureManager.GroupContactsForRCD()
        _boardStructureManager.EnsureAllStructureRecords()
        _panelBalanceManager.AddFeederRecords()

        _treeViewManager.RefreshTree()

        _DataGridViewManager.InitializeGridStructure()
        AppSettings.IsGridLoading = True
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
                                                             _cableCatalog,
                                                             _rcdCatalog)
        _boardStructureManager = New BoardStructureManager(_rcdCatalog)
        _panelBalanceManager = New PanelBalanceManager(_rcdCatalog,
                                                       _disconnectorCatalog,
                                                       _cableCatalog,
                                                       _calculationEngine)

        ' Йерархия и дървовидна структура (TreeView)
        _treeViewManager = New TreeViewManager(TreeView_Табло)

        _gridChangeManager = New DataGridViewChangeManager(_breakerCatalog,
                                                           _disconnectorCatalog,
                                                           _rcdCatalog,
                                                           _cableCatalog,
                                                           _calculationEngine)


        _DataGridViewManager = New DataGridViewManager(DataGridView1,
                                                       _disconnectorCatalog,
                                                       _breakerCatalog,
                                                       _cableCatalog,
                                                       _rcdCatalog,
                                                       _gridChangeManager)

    End Sub
#Region "⚙️ СВОЙСТВА ЗА ДОСТЪП ДО КАТАЛОЗИ И МЕНИДЖЪРИ"
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
    Public ReadOnly Property CalculationEngine As ElectricalCalculationEngine
        Get
            Return _calculationEngine
        End Get
    End Property
    Public ReadOnly Property BoardStructureManager As BoardStructureManager
        Get
            Return _boardStructureManager
        End Get
    End Property
    Public ReadOnly Property PanelBalanceManager As PanelBalanceManager
        Get
            Return _panelBalanceManager
        End Get
    End Property
#End Region
    ''' <summary>
    ''' Пълни ToolStripComboBox-а за производител
    ''' </summary>
    Private Sub FillManufacturerCombo()
        TscboManufacturer.ComboBox.Items.Clear()
        For Each brand As String In _breakerCatalog.Brand_For_combo
            TscboManufacturer.ComboBox.Items.Add(brand)
        Next
        If TscboManufacturer.ComboBox.Items.Count > 0 Then
            TscboManufacturer.ComboBox.SelectedIndex = 0
        End If
    End Sub
    Private Sub OnManufacturerChanged()
        UpdateGridCombosForNewManufacturer(_currentManufacturer)
    End Sub
    Private Sub TscboManufacturer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TscboManufacturer.SelectedIndexChanged
        Dim selectedBrand As String = TscboManufacturer.ComboBox.SelectedItem?.ToString()
        If Not String.IsNullOrEmpty(selectedBrand) Then
            CurrentManufacturer = selectedBrand
        End If
    End Sub
    Private Sub UpdateGridCombosForNewManufacturer(brandName As String)
        If _breakerCatalog Is Nothing Then Exit Sub
        _breakerCatalog.FilterComboLists(brandName)
    End Sub
    ' ============================================================
    ' ОБРАБОТКА НА СЪБИТИЯ ОТ TREEVIEW (DRAG & DROP)
    ' ============================================================
    ''' <summary>
    ''' Диригентски метод: Улавя преместването на обекти в дървото 
    ''' и отразява промените в централния ListTokow.
    ''' </summary>
    Private Sub _treeViewManager_RequestMoveObject(ByVal source As clsTokow, ByVal target As clsTokow) Handles _treeViewManager.RequestMoveObject
        If source Is Nothing OrElse target Is Nothing Then Exit Sub
        If source.Device = "Табло" Then
            ' Местим ЦЯЛО ТАБЛО (Промяна на захранващата структура)
            If target.Device = "Tablo" OrElse target.Device = "Табло" Then
                source.Табло_Родител = target.Tablo
                source.BuildingName = target.BuildingName
            ElseIf target.Device = "Сграда" Then
                source.Табло_Родител = ""
                source.BuildingName = target.BuildingName
            End If

        ElseIf source.Device = "Консуматор" Then
            ' Местим ТОКОВ КРЪГ (Прекачване от едно табло в друго)
            If target.Device = "Tablo" OrElse target.Device = "Табло" Then
                source.Tablo = target.Tablo
                source.BuildingName = target.BuildingName
            ElseIf target.Device = "Сграда" Then
                source.Tablo = ""
                source.Табло_Родител = ""
                source.BuildingName = target.BuildingName
            End If
        End If

        _panelBalanceManager.AddFeederRecords()
        ' Преначертаваме дървото
        _treeViewManager.RefreshTree()
    End Sub
    ' ============================================================
    ' ОБРАБОТКА НА СЪБИТИЕ: ЛЯВ КЛИК ВЪРХУ ВЪЗЕЛ В ДЪРВОТО
    ' ============================================================
    ''' <summary>
    ''' Изпълнява се, когато потребителят щракне с левия бутон на мишката върху възел.
    ''' Диригентът (формата) получава обекта clsTokow и решава как да реагира.
    ''' </summary>
    Private Sub _treeViewManager_NodeLeftClick(ByVal selectedObject As clsTokow) Handles _treeViewManager.NodeLeftClick
        ' Защита: Ако по някаква причина обектът е празен, излизаме безопасно
        If selectedObject Is Nothing Then Exit Sub
        AppSettings.IsGridLoading = True
        _DataGridViewManager.DisplayBoardStructure(selectedObject)
        GroupBox2.Text = "Детайли за табло -> " + selectedObject.Tablo
        ' Разпределяме логиката според това какъв обект е кликнат:
        Select Case selectedObject.Device
            Case "Табло"
            ' --------------------------------------------------------
            ' ПОТРЕБИТЕЛЯТ Е КЛИКНАЛ ВЪРХУ ТАБЛО (Ред "ОБЩО")
            ' --------------------------------------------------------
            ' Тук формата знае, че е избрано цяло табло (напр. "Т-1")
            ' Можеш да извлечеш името му чрез: selectedObject.Tablo
            ' Пример за бъдеща логика: Филтриране на таблицата само за това табло.

            Case "Консуматор"
            ' --------------------------------------------------------
            ' ПОТРЕБИТЕЛЯТ Е КЛИКНАЛ ВЪРХУ КОНКРЕТЕН ТОКОВ КРЪГ
            ' --------------------------------------------------------
            ' Тук формата знае кой токов кръг е избран (напр. "333")
            ' Можеш да извлечеш името/номера му чрез: selectedObject.ТоковКръг
            ' Пример за бъдеща логика: Позициониране на фокуса в Grid-а върху този кръг.

            Case "Сграда"
                ' --------------------------------------------------------
                ' ПОТРЕБИТЕЛЯТ Е КЛИКНАЛ ВЪРХУ КОРЕНА НА СГРАДАТА
                ' --------------------------------------------------------
                ' Избрана е самата сграда (напр. "Сграда_Block")
                ' Можеш да извлечеш името чрез: selectedObject.BuildingName

        End Select
        AppSettings.IsGridLoading = False
    End Sub
    ' =========================================================================
    ' ОБРАБОТКА НА ПРОМЕНИТЕ ОТ DATAGRIDVIEW (ОБНОВЯВАНЕ НА CLSTOKOW)
    ' =========================================================================
    ''' <summary>
    ''' Събитие за улавяне на промяна в клетка на Grid-a.
    ''' Събира суровата информация и я предава към DataGridViewManager.
    ''' </summary>
    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        ' 1. Защита при зареждане и невалидни индекси
        If AppSettings.IsGridLoading Then Exit Sub
        If e.RowIndex < 0 OrElse e.ColumnIndex < 0 Then Exit Sub
        ' 2. Вземаме името на променената колона
        Dim columnName As String = DataGridView1.Columns(e.ColumnIndex).Name
        ' Пропускаме служебните колони директно на ниво интерфейс
        If columnName = "colParameter" OrElse columnName = "colUnit" Then Exit Sub
        ' 3. Вземаме въведената нова стойност от клетката
        Dim cellValue As Object = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
        Dim newValue As String = If(cellValue IsNot Nothing, cellValue.ToString(), "")
        ' 4. ПРЕДАВАМЕ ВСИЧКО НА МЕНИДЖЪРА
        _DataGridViewManager.ProcessCellValueChanged(e.RowIndex, columnName, newValue)
    End Sub
    ''' <summary>
    ''' Принуждава ComboBox и CheckBox да реагират ВЕДНАГА при избор/цъкане, а не чак при излизане от клетката.
    ''' </summary>
    Private Sub DataGridView1_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs) Handles DataGridView1.CurrentCellDirtyStateChanged
        If TypeOf DataGridView1.CurrentCell Is DataGridViewComboBoxCell OrElse
           TypeOf DataGridView1.CurrentCell Is DataGridViewCheckBoxCell Then
            DataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub
End Class

