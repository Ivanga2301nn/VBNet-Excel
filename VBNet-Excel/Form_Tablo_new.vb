Imports System.Collections.Generic
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Linq
Imports System.Security.Cryptography
Imports System.Security.Cryptography.X509Certificates
Imports System.Text.RegularExpressions
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
    ' =================================================================
    ' === ОБЩИ НАСТРОЙКИ И ИЗТОЧНИК НА ИСТИНАТА ===
    ' =================================================================
    Public Property CurrentManufacturer As String = "Schneider"
    Public Const ROOT_NODE_TEXT As String = "Гл.Р.Т."
    ''' <summary>
    ''' ЦЕНТРАЛНИЯТ ИЗТОЧНИК НА ИСТИНАТА ЗА ЦЕЛИЯ ПРОЕКТ (Всички токови кръгове)
    ''' </summary>
    Public Property ListTokow As New List(Of clsTokow)
    ''' <summary>
    ''' Списък със суровите консуматори, извлечени от AutoCAD
    ''' </summary>
    Public Property ListKonsumator As New List(Of strKonsumator)
    ''' <summary>
    ''' Флаг, който спира събитията на Grid-а, докато трае зареждането на данни
    ''' </summary>
    Public Property IsGridLoading As Boolean = False
    ' =================================================================
    ' === ГЛОБАЛНИ КАТАЛОЗИ
    ' =================================================================
    Public Property RcdCatalog As RCDCatalog
    Public Property BreakerCatalog As BreakerCatalog
    Public Property CableCatalog As CableCatalog
    Public Property DisconnectorCatalog As DisconnectorCatalog
    Public Property MotorProtectionCatalog As MotorProtectionCatalog
    ' =================================================================
    ' === ВСИЧКИ КЛАСОВЕ И МЕНИДЖЪРИ ОТ СНИМКАТА (Живеят мирно тук) ===
    ' =================================================================
    ' Мениджъри за интерфейса (UI) и логиката на Grid / TreeView
    Public Property DataGridViewManager As DataGridViewManager
    Public Property DataGridViewChangeManager As DataGridViewChangeManager
    Public Property TreeViewManager As TreeViewManager
    Public Property FormSortPriority As Form_SortPriority
    ' Инженерни енджини и изчисления
    Public Property ElectricalCalculationEngine As ElectricalCalculationEngine
    Public Property PanelBalanceManager As PanelBalanceManager
    Public Property BoardStructureManager As BoardStructureManager
    ' Мениджъри за данни и интеграция (AutoCAD / Данни)
    Public Property ConsumerExtractor As ConsumerExtractor
    Public Property AutoCadInserter As Form_Tablo_new_AutoCadInserter
    Public Property BatchAddCircuits As Form_BatchAddCircuits
    Public Property ProjectPathResolver As ProjectPathResolver
    Public Property TargetDataGridView As DataGridView
    Public Property TargetTreeView As System.Windows.Forms.TreeView
End Module
Public Class Form_Tablo_new
    ' --- ГЛОБАЛНОТО СЪСТОЯНИЕ ЗА МАРКАТА (Пренасочено към централния склад) ---
    Public Property CurrentManufacturer As String
        Get
            Return AppSettings.CurrentManufacturer
        End Get
        Set(value As String)
            If AppSettings.CurrentManufacturer <> value Then
                AppSettings.CurrentManufacturer = value
                ' Методът се вика, ако имаш логика за преначертаване на интерфейса при смяна на марката
                OnManufacturerChanged()
            End If
        End Set
    End Property
    ' --- МЕНИДЖЪРИ НА ИНТЕРФЕЙСА (Държим WithEvents тук, за да хващаме събитията от TreeView-то) ---
    ' Закачаме го към AppSettings при инициализация
    Public WithEvents _treeViewManager As TreeViewManager

    ''' <summary>
    ''' Конструктор на формата - приема данните от AutoCAD
    ''' </summary>
    Public Sub New(ByVal consumersList As List(Of strKonsumator))
        ' 1. Наливаме извлечените консуматори директно в централния склад
        AppSettings.ListKonsumator = consumersList
        InitializeComponent()
        ' 2. ПЪРВО зареждаме каталозите и събуждаме мениджърите в AppSettings
        InitializeProjectComponents()
        ' 3. Свързваме локалния WithEvents мениджър с този в склада, за да работят събитията (кликове, селекции)
        Me._treeViewManager = AppSettings.TreeViewManager
    End Sub
    Private Sub Form_Tablo_new_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' 1. Размери и позициониране на формата
        Me.Height = 950
        Me.Width = 1600
        Me.StartPosition = FormStartPosition.CenterScreen
        ' 2. Напълване на Комбо бокса с производители (Schneider и т.н.)
        FillManufacturerCombo()
        ' =================================================================
        ' 3. ИНЖЕНЕРНИ ИЗЧИСЛЕНИЯ И ЛОГИКА (Всичко минава през AppSettings!)
        ' =================================================================
        ' Изчисления на токове, мощности и пада на напрежение
        AppSettings.ElectricalCalculationEngine.ExecuteCalculations()
        ' Сортиране на токовите кръгове по приоритети
        AppSettings.FormSortPriority.SortListTokow()
        ' Групиране на ДТЗ / контакти и подреждане на структурата на таблото
        AppSettings.BoardStructureManager.GroupContactsForRCD()
        AppSettings.BoardStructureManager.EnsureAllStructureRecords()
        ' Добавяне на захранващи линии / главни прекъсвачи в баланса на фазите
        AppSettings.PanelBalanceManager.AddFeederRecords()
        ' =================================================================
        ' 4. ВИЗУАЛИЗАЦИЯ НА ИНТЕРФЕЙСА (Дърво и Таблица)
        ' =================================================================
        ' Опресняваме TreeView структурата вляво
        AppSettings.TreeViewManager.RefreshTree()
        ' Инициализираме колоните и структурата на DataGridView (Grid-а)
        AppSettings.DataGridViewManager.InitializeGridStructure()
        ' Вдигаме флага, че зареждането приключи и Grid-ът вече е готов за работа
        AppSettings.IsGridLoading = True
    End Sub
    ''' <summary>
    ''' Инициализация на компонентите на проекта в централния склад AppSettings
    ''' </summary>
    Private Sub InitializeProjectComponents()
        TargetDataGridView = DataGridView1
        TargetTreeView = TreeView_Табло
        ' 1. Инициализираме КАТАЛОЗИТЕ директно в AppSettings.
        ' Когато се изпълни "New()", вътре в тях автоматично ще се зареди LoadCatalog()
        AppSettings.MotorProtectionCatalog = New MotorProtectionCatalog()
        AppSettings.CableCatalog = New CableCatalog()
        AppSettings.BreakerCatalog = New BreakerCatalog()
        AppSettings.DisconnectorCatalog = New DisconnectorCatalog()
        AppSettings.RcdCatalog = New RCDCatalog()

        ' 2. Инициализираме ИНЖЕНЕРНИТЕ ЕНДЖИНИ (Вече с празни конструктори!)
        AppSettings.ElectricalCalculationEngine = New ElectricalCalculationEngine()
        AppSettings.BoardStructureManager = New BoardStructureManager()
        AppSettings.PanelBalanceManager = New PanelBalanceManager()

        ' 3. Инициализираме МЕНИДЖЪРИТЕ ЗА ИНТЕРФЕЙСА
        ' Подаваме им само съответните контроли от формата, за да ги управляват
        AppSettings.TreeViewManager = New TreeViewManager()
        AppSettings.DataGridViewChangeManager = New DataGridViewChangeManager()
        AppSettings.DataGridViewManager = New DataGridViewManager()

        ' 4. Помощни класове и форми
        AppSettings.AutoCadInserter = New Form_Tablo_new_AutoCadInserter()
        AppSettings.FormSortPriority = New Form_SortPriority()

        ' (Ако ProjectPathResolver ти трябва, добави го и него тук)
        AppSettings.ProjectPathResolver = New ProjectPathResolver()
    End Sub
#Region "⚙️ СВОЙСТВА ЗА ДОСТЪП ДО КАТАЛОЗИ И МЕНИДЖЪРИ"
    ' --- СВОЙСТВА (Properties) за достъп от външни изчислителни класове ---
    ' Ако утре направиш друг клас, който ще смята, той ще иска достъп до тези каталози през формата:
    Public ReadOnly Property CableCatalog As CableCatalog
        Get
            Return AppSettings.CableCatalog
        End Get
    End Property
    Public ReadOnly Property BreakerCatalog As BreakerCatalog
        Get
            Return AppSettings.BreakerCatalog
        End Get
    End Property
    Public ReadOnly Property MotorCatalog As MotorProtectionCatalog
        Get
            Return AppSettings.MotorProtectionCatalog
        End Get
    End Property
    Public ReadOnly Property DisconnectorCatalog As DisconnectorCatalog
        Get
            Return AppSettings.DisconnectorCatalog
        End Get
    End Property
    Public ReadOnly Property RCDCatalog As RCDCatalog
        Get
            Return AppSettings.RcdCatalog
        End Get
    End Property
    Public ReadOnly Property CalculationEngine As ElectricalCalculationEngine
        Get
            Return AppSettings.ElectricalCalculationEngine
        End Get
    End Property
    Public ReadOnly Property BoardStructureManager As BoardStructureManager
        Get
            Return AppSettings.BoardStructureManager
        End Get
    End Property
    Public ReadOnly Property PanelBalanceManager As PanelBalanceManager
        Get
            Return AppSettings.PanelBalanceManager
        End Get
    End Property
#End Region
    ''' <summary>
    ''' Пълни ToolStripComboBox-а за производител
    ''' </summary>
    Private Sub FillManufacturerCombo()
        TscboManufacturer.ComboBox.Items.Clear()
        For Each brand As String In AppSettings.BreakerCatalog.Brand_For_combo
            TscboManufacturer.ComboBox.Items.Add(brand)
        Next
        If TscboManufacturer.ComboBox.Items.Count > 0 Then
            TscboManufacturer.ComboBox.SelectedIndex = 0
        End If
    End Sub
    Private Sub OnManufacturerChanged()
        UpdateGridCombosForNewManufacturer(CurrentManufacturer)
    End Sub
    Private Sub TscboManufacturer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TscboManufacturer.SelectedIndexChanged
        Dim selectedBrand As String = TscboManufacturer.ComboBox.SelectedItem?.ToString()
        If Not String.IsNullOrEmpty(selectedBrand) Then
            CurrentManufacturer = selectedBrand
        End If
    End Sub
    Private Sub UpdateGridCombosForNewManufacturer(brandName As String)
        If AppSettings.BreakerCatalog Is Nothing Then Exit Sub
        AppSettings.BreakerCatalog.FilterComboLists(brandName)
    End Sub
    ' ============================================================
    ' ОБРАБОТКА НА СЪБИТИЯ ОТ TREEVIEW (DRAG & DROP)
    ' ============================================================
    ''' <summary>
    ''' Диригентски метод: Улавя преместването на обекти в дървото 
    ''' и отразява промените в централния ListTokow.
    ''' </summary>
    Private Sub _treeViewManager_RequestMoveObject(ByVal source As clsTokow,
                                                   ByVal target As clsTokow) Handles _treeViewManager.RequestMoveObject
        If source Is Nothing OrElse target Is Nothing Then Exit Sub
        If source.Device = "Табло" Then
            If target.Device = "Tablo" OrElse target.Device = "Табло" Then
                source.Табло_Родител = target.Tablo
                source.BuildingName = target.BuildingName
            End If
            If target.Device = "Сграда" Then
                source.Табло_Родител = ""
                source.BuildingName = target.BuildingName
            End If
        End If
        If source.Device = "Консуматор" Then
            If target.Device = "Tablo" OrElse target.Device = "Табло" Then
                source.Tablo = target.Tablo
                source.BuildingName = target.BuildingName
            End If
            If target.Device = "Сграда" Then
                source.Tablo = ""
                source.Табло_Родител = ""
                source.BuildingName = target.BuildingName
            End If
        End If
        AppSettings.PanelBalanceManager.AddFeederRecords()
        AppSettings.TreeViewManager.RefreshTree()
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
        AppSettings.DataGridViewManager.DisplayBoardStructure(selectedObject)
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
        ' 2. Вземаме името на променената колона за бърза проверка на служебните колони
        Dim columnName As String = DataGridView1.Columns(e.ColumnIndex).Name
        If columnName = "colParameter" OrElse columnName = "colUnit" Then Exit Sub
        ' 3. Вземаме въведената нова стойност от клетката
        Dim cellValue As Object = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
        Dim newValue As String = If(cellValue IsNot Nothing, cellValue.ToString(), "")
        ' 4. ПРЕДАВАМЕ ВСИЧКО НА МЕНИДЖЪРА
        AppSettings.IsGridLoading = True
        AppSettings.DataGridViewManager.ProcessCellValueChanged(e.RowIndex, e.ColumnIndex, newValue)
        AppSettings.IsGridLoading = False
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
    Private Sub ToolStripButton_Вмъкни_Autocad_Click(sender As Object, e As EventArgs) Handles ToolStripButton_Вмъкни_Autocad.Click
        ' 1. Вземаме селектирания възел директно от TreeView контрола във формата
        Dim selectedNode As TreeNode = TreeView_Табло.SelectedNode
        If selectedNode Is Nothing Then Exit Sub
        ' 2. Първо разделяме оригиналния път по наклонена черта, за да не счупим нивата
        Dim rawPathParts As String() = selectedNode.FullPath.Split(New Char() {"\"c}, StringSplitOptions.RemoveEmptyEntries)
        ' Списък, в който ще съберем перфектно изчистените и тримнати стрингове
        Dim cleanParts As New List(Of String)()
        ' 3. Въртим цикъл и чистим всяка папка/табло поотделно
        For Each part As String In rawPathParts
            ' А) Махаме мощността в скобите " (XX.XX kW)"
            Dim cleanStr As String = Regex.Replace(part, "\s*\(.*?\)", "")
            ' Б) Махаме Unicode иконите (сгради, табла) отпред
            cleanStr = Regex.Replace(cleanStr, "[^а-яА-Яa-zA-Z0-9_\.\-\s]", "")
            ' В) Премахва коварния интервал, останал в началото след иконата!
            cleanStr = cleanStr.Trim()
            ' Записваме в чистия списък, ако не е празен
            If Not String.IsNullOrEmpty(cleanStr) Then
                cleanParts.Add(cleanStr)
            End If
        Next
        ' 4. Защита: Проверяваме дали масивът ни е валиден
        If cleanParts.Count = 0 Then Exit Sub
        ' Конвертираме обратно към масив от чисти стрингове
        Dim pathParts As String() = cleanParts.ToArray()
        ' 5. Извикваме инсъртера и му подаваме перфектно изчистения масив
        ' Сега pathParts(0) е чистото име на сградата, а последното е чистото име на таблото, без интервали!
        AppSettings.AutoCadInserter.ExecuteInsert(pathParts)
    End Sub
    Private Sub ToolStripButton_Сортиране_Click(sender As Object, e As EventArgs) Handles ToolStripButton_Сортиране.Click
        ' 1. Преди да отворим формата, вземаме текущо избрания възел в дървото
        Dim selectedNode As TreeNode = TreeView_Табло.SelectedNode
        If selectedNode Is Nothing Then Exit Sub
        ' 2. Отваряме формата за сортиране като диалог
        Using frm As New Form_SortPriority()
            If frm.ShowDialog() = DialogResult.OK Then
                ' --- ТУК ФОРМИРАМЕ selectedObject СЛЕД СОРТИРАНЕТО ---
                ' Вземаме актуалния бизнес обект от "джоба" (Tag) на избрания възел
                Dim selectedObject As clsTokow = TryCast(selectedNode.Tag, clsTokow)
                ' Проверяваме дали възелът наистина е валиден обект
                If selectedObject IsNot Nothing Then
                    ' Вдигаме флага, че гридът се зарежда (за да спрем събитията при пълнене)
                    AppSettings.IsGridLoading = True
                    ' Викаме мениджъра на таблицата.
                    ' Той ще прочете новосортирания списък 
                    ' и ще пренареди редовете на екрана веднага!
                    AppSettings.DataGridViewManager.DisplayBoardStructure(selectedObject)
                    ' Сваляме флага обратно
                    AppSettings.IsGridLoading = False
                End If
            End If
        End Using
    End Sub
End Class
