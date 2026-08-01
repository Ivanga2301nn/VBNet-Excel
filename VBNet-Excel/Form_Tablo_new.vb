Imports System.Text.RegularExpressions
Imports System.Windows.Forms
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


    ' =================================================================
    ' === ГЛОБАЛНИ ФЛАГОВЕ ЗА СЪСТОЯНИЕ НА ФОРМАТА ===
    ' =================================================================

    Private _gridUpdateDepth As Integer = 0

    ''' <summary>
    ''' True когато програмно потискаме Grid събитията.
    ''' </summary>
    Public ReadOnly Property IsGridLoading As Boolean
        Get
            Return _gridUpdateDepth > 0
        End Get
    End Property

    Public Sub EnterGridUpdate()
        System.Threading.Interlocked.Increment(_gridUpdateDepth)
    End Sub

    Public Sub ExitGridUpdate()
        If System.Threading.Interlocked.Decrement(_gridUpdateDepth) < 0 Then
            System.Threading.Interlocked.Exchange(_gridUpdateDepth, 0)
        End If
    End Sub

    ''' <summary>
    ''' Глобален флаг срещу рекурсия при обработка на RCD логика.
    ''' </summary>
    Public Property IsProcessingRcd As Boolean = False
End Module
Public Class Form_Tablo_new
    ' --- ГЛОБАЛНОТО СЪСТОЯНИЕ ЗА МАРКАТА ---
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
        AppSettings.EnterGridUpdate()
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
    ''' <summary>
    ''' Обработва промяната на избора в ToolStripComboBox `TscboManufacturer`.
    ''' Взема текущия избран елемент от вътрешния `ComboBox` и, ако не е празен,
    ''' го записва в глобалната/полето `CurrentManufacturer`.
    ''' </summary>
    ''' <param name="sender">Обектът-източник на събитието.</param>
    ''' <param name="e">Аргументи на събитието.</param>
    Private Sub TscboManufacturer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TscboManufacturer.SelectedIndexChanged
        ' Вземи избрания елемент като низ (пази се от Null чрез ?).
        Dim selectedBrand As String = TscboManufacturer.ComboBox.SelectedItem?.ToString()
        ' Ако има валидна (непразна) стойност — присвои я.
        If Not String.IsNullOrEmpty(selectedBrand) Then
            CurrentManufacturer = selectedBrand
        End If
    End Sub
    ''' <summary>
    ''' Обновява стойностите на Combo списъците в DataGridView за нов избран производител.
    ''' Ако каталогът на прекъсвачите не е инициализиран — прекратява изпълнението.
    ''' Делегира филтрирането към `AppSettings.BreakerCatalog.FilterComboLists`.
    ''' </summary>
    ''' <param name="brandName">Името на производителя, по което да се филтрират списъците.</param>
    Private Sub UpdateGridCombosForNewManufacturer(brandName As String)
        ' Ако няма наличен каталог — нищо не правим
        If AppSettings.BreakerCatalog Is Nothing Then Exit Sub
        ' Може да се добави проверка за празен/Null brandName:
        If String.IsNullOrWhiteSpace(brandName) Then Exit Sub
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
        AppSettings.EnterGridUpdate()
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
        AppSettings.ExitGridUpdate()
    End Sub
    ' =========================================================================
    ' ОБРАБОТКА НА ПРОМЕНИТЕ ОТ DATAGRIDVIEW (ОБНОВЯВАНЕ НА CLSTOKOW)
    ' =========================================================================
    ''' <summary>
    ''' Събитие за улавяне на промяна в клетка на Grid-a.
    ''' Събира суровата информация и я предава към DataGridViewManager.
    ''' </summary>
    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        If AppSettings.IsGridLoading Then Exit Sub
        If e.RowIndex < 0 OrElse e.ColumnIndex < 0 Then Exit Sub
        Dim columnName As String = DataGridView1.Columns(e.ColumnIndex).Name
        If columnName = "colParameter" OrElse columnName = "colUnit" Then Exit Sub
        Dim cellValue As Object = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
        Dim newValue As String = If(cellValue IsNot Nothing, cellValue.ToString(), "")

        AppSettings.EnterGridUpdate()
        Try
            AppSettings.DataGridViewManager.ProcessCellValueChanged(e.RowIndex, e.ColumnIndex, newValue)
        Finally
            AppSettings.ExitGridUpdate()
        End Try
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
    ''' <summary>
    ''' Обработва клик върху бутонa "Вмъкни Autocad".
    ''' Предава частите от селектирания път в TreeView към `AutoCadInserter`
    ''' за по-нататъшна обработка (напр. вмъкване на обекти в AutoCAD).
    ''' </summary>
    ''' <param name="sender">Източник на събитието.</param>
    ''' <param name="e">Аргументи на събитието.</param>
    Private Sub ToolStripButton_Вмъкни_Autocad_Click(sender As Object, e As EventArgs) Handles ToolStripButton_Вмъкни_Autocad.Click
        AppSettings.AutoCadInserter.ExecuteInsert(GetSelectedTreePathParts())
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
                    AppSettings.ExitGridUpdate()
                    ' Викаме мениджъра на таблицата.
                    ' Той ще прочете новосортирания списък 
                    ' и ще пренареди редовете на екрана веднага!
                    AppSettings.DataGridViewManager.DisplayBoardStructure(selectedObject)
                    ' Сваляме флага обратно
                    AppSettings.ExitGridUpdate()
                End If
            End If
        End Using
    End Sub
    ''' <summary>
    ''' Обработва клик върху бутонa "ШИНА".
    ''' Предава частите от селектирания път в TreeView към `BoardStructureManager`
    ''' за по-нататъшна обработка (напр. поставяне/премахване на шинна връзка).
    ''' </summary>
    ''' <param name="sender">Източник на събитието.</param>
    ''' <param name="e">Аргументи на събитието.</param>
    Private Sub ToolStripButton_ШИНА_Click(sender As Object, e As EventArgs) Handles ToolStripButton_ШИНА.Click
        ' Вземи текущо избраните части от пътя в дървото (например табло / подсекция)
        Dim pathParts = GetSelectedTreePathParts()
        ' Предай ги на BoardStructureManager за обработка на шината/връзките.
        AppSettings.BoardStructureManager.ProcessBusAndSwitch(pathParts)
    End Sub
    ''' <summary>
    ''' Връща масив с имената от пътя на избрания възел в TreeView,
    ''' като премахва иконите, мощността в скоби и излишните интервали.
    ''' </summary>
    ''' <returns>
    ''' Масив от почистени имена или Nothing, ако няма избран възел.
    ''' </returns>
    Private Function GetSelectedTreePathParts() As String()
        ' Вземаме избрания възел
        Dim selectedNode As TreeNode = TreeView_Табло.SelectedNode
        If selectedNode Is Nothing Then Return Nothing
        ' Разделяме пътя по "\"
        Dim rawPathParts As String() = selectedNode.FullPath.Split(
            New Char() {"\"c},
            StringSplitOptions.RemoveEmptyEntries)
        ' Списък с почистените имена
        Dim cleanParts As New List(Of String)
        For Each part As String In rawPathParts
            ' Премахваме мощността в скоби
            Dim cleanStr As String = Regex.Replace(part, "\s*\(.*?\)", "")
            ' Премахваме Unicode иконите
            cleanStr = Regex.Replace(cleanStr, "[^а-яА-Яa-zA-Z0-9_\.\-\s]", "")
            ' Премахваме водещи и следващи интервали
            cleanStr = cleanStr.Trim()
            If cleanStr <> "" Then
                cleanParts.Add(cleanStr)
            End If
        Next
        If cleanParts.Count = 0 Then Return Nothing
        Return cleanParts.ToArray()
    End Function
End Class
