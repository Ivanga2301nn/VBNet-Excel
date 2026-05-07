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
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Org.BouncyCastle.Asn1.Cmp
Imports Org.BouncyCastle.Math.EC.ECCurve
Imports Button = System.Windows.Forms.Button
Imports Font = System.Drawing.Font

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
#Region "ОСНОВНА ПОСЛЕДОВАТЕЛНОСТ НА ОБРАБОТКА НА ДАННИТЕ"
        ' Този блок от извиквания представлява основния работен поток
        ' на програмата при анализ на електрическите консуматори,
        ' изчисляване на токовете и подготовка на визуализацията
        ' (TreeView и DataGridView).        '
        ' Стъпките се изпълняват в строго определен ред, защото
        ' всяка следваща процедура използва резултатите от предходната.
        ' Инициализира каталозите с електрическа апаратура:
        ' - прекъсвачи (BreakerInfo)
        ' - възможни серии и номинални токове
        ' - други каталожни данни, използвани при оразмеряване
        '
        ' Тази процедура подготвя данните, които по-късно ще се
        ' използват от SelectBreaker() и други функции.
#End Region
        SetCatalog()
#Region "Извличане на консуматорите от AutoCAD"
        ' Обхожда всички избрани блокове в чертежа и:
        ' - прочита техните атрибути
        ' - извлича Dynamic Block properties
        ' - изчислява мощността
        ' - създава обекти strKonsumator
        '
        ' Резултатът се записва в списъка:
        ' ListKonsumator
#End Region
        GetKonsumatori()
        If ListKonsumator.Count = 0 Then
            Dim messages As String() = {
                "Кафе-пауза? Списъкът е празен, няма нищо за вършене тук!",
                "Изпратихме детективи, но не открихме нито един консуматор в чертежа...",
                "Списъкът е толкова празен, колкото хладилник в понеделник сутрин.",
                "Пълна тишина... Списъкът е самотен и празен. " & vbCrLf & "Начертай нещо, за да му вдъхнеш живот!",
                "Гледах наляво, гледах надясно... консуматори няма. " & vbCrLf & "Отивам да почина, докато ги намериш!",
                "Ракетата е готова, но няма пътници! " & vbCrLf & "Добави консуматори в списъка и ще излетим заедно.",
                "Грешка в матрицата: Консуматорите се оказаха илюзия. " & vbCrLf & "Опитай пак, когато реалността се стабилизира!",
                "Списъкът е по-празен от фитнес зала на 1-ви януари!",
                "Нищо за правене... Да отидем за бира?",
                "Консуматорите си взеха отпуск без да кажат.",
                "404: Консуматори не са открити в тази вселена."
            }
            ' Генерираме произволен индекс, за да е изненада всеки път
            Dim rnd As New Random()
            Dim index As Integer = rnd.Next(0, messages.Length)
            ' Показваме избраното съобщение
            MessageBox.Show(messages(index) & vbCrLf & vbCrLf & "Ще затворя прозореца, за да не си пречим.",
                    "Мисията невъзможна",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information)
            Me.Close()
            Return
            ' Сбогом, форма!
            ' Затваряне на формата
            Me.Close()
            ' Използваме Exit Sub (или Return), за да сме сигурни, че 
            ' кодът след този блок няма да се изпълни
            Exit Sub
        End If
#Region "Създаване на списък с токови кръгове"
        ' Групира всички консуматори по:
        ' - табло (ТАБЛО)
        ' - токов кръг (КРЪГ)
        '
        ' За всяка уникална комбинация се създава обект strTokow,
        ' който съдържа всички консуматори от този кръг.
        '
        ' Резултатът се записва в:
        ' ListTokow
#End Region
        CreateTokowList()
#Region "Инициализация на конфигурациите на блоковете"
        ' Създава и попълва списъка BlockConfigs, който описва:
        ' - типовете блокове
        ' - категории (Lamp, Contact, Device)
        ' - стандартни кабели
        ' - стандартни прекъсвачи
        ' - правила според Visibility
        '
        ' Тази информация се използва при анализа на консуматорите.
#End Region
        InitializeBlockConfigs()
#Region "Изчисляване на натоварванията на токовите кръгове"
        ' За всеки токов кръг:
        ' - обработва всички консуматори
        ' - изчислява общата мощност
        ' - изчислява номиналния ток
        ' - избира подходящ прекъсвач
        ' - попълва параметрите на кръга
        '
        ' Това е основната електротехническа част на алгоритъма.
#End Region
        CalculateCircuitLoads()
#Region "Изчисляване на ДТЗ (RCD)"
        ' Определя дали за даден токов кръг е необходимо
        ' дефектнотоково защитно устройство (RCD) и ако е нужно:
        '
        ' - определя типа (AC / A / F)
        ' - определя чувствителността (30mA, 100mA, 300mA)
        ' - определя номиналния ток
        '
        ' Решението се базира на типа на консуматорите
        ' (например контакти изискват ДТЗ).
#End Region
        CalculateRCD()
#Region "Сортиране на токовите кръгове"
        ' Подрежда кръговете в удобен ред за визуализация
        ' и за последващо генериране на табла.
        '
        ' Сортирането може да бъде например по:
        ' - табло
        ' - номер на кръг
        ' - тип товар
#End Region
        SortCircuits()
        ' 8. ✅ ГРУПИРАНЕ НА КОНТАКТИТЕ ПО ДЗТ (НОВО!)
        GroupContactsForRCD()
#Region "Създаване на TreeView структура"
        ' Изгражда йерархично представяне на консуматорите:
        '
        ' Табло
        '   ├── Токов кръг
        '   │     ├── Консуматор 1
        '   │     ├── Консуматор 2
        '   │     └── ...
        '
        ' TreeView се използва за визуална навигация
        ' и бърз преглед на структурата на таблото.
#End Region
        InitializePanelParents(ROOT_NODE_TEXT)
        BuildTreeViewFromKonsumatori()
#Region "Подготовка на DataGridView"
        ' Конфигурира таблицата за редактиране на параметрите
        ' на токовите кръгове:
        '
        ' - създава редове с параметри
        ' - добавя ComboBox клетки
        ' - зарежда стойности от каталозите
        '
        ' DataGridView позволява на потребителя
        ' да променя ръчно параметри като:
        ' - тип прекъсвач
        ' - номинален ток
        ' - ДТЗ
        ' - шина
#End Region
        SetupDataGridView()
        SetupDataGridView_Total()
        calcBreaker = False
    End Sub
    Dim PI As Double = 3.1415926535897931
    Dim cu As CommonUtil = New CommonUtil()
    Private ListKonsumator As New List(Of strKonsumator)
    ' Списък за токовите кръгове
    Dim ListTokow As New List(Of strTokow)
    Private Const ZnakX As String = "х" ' Напиши го веднъж тук (на кирилица)
    ' ============================================================
    ' КАТАЛОЖНИ ПРОМЕНЛИВИ (на ниво форма)
    ' ============================================================
    Private BlockConfigs As New List(Of BlockConfig)
    Private Breakers As New List(Of BreakerInfo)
    Private Cables As New List(Of CableInfo)
    Private Busbars_Cu As New List(Of BusbarInfo)
    Private Busbars_Al As New List(Of BusbarInfo)
    Private RCD_Catalog As New List(Of RCDInfo)
    Private GV_Database As New List(Of GV_Entry)

    Private IcableDict As New Dictionary(Of String, Integer())
    Private Catalog_Cables As New List(Of CableInfo)
    Public Class CableInfo
        Public Property PhaseSize As String         ' "2,5", "4", и т.н.
        Public Property NeutralSize As String       ' "0", "1,5", "2,5", и т.н.
        Public Property MaxCurrent_Air As Double    ' Допустим ток във въздух
        Public Property MaxCurrent_Ground As Double ' Допустим ток в земя
        Public Property Material As String          ' "Cu", "Al"
        Public Property CableType As String         ' "СВТ", "САВТ", "Al/R"
        Public Property MaxWorkingTemp As Double    ' ← ← ← НОВО! (65, 70, 90°C)
        Public Property InsulationType As String    ' ← ← ← НОВО! ("ПВЦ", "XLPE", "GUM")
    End Class
    Public Structure strMountMethod
        Dim Simbol As String
        Dim Text As String
    End Structure
    Dim LiMountMethod As New List(Of strMountMethod)
    Private Cable_AlR_2 As New Dictionary(Of Integer, String)
    Private Cable_AlR_4 As New Dictionary(Of Integer, String)
    ''' <summary>
    ''' Съдържа всички разединители (обекти с параметри)
    ''' Използва се като основна база данни
    ''' </summary>
    Private Disconnectors As New List(Of DisconnectorInfo)
    ''' <summary>
    ''' Списък с кабели (за избор в ComboBox)
    ''' </summary>
    Private Cable_For_combo As List(Of String)
    ''' <summary>
    ''' Списък с автоматични прекъсвачи
    ''' </summary>
    Private Breakers_For_combo As List(Of String)
    ''' <summary>
    ''' Списък с токови защити (Trip Unit)
    ''' </summary>
    Private TripUnit_For_combo As List(Of String)
    ''' <summary>
    ''' Списък с характеристики на прекъсвачи (B, C, D и др.)
    ''' </summary>
    Private Curve_For_combo As List(Of String)
    ''' <summary>
    ''' Списък с типове разединители (за избор)
    ''' </summary>
    Private Disconnectors_For_combo As List(Of String)
    ''' <summary>
    ''' Списък с токове на разединители (например 25A, 40A...)
    ''' </summary>
    Private Discon_Tok_For_combo As List(Of String)
    ''' <summary>
    ''' Глобални променливи за таблата
    ''' </summary>
    Public widthColom As Double = 120      ' Ширина на всяка колона в таблицата
    Public heightRow As Double = 25        ' Височина на редовете
    Public widthText As Double = 140       ' Ширина на колоната за текст (напр. "Токов кръг")
    Public widthTextDim As Double = 40     ' Допълнителна ширина за текстова колона (напр. за единици)
    Public lengthProw As Double = 90       ' Дължина на вертикалните линии между текст и блокове
    Public lengthProwBlock As Double = 0   ' Дължина на линията под блока за прекъсвач (ще се изчислява по-късно)
    Public padingText As Double = 3        ' Отстояние на текста от линиите
    Public widthTablo As Double = 410      ' Ширина на цялото табло (за блокове и линии)
    Public heightText As Double = 12       ' Височина на текста, използван в блоковете
    Public Y_Шина As Double = 620          ' Вертикална позиция на шината (Y координата)
    ''' <summary>
    ''' Структура за дефиниция на линия за чертане
    ''' </summary>
    Private Structure LineDefinition
        Public StartPoint As Point3d
        Public EndPoint As Point3d
        Public Layer As String
        Public LineWeightValue As Integer
        Public LineType As String
        Public ColorIndex As Integer
        Public Sub New(startPoint As Point3d, endPoint As Point3d, layer As String,
                   lineWeightValue As Integer,
                   lineType As String, Optional colorIndex As Integer = -1)
            Me.StartPoint = startPoint
            Me.EndPoint = endPoint
            Me.Layer = layer
            Me.LineWeightValue = lineWeightValue
            Me.LineType = lineType
            Me.ColorIndex = colorIndex
        End Sub
    End Structure
    ' ============================================================
    ' НОВИ ПРОМЕНЛИВИ ЗА СЪСТОЯНИЕТО (State Variables)
    ' ============================================================
    Private twoBus As Boolean = False
    Private hasDisconnector As Boolean = False
    Private Faza_Tablo As Boolean = False
    Private brTokKrygoweNa6ina As Integer = 0
    Private selectedTablo As String = "" ' За да е достъпно във всички процедури
    ' ─────────────────────────────────────────────────────────────
    ' КОНСТАНТИ & ПРОМЕНЛИВИ ЗА ВИЗУАЛНА МАРКИРОВКА
    ' ─────────────────────────────────────────────────────────────
    Private ROOT_NODE_NAME As String = "__ROOT__"
    Private ROOT_NODE_TEXT As String = "Гл.Р.Т."
    Private highlightNode As TreeNode = Nothing
    Private originalBackColor As Color = SystemColors.Window
    Private originalForeColor As Color = SystemColors.WindowText
    ' Флаг, указващ дали трябва да се извърши изчисление на прекъсвача.
    Private calcBreaker As Boolean = True
    'Private isUpdatingGrid As Boolean = False
    ' ============================================================
    ' КАТАЛОЖНИ СТРУКТУРИ
    ' ============================================================
    ' =====================================================
    ' 4. РЕДОВЕ: Параметри с мерни единици и типове клетки
    ' =====================================================
    ' Структура: {Параметър, Мерна единица, Тип клетка}
    ' Тип клетка: "Text", "Combo", "Check"
    Dim rowData As String()() = {
        New String() {"Прекъсвач", "", "Text"},
        New String() {"Изчислен ток", "A", "Text"},
        New String() {"Тип на апарата", "", "Combo"},
        New String() {"Номинален ток", "A", "Combo"},
        New String() {"Изкл. възможн.", "kA", "Text"},
        New String() {"Крива", "", "Combo"},
        New String() {"Защитен блок", "", "Combo"},
        New String() {"Брой полюси", "бр.", "Text"},
        New String() {"ДТЗ (RCD)", "", "Text"},
        New String() {"ДТЗ Нула", "", "Text"},
        New String() {"Вид на апарата", "", "Text"},
        New String() {"Клас на апарата", "", "Text"},
        New String() {"ДТЗ(RCD) Ном. ток", "A", "Text"},
        New String() {"Чувствителност", "mA", "Text"},
        New String() {"ДТЗ(RCD) полюси", "бр.", "Text"},
        New String() {"---------", "", "Text"},
        New String() {"Брой лампи", "бр.", "Text"},
        New String() {"Брой контакти", "бр.", "Text"},
        New String() {"Инст. мощност", "kW", "Text"},
        New String() {"---------", "", "Text"},
        New String() {"Кабел", "", "Text"},
        New String() {"Начин на монтаж", "--", "Combo"},
        New String() {"Начин на полагане", "--", "Combo"},
        New String() {"Паралелни кабели (фаза): ", "бр.", "Text"},
        New String() {"Съседни кабели (група):", "бр.", "Text"},
        New String() {"Тип кабел", "---", "Combo"},
        New String() {"Сечение", "mm²", "Text"},
        New String() {"---------", "", "Text"},
        New String() {"Фаза", "", "Text"},
        New String() {"Консуматор", "---", "Text"},
        New String() {"предназначение", "---", "Text"},
        New String() {"Управление", "---", "Combo"},
        New String() {"---------", "", "Text"},
        New String() {"Шина", "---", "Check"},
        New String() {"Постави ДТЗ (RCD)", "---", "Check"}
    }
    ' Клас за съхранение на техническите данни за моторана защита от каталога
    Public Class GV_Entry
        Public MinCurrent As Double  ' Долна граница на тока (A)
        Public MaxCurrent As Double  ' Горна граница на тока (A)
        Public Type As String        ' Модел (напр. GV2-ME)
        Public MotorPower As String  ' Мощност на двигателя (Pдвиг)
        Public SettingRange As String ' Диапазон на настройка (Наст)

        Public Sub New(min As Double, max As Double, t As String, p As String, n As String)
            Me.MinCurrent = min
            Me.MaxCurrent = max
            Me.Type = t
            Me.MotorPower = p
            Me.SettingRange = n
        End Sub
    End Class
    ' ============================================================
    ' РЕЧНИК ЗА УПРАВЛЕНИЕ -> БЛОК
    ' ============================================================
    Dim ControlBlockMap As New Dictionary(Of String, String) From {
            {"Импулсно реле", "s_tl"},
            {"Контактор", "s_ct_cont_no"},
            {"Моторна защита", "s_tesys_cont_no"},
            {"Моторен механизъм", "s_ns100_motor_fixed"},
            {"Честотен регулатор", "s_altivar"},
            {"Стълбищен автомат", "s_min"},
            {"Електромер", "s_Wh_meter"},
            {"Фото реле", "s_switch_light_sens"}
        }
    ''' <summary>
    ''' Конфигурация за управляващо устройство
    ''' </summary>
    Private Structure ControlDeviceConfig
        Public Str_1 As String
        Public Str_2 As String
        Public Str_3 As String
        Public Str_4 As String
        Public Str_5 As String
        Public ShortName As String
        Public Sub New(str_1 As String, str_2 As String, str_3 As String,
                   str_4 As String, str_5 As String, shortName As String)
            Me.Str_1 = str_1
            Me.Str_2 = str_2
            Me.Str_3 = str_3
            Me.Str_4 = str_4
            Me.Str_5 = str_5
            Me.ShortName = shortName
        End Sub
    End Structure
    ''' <summary>
    ''' Клас за контактор от каталога
    ''' </summary>
    Public Class ContactorEntry
        Public Property PartNumber As String        ' Уникален артикулен/каталожен номер (напр. "LC1D12M7")
        Public Property FrameSize As String         ' Размер на рамката/серията (напр. "09", "12", "18")
        Public Property RatedCurrent_AC1 As Double  ' Номинален ток за AC-1 [A] (активни товари)
        Public Property RatedCurrent_AC3 As Double  ' Номинален ток за AC-3 [A] (двигатели - най-често използван)
        Public Property RatedCurrent_AC4 As Double  ' Номинален ток за AC-4 [A] (тежки режими, реверс)
        Public Property MaxPower_AC3_400V As Double ' Макс. мощност на двигател при 400V AC-3 [kW]
        Public Property AvailableCoils As List(Of String) ' Списък с налични кодове за бобини (напр. {"24DC", "230AC"})
        Public Property HasAuxContacts As Boolean ' Флаг дали поддържа допълнителни помощни контакти
        Public Property MaxAuxContacts As Integer ' Макс. брой помощни контакти (вградени + добавяеми)
        Public Property CompatibleRelay As String ' Артикулен номер на съвместимо термично реле
        Public Property DeratingFactor As Dictionary(Of Double, Double) ' Речник за дерейтинг: Key=Температура [°C], Value=Коефициент [0.0-1.0]
    End Class
    ' Shared гарантира, че функцията се вика само веднъж при първото обръщение
    Private Shared Catalog_Contactor As Dictionary(Of String, ContactorEntry)
    ''' <summary>
    ''' Структура: DisconnectorInfo
    ''' </summary>
    ''' <remarks>
    ''' Тази структура съдържа информация за прекъсвач (изключвател/разединител),
    ''' използван в електрически инсталации или автоматизация на табла.
    ''' Структурата се използва за съхранение на основни характеристики на прекъсвача,
    ''' необходими за избор, документиране и електрически изчисления.
    ''' </remarks>
    Public Structure DisconnectorInfo
        ''' <summary>
        ''' Номинален ток на прекъсвача в ампери.
        ''' </summary>
        ''' <remarks>
        ''' Стойности като 20, 32, 40 и т.н. съответстват на тока, при който прекъсвачът
        ''' може безопасно да работи без да се изключва.
        ''' Използва се за:
        ''' - оразмеряване на електрическата линия
        ''' - избор на подходящ прекъсвач за токовата натовареност
        ''' Потенциален проблем:
        ''' - стойността трябва да съвпада с допустимите стойности на производителя.
        ''' </remarks>
        Dim NominalCurrent As Integer
        ''' <summary>
        ''' Тип на прекъсвача.
        ''' </summary>
        ''' <remarks>
        ''' Например:
        ''' - "iSW" – изключвател/разединител с индикатор
        ''' - "INS" – разединител с термична защита
        ''' - "IN" – стандартен изключвател
        ''' Типът влияе върху функционалността, монтаж и съвместимост с таблото.
        ''' </remarks>
        Dim Type As String
        ''' <summary>
        ''' Марка на прекъсвача.
        ''' </summary>
        ''' <remarks>
        ''' Примерни стойности:
        ''' - "Acti9"
        ''' - "Easy9"
        ''' Марката се използва за документиране и избор на съвместими компоненти.
        ''' </remarks>
        Dim Brand As String
        ''' <summary>
        ''' Брой полюси на прекъсвача.
        ''' </summary>
        ''' <remarks>
        ''' Стойности като 2, 3, 4 показват колко отделни вериги прекъсвачът управлява.
        ''' Използва се за:
        ''' - определяне на конфигурацията на инсталацията
        ''' - избор на подходящ тип за еднофазни или трифазни вериги
        ''' - осигуряване на баланс между фазите
        ''' Потенциален проблем:
        ''' - несъответствие между брой полюси и фазовостта на консуматора може да доведе до неправилна работа.
        ''' </remarks>
        Dim Poles As Integer
    End Structure
    ''' <summary>
    ''' Структура: BusbarInfo
    ''' </summary>
    ''' <remarks>
    ''' Тази структура съдържа информация за шинопровод (Busbar) в електрическо табло или инсталация.
    ''' Използва се за оразмеряване, документиране и проверка на допустимия ток на шините.
    ''' </remarks>
    Public Structure BusbarInfo
        ''' <summary>
        ''' Допустим ток на шината в ампери.
        ''' </summary>
        ''' <remarks>
        ''' Стойността показва максималния ток, който шината може да поеме без риск от прегряване.
        ''' Използва се при:
        ''' - изчисляване на токове на веригата
        ''' - избор на подходящ прекъсвач
        ''' - проверка на съвместимост между шинопровода и консуматорите.
        ''' Потенциален проблем: ако токовата натовареност на шината се надвиши, може да се получи повреда.
        ''' </remarks>
        Dim CurrentCapacity As Integer
        ''' <summary>
        ''' Сечение на шината.
        ''' </summary>
        ''' <remarks>
        ''' Представено във формат "ширина x дебелина", напр. "30x4", "50x5" (мм).
        ''' Използва се за определяне на допустим ток и механична здравина.
        ''' Сечението влияе върху:
        ''' - токовата проводимост
        ''' - падовете на напрежение
        ''' - избор на табло и монтажни аксесоари.
        ''' </remarks>
        Dim Section As String
        ''' <summary>
        ''' Материал на шината.
        ''' </summary>
        ''' <remarks>
        ''' Възможни стойности:
        ''' - "Cu" – мед
        ''' - "Al" – алуминий
        ''' Материалът определя електрическата проводимост, допустимия ток и механичните свойства.
        ''' </remarks>
        Dim Material As String
    End Structure
    ''' <summary>
    ''' Структура: RCDInfo
    ''' </summary>
    ''' <remarks>
    ''' Тази структура съхранява информация за защитно устройство от тип диференциална токова защита (ДТЗ/RCD),
    ''' използвана за защита на хора и оборудване от токови утечки.
    ''' Структурата описва основните характеристики на RCD, необходими за избор, монтаж и документиране.
    ''' </remarks>
    Public Structure RCDInfo
        ''' <summary>
        ''' Производител на RCD устройството.
        ''' </summary>
        ''' <remarks>
        ''' Примерни стойности:
        ''' - "Schneider"
        ''' - "ABB"
        ''' - "Legrand"
        ''' Марката се използва за документиране, избор на съвместими компоненти
        ''' и осигуряване на надеждност според стандартите.
        ''' </remarks>
        Dim Brand As String
        ''' <summary>
        ''' Номинален ток на RCD в ампери.
        ''' </summary>
        ''' <remarks>
        ''' Обикновено стойности като 25, 40, 63 и т.н.
        ''' Определя максималния ток, при който устройството може да работи без да се изключва.
        ''' Използва се за оразмеряване на веригата и съвместимост с прекъсвачи.
        ''' </remarks>
        Dim NominalCurrent As Integer
        ''' <summary>
        ''' Тип на чувствителността на RCD спрямо диференциален ток.
        ''' </summary>
        ''' <remarks>
        ''' Примерни стойности:
        ''' - "AC" – реагира на синусоидален променлив ток
        ''' - "A" – реагира на променлив и пулсиращ постоянен ток
        ''' - "F" – висока чувствителност, бърза реакция на различни видове ток
        ''' Типът влияе на приложимостта в различни електрически инсталации.
        ''' </remarks>
        Dim Type As String
        ''' <summary>
        ''' Брой полюси на устройството.
        ''' </summary>
        ''' <remarks>
        ''' Стойности като:
        ''' - "2p" – двуполюсен
        ''' - "4p" – четириполюсен
        ''' Броят на полюсите определя фазовостта на веригата, която RCD защитава.
        ''' </remarks>
        Dim Poles As String
        ''' <summary>
        ''' Чувствителност на RCD в милиампери (mA).
        ''' </summary>
        ''' <remarks>
        ''' Примерни стойности:
        ''' - 10, 30, 100, 300, 500
        ''' Определя токът на утечка, при който устройството ще сработи.
        ''' Използва се за защита на хора (обикновено 30 mA) или оборудване (100–500 mA).
        ''' </remarks>
        Dim Sensitivity As Integer
        ''' <summary>
        ''' Вид на устройството.
        ''' </summary>
        ''' <remarks>
        ''' Примерни стойности:
        ''' - "RCCB" – само диференциална защита без прекъсвач
        ''' - "RCBO" – комбинира диференциална защита с прекъсвач
        ''' - "iID" – интелигентен диференциален изключвател
        ''' Видът определя функционалността и начина на защита на веригата.
        ''' </remarks>
        Dim DeviceType As String
        ''' <summary>
        ''' Показва дали устройството има вграден прекъсвач.
        ''' </summary>
        ''' <remarks>
        ''' True – устройството е RCBO (комбиниран прекъсвач + ДТЗ)
        ''' False – устройството е RCCB (само диференциална защита)
        ''' Това поле се използва при изчисления и избор на защита за конкретна електрическа линия.
        ''' </remarks>
        Dim Breaker As Boolean
    End Structure
    ''' <summary>
    ''' Структура: strKonsumator
    ''' </summary>
    ''' <remarks>
    ''' Тази структура описва един електрически консуматор (товар), извлечен от блок в AutoCAD чертеж.
    ''' Използва се като контейнер за съхранение на информацията, прочетена от атрибутите на блока,
    ''' както и на допълнително изчислени или обработени стойности.
    '''
    ''' Всеки елемент от този тип представлява един обект от електрическата инсталация
    ''' (например осветително тяло, контакт, консуматор, LED лента и др.).
    '''
    ''' Данните в структурата се използват по-късно за:
    ''' - изчисляване на мощности
    ''' - групиране по токови кръгове
    ''' - определяне на табла
    ''' - електрически изчисления
    ''' - създаване на таблици и отчети
    '''
    ''' Структурата е създадена като Value Type (Structure), което означава:
    ''' - съхранява се по стойност
    ''' - при копиране се копират всички полета
    ''' - използва се удобно в списъци (List(Of strKonsumator)) при обработка на множество консуматори
    ''' </remarks>
    Public Structure strKonsumator
        ''' <summary>
        ''' Име на блока в AutoCAD.
        ''' </summary>
        ''' <remarks>
        ''' Това е името на блока, който представлява консуматора в чертежа.
        ''' Често чрез него може да се определи типът на консуматора
        ''' (например осветително тяло, контакт, LED лента и др.).
        '''
        ''' Полето се използва основно за:
        ''' - идентификация на типа елемент
        ''' - филтриране на блокове
        ''' - диагностични проверки
        ''' </remarks>
        Dim Name As String
        ''' <summary>
        ''' Уникален идентификатор на блока в AutoCAD.
        ''' </summary>
        ''' <remarks>
        ''' Прекият идентификатор на блока в AutoCAD за текущата сесия.
        ''' Това поле (ID_Block) представлява директна връзка между структурата и реалния AutoCAD обект.
        ''' Използва се само по време на runtime, за да се достъпят атрибутите на блока или да се извършат промени в чертежа.
        ''' 
        ''' ВНИМАНИЕ: ObjectId не е постоянен между различни сесии на AutoCAD.
        ''' След затваряне и отваряне на файла, ID_Block може да стане невалиден.
        ''' За дългосрочно съхранение се използва Handle_Block.
        ''' </remarks>
        Dim ID_Block As ObjectId
        ''' <summary>
        ''' Постоянният идентификатор на блока, който може да се записва и чете от файл (JSON).
        ''' </summary>
        ''' <remarks>
        ''' Handle_Block съхранява Handle на блока като низ. Той е стабилен и остава валиден след затваряне и отваряне на DWG файла.
        ''' След зареждане на данните, Handle_Block може да се използва за възстановяване на ID_Block чрез AutoCAD API.
        ''' </remarks>
        Dim Handle_Block As String
        ''' <summary>
        ''' Име или номер на токовия кръг.
        ''' </summary>
        ''' <remarks>
        ''' Токовият кръг определя към коя електрическа линия е свързан консуматорът.
        ''' Тази информация обикновено се извлича от атрибут на блока.
        '''
        ''' Използва се за:
        ''' - групиране на консуматорите по кръгове
        ''' - изчисляване на общата мощност на кръга
        ''' - избор на прекъсвачи
        ''' - избор на кабели
        ''' </remarks>
        Dim ТоковКръг As String
        ''' <summary>
        ''' Мощност на консуматора като текст.
        ''' </summary>
        ''' <remarks>
        ''' Стойността се чете директно от атрибут на блока в AutoCAD.
        ''' Обикновено съдържа текстова стойност като:
        ''' - "60W"
        ''' - "0.1 kW"
        ''' - "12"
        '''
        ''' Тази стойност се съхранява в оригиналния си вид, както е записана в чертежа.
        ''' По-късно тя може да бъде обработена и конвертирана в числова стойност.
        '''
        ''' Потенциален проблем:
        ''' Ако текстът съдържа нестандартни символи или различни единици,
        ''' може да се наложи допълнителна обработка преди конвертиране.
        ''' </remarks>
        Dim strМОЩНОСТ As String
        ''' <summary>
        ''' Мощност на консуматора като числова стойност.
        ''' </summary>
        ''' <remarks>
        ''' Това е числовата версия на мощността, извлечена от strМОЩНОСТ.
        '''
        ''' След обработка текстовата стойност се преобразува в Double,
        ''' което позволява извършването на математически операции:
        ''' - сумиране на мощности
        ''' - изчисляване на токове
        ''' - оразмеряване на кабели
        '''
        ''' Единицата на измерване трябва да бъде предварително унифицирана
        ''' (например всички стойности да са във ватове или киловати).
        ''' </remarks>
        Dim doubМОЩНОСТ As Double
        ''' <summary>
        ''' Име или обозначение на електрическото табло.
        ''' </summary>
        ''' <remarks>
        ''' Показва към кое електрическо табло е свързан консуматорът.
        '''
        ''' Използва се за:
        ''' - групиране на консуматорите по табла
        ''' - изчисляване на общата мощност на таблото
        ''' - създаване на електрически схеми
        ''' </remarks>
        Dim ТАБЛО As String
        ''' <summary>
        ''' Основно предназначение на консуматора.
        ''' </summary>
        ''' <remarks>
        ''' Това поле описва функционалното предназначение на консуматора.
        ''' Например:
        ''' - Осветление
        ''' - Контакти
        ''' - Технологично оборудване
        '''
        ''' Данните обикновено се извличат от атрибут на блока.
        ''' </remarks>
        Dim Pewdn As String
        ''' <summary>
        ''' Допълнително предназначение на консуматора.
        ''' </summary>
        ''' <remarks>
        ''' Това поле служи за уточняване или допълване на основното предназначение.
        ''' Например:
        ''' - тип осветление
        ''' - зона на използване
        ''' - специфична функция
        '''
        ''' Може да бъде празно, ако няма допълнителна информация.
        ''' </remarks>
        Dim PEWDN1 As String
        ''' <summary>
        ''' Дължина на LED лента.
        ''' </summary>
        ''' <remarks>
        ''' Използва се когато консуматорът представлява LED осветление тип лента.
        '''
        ''' Стойността представлява физическата дължина на лентата,
        ''' която по-късно може да се използва за:
        ''' - изчисляване на мощност
        ''' - определяне на захранване
        ''' - оразмеряване на кабели.
        '''
        ''' Ако консуматорът не е LED лента,
        ''' това поле обикновено остава със стойност 0.
        ''' </remarks>
        Dim Dylvina_Led As Double
        ''' <summary>
        ''' Състояние на Visibility параметъра на динамичен блок.
        ''' </summary>
        ''' <remarks>
        ''' Някои AutoCAD блокове са динамични и съдържат параметър Visibility,
        ''' който определя визуалния вариант на блока.
        '''
        ''' Това поле съхранява текущото състояние на този параметър.
        '''
        ''' Използва се например за:
        ''' - определяне на типа осветително тяло
        ''' - избор между различни конфигурации на блока
        ''' </remarks>
        Dim Visibility As String
        ''' <summary>
        ''' Брой фази на консуматора.
        ''' </summary>
        ''' <remarks>
        ''' Определя дали консуматорът е:
        ''' - еднофазен (1)
        ''' - трифазен (3)
        '''
        ''' Тази информация е важна за:
        ''' - изчисляване на токове
        ''' - балансиране на фазите
        ''' - избор на защити
        ''' </remarks>
        Dim Phase As Integer
    End Structure
    ''' <summary>
    ''' Структура токов кръг в електрическо табло.
    ''' Съдържа идентификация, мощност, ток, кабел,
    ''' защитна апаратура (прекъсвач и ДТЗ),
    ''' както и списък с консуматори в кръга.
    ''' Използва се като логическа структура за обработка
    ''' и визуализация (напр. в DataGridView).
    ''' </summary>
    Public Class strTokow
        ' ============================================================
        ' ИДЕНТИФИКАЦИЯ
        ' ============================================================
        Public Device As String                     ' какъв тип консуматор е (осветление, контакти, технологично оборудване)
        Public Tablo As String                      ' Табло към което принадлежи кръгът
        Public ТоковКръг As String                  ' Име или номер на токовия кръг
        Public Табло_Родител As String              ' Името на таблото което захранва таблото в което е монтиран токовия кръг"
        Public Мощност As Double                    ' Обща мощност на кръга (kW)
        Public Ток As Double                        ' Изчислен ток (A)
        ' ============================================================
        ' БРОЯЧИ
        ' ============================================================
        Public brLamp As Integer                    ' Брой лампи в кръга
        Public brKontakt As Integer                 ' Брой контакти в кръга
        Public Фаза As String                       ' Фаза: "1P", "3P", "L1", "L2", "L3"
        Public Брой_Полюси As Integer               ' Брой на фазите на токовия кръг
        ' ============================================================
        ' КАБЕЛ
        ' ============================================================
        Public Кабел_Монтаж As String               ' Начин на монтаж                       "A1"=гипсокартон, "B2"=под мазилка, "C"=над таван   
        Public Кабел_Полагане As String             ' Въздух или земя                       0=въздух (35°C), 1=земя (15°C)
        Public Кабел_Сечение As String              ' Сечение на кабела (пример: "3x2.5")
        Public Кабел_Тип As String                  ' Тип кабел (NYM, YJV, CBT и др.)
        Public Кабел_Брой_Фаза As String            ' Брой на Паралелни Жила на Фаза
        Public Кабел_Брой_Група As String           ' Брой на Паралелни кабели по скара
        ' ============================================================
        ' ЗАЩИТА (ПРЕКЪСВАЧ)
        ' ============================================================
        Public Breaker_Тип_Апарат As String         ' Серия апарат (EZ9, C120, NSX, MTZ)
        Public Breaker_Крива As String              ' Характеристика (B, C, D)
        Public Breaker_Номинален_Ток As String      ' Номинален ток (пример: "16A")
        Public Breaker_Изкл_Възможност As String    ' Изключвателна способност ("6000A", "10000A")
        Public Breaker_Защитен_блок As String       ' Изключвателна способност ("6000A", "10000A")
        ' ============================================================
        ' ДТЗ (RCD)
        ' ============================================================
        Public RCD_Бранд As String                  ' Производител на ДТЗ
        Public RCD_Клас As String                   ' Тип ДТЗ (AC, A, F)
        Public RCD_Тип As String                    ' EZ9 RCCB, EZ9 RCBO, iID
        Public RCD_Чувствителност As String         ' Чувствителност ("30mA", "100mA", "300mA")
        Public RCD_Ток As String                    ' Номинален ток на ДТЗ ("25A", "40A", "63A")
        Public RCD_Полюси As String                 ' Полюси на ДТЗ ("2p", "4p")
        Public RCD_Нула As String                   ' номер 
        Public RCD_Автомат As Boolean               ' трябва ли ДТЗ да е RCBO (с вграден прекъсвач) или може да е RCCB (само ДТЗ) 
        ' ============================================================
        ' ОПИСАНИЕ / ТЕКСТОВЕ
        ' ============================================================
        Public Консуматор As String                 ' Обобщен текст за консуматора
        Public предназначение As String             ' Предназначение на кръга
        ' ============================================================
        ' ДОПЪЛНИТЕЛНИ ФЛАГОВЕ
        ' ============================================================
        Public Управление As String                 ' Тип управление (ако има)
        Public Шина As Boolean                      ' Дали кръгът е на отделна шина
        Public ДТЗ_RCD As Boolean                   ' Дали има задължително трявба да има ДТЗ
        ' ============================================================
        ' КОНСУМАТОРИ В КРЪГА
        ' ============================================================
        Public Konsumator As List(Of strKonsumator)
        ' Списък с всички реални консуматори,
        ' принадлежащи към този токов кръг.
        ''' <summary>
        ''' Създава независимо копие на записа (идеално за Class типове)
        ''' </summary>
        Public Function Clone() As strTokow
            ' MemberwiseClone копира всички стойности (String, Integer, Boolean и др.)
            ' в нов обект от същия тип. За примитиви и String това е напълно безопасно.
            If Me.Tablo = "__ROOT__" OrElse Me.ТоковКръг = "__ROOT__" Then
                ' СЛОЖИ BREAKPOINT ТУК!
                Debug.WriteLine("Клонира се обект, съдържащ __ROOT__")
            End If
            Return DirectCast(Me.MemberwiseClone(), strTokow)
        End Function
    End Class
    ''' <summary>
    ''' КАТАЛОГ автоматичен прекъсвач – MCB, MCCB или ACB.
    ''' Може да се използва за избор на прекъсвач за генераторни табла,
    ''' както и за по-сложни сценарии с селективност и късо съединение.
    ''' </summary>
    Public Class BreakerInfo
        Public Brand As String                      ' Производител на прекъсвача (например "Schneider").
        Public Series As String                     ' Серия или модел на прекъсвача (например "EZ9", "C120", "NSX", "MTZ").
        ''' <summary>
        ''' Категория на прекъсвача:
        ''' - "MCB" – миниатюрен автоматичен прекъсвач
        ''' - "MCCB" – корпусен прекъсвач
        ''' - "ACB" – въздушен прекъсвач
        ''' </summary>
        Public Category As String
        Public NominalCurrent As Integer            ' Номинален ток на прекъсвача в ампери.
        Public Poles As Integer                     ' Брой полюси (1P, 2P, 3P или 4P).
        ''' <summary>
        ''' Работна прекъсвателна способност (Ics) в kA.
        ''' Това е стойността, до която прекъсвачът може да изключва многократно.
        ''' </summary>
        Public Ics_kA As Decimal
        ''' <summary>
        ''' Крива на MCB (B, C или D). 
        ''' Само за миниатюрни автоматични прекъсвачи.  
        ''' Определя характеристиката на изключване при късо съединение.
        ''' MCCB и ACB не използват това поле.
        ''' </summary>
        Public Curve As String
        ''' <summary>
        ''' Тип на защитната единица (Trip Unit) – TM-D, Micrologic и т.н.
        ''' Само за MCCB и ACB.  
        ''' Определя електронната или термомагнитната защита.
        ''' </summary>
        Public TripUnit As String
    End Class
    ''' <summary>
    ''' Конфигурация за всеки тип блок
    ''' </summary>
    Public Class BlockConfig
        Public BlockNames As List(Of String)        ' Възможни имена на блока
        Public Category As String                   ' "Lamp", "Contact", "Device", "Panel"
        Public DefaultPoles As Integer              ' "1p" или "3p"
        Public DefaultCable As String               ' "3x1.5", "3x2.5", "5x2.5"
        Public DefaultBreaker As String             ' "10", "16", "20"
        Public DefaultBreakerType As String         ' "10", "16", "20"
        Public DefaultPrednaz As String             ' Предназначение 
        Public DefaultPrednaz1 As String            ' Предназначение 
        Public VisibilityRules As List(Of VisRule)  ' Правила за visibility
    End Class
    ''' <summary>
    ''' Правило за конкретна visibility стойност
    ''' </summary>
    Public Class VisRule
        Public VisibilityPattern As String        ' "3P", "Двугнездов", "Проточен"
        Public Poles As Integer                    ' "1p" или "3p"
        Public Cable As String                    ' "3x2.5", "5x4"
        Public Breaker As String                  ' "16", "25", "32"
        Public Phase As String                    ' "L" или "L1,L2,L3"
        Public BreakerType As String              ' опционално за специфични правила
        Public ContactCount As Integer            ' Колко контакта добавя (1, 2, 3)
    End Class
    ''' <summary>
    ''' Клас за групиране на токови кръгове за балансиране на фазите
    ''' </summary>
    Public Class BalanceGroup
        ''' <summary>
        ''' Списък с токови кръгове в групата
        ''' </summary>
        Public Circuits As List(Of strTokow)
        ''' <summary>
        ''' Тип на групата: "ThreePhase", "RCD", "SmallBus", "LargeBus", "Normal"
        ''' </summary>
        Public GroupType As String
        ''' <summary>
        ''' Ключ на групата: RCD_Нула (N1, N2...), "Bus" или Nothing
        ''' </summary>
        Public GroupKey As String
        ''' <summary>
        ''' Сумарен ток на групата (сума от токовете на всички ТК)
        ''' </summary>
        Public TotalCurrent As Double
        ''' <summary>
        ''' Зададена фаза след балансиране (L1, L2, L3 или "L1,L2,L3")
        ''' </summary>
        Public AssignedPhase As String
        ''' <summary>
        ''' Конструктор - инициализира списъка с ТК
        ''' </summary>
        Public Sub New()
            Circuits = New List(Of strTokow)
        End Sub
        ''' <summary>
        ''' Брой токови кръгове в групата
        ''' </summary>
        Public ReadOnly Property CircuitCount As Integer
            Get
                Return Circuits.Count
            End Get
        End Property
        ''' <summary>
        ''' Сумарна мощност на групата
        ''' </summary>
        Public ReadOnly Property TotalPower As Double
            Get
                Return Circuits.Sum(Function(t) t.Мощност)
            End Get
        End Property
    End Class
    ''' <summary>
    ''' Извлича информация за всички избрани блокове от AutoCAD,
    ''' създава обекти от тип strKonsumator и ги добавя в списъка ListKonsumator.
    '''
    ''' Основна идея:
    ''' Процедурата обхожда избраните блокове (INSERT) в чертежа,
    ''' извлича атрибути и динамични свойства от всеки блок и
    ''' попълва структурата strKonsumator.
    '''
    ''' Основни стъпки:
    ''' 1) Потребителят избира блокове от чертежа.
    ''' 2) За всеки блок се прочитат:
    '''    - Атрибутите (ТАБЛО, КРЪГ, МОЩНОСТ и др.)
    '''    - Dynamic properties (Visibility, Дължина)
    ''' 3) Изчислява се мощността на консуматора.
    ''' 4) Определя се типът на блока чрез ProcessBlockByType().
    ''' 5) Ако консуматорът има валидна мощност → добавя се в ListKonsumator.
    '''
    ''' Допълнително:
    ''' - Показва прогрес чрез ToolStripProgressBar.
    ''' - Работи в Transaction за безопасен достъп до AutoCAD Database.
    ''' </summary>
    Private Sub GetKonsumatori()
        ' ------------------------------------------------------------
        ' 1) Вземане на текущия AutoCAD документ, редактор и база
        ' ------------------------------------------------------------
        Dim acDoc As Document =
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim edt As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database
        ' ------------------------------------------------------------
        ' 2) Избор на обекти от потребителя
        ' ------------------------------------------------------------
        ' cu.GetObjects() връща всички избрани обекти от тип INSERT (блокове)
        Dim SelectedSet = cu.GetObjects("INSERT", "Изберете блок")
        ' Ако не е избран нито един блок → прекратяваме процедурата
        If SelectedSet Is Nothing Then
            Me.Close()
            Exit Sub
        End If
        ' ObjectId на текущия блок
        Dim blkRecId As ObjectId = ObjectId.Null
        ' ------------------------------------------------------------
        ' 3) Стартиране на Transaction за работа с AutoCAD Database
        ' ------------------------------------------------------------
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                ' Отваряне на BlockTable за четене
                Dim acBlkTbl As BlockTable
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                ' --------------------------------------------------------
                ' Инициализация на ProgressBar
                ' --------------------------------------------------------
                ToolStripProgressBar1.Maximum = SelectedSet.Count
                ToolStripProgressBar1.Value = 0
                ' --------------------------------------------------------
                ' 4) Обработка на всеки избран блок
                ' --------------------------------------------------------
                For Each sObj As SelectedObject In SelectedSet
                    ' Вземаме ObjectId на блока
                    blkRecId = sObj.ObjectId
                    ' Отваряме блока за четене
                    Dim acBlkRef As BlockReference =
                    DirectCast(acTrans.GetObject(blkRecId, OpenMode.ForRead), BlockReference)
                    ' Колекция с атрибутите на блока
                    Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                    ' Колекция с Dynamic Properties
                    Dim props As DynamicBlockReferencePropertyCollection =
                    acBlkRef.DynamicBlockReferencePropertyCollection
                    Dim Visibility As String = ""
                    Dim Dylvina_Led As Double = 0
                    Dim nameBlock As String = (CType(acBlkRef.DynamicBlockTableRecord.GetObject(OpenMode.ForRead),
                    BlockTableRecord)).Name
                    For Each prop As DynamicBlockReferenceProperty In props
                        ' Вземане на стойностите според името на свойството
                        Select Case prop.PropertyName
                            Case "Visibility", "Visibility1"
                                Visibility = prop.Value
                            Case "Дължина"  ' Дължина (използва се при LED линии)
                                Dylvina_Led = prop.Value
                        End Select
                    Next
                    Select Case Visibility
                        Case "Само ключ", "текст",
                             "Лампион - рошав", "Лампион", "Настолна лампа - рошава",
                             "Настолна лампа", "Фотодатчик", "Датчик 360°", "Датчик насочен",
                             "Драйвер", "ПВ", "Линии", "Само текст", "Табло_Ново"
                            Continue For
                    End Select
                    ' Създаваме нов обект за консуматора
                    Dim Kons As New strKonsumator
                    Kons.Visibility = Visibility
                    Kons.Dylvina_Led = Dylvina_Led
                    ' ----------------------------------------------------
                    ' Извличане на името на блока
                    ' ----------------------------------------------------
                    Kons.Name = nameBlock
                    ' Записване на ObjectId за runtime (само текущата сесия)
                    Kons.ID_Block = blkRecId
                    ' Записване на Handle за дългосрочно съхранение и JSON
                    Kons.Handle_Block = blkRecId.Handle.ToString()
                    ' ----------------------------------------------------
                    ' 5) Четене на всички атрибути на блока
                    ' ----------------------------------------------------
                    For Each attId As ObjectId In attCol
                        ' Отваряме обекта
                        Dim dbObj As DBObject = acTrans.GetObject(attId, OpenMode.ForRead)
                        ' Преобразуваме го в AttributeReference
                        Dim acAttRef As AttributeReference = dbObj
                        ' ------------------------------------------------
                        ' Проверка на Tag на атрибута и записване
                        ' ------------------------------------------------
                        Select Case acAttRef.Tag
                            Case "ТАБЛО" : Kons.ТАБЛО = acAttRef.TextString
                            Case "КРЪГ" : Kons.ТоковКръг = acAttRef.TextString
                            Case "Pewdn" : Kons.Pewdn = acAttRef.TextString
                            Case "PEWDN1" : Kons.PEWDN1 = acAttRef.TextString
                            Case "LED", "МОЩНОСТ" ' Двете стойности водят до едно и също действие
                                Kons.strМОЩНОСТ = acAttRef.TextString
                        End Select
                    Next
                    If Kons.strМОЩНОСТ Is Nothing Then Continue For
                    ' ----------------------------------------------------
                    ' 7) Изчисляване на мощността
                    ' ----------------------------------------------------
                    ' CalcPower() изчислява реалната мощност на консуматора
                    ' според текста на мощността и дължината (ако има)
                    Kons.doubМОЩНОСТ = CalcPower(Kons.strМОЩНОСТ, Kons.Dylvina_Led)
                    ' ----------------------------------------------------
                    ' 8) Допълнителна обработка според типа на блока
                    ' ----------------------------------------------------
                    ProcessBlockByType(Kons, Kons.Name, Kons.Visibility)
                    ' ----------------------------------------------------
                    ' 9) Добавяне в списъка с консуматори
                    ' ----------------------------------------------------
                    ' Добавяме само ако има валидна мощност
                    If Kons.doubМОЩНОСТ > 0 Then ListKonsumator.Add(Kons)
                    ' ----------------------------------------------------
                    ' 10) Обновяване на ProgressBar
                    ' ----------------------------------------------------
                    ToolStripProgressBar1.Value += 1
                Next
                ' --------------------------------------------------------
                ' Потвърждаване на Transaction
                ' --------------------------------------------------------
                acTrans.Commit()
            Catch ex As Exception
                ' --------------------------------------------------------
                ' Обработка на грешки
                ' --------------------------------------------------------
                MsgBox("Възникна грешка:  " &
                   ex.Message &
                   vbCrLf & vbCrLf &
                   ex.StackTrace.ToString)
                ' Прекратяване на Transaction
                acTrans.Abort()
            End Try
        End Using
        ToolStripProgressBar1.Value = ToolStripProgressBar1.Minimum
    End Sub
    ''' <summary>
    ''' Универсална функция за изчисляване на мощност.
    ''' Разпознава автоматично формата на входа:
    ''' - LED ленти: "60 led/m" + дължина
    ''' - Директна мощност: "3500" или "3.5"
    ''' - Контакти/Консуматори: "2х100", "3х100", "100"
    ''' </summary>
    ''' <param name="strМОЩНОСТ">Текст от атрибута "МОЩНОСТ"</param>
    ''' <param name="Dylvina_Led">Дължина на LED лента в метри (ако е приложимо)</param>
    ''' <returns>Обща мощност във Watt</returns>
    Private Function CalcPower(strМОЩНОСТ As String,
                       Optional Dylvina_Led As Double = 0) As Double
        ' --- 1. Валидация ---
        If String.IsNullOrEmpty(strМОЩНОСТ) Then
            Return 0.0
        End If
        Dim input As String = strМОЩНОСТ.Trim().ToLower()
        ' --- 2. LED ЛЕНТИ (формат: "60 led/m", "120led/m") ---
        If input.Contains("led/m") Then
            ' Проверка дали текстът съдържа "led/m" (т.е. LED лента)
            ' Вземаме числото пред "led/m", което показва броя диоди на метър
            ' Превръщаме текста в малки букви, махаме "led/m" и изтриваме интервали
            Dim диоди As Double = Val(strМОЩНОСТ.ToLower().Replace("led/m", "").Trim())
            ' Декларираме променлива за мощността на метър (W/m)
            Dim мощностНаМетър As Double
            ' Определяме мощността на метър според таблица с известни стойности
            ' Ако броят диоди не е стандартен, използваме средна мощност на диод (0.24 W/диод)
            Select Case диоди
                Case 30
                    мощностНаМетър = 7.2       ' 30 диода/м → 7.2 W/м
                Case 60
                    мощностНаМетър = 14.4      ' 60 диода/м → 14.4 W/м
                Case 72
                    мощностНаМетър = 17.28     ' 72 диода/м → 17.28 W/м
                Case 120
                    мощностНаМетър = 28.8      ' 120 диода/м → 28.8 W/м
                Case Else
                    ' За непознат брой диоди използваме средна мощност на диод 0.24 W/диод
                    мощностНаМетър = диоди * 0.24
            End Select
            ' Изчисляваме мощността за реалната дължина на лентата (Dylvina_Led в см)
            Return (Dylvina_Led / 100) * мощностНаМетър
        End If
        ' --- 3. КОНТАКТИ/КОНСУМАТОРИ (формат: "2х100", "3х100", "100") ---
        ' Поддържа различни разделители: "х", "x", "*", "Х"
        Dim separators As String() = {"х", "x", "*", "Х", "X"}
        For Each sep As String In separators
            If input.Contains(sep) Then
                Dim parts As String() = input.Split(sep)
                If parts.Length = 2 Then
                    Dim count As Double = 0.0
                    Dim power As Double = 0.0
                    If Double.TryParse(parts(0).Trim(), count) AndAlso
                    Double.TryParse(parts(1).Trim(), power) Then
                        Return count * power  ' Брой × Мощност на бройка
                    End If
                End If
            End If
        Next
        ' --- 5. ОБИКНОВЕНО ЧИСЛО (формат: "3500", "3.5") ---
        Dim numericValue As Double = 0.0
        If Double.TryParse(input, numericValue) Then
            Return numericValue  ' Предполагаме W
        End If
        ' --- 6. НЕУСПЕШНО РАЗПОЗНАВАНЕ ---
        Return 0.0
    End Function
    ''' <summary>
    ''' Изгражда TreeView структура на база списъка ListTokow.
    ''' Логиката:
    ''' - създава коренен възел
    ''' - групира елементите по табло
    ''' - добавя таблата като деца на корена
    ''' - създава отделна група за записи без табло
    ''' </summary>
    Private Sub BuildTreeViewFromKonsumatori()
        ' Изчиства всички възли от TreeView
        TreeView1.Nodes.Clear()
        ' Създава коренен възел
        Dim rootNode As New TreeNode(ROOT_NODE_TEXT)
        rootNode.Name = ROOT_NODE_NAME
        rootNode.ForeColor = Color.DarkBlue
        ' Извлича всички записи, които принадлежат към корена
        Dim rootRecords = ListTokow.Where(Function(x)
                                              Return String.Equals(x.Tablo, ROOT_NODE_TEXT, StringComparison.OrdinalIgnoreCase)
                                          End Function).ToList()
        ' Брой уникални токови кръгове в корена
        Dim rootCircuitCount = rootRecords.Select(Function(r) r.ТоковКръг).Distinct().Count()
        ' Търси записа за таблото (обобщен запис)
        Dim rootMaster = rootRecords.FirstOrDefault(Function(r) r.Device = "Табло")
        ' Взима общата мощност (ако има такъв запис)
        Dim rootPower As Double = If(rootMaster IsNot Nothing, rootMaster.Мощност, 0)
        ' Задава текст на корена с мощност
        rootNode.Text = $"{ROOT_NODE_TEXT} ({rootPower:F1}kW)"
        ' Съхранява свързаните записи в Tag
        rootNode.Tag = rootRecords
        ' Добавя корена в TreeView
        TreeView1.Nodes.Add(rootNode)
        ' Групира всички записи по табло
        Dim panels = ListTokow.GroupBy(Function(k) k.Tablo).ToList()
        ' Филтрира валидните табла (без празни и без корена)
        Dim validPanels = panels.Where(Function(p)
                                           Return Not String.IsNullOrWhiteSpace(p.Key) AndAlso
                                              Not String.Equals(p.Key.Trim(), ROOT_NODE_TEXT, StringComparison.OrdinalIgnoreCase)
                                       End Function).
                                       OrderBy(Function(p) p.Key.Trim()).ToList()
        ' Групи без име (празно табло)
        Dim emptyPanels = panels.Where(Function(p) String.IsNullOrWhiteSpace(p.Key)).ToList()
        ' Добавя валидните табла като деца на корена
        For Each panelGroup In validPanels
            Dim panelName As String = panelGroup.Key.Trim()

            ' 1. Намираме главния запис за мощността на таблото
            Dim tableDeviceRecord = panelGroup.FirstOrDefault(Function(k)
                                                                  Return String.Equals(k.Device, "Табло", StringComparison.OrdinalIgnoreCase)
                                                              End Function)
            Dim totalPower As Double = If(tableDeviceRecord IsNot Nothing, tableDeviceRecord.Мощност, 0)

            ' 2. Създаваме основния възел за ТАБЛОТО
            Dim panelNode As New TreeNode($"{panelName} ({totalPower:F1} kW)")
            panelNode.Name = panelName
            panelNode.Tag = panelGroup.ToList()

            ' 3. ГРУПИРАМЕ данните за токовите кръгове
            Dim circuits = panelGroup.Where(Function(k) Not String.Equals(k.Device, "Табло", StringComparison.OrdinalIgnoreCase)) _
                             .GroupBy(Function(k) k.ТоковКръг)
            ' 4. СЪЗДАВАМЕ ЕДИН ЕДИНСТВЕН ВЪЗЕЛ-КОНТЕЙНЕР С ОБЩА МОЩНОСТ
            If circuits.Any() Then
                ' Изчисляваме сумата от мощностите на всички записи, които не са "Табло"
                Dim totalCircuitsPower As Double = panelGroup.Where(Function(k) Not String.Equals(k.Device, "Табло", StringComparison.OrdinalIgnoreCase)) _
                                                     .Sum(Function(k) k.Мощност)

                ' Създаваме заглавния възел с общата сума
                Dim circuitsFolderNode As New TreeNode($"🔌 ТК ({totalCircuitsPower:F1} kW)")
                circuitsFolderNode.ForeColor = Color.DarkBlue ' Тъмно син цвят за акцент
                circuitsFolderNode.NodeFont = New Font(TreeView1.Font, FontStyle.Bold) ' Удебелен шрифт

                ' Добавяме всеки отделен кръг вътре в папката
                For Each circuitGroup In circuits
                    Dim circuitName As String = circuitGroup.Key
                    Dim circuitPower As Double = circuitGroup.Sum(Function(k) k.Мощност)

                    Dim circuitNode As New TreeNode($"{circuitName} ({circuitPower:F2} kW)")
                    circuitNode.Tag = circuitGroup.ToList()

                    circuitsFolderNode.Nodes.Add(circuitNode)
                Next

                ' Добавяме "папката" към възела на таблото
                panelNode.Nodes.Add(circuitsFolderNode)
            End If

            ' Добавяме таблото към корена
            rootNode.Nodes.Add(panelNode)
        Next
        ' Добавя група "Без име", ако има такива записи
        If emptyPanels.Any() Then
            ' Общ брой кръгове
            Dim totalEmptyCircuits =
            emptyPanels.Sum(Function(p)
                                Return p.Select(Function(k) k.ТоковКръг).Distinct().Count()
                            End Function)
            ' Обща мощност
            Dim totalEmptyPower =
            emptyPanels.Sum(Function(p)
                                Return p.Sum(Function(k) k.Мощност)
                            End Function)
            ' Създава възел "Без име"
            Dim emptyNode As New TreeNode($"Без име ({totalEmptyPower:F1}kW)")
            emptyNode.Name = "__EMPTY__"
            emptyNode.ForeColor = Color.OrangeRed
            ' Обединява всички записи без табло
            emptyNode.Tag = emptyPanels.SelectMany(Function(p) p).ToList()
            ' Добавя към корена
            rootNode.Nodes.Add(emptyNode)
        End If
        ' Позволява Drag & Drop операции
        TreeView1.AllowDrop = True
        ' Разгъва корена
        rootNode.Expand()
    End Sub
    ''' <summary>
    ''' Стартира операция по влачене (Drag) на възел от TreeView.
    ''' Забранява влачене на кореновия възел и системния възел "__EMPTY__".
    ''' При валиден възел започва Drag&amp;Drop с ефект Move.
    ''' </summary>
    Private Sub TreeView1_ItemDrag(sender As Object, e As ItemDragEventArgs) Handles TreeView1.ItemDrag
        ' Взимаме влачения възел от събитието
        Dim draggedNode As TreeNode = DirectCast(e.Item, TreeNode)
        ' Забраняваме влачене на корена и специалния възел "Без име"
        If draggedNode.Name = ROOT_NODE_NAME OrElse draggedNode.Name = "__EMPTY__" Then Return
        ' Стартираме Drag&Drop операция
        TreeView1.DoDragDrop(draggedNode, DragDropEffects.Move)
    End Sub
    ''' <summary>
    ''' Обработва влизане на влачен обект в TreeView.
    ''' По подразбиране забранява всякакъв Drop, докато не се валидира в DragOver.
    ''' </summary>
    Private Sub TreeView1_DragEnter(sender As Object, e As DragEventArgs) Handles TreeView1.DragEnter
        ' По подразбиране не позволяваме пускане
        e.Effect = DragDropEffects.None
    End Sub
    ''' <summary>
    ''' Обработва движението на влачен възел върху TreeView.
    ''' Проверява дали целевият възел е валиден и ако е:
    ''' - маркира го визуално
    ''' - разрешава операцията Move
    ''' Забранява:
    ''' - пускане върху "__EMPTY__"
    ''' - пускане върху самия себе си
    ''' - създаване на циклична йерархия
    ''' </summary>
    Private Sub TreeView1_DragOver(sender As Object, e As DragEventArgs) Handles TreeView1.DragOver
        ' По подразбиране забраняваме операцията
        e.Effect = DragDropEffects.None
        ' Нулираме предишната визуална маркировка
        ResetNodeHighlight()
        ' Опитваме се да извлечем влачения възел
        Dim draggedNode As TreeNode = TryCast(e.Data.GetData(GetType(TreeNode)), TreeNode)
        ' Ако няма валиден възел → прекратяваме
        If draggedNode Is Nothing Then Return
        ' Преобразуваме координатите на мишката към TreeView
        Dim targetPoint As Point = TreeView1.PointToClient(New Point(e.X, e.Y))
        ' Взимаме възела под курсора
        Dim targetNode As TreeNode = TreeView1.GetNodeAt(targetPoint)
        ' Ако няма целеви възел → прекратяваме
        If targetNode Is Nothing Then Return
        ' Забраняваме пускане върху "Без име"
        If targetNode.Name = "__EMPTY__" Then Return
        ' Забраняваме пускане върху самия себе си
        If targetNode.Name = draggedNode.Name Then Return
        ' Забраняваме създаване на цикъл (родител → дете)
        If IsAlreadyChildOf(targetNode, draggedNode) Then Return
        ' Ако сме тук → целта е валидна
        ' Маркираме визуално възела
        MarkNodeHighlight(targetNode)
        ' Позволяваме преместване
        e.Effect = DragDropEffects.Move
    End Sub
    ''' <summary>
    ''' Обработва напускане на зоната на TreeView по време на Drag&amp;Drop.
    ''' Премахва визуалната маркировка от възлите.
    ''' </summary>
    Private Sub TreeView1_DragLeave(sender As Object, e As EventArgs) Handles TreeView1.DragLeave
        ' Премахваме всякаква визуална маркировка
        ResetNodeHighlight()
    End Sub
    ''' <summary>
    ''' Събитие при кликване върху възел в TreeView. 
    ''' Ако е десен бутон - маркираме възела и показваме контекстно меню.
    ''' </summary>
    Private Sub TreeView1_NodeMouseClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseClick
        ' Проверяваме дали е натиснат десният бутон
        If e.Button = MouseButtons.Right Then
            ' Важно: Маркираме възела, върху който сме кликнали (Visual feedback)
            TreeView1.SelectedNode = e.Node
            ' Създаваме контекстното меню "в движение"
            ' (За по-големи проекти е по-добре да е предварително дефинирано в дизайнера)
            Dim ctxMenu As New ContextMenuStrip()
            ' Създаваме елемента за обновяване
            Dim refreshItem As New ToolStripMenuItem("🔄 Обновяване на списъка", Nothing, AddressOf RefreshTree_Click)
            ctxMenu.Items.Add(refreshItem)
            ' Показваме менюто точно на позицията на мишката
            ctxMenu.Show(TreeView1, e.Location)
        End If
    End Sub
    ''' <summary>
    ''' Логика за преизграждане на дървото.
    ''' </summary>
    Private Sub RefreshTree_Click(sender As Object, e As EventArgs)
        ' Сменяме курсора на "изчакване", за да знае потребителят, че нещо се случва
        Cursor = Cursors.WaitCursor
        ' Спираме прерисуването на контролата, за да няма трептене (flickering)
        TreeView1.BeginUpdate()
        Try
            ' 1. Изчистваме старите данни (ако BuildTreeView го изисква)
            TreeView1.Nodes.Clear()
            ' 2. Извикваме твоята основна процедура за пълнене на данни
            BuildTreeViewFromKonsumatori()
            ' 3. По желание: Разгъваме първия възел и го избираме
            If TreeView1.Nodes.Count > 0 Then
                TreeView1.Nodes(0).Expand()
                TreeView1.SelectedNode = TreeView1.Nodes(0)
            End If
        Catch ex As Exception
            ' Показваме съобщение при грешка (напр. проблем с базата данни)
            MessageBox.Show("Грешка при обновяване на данните: " & vbCrLf & ex.Message,
                        "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' ВИНАГИ пускаме прерисуването обратно и връщаме нормалния курсор
            TreeView1.EndUpdate()
            Cursor = Cursors.Default
        End Try
    End Sub
    ''' <summary>
    ''' Обработва операцията Drop в TreeView.
    ''' Премества табло в нов родител, като:
    ''' - валидира операцията
    ''' - премахва стария "фийдър"
    ''' - добавя нов "фийдър"
    ''' - преизчислява засегнатите табла
    ''' - обновява визуализацията
    ''' </summary>
    Private Sub TreeView1_DragDrop(sender As Object, e As DragEventArgs) Handles TreeView1.DragDrop
        ' Премахва визуалната маркировка от предишно Drag
        ResetNodeHighlight()
        ' Взима влачения възел от Drag данните
        Dim draggedNode As TreeNode = TryCast(e.Data.GetData(GetType(TreeNode)), TreeNode)
        ' Определя позицията на курсора спрямо TreeView
        Dim targetPoint As Point = TreeView1.PointToClient(New Point(e.X, e.Y))
        ' Взима възела под курсора (целевия родител)
        Dim targetNode As TreeNode = TreeView1.GetNodeAt(targetPoint)
        ' Валидации – прекратява при невалидни условия
        If draggedNode Is Nothing OrElse targetNode Is Nothing Then Return
        If targetNode.Name = "__EMPTY__" OrElse targetNode.Name = draggedNode.Name Then Return
        If IsAlreadyChildOf(targetNode, draggedNode) Then Return
        ' Взима имената на стария и новия родител
        Dim rawOldParent As String = If(draggedNode.Parent IsNot Nothing, draggedNode.Parent.Name, "")
        Dim rawNewParent As String = targetNode.Name
        ' Нормализира имената (специално за корена)
        Dim oldParentName As String = If(rawOldParent = ROOT_NODE_NAME, ROOT_NODE_TEXT, rawOldParent)
        Dim newParentName As String = If(rawNewParent = ROOT_NODE_NAME, ROOT_NODE_TEXT, rawNewParent)
        ' Ако няма реална промяна → прекратява
        If String.IsNullOrEmpty(oldParentName) OrElse oldParentName = newParentName Then Return
        ' Търси стария "фийдър" (връзката между стар родител и табло)
        Dim oldFeeder = ListTokow.FirstOrDefault(
                                Function(x)
                                    Return x.Device = "Дете" AndAlso
                                           x.ТоковКръг = draggedNode.Name AndAlso
                                           (x.Tablo = oldParentName Or x.Tablo = rawOldParent)
                                End Function)
        ' Ако е намерен → премахва го
        If oldFeeder IsNot Nothing Then ListTokow.Remove(oldFeeder)
        ' Запазва референция към стария родител (за по-късно обновяване)
        Dim oldParentNode As TreeNode = draggedNode.Parent
        ' Премества възела визуално в TreeView
        draggedNode.Remove()
        targetNode.Nodes.Add(draggedNode)
        targetNode.Expand()
        TreeView1.SelectedNode = draggedNode
        ' Добавя нов "фийдър" към новия родител
        PrepareSourcePanelData(draggedNode.Name, newParentName)
        ' Преизчислява стария родител (намалява натоварването)
        BuildPanelSummaryRecord(oldParentName)
        If oldParentNode IsNot Nothing Then RefreshNodeText(oldParentNode)
        ' Преизчислява новия родител (увеличава натоварването)
        BuildPanelSummaryRecord(newParentName)
        RefreshNodeText(targetNode)
        ' Преизчислява самото преместено табло
        BuildPanelSummaryRecord(draggedNode.Name)
        RefreshNodeText(draggedNode)
        ' Обновява всички родители нагоре по дървото
        If oldParentNode IsNot Nothing Then UpdatePathToRoot(oldParentNode)
        UpdatePathToRoot(targetNode)
        ' Финално сортиране и обновяване на визуализацията
        SortCircuits()
        SetupDataGridView_Total()
    End Sub
    ''' <summary>
    ''' Обновява всички табла по веригата от даден възел до корена.
    ''' За всяко табло:
    ''' - преизчислява агрегираните стойности
    ''' - обновява визуалния текст в TreeView
    ''' </summary>
    Private Sub UpdatePathToRoot(startNode As TreeNode)
        ' Започваме от подадения възел
        Dim current As TreeNode = startNode
        ' Обхождаме нагоре по дървото до корена
        While current IsNot Nothing
            ' Име на текущия възел в TreeView
            Dim panelName As String = current.Name
            ' Нормализиране на името за работа с ListTokow
            ' Ако е корен → използваме текстовото име
            Dim dataName As String =
            If(panelName = ROOT_NODE_NAME, ROOT_NODE_TEXT, panelName)
            ' Преизчислява обобщените данни за текущото табло
            BuildPanelSummaryRecord(dataName)
            ' Обновява текста на възела в TreeView (мощност, брой кръгове и др.)
            RefreshNodeText(current)
            ' Ако сме достигнали корена → прекратяваме
            If panelName = ROOT_NODE_NAME Then Exit While
            ' Преминаваме към родителския възел
            current = current.Parent
        End While
    End Sub
    ''' <summary>
    ''' Обновява текста на конкретен възел според актуалните данни в ListTokow
    ''' </summary>
    Private Sub RefreshNodeText(node As TreeNode)
        If node Is Nothing OrElse node.Name = ROOT_NODE_NAME OrElse node.Name = "__EMPTY__" Then Return
        Dim panelName = node.Name
        Dim records = ListTokow.Where(Function(x) String.Equals(x.Tablo, panelName, StringComparison.OrdinalIgnoreCase)).ToList()
        ' Брой уникални кръгове
        Dim circuitCount = records.Select(Function(r) r.ТоковКръг).Distinct().Count()
        ' Мощност: четем я от главния запис (Device="Табло"), както е в BuildTreeViewFromKonsumatori
        Dim masterRec = records.FirstOrDefault(Function(r) String.Equals(r.Device, "Табло", StringComparison.OrdinalIgnoreCase))
        Dim totalPower As Double = If(masterRec IsNot Nothing, masterRec.Мощност, 0)
        ' Прилагаме същия формат, който вече използваш
        'node.Text = $"{panelName} ({circuitCount} {If(circuitCount = 1, "кръг", "кръга")}, {totalPower:F1}kW)"
        node.Text = $"{panelName} ({totalPower:F1}kW)"
    End Sub
    ''' <summary>
    ''' Намира, валидира и копира в паметта записа на изходното табло.
    ''' </summary>
    Private Sub PrepareSourcePanelData(sourceName As String, targetName As String)
        Dim matches = ListTokow.Where(Function(x) x.Tablo = sourceName AndAlso x.Device = "Табло").ToList()
        If matches.Count = 0 Then
            MessageBox.Show($"Не е намерен запис за табло '{sourceName}' с device = 'Табло'.", "Липсващи данни", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        If matches.Count > 1 Then
            MessageBox.Show($"Намерени са {matches.Count} записа за табло '{sourceName}'. Очаква се точно един.", "Дублирани данни", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        ' 1. Създаваме независимо копие
        Dim feederRecord As strTokow = matches(0).Clone()
        ' 2. Променяме САМО копието за връзката
        feederRecord.ТоковКръг = sourceName
        feederRecord.Tablo = targetName
        feederRecord.Device = "Дете"
        feederRecord.Табло_Родител = ""
        feederRecord.Консуматор = "Табло"
        feederRecord.предназначение = sourceName
        ' 3. Добавяме в списъка
        ListTokow.Add(feederRecord)
        ' Оразмеряване на прекъсвач и кабел за новия ТК
        Try
            calcBreaker = True
            CalculateBreaker(feederRecord)
        Finally
            calcBreaker = False ' Гарантираме нулиране дори при грешка
        End Try
        CalculateCable(feederRecord)
        SortCircuits()
        SetupDataGridView_Total()
    End Sub
    ''' <summary>
    ''' Инициализира йерархията още при стартиране:
    ''' 1. Създава изрични връзки ("Дете") в родителското табло за всяко уникално табло.
    ''' 2. Уверява се, че всяко табло има изчислен "ОБЩО" запис.
    ''' 3. Преизчислява родителското табло, за да сумира всички връзки.
    ''' </summary>
    ''' <param name="parentName">Име на родителското табло (напр. "Електромерно табло")</param>
    Private Sub InitializePanelParents(parentName As String)
        If ListTokow Is Nothing OrElse ListTokow.Count = 0 Then Return
        ' ─────────────────────────────────────────────────────
        ' 2. СЪЗДАВАНЕ НА ВРЪЗКИ (ФИЙДЪРИ) В РОДИТЕЛСКОТО ТАБЛО
        ' ─────────────────────────────────────────────────────
        Dim uniquePanels = ListTokow.Where(Function(x) Not String.IsNullOrWhiteSpace(x.Tablo)).
                                     Select(Function(x) x.Tablo.Trim()).Distinct().ToList()
        For Each pName In uniquePanels
            ' Пропускаме самото родителско табло
            If String.Equals(pName, parentName, StringComparison.OrdinalIgnoreCase) Then Continue For
            ' Проверяваме дали връзката вече съществува в родителското табло
            Dim linkExists = ListTokow.Any(Function(x) x.Tablo = parentName AndAlso
                                           x.Device = "Дете" AndAlso
                                           String.Equals(x.ТоковКръг, pName, StringComparison.OrdinalIgnoreCase))
            If Not linkExists Then
                Dim childMaster = ListTokow.FirstOrDefault(Function(x) x.Tablo = pName AndAlso x.Device = "Табло")
                If childMaster IsNot Nothing Then
                    ' ✅ Създаваме независимо копие чрез Clone()
                    ' Това автоматично копира Мощност, Ток, Кабел, Полюси и т.н.
                    Dim feeder As strTokow = childMaster.Clone()
                    ' Променяме САМО полетата, които превръщат записа във връзка (фийдър)
                    feeder.Tablo = parentName
                    feeder.Device = "Дете"
                    feeder.ТоковКръг = pName
                    feeder.Табло_Родител = ""
                    feeder.Консуматор = "Табло"
                    feeder.предназначение = pName
                    ' Добавяме новата връзка в общия списък
                    ListTokow.Add(feeder)
                End If
            End If
        Next
        ' ─────────────────────────────────────────────────────
        ' 3. ПРЕИЗЧИСЛЯВАНЕ НА РОДИТЕЛСКОТО ТАБЛО
        ' ─────────────────────────────────────────────────────
        ' Сега то ще сумира всички току-що създадени фийдъри
        BuildPanelSummaryRecord(parentName)
    End Sub
    ' ─────────────────────────────────────────────────────────────
    ' ПОМОЩНИ МЕТОДИ
    ' ─────────────────────────────────────────────────────────────
    Private Sub MarkNodeHighlight(node As TreeNode)
        If node Is highlightNode Then Return ' Вече е маркирано
        highlightNode = node
        originalBackColor = node.BackColor
        originalForeColor = node.ForeColor
        node.BackColor = Color.LightGreen
        node.ForeColor = Color.DarkGreen
        node.EnsureVisible()
    End Sub
    Private Sub ResetNodeHighlight()
        If highlightNode IsNot Nothing Then
            highlightNode.BackColor = originalBackColor
            highlightNode.ForeColor = originalForeColor
            highlightNode = Nothing
        End If
    End Sub
    Private Function IsAlreadyChildOf(nodeA As TreeNode, nodeB As TreeNode) As Boolean
        Dim current As TreeNode = nodeA.Parent
        While current IsNot Nothing
            If current.Name = nodeB.Name Then Return True
            current = current.Parent
        End While
        Return False
    End Function
    ''' <summary>
    ''' Създава и конфигурира основната структура на DataGridView за показване на електрически табла и кръгове.
    ''' Изграждат се фиксирани колони (Параметър, Мерна единица, ОБЩО) и динамични колони според rowData.
    ''' Клетките се генерират според тип (ComboBox, CheckBox, TextBox), след което се прилага визуално форматиране.
    ''' </summary>
    Private Sub SetupDataGridView()
        ' Изчистване на старата структура
        DataGridView1.Columns.Clear()
        DataGridView1.Rows.Clear()
        DataGridView1.RowHeadersVisible = False
        ' =====================================================
        ' 1. ПЪРВА КОЛОНА: Параметри (описателна колона)
        ' =====================================================
        Dim colParam As New DataGridViewTextBoxColumn()
        colParam.Name = "colParameter"
        colParam.HeaderText = "Параметър"
        colParam.Width = 200
        colParam.Frozen = True
        colParam.DefaultCellStyle.Font = New Font("Arial", 10, FontStyle.Bold)
        colParam.DefaultCellStyle.BackColor = Color.FromArgb(200, 220, 255)
        colParam.SortMode = DataGridViewColumnSortMode.NotSortable
        DataGridView1.Columns.Add(colParam)
        ' =====================================================
        ' 2. ВТОРА КОЛОНА: Мерни единици
        ' =====================================================
        Dim colUnit As New DataGridViewTextBoxColumn()
        colUnit.Name = "colUnit"
        colUnit.HeaderText = ""
        colUnit.Width = 50
        colUnit.Frozen = True
        colUnit.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        colUnit.DefaultCellStyle.Font = New Drawing.Font("Arial", 10, FontStyle.Italic)
        colUnit.DefaultCellStyle.ForeColor = Color.Gray
        colUnit.SortMode = DataGridViewColumnSortMode.NotSortable
        DataGridView1.Columns.Add(colUnit)
        ' =====================================================
        ' 3. КОЛОНА: ОБЩО (резултатна колона)
        ' =====================================================
        Dim colTotal As New DataGridViewTextBoxColumn()
        colTotal.Name = "colTotal"
        colTotal.HeaderText = "ОБЩО"
        colTotal.Width = 130
        colTotal.DefaultCellStyle.Font = New Font("Arial", 10, FontStyle.Bold)
        colTotal.DefaultCellStyle.BackColor = Color.FromArgb(230, 240, 255)
        colTotal.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        colTotal.SortMode = DataGridViewColumnSortMode.NotSortable
        DataGridView1.Columns.Add(colTotal)
        ' =====================================================
        ' 4. РЕДОВЕ: попълване от rowData шаблона
        ' =====================================================
        For Each row As String() In rowData
            Dim dgvRow As New DataGridViewRow()
            dgvRow.CreateCells(DataGridView1)
            ' Параметър
            dgvRow.Cells(0).Value = row(0)
            ' Мерна единица
            dgvRow.Cells(1).Value = row(1)
            ' Тип на клетките за останалите колони
            Dim cellType As String = row(2)
            ' Генериране на клетки за динамичните колони
            For colIndex As Integer = 2 To DataGridView1.Columns.Count - 2
                Dim cell As DataGridViewCell = Nothing
                Select Case cellType
                    Case "Combo"
                        cell = New DataGridViewComboBoxCell()
                        SetupComboBoxCell(cell, row(0), False)
                    Case "Check"
                        cell = New DataGridViewCheckBoxCell()
                    Case Else
                        cell = New DataGridViewTextBoxCell()
                End Select
                cell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgvRow.Cells(colIndex) = cell
            Next
            ' =====================================================
            ' Оцветяване на редове според типа параметър
            ' =====================================================
            Select Case row(0).ToString()
                Case "---------"
                    dgvRow.DefaultCellStyle.BackColor = Color.FromArgb(220, 220, 220)
                Case "Прекъсвач", "ДТЗ (RCD)", "Кабел"
                    dgvRow.DefaultCellStyle.BackColor = Color.FromArgb(180, 200, 255)
                    dgvRow.DefaultCellStyle.Font = New Font("Arial", 10, FontStyle.Bold)
                Case Else
                    ' стандартен стил
            End Select
            DataGridView1.Rows.Add(dgvRow)
        Next
        ' =====================================================
        ' 5. НАСТРОЙКИ
        ' =====================================================
        DataGridView1.AllowUserToAddRows = False                                    ' Забранява на потребителя да добавя празен нов ред в края на таблицата
        DataGridView1.AllowUserToDeleteRows = False                                 ' Забранява на потребителя да изтрива редове с натискане на Delete
        DataGridView1.ReadOnly = False                                              ' Позволява редакция на клетките (важно за ComboBox и CheckBox)
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None    ' Изключва автоматичното оразмеряване (разчита на зададен Width)
        DataGridView1.ColumnHeadersDefaultCellStyle.Font = New Font("Arial", 10, FontStyle.Bold) ' Задава шрифт Bold за заглавния ред
        DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter ' Центрира текста в заглавията на колоните
        DataGridView1.ColumnHeadersHeight = 25                                      ' Фиксира височината на заглавната лента на 170 пиксела
        DataGridView1.RowTemplate.Height = 25                                       ' Задава стандартна височина на всеки нов ред с данни
        DataGridView1.BackgroundColor = Color.White                                 ' Променя цвета на фона на самата контрола (зад редовете) на бял
        DataGridView1.GridColor = Color.Gray                                        ' Задава сив цвят за линиите на мрежата между клетките
        DataGridView1.BorderStyle = BorderStyle.Fixed3D                             ' Прави рамката на цялата таблица да изглежда обемна (3D)
        DataGridView1.CellBorderStyle = DataGridViewCellBorderStyle.Single          ' Задава единична тънка линия за граница между отделните клетки
    End Sub
    ''' <summary>
    ''' Конфигурира и попълва специалните колони в DataGridView1 за обобщен изглед.
    ''' Логиката:
    ''' 1. Определя кои колони (colTotal, colDiscon) съществуват в грида
    ''' 2. Обхожда всички редове и ги синхронизира с данните от rowData
    ''' 3. За всяка целева колона създава подходящ тип клетка (ComboBox, CheckBox или TextBox)
    ''' 4. Попълва ComboBox клетки според контекста на реда
    ''' 5. Прилага форматиране (цветове и стилове) според типа ред
    ''' </summary>
    Private Sub SetupDataGridView_Total()
        ' Списък с индекси на целевите колони, които ще се обработват
        Dim targetColumns As New List(Of Integer)
        ' Имена на колоните, които търсим в DataGridView
        Dim colNames() As String = {"colTotal", "colDiscon"}
        ' Проверка дали колоните съществуват в грида и взимане на индексите им
        For Each colName In colNames
            If DataGridView1.Columns.Contains(colName) Then
                targetColumns.Add(DataGridView1.Columns(colName).Index)
            End If
        Next
        ' Обхождане на всички редове в DataGridView
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            Dim dgvRow As DataGridViewRow = DataGridView1.Rows(i)
            ' Защита от несъответствие между визуалните редове и източника на данни
            If i >= rowData.Count Then Continue For
            ' Взима съответния ред от източника на данни
            Dim data As String() = rowData(i)
            ' Тип на реда (определя какви клетки ще се създадат)
            Dim cellType As String = data(2)
            ' Обхождане на целевите колони (Total / Disconnector)
            For Each colIndex In targetColumns
                Dim specialCell As DataGridViewCell = Nothing
                ' Определяне на типа клетка според cellType
                Select Case cellType
                    Case "Combo"
                        ' Клетка тип ComboBox
                        Dim comboCell As New DataGridViewComboBoxCell()
                        ' Допълнителна логика според първия елемент на реда (data(0))
                        Select Case data(0).ToString()
                            Case "Управление"
                                ' Специална инициализация за управление
                                SetupComboBoxCell(comboCell, data(0), True)
                            Case "Тип на апарата"
                                ' Попълване от списък с прекъсвачи
                                comboCell.Items.Clear()
                                comboCell.Items.AddRange(Disconnectors_For_combo.ToArray())
                            Case "Номинален ток"
                                ' Попълване на номинални токове
                                comboCell.Items.Clear()
                                comboCell.Items.AddRange(Discon_Tok_For_combo.ToArray())
                            Case "Тип кабел"
                                ' Попълване на кабели
                                comboCell.Items.Clear()
                                comboCell.Items.AddRange(Cable_For_combo.ToArray())
                            Case "Начин на монтаж"
                                ' Попълване от дефиниран списък с монтажи
                                comboCell.Items.Clear()
                                comboCell.Items.AddRange(LiMountMethod.Select(Function(m) m.Text).ToArray())
                            Case "Начин на полагане"
                                ' Фиксиран списък за полагане
                                Dim valuesLaying As New List(Of String) From {"Във въздух", "В земя"}
                                comboCell.Items.Clear()
                                comboCell.Items.AddRange(valuesLaying.ToArray())
                            Case Else
                                ' Няма дефинирана логика за този случай
                        End Select
                        specialCell = comboCell
                    Case "Check"
                        ' Клетка тип Checkbox
                        specialCell = New DataGridViewCheckBoxCell()
                        ' Центриране на съдържанието
                        specialCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                        ' Задаване на начална стойност
                        specialCell.Value = False
                    Case Else
                        ' Default: текстова клетка
                        specialCell = New DataGridViewTextBoxCell()
                        specialCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                End Select
                ' Подмяна на клетката в конкретния ред и колона
                dgvRow.Cells(colIndex) = specialCell
            Next
            ' Вземане на стойността от първата колона (за определяне на стил на реда)
            Dim firstVal As String = If(dgvRow.Cells(0).Value IsNot Nothing,
                                    dgvRow.Cells(0).Value.ToString(),
                                    "")
            ' Форматиране на целия ред според типа съдържание
            Select Case firstVal
                Case "---------"
                    ' Сив разделителен ред
                    dgvRow.DefaultCellStyle.BackColor = Color.FromArgb(220, 220, 220)
                Case "Прекъсвач", "ДТЗ (RCD)", "Кабел"
                    ' Акцентни редове за основни компоненти
                    dgvRow.DefaultCellStyle.BackColor = Color.FromArgb(180, 200, 255)
                    dgvRow.DefaultCellStyle.Font = New Font("Arial", 10, FontStyle.Bold)
            End Select
        Next
    End Sub
    ''' <summary>
    ''' Настройва DataGridViewComboBoxCell с подходящи елементи според подаден параметър.
    ''' </summary>
    ''' <param name="cell">Клетката от DataGridView, която ще бъде превърната в ComboBox.</param>
    ''' <param name="parameter">
    ''' Параметърът определя какви стойности да се добавят в ComboBox:
    ''' - "Тип на апарата" – зарежда списък с прекъсвачи (Breakers_For_combo)
    ''' - "Номинален ток" – зарежда стандартни номинални токове
    ''' - "Крива" – зарежда типове токови криви B, C, D
    ''' - "Управление" – зарежда възможни видове управление (например импулсно реле, моторна защита и др.)
    ''' - "Тип" – може да зарежда типове кабели (коментарът показва, че е оставено за бъдеща имплементация)
    ''' </param>
    ''' <remarks>
    ''' ✅ Добра практика:
    ''' - Изчиства старите елементи в ComboBox преди добавяне на нови.
    ''' - Задава първия елемент като стойност на клетката, за да се избегне празна стойност.
    ''' - Настройва DisplayStyle на DropDown, позволявайки потребителя да вижда падащия списък.
    '''
    ''' Потенциални фрапиращи моменти:
    ''' - Ако Breakers_For_combo е празен, клетката ще зададе стойност на първия елемент, което може да хвърли изключение.
    ''' - Коментарът при "Тип" показва, че добавянето на Kable_Type не е реализирано, което може да доведе до липсващи данни.
    ''' - Не се проверява дали подаденият cell е наистина DataGridViewComboBoxCell извън CType кастинга.
    ''' </remarks>
    Private Sub SetupComboBoxCell(cell As DataGridViewCell,
                                  parameter As String,
                                  Discon As Boolean
                                  )
        ' Преобразуваме клетката към ComboBoxCell
        Dim comboCell As DataGridViewComboBoxCell = CType(cell, DataGridViewComboBoxCell)
        ' Изчистваме всички стари елементи, за да не се дублират
        comboCell.Items.Clear()
        ' Добавяме елементи според типа параметър
        Select Case parameter
            Case "Тип на апарата"
                If Discon Then
                    comboCell.Items.AddRange(Disconnectors_For_combo.ToArray())
                Else
                    comboCell.Items.AddRange(Breakers_For_combo.ToArray())
                End If
            Case "Номинален ток"
                If Discon Then
                    comboCell.Items.AddRange(Discon_Tok_For_combo.ToArray())
                Else
                    comboCell.Items.AddRange("6", "10", "16", "20", "25", "32", "40", "50", "63")
                End If
            Case "Крива"
                If Discon Then
                    comboCell.Items.AddRange("-")
                Else
                    comboCell.Items.AddRange("C", "B", "D")
                End If
            Case "Управление"
                If Discon Then
                    comboCell.Items.AddRange("Няма")
                Else
                    comboCell.Items.AddRange("Няма",
                     "Фото реле",
                     "Стълбищен автомат",
                     "Импулсно реле",
                     "Контактор",
                     "Моторна защита",
                     "Моторен механизъм",
                     "Честотен регулатор",
                     "Електромер"
                     )
                End If
            Case "Тип кабел"
                comboCell.Items.AddRange(Cable_For_combo.ToArray())
                ' Възможно зареждане на Kable_Type в бъдеще
        End Select
        ' ✅ Задаваме първия елемент като стойност, за да избегнем празна клетка
        If comboCell.Items.Count > 0 Then comboCell.Value = comboCell.Items(0)
        ' Настройка на вида на ComboBox:
        ' - DropDown позволява избор от списъка с възможност за писане
        ' - DropDownList ограничава избора само до елементите от списъка
        comboCell.DisplayStyle = ComboBoxStyle.DropDown
    End Sub
    ''' <summary>
    ''' Обработва грешки, възникнали при въвеждане или обработка на данни в DataGridView.
    '''
    ''' DataError събитието се извиква когато:
    ''' - въведената стойност не може да се конвертира към типа на колоната
    '''   (например текст в числова колона)
    ''' - стойността на ComboBox не съществува в списъка
    ''' - възникне грешка при запис в DataSource
    ''' - има проблем при parsing или validation на данните
    '''
    ''' В тази реализация грешките се потискат, за да не се показва
    ''' стандартният диалог на DataGridView и да не се прекъсва работата
    ''' на приложението.
    '''
    ''' Свойства:
    ''' e.ThrowException = False
    '''     → предотвратява хвърлянето на exception след обработката
    '''       на събитието. Това спира стандартния crash или popup. :contentReference[oaicite:0]{index=0}
    '''
    ''' e.Cancel = True
    '''     → отменя текущата операция върху клетката и предотвратява
    '''       продължаването на грешната операция.
    '''
    ''' Цел:
    ''' Да се избегнат runtime грешки в DataGridView при невалидни
    ''' стойности (например при ComboBox клетки или при грешни типове).
    '''
    ''' Забележка:
    ''' Ако е необходимо, тук може да се добави логика за:
    ''' - показване на MessageBox
    ''' - записване в лог
    ''' - визуално маркиране на грешната клетка
    ''' </summary>
    Private Sub DataGridView1_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
        ' ------------------------------------------------------------
        ' Забранява хвърлянето на exception след обработката на събитието
        ' ------------------------------------------------------------
        e.ThrowException = False
        ' ------------------------------------------------------------
        ' Отменя текущата операция върху клетката
        ' ------------------------------------------------------------
        e.Cancel = True
    End Sub
    ' ФУНКЦИЯ ЗА ЗАРЕЖДАНЕ НА КАТАЛОЗИТЕ
    Private Sub SetCatalog()
        ' Речник за всички кабели
        FillCables()
        ' Речник за всички автоматични прекъсвачи
        ' Инициализиране на списъка
        Breakers = New List(Of BreakerInfo)
        ' Попълване на всички прекъсвачи чрез отделната процедура
        FillBreakers()
        ' ============================================================
        ' РАЗЕДИНИТЕЛИ (Товарови прекъсвачи)
        ' ============================================================
        Disconnectors = New List(Of DisconnectorInfo) From {
                            New DisconnectorInfo With {.NominalCurrent = 20, .Type = "iSW", .Brand = "Acti9", .Poles = 1},
                            New DisconnectorInfo With {.NominalCurrent = 25, .Type = "iSW", .Brand = "Acti9", .Poles = 1},
                            New DisconnectorInfo With {.NominalCurrent = 32, .Type = "iSW", .Brand = "Acti9", .Poles = 1},
                            New DisconnectorInfo With {.NominalCurrent = 40, .Type = "iSW", .Brand = "Acti9", .Poles = 1},
                            New DisconnectorInfo With {.NominalCurrent = 63, .Type = "iSW", .Brand = "Acti9", .Poles = 1},
                            New DisconnectorInfo With {.NominalCurrent = 80, .Type = "iSW", .Brand = "Acti9", .Poles = 1},
                            New DisconnectorInfo With {.NominalCurrent = 100, .Type = "iSW", .Brand = "Acti9", .Poles = 1},
                            New DisconnectorInfo With {.NominalCurrent = 125, .Type = "iSW", .Brand = "Acti9", .Poles = 1},
                            New DisconnectorInfo With {.NominalCurrent = 20, .Type = "iSW", .Brand = "Acti9", .Poles = 2},
                            New DisconnectorInfo With {.NominalCurrent = 25, .Type = "iSW", .Brand = "Acti9", .Poles = 2},
                            New DisconnectorInfo With {.NominalCurrent = 32, .Type = "iSW", .Brand = "Acti9", .Poles = 2},
                            New DisconnectorInfo With {.NominalCurrent = 40, .Type = "iSW", .Brand = "Acti9", .Poles = 2},
                            New DisconnectorInfo With {.NominalCurrent = 63, .Type = "iSW", .Brand = "Acti9", .Poles = 2},
                            New DisconnectorInfo With {.NominalCurrent = 80, .Type = "iSW", .Brand = "Acti9", .Poles = 2},
                            New DisconnectorInfo With {.NominalCurrent = 100, .Type = "iSW", .Brand = "Acti9", .Poles = 2},
                            New DisconnectorInfo With {.NominalCurrent = 125, .Type = "iSW", .Brand = "Acti9", .Poles = 2},
                            New DisconnectorInfo With {.NominalCurrent = 20, .Type = "iSW", .Brand = "Acti9", .Poles = 3},
                            New DisconnectorInfo With {.NominalCurrent = 25, .Type = "iSW", .Brand = "Acti9", .Poles = 3},
                            New DisconnectorInfo With {.NominalCurrent = 32, .Type = "iSW", .Brand = "Acti9", .Poles = 3},
                            New DisconnectorInfo With {.NominalCurrent = 40, .Type = "iSW", .Brand = "Acti9", .Poles = 3},
                            New DisconnectorInfo With {.NominalCurrent = 63, .Type = "iSW", .Brand = "Acti9", .Poles = 3},
                            New DisconnectorInfo With {.NominalCurrent = 80, .Type = "iSW", .Brand = "Acti9", .Poles = 3},
                            New DisconnectorInfo With {.NominalCurrent = 100, .Type = "iSW", .Brand = "Acti9", .Poles = 3},
                            New DisconnectorInfo With {.NominalCurrent = 125, .Type = "iSW", .Brand = "Acti9", .Poles = 3},
                            New DisconnectorInfo With {.NominalCurrent = 20, .Type = "iSW", .Brand = "Acti9", .Poles = 4},
                            New DisconnectorInfo With {.NominalCurrent = 25, .Type = "iSW", .Brand = "Acti9", .Poles = 4},
                            New DisconnectorInfo With {.NominalCurrent = 32, .Type = "iSW", .Brand = "Acti9", .Poles = 4},
                            New DisconnectorInfo With {.NominalCurrent = 40, .Type = "iSW", .Brand = "Acti9", .Poles = 4},
                            New DisconnectorInfo With {.NominalCurrent = 63, .Type = "iSW", .Brand = "Acti9", .Poles = 4},
                            New DisconnectorInfo With {.NominalCurrent = 80, .Type = "iSW", .Brand = "Acti9", .Poles = 4},
                            New DisconnectorInfo With {.NominalCurrent = 100, .Type = "iSW", .Brand = "Acti9", .Poles = 4},
                            New DisconnectorInfo With {.NominalCurrent = 125, .Type = "iSW", .Brand = "Acti9", .Poles = 4},
                            New DisconnectorInfo With {.NominalCurrent = 100, .Type = "INS", .Brand = "Easy9", .Poles = 3},
                            New DisconnectorInfo With {.NominalCurrent = 125, .Type = "INS", .Brand = "Easy9", .Poles = 3},
                            New DisconnectorInfo With {.NominalCurrent = 160, .Type = "INS", .Brand = "Easy9", .Poles = 3},
                            New DisconnectorInfo With {.NominalCurrent = 200, .Type = "INS", .Brand = "Easy9", .Poles = 3},
                            New DisconnectorInfo With {.NominalCurrent = 250, .Type = "INS", .Brand = "Easy9", .Poles = 3},
                            New DisconnectorInfo With {.NominalCurrent = 315, .Type = "INS", .Brand = "Easy9", .Poles = 3},
                            New DisconnectorInfo With {.NominalCurrent = 400, .Type = "INS", .Brand = "Easy9", .Poles = 3},
                            New DisconnectorInfo With {.NominalCurrent = 500, .Type = "INS", .Brand = "Easy9", .Poles = 3},
                            New DisconnectorInfo With {.NominalCurrent = 630, .Type = "INS", .Brand = "Easy9", .Poles = 3},
                            New DisconnectorInfo With {.NominalCurrent = 800, .Type = "INS", .Brand = "Easy9", .Poles = 3},
                            New DisconnectorInfo With {.NominalCurrent = 1000, .Type = "INS", .Brand = "Easy9", .Poles = 3},
                            New DisconnectorInfo With {.NominalCurrent = 1250, .Type = "INS", .Brand = "Easy9", .Poles = 3},
                            New DisconnectorInfo With {.NominalCurrent = 1600, .Type = "INS", .Brand = "Easy9", .Poles = 3},
                            New DisconnectorInfo With {.NominalCurrent = 100, .Type = "INS", .Brand = "Easy9", .Poles = 4},
                            New DisconnectorInfo With {.NominalCurrent = 125, .Type = "INS", .Brand = "Easy9", .Poles = 4},
                            New DisconnectorInfo With {.NominalCurrent = 160, .Type = "INS", .Brand = "Easy9", .Poles = 4},
                            New DisconnectorInfo With {.NominalCurrent = 200, .Type = "INS", .Brand = "Easy9", .Poles = 4},
                            New DisconnectorInfo With {.NominalCurrent = 250, .Type = "INS", .Brand = "Easy9", .Poles = 4},
                            New DisconnectorInfo With {.NominalCurrent = 315, .Type = "INS", .Brand = "Easy9", .Poles = 4},
                            New DisconnectorInfo With {.NominalCurrent = 400, .Type = "INS", .Brand = "Easy9", .Poles = 4},
                            New DisconnectorInfo With {.NominalCurrent = 500, .Type = "INS", .Brand = "Easy9", .Poles = 4},
                            New DisconnectorInfo With {.NominalCurrent = 630, .Type = "INS", .Brand = "Easy9", .Poles = 4},
                            New DisconnectorInfo With {.NominalCurrent = 800, .Type = "INS", .Brand = "Easy9", .Poles = 4},
                            New DisconnectorInfo With {.NominalCurrent = 1000, .Type = "INS", .Brand = "Easy9", .Poles = 4},
                            New DisconnectorInfo With {.NominalCurrent = 1250, .Type = "INS", .Brand = "Easy9", .Poles = 4},
                            New DisconnectorInfo With {.NominalCurrent = 1600, .Type = "INS", .Brand = "Easy9", .Poles = 4},
                            New DisconnectorInfo With {.NominalCurrent = 1600, .Type = "IN", .Brand = "Acti9", .Poles = 3},
                            New DisconnectorInfo With {.NominalCurrent = 2000, .Type = "IN", .Brand = "Acti9", .Poles = 3},
                            New DisconnectorInfo With {.NominalCurrent = 2500, .Type = "IN", .Brand = "Acti9", .Poles = 3},
                            New DisconnectorInfo With {.NominalCurrent = 1600, .Type = "IN", .Brand = "Acti9", .Poles = 4},
                            New DisconnectorInfo With {.NominalCurrent = 2000, .Type = "IN", .Brand = "Acti9", .Poles = 4},
                            New DisconnectorInfo With {.NominalCurrent = 2500, .Type = "IN", .Brand = "Acti9", .Poles = 4}
        }
        Disconnectors_For_combo = Disconnectors.Select(Function(b) b.Type).Distinct().ToList()
        Discon_Tok_For_combo = Disconnectors.
                       Select(Function(d) d.NominalCurrent).
                       Distinct().
                       OrderBy(Function(n) n).
                       Select(Function(n) n.ToString()).
                       ToList()
        ' --- 4. МЕДНИ ШИНИ ---
        Busbars_Cu = New List(Of BusbarInfo) From {
        New BusbarInfo With {.CurrentCapacity = 210, .Section = "15x3", .Material = "Cu"},
        New BusbarInfo With {.CurrentCapacity = 275, .Section = "20x3", .Material = "Cu"},
        New BusbarInfo With {.CurrentCapacity = 340, .Section = "25x3", .Material = "Cu"},
        New BusbarInfo With {.CurrentCapacity = 475, .Section = "30x4", .Material = "Cu"},
        New BusbarInfo With {.CurrentCapacity = 625, .Section = "40x4", .Material = "Cu"},
        New BusbarInfo With {.CurrentCapacity = 700, .Section = "40x5", .Material = "Cu"},
        New BusbarInfo With {.CurrentCapacity = 860, .Section = "50x5", .Material = "Cu"},
        New BusbarInfo With {.CurrentCapacity = 955, .Section = "50x6", .Material = "Cu"},
        New BusbarInfo With {.CurrentCapacity = 1125, .Section = "60x6", .Material = "Cu"},
        New BusbarInfo With {.CurrentCapacity = 1480, .Section = "80x6", .Material = "Cu"},
        New BusbarInfo With {.CurrentCapacity = 1810, .Section = "100x6", .Material = "Cu"}
    }
        ' --- 5. АЛУМИНИЕВИ ШИНИ ---
        Busbars_Al = New List(Of BusbarInfo) From {
        New BusbarInfo With {.CurrentCapacity = 165, .Section = "15x3", .Material = "Al"},
        New BusbarInfo With {.CurrentCapacity = 215, .Section = "20x3", .Material = "Al"},
        New BusbarInfo With {.CurrentCapacity = 265, .Section = "25x3", .Material = "Al"},
        New BusbarInfo With {.CurrentCapacity = 365, .Section = "30x4", .Material = "Al"},
        New BusbarInfo With {.CurrentCapacity = 480, .Section = "40x4", .Material = "Al"},
        New BusbarInfo With {.CurrentCapacity = 540, .Section = "40x5", .Material = "Al"},
        New BusbarInfo With {.CurrentCapacity = 665, .Section = "50x5", .Material = "Al"},
        New BusbarInfo With {.CurrentCapacity = 740, .Section = "50x6", .Material = "Al"},
        New BusbarInfo With {.CurrentCapacity = 870, .Section = "60x6", .Material = "Al"},
        New BusbarInfo With {.CurrentCapacity = 1150, .Section = "80x6", .Material = "Al"},
        New BusbarInfo With {.CurrentCapacity = 1425, .Section = "100x6", .Material = "Al"}
    }
        ' --- 6. ДТЗ / RCD ---
        RCD_Catalog = New List(Of RCDInfo) From {
                            New RCDInfo With {.Brand = "Schneider", .NominalCurrent = 25, .Type = "AC", .Poles = "2p", .Sensitivity = 30, .DeviceType = "EZ9 RCCB", .Breaker = False},
                            New RCDInfo With {.Brand = "Schneider", .NominalCurrent = 25, .Type = "AC", .Poles = "4p", .Sensitivity = 30, .DeviceType = "EZ9 RCCB", .Breaker = False},
                            New RCDInfo With {.Brand = "Schneider", .NominalCurrent = 40, .Type = "AC", .Poles = "2p", .Sensitivity = 30, .DeviceType = "EZ9 RCCB", .Breaker = False},
                            New RCDInfo With {.Brand = "Schneider", .NominalCurrent = 40, .Type = "AC", .Poles = "4p", .Sensitivity = 30, .DeviceType = "EZ9 RCCB", .Breaker = False},
                            New RCDInfo With {.Brand = "Schneider", .NominalCurrent = 63, .Type = "AC", .Poles = "2p", .Sensitivity = 30, .DeviceType = "EZ9 RCCB", .Breaker = False},
                            New RCDInfo With {.Brand = "Schneider", .NominalCurrent = 63, .Type = "AC", .Poles = "4p", .Sensitivity = 30, .DeviceType = "EZ9 RCCB", .Breaker = False},
                            New RCDInfo With {.Brand = "Schneider", .NominalCurrent = 6, .Type = "AC", .Poles = "2p", .Sensitivity = 30, .DeviceType = "EZ9 RCBO", .Breaker = True},
                            New RCDInfo With {.Brand = "Schneider", .NominalCurrent = 10, .Type = "AC", .Poles = "2p", .Sensitivity = 30, .DeviceType = "EZ9 RCBO", .Breaker = True},
                            New RCDInfo With {.Brand = "Schneider", .NominalCurrent = 16, .Type = "AC", .Poles = "2p", .Sensitivity = 30, .DeviceType = "EZ9 RCBO", .Breaker = True},
                            New RCDInfo With {.Brand = "Schneider", .NominalCurrent = 20, .Type = "AC", .Poles = "2p", .Sensitivity = 30, .DeviceType = "EZ9 RCBO", .Breaker = True},
                            New RCDInfo With {.Brand = "Schneider", .NominalCurrent = 25, .Type = "AC", .Poles = "2p", .Sensitivity = 30, .DeviceType = "EZ9 RCBO", .Breaker = True},
                            New RCDInfo With {.Brand = "Schneider", .NominalCurrent = 32, .Type = "AC", .Poles = "2p", .Sensitivity = 30, .DeviceType = "EZ9 RCBO", .Breaker = True},
                            New RCDInfo With {.Brand = "Schneider", .NominalCurrent = 40, .Type = "AC", .Poles = "2p", .Sensitivity = 30, .DeviceType = "EZ9 RCBO", .Breaker = True},
                            New RCDInfo With {.Brand = "Schneider", .NominalCurrent = 25, .Type = "si", .Poles = "2p", .Sensitivity = 300, .DeviceType = "iID", .Breaker = False},
                            New RCDInfo With {.Brand = "Schneider", .NominalCurrent = 40, .Type = "si", .Poles = "2p", .Sensitivity = 300, .DeviceType = "iID", .Breaker = False},
                            New RCDInfo With {.Brand = "Schneider", .NominalCurrent = 63, .Type = "si", .Poles = "2p", .Sensitivity = 300, .DeviceType = "iID", .Breaker = False},
                            New RCDInfo With {.Brand = "Schneider", .NominalCurrent = 25, .Type = "si", .Poles = "4p", .Sensitivity = 300, .DeviceType = "iID", .Breaker = False},
                            New RCDInfo With {.Brand = "Schneider", .NominalCurrent = 40, .Type = "si", .Poles = "4p", .Sensitivity = 300, .DeviceType = "iID", .Breaker = False},
                            New RCDInfo With {.Brand = "Schneider", .NominalCurrent = 63, .Type = "si", .Poles = "4p", .Sensitivity = 300, .DeviceType = "iID", .Breaker = False},
                            New RCDInfo With {.Brand = "Schneider", .NominalCurrent = 25, .Type = "AC", .Poles = "4p", .Sensitivity = 30, .DeviceType = "iID", .Breaker = False},
                            New RCDInfo With {.Brand = "Schneider", .NominalCurrent = 40, .Type = "AC", .Poles = "4p", .Sensitivity = 30, .DeviceType = "iID", .Breaker = False},
                            New RCDInfo With {.Brand = "Schneider", .NominalCurrent = 80, .Type = "AC", .Poles = "4p", .Sensitivity = 30, .DeviceType = "iID", .Breaker = False},
                            New RCDInfo With {.Brand = "Schneider", .NominalCurrent = 100, .Type = "AC", .Poles = "4p", .Sensitivity = 30, .DeviceType = "iID", .Breaker = False},
                            New RCDInfo With {.Brand = "Schneider", .NominalCurrent = 25, .Type = "AC", .Poles = "2p", .Sensitivity = 30, .DeviceType = "Vigi iC60", .Breaker = True},
                            New RCDInfo With {.Brand = "Schneider", .NominalCurrent = 40, .Type = "AC", .Poles = "2p", .Sensitivity = 30, .DeviceType = "Vigi iC60", .Breaker = True},
                            New RCDInfo With {.Brand = "Schneider", .NominalCurrent = 63, .Type = "AC", .Poles = "2p", .Sensitivity = 30, .DeviceType = "Vigi iC60", .Breaker = True},
                            New RCDInfo With {.Brand = "Schneider", .NominalCurrent = 25, .Type = "AC", .Poles = "4p", .Sensitivity = 30, .DeviceType = "Vigi iC60", .Breaker = True},
                            New RCDInfo With {.Brand = "Schneider", .NominalCurrent = 40, .Type = "AC", .Poles = "4p", .Sensitivity = 30, .DeviceType = "Vigi iC60", .Breaker = True},
                            New RCDInfo With {.Brand = "Schneider", .NominalCurrent = 63, .Type = "AC", .Poles = "4p", .Sensitivity = 30, .DeviceType = "Vigi iC60", .Breaker = True}
        }
        LiMountMethod = New List(Of strMountMethod) From {
            New strMountMethod With {.Simbol = "A1", .Text = "В изолация"},
            New strMountMethod With {.Simbol = "B1", .Text = "Тръба (стена)"},
            New strMountMethod With {.Simbol = "C", .Text = "Върху стена"},
            New strMountMethod With {.Simbol = "D1", .Text = "Тръба (земя)"},
            New strMountMethod With {.Simbol = "D2", .Text = "Кабел (земя)"},
            New strMountMethod With {.Simbol = "E", .Text = "Кабелна скара"},
            New strMountMethod With {.Simbol = "F", .Text = "Многож. скара"},
            New strMountMethod With {.Simbol = "G", .Text = "Свободен въздух"}
            }
        ' 1. От Символ към Текст (за запълване на таблицата)
        ' Return mountMethod.FirstOrDefault(Function(m) m.Simbol = simbol).Text
        ' 2. От Текст към Символ (ако ти трябва да разбереш кода при избран елемент в UI)
        ' Return mountMethod.FirstOrDefault(Function(m) m.Text = Text).Simbol
        Catalog_Contactor = LoadContactorCatalog()
        ' Добавяне на записи за серия GV2-ME
        GV_Database.Add(New GV_Entry(0.1, 0.16, "GV2-ME", "<0,06kW", "0.1-0.16A"))
        GV_Database.Add(New GV_Entry(0.16, 0.25, "GV2-ME", "0,06kW", "0.16-0.25A"))
        GV_Database.Add(New GV_Entry(0.25, 0.4, "GV2-ME", "0,09kW", "0.25-0.40A"))
        GV_Database.Add(New GV_Entry(0.4, 0.63, "GV2-ME", "0,12kW", "0.4-0.63A"))
        GV_Database.Add(New GV_Entry(0.63, 1.0, "GV2-ME", "0,25kW", "0.63-1.0A"))
        GV_Database.Add(New GV_Entry(1.0, 1.6, "GV2-ME", "0,37kW", "1.0-1.6A"))
        GV_Database.Add(New GV_Entry(1.6, 2.5, "GV2-ME", "0,75kW", "1.6-2.5A"))
        GV_Database.Add(New GV_Entry(2.5, 4.0, "GV2-ME", "1,1kW", "2.5-4.0A"))
        GV_Database.Add(New GV_Entry(4.0, 6.3, "GV2-ME", "2,2kW", "4.0-6.3A"))
        GV_Database.Add(New GV_Entry(6.0, 10.0, "GV2-ME", "4,0kW", "6.0-10A"))
        GV_Database.Add(New GV_Entry(9.0, 14.0, "GV2-ME", "5,5kW", "9.0-14A"))
        GV_Database.Add(New GV_Entry(13.0, 18.0, "GV2-ME", "7,5kW", "13-18A"))
        GV_Database.Add(New GV_Entry(17.0, 23.0, "GV2-ME", "9,0kW", "17-23A"))
        ' Добавяне на записи за серия GV3-P
        GV_Database.Add(New GV_Entry(17.0, 25.0, "GV3-P", "11,0kW", "17-25A"))
        GV_Database.Add(New GV_Entry(23.0, 32.0, "GV3-P", "15,0kW", "23-32A"))
        GV_Database.Add(New GV_Entry(30.0, 40.0, "GV3-P", "18,5kW", "30-40A"))
        ' Добавяне на записи за серия GV4P
        GV_Database.Add(New GV_Entry(20.0, 50.0, "GV4P", "11-22kW", "20-50A"))
        GV_Database.Add(New GV_Entry(40.0, 80.0, "GV4P", "22-37kW", "40-80A"))
        GV_Database.Add(New GV_Entry(65.0, 115.0, "GV4P", "37-55kW", "65-115A"))
    End Sub
    ' Пример за зареждане на каталога с реални данни от Schneider Electric TeSys D
    Private Function LoadContactorCatalog() As Dictionary(Of String, ContactorEntry)
        Dim catalog As New Dictionary(Of String, ContactorEntry)
        ' LC1D09 - 9A AC-3
        catalog.Add("LC1D09", New ContactorEntry With {
        .PartNumber = "LC1D09",
        .FrameSize = "09",
        .RatedCurrent_AC1 = 25, ' AC-1: 25A при 440V [[30]]
        .RatedCurrent_AC3 = 9,  ' AC-3: 9A при 440V [[30]]
        .RatedCurrent_AC4 = 6,  ' AC-4: приблизително 2/3 от AC-3
        .MaxPower_AC3_400V = 4.0, ' 4 kW при 400V AC-3 [[24]]
        .AvailableCoils = New List(Of String) From {"24AC", "48AC", "110AC", "220AC", "230AC", "24DC", "48DC", "110DC", "220DC"},
        .HasAuxContacts = True,
        .MaxAuxContacts = 4, ' 1NO+1NC вградени + допълнителни блокове
        .CompatibleRelay = "LRD08", ' LRD08 за 2.5-4A [[52]]
        .DeratingFactor = New Dictionary(Of Double, Double) From {
            {40, 1.0},
            {50, 0.95},
            {60, 0.9}
        }
    })

        ' LC1D12 - 12A AC-3
        catalog.Add("LC1D12", New ContactorEntry With {
        .PartNumber = "LC1D12",
        .FrameSize = "12",
        .RatedCurrent_AC1 = 32, ' AC-1: 32A при 440V [[39]]
        .RatedCurrent_AC3 = 12, ' AC-3: 12A при 440V [[25]]
        .RatedCurrent_AC4 = 8,  ' AC-4: приблизително
        .MaxPower_AC3_400V = 5.5, ' 5.5 kW при 400V AC-3 [[25]]
        .AvailableCoils = New List(Of String) From {"24AC", "48AC", "110AC", "220AC", "230AC", "24DC", "48DC", "110DC", "220DC"},
        .HasAuxContacts = True,
        .MaxAuxContacts = 4, ' 1NO+1NC вградени + допълнителни
        .CompatibleRelay = "LRD10", ' LRD10 за 4-6A [[52]]
        .DeratingFactor = New Dictionary(Of Double, Double) From {
            {40, 1.0},
            {50, 0.95},
            {60, 0.9}
        }
    })

        ' LC1D18 - 18A AC-3
        catalog.Add("LC1D18", New ContactorEntry With {
        .PartNumber = "LC1D18",
        .FrameSize = "18",
        .RatedCurrent_AC1 = 40, ' AC-1: 40A при 440V
        .RatedCurrent_AC3 = 18, ' AC-3: 18A при 440V [[43]]
        .RatedCurrent_AC4 = 12, ' AC-4: приблизително
        .MaxPower_AC3_400V = 7.5, ' 7.5 kW при 400V AC-3
        .AvailableCoils = New List(Of String) From {"24AC", "110AC", "220AC", "230AC", "24DC", "110DC", "220DC"},
        .HasAuxContacts = True,
        .MaxAuxContacts = 4,
        .CompatibleRelay = "LRD14", ' LRD14 за 7-10A [[52]]
        .DeratingFactor = New Dictionary(Of Double, Double) From {
            {40, 1.0},
            {50, 0.95},
            {60, 0.9}
        }
    })

        ' LC1D25 - 25A AC-3
        catalog.Add("LC1D25", New ContactorEntry With {
        .PartNumber = "LC1D25",
        .FrameSize = "25",
        .RatedCurrent_AC1 = 40, ' AC-1: 40A при 440V [[24]]
        .RatedCurrent_AC3 = 25, ' AC-3: 25A при 440V [[24]]
        .RatedCurrent_AC4 = 16, ' AC-4: приблизително
        .MaxPower_AC3_400V = 11.0, ' 11 kW при 400V AC-3 [[24]]
        .AvailableCoils = New List(Of String) From {"24AC", "48AC", "110AC", "220AC", "230AC", "400AC", "24DC", "48DC", "110DC", "220DC"},
        .HasAuxContacts = True,
        .MaxAuxContacts = 6, ' 1NO+1NC вградени + повече допълнителни
        .CompatibleRelay = "LRD16", ' LRD16 за 9-13A [[53]]
        .DeratingFactor = New Dictionary(Of Double, Double) From {
            {40, 1.0},
            {50, 0.95},
            {60, 0.9}
        }
    })

        ' LC1D32 - 32A AC-3
        catalog.Add("LC1D32", New ContactorEntry With {
        .PartNumber = "LC1D32",
        .FrameSize = "32",
        .RatedCurrent_AC1 = 50, ' AC-1: 50A
        .RatedCurrent_AC3 = 32, ' AC-3: 32A при 440V [[30]]
        .RatedCurrent_AC4 = 20, ' AC-4: приблизително
        .MaxPower_AC3_400V = 15.0, ' 15 kW при 400V AC-3
        .AvailableCoils = New List(Of String) From {"24AC", "110AC", "220AC", "230AC", "24DC", "110DC", "220DC"},
        .HasAuxContacts = True,
        .MaxAuxContacts = 6,
        .CompatibleRelay = "LRD22", ' LRD22 за 16-24A [[57]]
        .DeratingFactor = New Dictionary(Of Double, Double) From {
            {40, 1.0},
            {50, 0.95},
            {60, 0.9}
        }
    })

        ' LC1D38 - 38A AC-3
        catalog.Add("LC1D38", New ContactorEntry With {
        .PartNumber = "LC1D38",
        .FrameSize = "38",
        .RatedCurrent_AC1 = 60, ' AC-1: 60A
        .RatedCurrent_AC3 = 38, ' AC-3: 38A при 440V
        .RatedCurrent_AC4 = 25, ' AC-4: приблизително
        .MaxPower_AC3_400V = 18.5, ' 18.5 kW при 400V AC-3
        .AvailableCoils = New List(Of String) From {"24AC", "110AC", "220AC", "230AC", "24DC", "110DC", "220DC"},
        .HasAuxContacts = True,
        .MaxAuxContacts = 6,
        .CompatibleRelay = "LRD32", ' LRD32 за 23-32A [[53]]
        .DeratingFactor = New Dictionary(Of Double, Double) From {
            {40, 1.0},
            {50, 0.95},
            {60, 0.9}
        }
    })

        ' LC1D50 - 50A AC-3
        catalog.Add("LC1D50", New ContactorEntry With {
        .PartNumber = "LC1D50",
        .FrameSize = "50",
        .RatedCurrent_AC1 = 80, ' AC-1: 80A
        .RatedCurrent_AC3 = 50, ' AC-3: 50A при 440V [[47]]
        .RatedCurrent_AC4 = 33, ' AC-4: приблизително
        .MaxPower_AC3_400V = 22.0, ' 22 kW при 400V AC-3 [[47]]
        .AvailableCoils = New List(Of String) From {"24AC", "110AC", "220AC", "230AC", "400AC", "24DC", "110DC", "220DC"},
        .HasAuxContacts = True,
        .MaxAuxContacts = 8,
        .CompatibleRelay = "LRD35", ' LRD35 за 30-38A
        .DeratingFactor = New Dictionary(Of Double, Double) From {
            {40, 1.0},
            {50, 0.95},
            {60, 0.9}
        }
    })

        Return catalog
    End Function
    Private Sub FillCables()
        Catalog_Cables.Clear()


        Catalog_Cables.Add(New CableInfo With {.CableType = "СВТ", .Material = "Cu", .PhaseSize = "1,5", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 19, .MaxCurrent_Ground = 25, .NeutralSize = "1,5"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "СВТ", .Material = "Cu", .PhaseSize = "2,5", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 25, .MaxCurrent_Ground = 34, .NeutralSize = "2,5"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "СВТ", .Material = "Cu", .PhaseSize = "4", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 34, .MaxCurrent_Ground = 45, .NeutralSize = "4"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "СВТ", .Material = "Cu", .PhaseSize = "6", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 43, .MaxCurrent_Ground = 55, .NeutralSize = "6"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "СВТ", .Material = "Cu", .PhaseSize = "10", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 59, .MaxCurrent_Ground = 76, .NeutralSize = "10"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "СВТ", .Material = "Cu", .PhaseSize = "16", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 79, .MaxCurrent_Ground = 96, .NeutralSize = "16"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "СВТ", .Material = "Cu", .PhaseSize = "25", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 105, .MaxCurrent_Ground = 126, .NeutralSize = "16"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "СВТ", .Material = "Cu", .PhaseSize = "35", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 126, .MaxCurrent_Ground = 151, .NeutralSize = "16"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "СВТ", .Material = "Cu", .PhaseSize = "50", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 157, .MaxCurrent_Ground = 178, .NeutralSize = "25"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "СВТ", .Material = "Cu", .PhaseSize = "70", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 199, .MaxCurrent_Ground = 225, .NeutralSize = "35"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "СВТ", .Material = "Cu", .PhaseSize = "95", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 246, .MaxCurrent_Ground = 270, .NeutralSize = "50"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "СВТ", .Material = "Cu", .PhaseSize = "120", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 285, .MaxCurrent_Ground = 306, .NeutralSize = "70"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "СВТ", .Material = "Cu", .PhaseSize = "150", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 326, .MaxCurrent_Ground = 346, .NeutralSize = "70"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "СВТ", .Material = "Cu", .PhaseSize = "185", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 374, .MaxCurrent_Ground = 390, .NeutralSize = "95"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "СВТ", .Material = "Cu", .PhaseSize = "240", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 445, .MaxCurrent_Ground = 458, .NeutralSize = "120"})


        Catalog_Cables.Add(New CableInfo With {.CableType = "САВТ", .Material = "Al", .PhaseSize = "1,5", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 0, .MaxCurrent_Ground = 0, .NeutralSize = "1,5"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "САВТ", .Material = "Al", .PhaseSize = "2,5", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 20, .MaxCurrent_Ground = 25, .NeutralSize = "2,5"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "САВТ", .Material = "Al", .PhaseSize = "4", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 26, .MaxCurrent_Ground = 32, .NeutralSize = "4"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "САВТ", .Material = "Al", .PhaseSize = "6", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 34, .MaxCurrent_Ground = 42, .NeutralSize = "6"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "САВТ", .Material = "Al", .PhaseSize = "10", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 43, .MaxCurrent_Ground = 53, .NeutralSize = "10"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "САВТ", .Material = "Al", .PhaseSize = "16", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 64, .MaxCurrent_Ground = 75, .NeutralSize = "16"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "САВТ", .Material = "Al", .PhaseSize = "25", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 82, .MaxCurrent_Ground = 92, .NeutralSize = "16"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "САВТ", .Material = "Al", .PhaseSize = "35", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 100, .MaxCurrent_Ground = 110, .NeutralSize = "16"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "САВТ", .Material = "Al", .PhaseSize = "50", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 119, .MaxCurrent_Ground = 134, .NeutralSize = "25"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "САВТ", .Material = "Al", .PhaseSize = "70", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 152, .MaxCurrent_Ground = 170, .NeutralSize = "35"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "САВТ", .Material = "Al", .PhaseSize = "95", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 185, .MaxCurrent_Ground = 210, .NeutralSize = "50"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "САВТ", .Material = "Al", .PhaseSize = "120", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 215, .MaxCurrent_Ground = 245, .NeutralSize = "70"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "САВТ", .Material = "Al", .PhaseSize = "150", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 245, .MaxCurrent_Ground = 274, .NeutralSize = "70"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "САВТ", .Material = "Al", .PhaseSize = "185", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 285, .MaxCurrent_Ground = 310, .NeutralSize = "95"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "САВТ", .Material = "Al", .PhaseSize = "240", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 338, .MaxCurrent_Ground = 360, .NeutralSize = "120"})


        Catalog_Cables.Add(New CableInfo With {.CableType = "Al/R", .Material = "Al", .PhaseSize = "16", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 83, .MaxCurrent_Ground = 0, .NeutralSize = "16"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "Al/R", .Material = "Al", .PhaseSize = "25", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 111, .MaxCurrent_Ground = 0, .NeutralSize = "25"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "Al/R", .Material = "Al", .PhaseSize = "35", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 138, .MaxCurrent_Ground = 0, .NeutralSize = "35"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "Al/R", .Material = "Al", .PhaseSize = "35", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 138, .MaxCurrent_Ground = 0, .NeutralSize = "54"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "Al/R", .Material = "Al", .PhaseSize = "50", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 164, .MaxCurrent_Ground = 0, .NeutralSize = "54"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "Al/R", .Material = "Al", .PhaseSize = "70", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 213, .MaxCurrent_Ground = 0, .NeutralSize = "54"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "Al/R", .Material = "Al", .PhaseSize = "95", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 258, .MaxCurrent_Ground = 0, .NeutralSize = "70"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "Al/R", .Material = "Al", .PhaseSize = "150", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 344, .MaxCurrent_Ground = 0, .NeutralSize = "70"})


        Catalog_Cables.Add(New CableInfo With {.CableType = "ПВ-А1", .Material = "Cu", .PhaseSize = "1,5", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 20, .MaxCurrent_Ground = 29, .NeutralSize = "0"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "ПВ-А1", .Material = "Cu", .PhaseSize = "2,5", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 27, .MaxCurrent_Ground = 38, .NeutralSize = "0"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "ПВ-А1", .Material = "Cu", .PhaseSize = "4", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 36, .MaxCurrent_Ground = 49, .NeutralSize = "0"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "ПВ-А1", .Material = "Cu", .PhaseSize = "6", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 45, .MaxCurrent_Ground = 62, .NeutralSize = "0"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "ПВ-А1", .Material = "Cu", .PhaseSize = "10", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 63, .MaxCurrent_Ground = 83, .NeutralSize = "0"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "ПВ-А1", .Material = "Cu", .PhaseSize = "16", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 82, .MaxCurrent_Ground = 104, .NeutralSize = "0"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "ПВ-А1", .Material = "Cu", .PhaseSize = "25", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 113, .MaxCurrent_Ground = 136, .NeutralSize = "0"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "ПВ-А1", .Material = "Cu", .PhaseSize = "35", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 138, .MaxCurrent_Ground = 162, .NeutralSize = "0"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "ПВ-А1", .Material = "Cu", .PhaseSize = "50", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 168, .MaxCurrent_Ground = 192, .NeutralSize = "0"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "ПВ-А1", .Material = "Cu", .PhaseSize = "70", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 210, .MaxCurrent_Ground = 236, .NeutralSize = "0"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "ПВ-А1", .Material = "Cu", .PhaseSize = "95", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 262, .MaxCurrent_Ground = 285, .NeutralSize = "0"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "ПВ-А1", .Material = "Cu", .PhaseSize = "120", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 307, .MaxCurrent_Ground = 322, .NeutralSize = "0"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "ПВ-А1", .Material = "Cu", .PhaseSize = "150", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 352, .MaxCurrent_Ground = 363, .NeutralSize = "0"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "ПВ-А1", .Material = "Cu", .PhaseSize = "185", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 405, .MaxCurrent_Ground = 410, .NeutralSize = "0"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "ПВ-А1", .Material = "Cu", .PhaseSize = "240", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 482, .MaxCurrent_Ground = 475, .NeutralSize = "0"})


        Catalog_Cables.Add(New CableInfo With {.CableType = "NYY", .Material = "Cu", .PhaseSize = "1,5", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 19.5, .MaxCurrent_Ground = 27, .NeutralSize = "1,5"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NYY", .Material = "Cu", .PhaseSize = "2,5", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 25, .MaxCurrent_Ground = 36, .NeutralSize = "2,5"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NYY", .Material = "Cu", .PhaseSize = "4", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 34, .MaxCurrent_Ground = 47, .NeutralSize = "4"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NYY", .Material = "Cu", .PhaseSize = "6", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 43, .MaxCurrent_Ground = 59, .NeutralSize = "6"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NYY", .Material = "Cu", .PhaseSize = "10", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 59, .MaxCurrent_Ground = 79, .NeutralSize = "10"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NYY", .Material = "Cu", .PhaseSize = "16", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 79, .MaxCurrent_Ground = 102, .NeutralSize = "16"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NYY", .Material = "Cu", .PhaseSize = "25", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 106, .MaxCurrent_Ground = 133, .NeutralSize = "16"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NYY", .Material = "Cu", .PhaseSize = "35", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 129, .MaxCurrent_Ground = 159, .NeutralSize = "16"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NYY", .Material = "Cu", .PhaseSize = "50", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 157, .MaxCurrent_Ground = 188, .NeutralSize = "25"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NYY", .Material = "Cu", .PhaseSize = "70", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 199, .MaxCurrent_Ground = 232, .NeutralSize = "35"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NYY", .Material = "Cu", .PhaseSize = "95", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 246, .MaxCurrent_Ground = 280, .NeutralSize = "50"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NYY", .Material = "Cu", .PhaseSize = "120", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 285, .MaxCurrent_Ground = 318, .NeutralSize = "70"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NYY", .Material = "Cu", .PhaseSize = "150", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 326, .MaxCurrent_Ground = 359, .NeutralSize = "70"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NYY", .Material = "Cu", .PhaseSize = "185", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 374, .MaxCurrent_Ground = 406, .NeutralSize = "95"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NYY", .Material = "Cu", .PhaseSize = "240", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 445, .MaxCurrent_Ground = 473, .NeutralSize = "120"})


        Catalog_Cables.Add(New CableInfo With {.CableType = "NAYY", .Material = "Al", .PhaseSize = "1,5", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 0, .MaxCurrent_Ground = 0, .NeutralSize = "1,5"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NAYY", .Material = "Al", .PhaseSize = "2,5", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 0, .MaxCurrent_Ground = 0, .NeutralSize = "2,5"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NAYY", .Material = "Al", .PhaseSize = "4", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 0, .MaxCurrent_Ground = 0, .NeutralSize = "4"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NAYY", .Material = "Al", .PhaseSize = "6", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 0, .MaxCurrent_Ground = 0, .NeutralSize = "6"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NAYY", .Material = "Al", .PhaseSize = "10", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 0, .MaxCurrent_Ground = 0, .NeutralSize = "10"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NAYY", .Material = "Al", .PhaseSize = "16", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 0, .MaxCurrent_Ground = 0, .NeutralSize = "16"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NAYY", .Material = "Al", .PhaseSize = "25", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 82, .MaxCurrent_Ground = 102, .NeutralSize = "16"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NAYY", .Material = "Al", .PhaseSize = "35", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 100, .MaxCurrent_Ground = 123, .NeutralSize = "16"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NAYY", .Material = "Al", .PhaseSize = "50", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 119, .MaxCurrent_Ground = 144, .NeutralSize = "25"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NAYY", .Material = "Al", .PhaseSize = "70", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 152, .MaxCurrent_Ground = 179, .NeutralSize = "35"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NAYY", .Material = "Al", .PhaseSize = "95", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 186, .MaxCurrent_Ground = 215, .NeutralSize = "50"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NAYY", .Material = "Al", .PhaseSize = "120", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 216, .MaxCurrent_Ground = 245, .NeutralSize = "70"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NAYY", .Material = "Al", .PhaseSize = "150", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 246, .MaxCurrent_Ground = 275, .NeutralSize = "70"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NAYY", .Material = "Al", .PhaseSize = "185", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 285, .MaxCurrent_Ground = 313, .NeutralSize = "95"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NAYY", .Material = "Al", .PhaseSize = "240", .MaxWorkingTemp = 70, .InsulationType = "PVC", .MaxCurrent_Air = 338, .MaxCurrent_Ground = 364, .NeutralSize = "120"})


        Catalog_Cables.Add(New CableInfo With {.CableType = "N2XY", .Material = "Cu", .PhaseSize = "1,5", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 24, .MaxCurrent_Ground = 31, .NeutralSize = "1,5"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "N2XY", .Material = "Cu", .PhaseSize = "2,5", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 32, .MaxCurrent_Ground = 40, .NeutralSize = "2,5"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "N2XY", .Material = "Cu", .PhaseSize = "4", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 42, .MaxCurrent_Ground = 52, .NeutralSize = "4"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "N2XY", .Material = "Cu", .PhaseSize = "6", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 53, .MaxCurrent_Ground = 64, .NeutralSize = "6"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "N2XY", .Material = "Cu", .PhaseSize = "10", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 74, .MaxCurrent_Ground = 86, .NeutralSize = "10"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "N2XY", .Material = "Cu", .PhaseSize = "16", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 98, .MaxCurrent_Ground = 112, .NeutralSize = "16"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "N2XY", .Material = "Cu", .PhaseSize = "25", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 133, .MaxCurrent_Ground = 145, .NeutralSize = "16"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "N2XY", .Material = "Cu", .PhaseSize = "35", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 162, .MaxCurrent_Ground = 174, .NeutralSize = "16"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "N2XY", .Material = "Cu", .PhaseSize = "50", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 197, .MaxCurrent_Ground = 206, .NeutralSize = "25"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "N2XY", .Material = "Cu", .PhaseSize = "70", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 250, .MaxCurrent_Ground = 254, .NeutralSize = "35"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "N2XY", .Material = "Cu", .PhaseSize = "95", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 308, .MaxCurrent_Ground = 305, .NeutralSize = "50"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "N2XY", .Material = "Cu", .PhaseSize = "120", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 359, .MaxCurrent_Ground = 348, .NeutralSize = "70"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "N2XY", .Material = "Cu", .PhaseSize = "150", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 412, .MaxCurrent_Ground = 392, .NeutralSize = "70"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "N2XY", .Material = "Cu", .PhaseSize = "185", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 475, .MaxCurrent_Ground = 444, .NeutralSize = "95"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "N2XY", .Material = "Cu", .PhaseSize = "240", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 564, .MaxCurrent_Ground = 517, .NeutralSize = "120"})


        Catalog_Cables.Add(New CableInfo With {.CableType = "NA2XY", .Material = "Al", .PhaseSize = "1,5", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 0, .MaxCurrent_Ground = 0, .NeutralSize = "1,5"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NA2XY", .Material = "Al", .PhaseSize = "2,5", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 0, .MaxCurrent_Ground = 0, .NeutralSize = "2,5"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NA2XY", .Material = "Al", .PhaseSize = "4", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 0, .MaxCurrent_Ground = 0, .NeutralSize = "4"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NA2XY", .Material = "Al", .PhaseSize = "6", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 0, .MaxCurrent_Ground = 0, .NeutralSize = "6"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NA2XY", .Material = "Al", .PhaseSize = "10", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 0, .MaxCurrent_Ground = 0, .NeutralSize = "10"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NA2XY", .Material = "Al", .PhaseSize = "16", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 0, .MaxCurrent_Ground = 0, .NeutralSize = "16"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NA2XY", .Material = "Al", .PhaseSize = "25", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 102, .MaxCurrent_Ground = 112, .NeutralSize = "16"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NA2XY", .Material = "Al", .PhaseSize = "35", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 126, .MaxCurrent_Ground = 135, .NeutralSize = "16"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NA2XY", .Material = "Al", .PhaseSize = "50", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 149, .MaxCurrent_Ground = 158, .NeutralSize = "25"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NA2XY", .Material = "Al", .PhaseSize = "70", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 191, .MaxCurrent_Ground = 196, .NeutralSize = "35"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NA2XY", .Material = "Al", .PhaseSize = "95", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 234, .MaxCurrent_Ground = 234, .NeutralSize = "50"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NA2XY", .Material = "Al", .PhaseSize = "120", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 273, .MaxCurrent_Ground = 268, .NeutralSize = "70"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NA2XY", .Material = "Al", .PhaseSize = "150", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 311, .MaxCurrent_Ground = 300, .NeutralSize = "70"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NA2XY", .Material = "Al", .PhaseSize = "185", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 360, .MaxCurrent_Ground = 342, .NeutralSize = "95"})
        Catalog_Cables.Add(New CableInfo With {.CableType = "NA2XY", .Material = "Al", .PhaseSize = "240", .MaxWorkingTemp = 90, .InsulationType = "XLPE", .MaxCurrent_Air = 427, .MaxCurrent_Ground = 398, .NeutralSize = "120"})

        Cable_For_combo = Catalog_Cables.Select(Function(b) b.CableType).Distinct().ToList()
    End Sub
    ''' <summary>
    ''' Процедура за добавяне на всички прекъсвачи.
    ''' Тук се генерират MCB, MCCB и ACB.
    ''' </summary>
    Private Sub FillBreakers()
        Breakers.Clear()
        ' ==========================
        ' MCB – EZ9
        ' ==========================
        Dim EZ9_Currents = {6, 10, 16, 20, 25, 32, 40, 50, 63}
        Dim EZ9_Curves = {"C", "B", "D"}
        Dim EZ9_Poles = {1, 3}
        For Each Inom In EZ9_Currents
            For Each curve In EZ9_Curves
                For Each poles In EZ9_Poles
                    Breakers.Add(New BreakerInfo With {
                    .Brand = "Schneider",
                    .Series = "EZ9 MCB",
                    .Category = "MCB",
                    .NominalCurrent = Inom,
                    .Poles = poles,
                    .Curve = curve,
                    .Ics_kA = 6,
                    .TripUnit = Nothing
                })
                Next
            Next
        Next
        ' ==========================
        ' MCB – Acti9 iC60N (6kA / 10kA)
        ' ==========================
        ' iC60N предлага изключително малки токове за защита на контролни вериги
        Dim iC60_Currents = {2, 3, 4, 6, 10, 16, 20, 25, 32, 40, 50, 63}
        Dim iC60_Curves = {"C", "B", "D"}
        Dim iC60_Poles = {1, 3}
        For Each Inom In iC60_Currents
            For Each curve In iC60_Curves
                For Each poles In iC60_Poles
                    Breakers.Add(New BreakerInfo With {
                .Brand = "Schneider",
                .Series = "iC60N",
                .Category = "MCB",
                .NominalCurrent = Inom,
                .Poles = poles,
                .Curve = curve,
                .Ics_kA = 6,
                .TripUnit = Nothing
            })
                Next
            Next
        Next
        ' ==========================
        ' MCB – C120
        ' ==========================
        Dim C120_Currents = {80, 100, 125}
        Dim C120_Curves = {"C", "D"}
        Dim C120_Poles = {1, 3}
        For Each Inom In C120_Currents
            For Each curve In C120_Curves
                For Each poles In C120_Poles
                    Breakers.Add(New BreakerInfo With {
                    .Brand = "Schneider",
                    .Series = "C120",
                    .Category = "MCB",
                    .NominalCurrent = Inom,
                    .Poles = poles,
                    .Curve = curve,
                    .Ics_kA = 10,
                    .TripUnit = Nothing
                })
                Next
            Next
        Next
        ' ==========================
        ' MCCB – ComPacT NSXm (16A до 160A)
        ' ==========================
        Dim NSXm_Currents = {16, 25, 32, 40, 50, 63, 80, 100, 125, 160}
        Dim NSXm_Curves = {"E", "B", "F", "N", "H"}
        Dim NSXm_TripUnits = {"TM-D", "TM-DC"}
        For Each Inom In NSXm_Currents
            For Each curve In NSXm_Curves
                For Each trip In NSXm_TripUnits
                    Breakers.Add(New BreakerInfo With {
                       .Brand = "Schneider",
                        .Series = "NSXm",
                        .Category = "MCCB",
                        .NominalCurrent = Inom,
                        .Poles = 3,
                        .TripUnit = trip,
                        .Ics_kA = 25,
                        .Curve = curve
                        })
                Next
            Next
        Next
        ' NSX100 – TM‑D, TM‑DC
        Dim NSX100_Currents = {16, 25, 32, 40, 63, 80, 100}
        Dim NSX100_Curves = {"B", "F", "N", "H", "S", "L"}
        Dim NSX100_TripUnits = {"TM-D", "TM-DC"}
        For Each Inom In NSX100_Currents
            For Each curve In NSX100_Curves
                For Each trip In NSX100_TripUnits
                    Breakers.Add(New BreakerInfo With {
                        .Brand = "Schneider",
                        .Series = "NSX100",
                        .Category = "MCCB",
                        .NominalCurrent = Inom,
                        .Poles = 3,
                        .TripUnit = trip,
                        .Ics_kA = 25,
                        .Curve = curve
                    })
                Next
            Next
        Next
        ' NSX160 – TM‑D, TM‑DC
        Dim NSX160_Currents = {80, 100, 125, 160}
        Dim NSX160_Curves = {"B", "F", "N", "H", "S", "L"}
        Dim NSX160_TripUnits = {"TM-D"}
        For Each Inom In NSX160_Currents
            For Each curve In NSX160_Curves
                For Each trip In NSX160_TripUnits
                    Breakers.Add(New BreakerInfo With {
                        .Brand = "Schneider",
                        .Series = "NSX160",
                        .Category = "MCCB",
                        .NominalCurrent = Inom,
                        .Poles = 3,
                        .TripUnit = trip,
                        .Ics_kA = 36,
                        .Curve = curve
                    })
                Next
            Next
        Next
        ' NSX250 – Micrologic (по‑големи токове обикновено с електронна защита)
        Dim NSX250_Currents = {125, 160, 200, 250}
        Dim NSX250_Curves = {"B", "F", "N", "H", "S", "L"}
        Dim NSX250_TripUnits = {"TM-D", "Micrologic 2.0", "Micrologic 5.0"}
        For Each Inom In NSX250_Currents
            For Each curve In NSX250_Curves
                For Each trip In NSX250_TripUnits
                    Breakers.Add(New BreakerInfo With {
                        .Brand = "Schneider",
                        .Series = "NSX250",
                        .Category = "MCCB",
                        .NominalCurrent = Inom,
                        .Poles = 3,
                        .TripUnit = trip,
                        .Ics_kA = 50,
                        .Curve = curve
                    })
                Next
            Next
        Next
        ' NSX400/NSX630 – Micrologic
        Dim NSX400_Currents = {250, 320, 400}
        Dim NSX400_Curves = {"F", "N", "H", "S", "L"}
        Dim NSX_High_TripUnits = {"Micrologic 2.3"}
        For Each Inom In NSX400_Currents
            For Each curve In NSX400_Curves
                For Each trip In NSX_High_TripUnits
                    Breakers.Add(New BreakerInfo With {
                        .Brand = "Schneider",
                        .Series = "NSX400",
                        .Category = "MCCB",
                        .NominalCurrent = Inom,
                        .Poles = 3,
                        .TripUnit = trip,
                        .Ics_kA = 70,
                        .Curve = curve
                    })
                Next
            Next
        Next
        Dim NSX630_Currents = {400, 500, 630}
        For Each Inom In NSX630_Currents
            For Each curve In NSX400_Curves
                For Each trip In NSX_High_TripUnits
                    Breakers.Add(New BreakerInfo With {
                        .Brand = "Schneider",
                        .Series = "NSX630",
                        .Category = "MCCB",
                        .NominalCurrent = Inom,
                        .Poles = 3,
                        .TripUnit = trip,
                        .Ics_kA = 100,
                        .Curve = curve
                    })
                Next
            Next
        Next
        ' ACB – MTZ
        Dim MTZ_Currents = {800, 1000, 1250, 1600, 2000, 2500, 3200, 4000, 5000, 6300}
        Dim MTZ_Icu = {42, 65, 100}
        Dim MTZ_Poles = {3, 4}
        For Each Inom In MTZ_Currents
            For Each icuValue In MTZ_Icu
                For Each poles In MTZ_Poles
                    Breakers.Add(New BreakerInfo With {
                    .Brand = "Schneider",
                    .Series = "MTZ",
                    .Category = "ACB",
                    .NominalCurrent = Inom,
                    .Poles = poles,
                    .TripUnit = "Micrologic 6.0",
                    .Ics_kA = icuValue,
                    .Curve = "MTZ"
                })
                Next
            Next
        Next
        ' ✅ Попълни ComboBox стойностите от Breakers
        Breakers_For_combo = Breakers.Select(Function(b) b.Series).Distinct().ToList()
        TripUnit_For_combo = Breakers.Select(Function(b) b.TripUnit).Distinct().ToList()
        Curve_For_combo = Breakers.Select(Function(b) b.Curve).Distinct().ToList()
    End Sub
    ''' <summary>
    ''' Обработва блока според неговото име и Visibility свойство
    ''' Определя типа и брой фази (1 или 3)
    ''' НЕ определя брой лампи/контакти и НЕ ползва мощност за фази
    ''' </summary>
    Private Sub ProcessBlockByType(ByRef Kons As strKonsumator,
                            blockName As String,
                            visibility As String)
        ' ============================================================
        ' ПРОВЕРКА ЗА NOTHING VISIBILITY
        ' ============================================================
        Dim vis As String = ""
        If visibility IsNot Nothing Then
            vis = visibility.ToUpper()
        End If
        ' Нормализирай имената (главни букви)
        Dim name As String = blockName.ToUpper()
        ' По подразбиране - 1 фаза
        Kons.Phase = 1
        ' ============================================================
        ' SELECT CASE ПО ИМЕ НА БЛОКА
        ' ============================================================
        Select Case True
        ' ============================================================
        ' 1. БЛОКОВЕ КОИТО МОГАТ ДА СА 1P ИЛИ 3P (проверка visibility)
        ' ============================================================
        ' --- БОЙЛЕРИ ---
            Case name.Contains("БОЙЛЕР")
                Select Case True
                    Case vis.Contains("ПРОТОЧЕН - 380V"), vis.Contains("ХОРИЗОНТАЛЕН - 380V"),
                     vis.Contains("ВЕРТИКАЛЕН - 380V"), vis.Contains("Изход 3p")
                        Kons.Phase = 3
                    Case Else
                        Kons.Phase = 1
                End Select
        ' --- ВЕНТИЛАЦИИ / КЛИМАТИЦИ ---
            Case name.Contains("ВЕНТИЛАЦИИ"),
                 name.Contains("ВЕНТИЛАТОР"),
                 name.Contains("КЛИМАТИК"),
                 name.Contains("КОНВЕКТОР"),
                 name.Contains("ГОРЕЛКА"),
                 name.Contains("НАГРЕВАТЕЛ"),
                 name.Contains("ЕЛ. ЛИРА")
                Select Case True
                    Case vis.Contains("ПРОЗОРЧЕН 3P"), vis.Contains("КАНАЛЕН 3P")
                        Kons.Phase = 3
                    Case Else
                        Kons.Phase = 1
                End Select
        ' --- КОНТАКТИ ---
            Case name.Contains("КОНТАКТ")
                Select Case True
                    Case vis.Contains("ТРИФАЗЕН"), vis.Contains("ТР+2МФ"), vis.Contains("3P")
                        Kons.Phase = 3
                    Case Else
                        Kons.Phase = 1
                End Select
        ' ============================================================
        ' 2. ВСИЧКИ ОСТАНАЛИ БЛОКОВЕ - ВИНАГИ 1 ФАЗА
        ' ============================================================
            Case name.Contains("LED_DENIMA"), name.Contains("LED_LENTA"),
                 name.Contains("LED_ULTRALUX"), name.Contains("LED_ЛУНА"),
                 name.Contains("АВАРИЯ"), name.Contains("БОЙЛЕРНО ТАБЛО"),
                 name.Contains("ЛАМПИ_СПАЛНЯ"), name.Contains("ЛИНИЯ МХЛ"),
                 name.Contains("ЛУМИНЕСЦЕНТНА"), name.Contains("МЕТАЛХАЛОГЕННА"),
                 name.Contains("ПЛАФОНИ"), name.Contains("АПЛИК"),
                 name.Contains("ПЕНДЕЛ"), name.Contains("ЛАМПИОН"),
                 name.Contains("НАСТОЛНА ЛАМПА"), name.Contains("ФАСАДНО"),
                 name.Contains("БАНСКИ АПЛИК"), name.Contains("ДАТЧИК"),
                 name.Contains("ФОТОДАТЧИК"), name.Contains("ПОЛИЛЕЙ"),
                 name.Contains("ПРОЖЕКТОР")
                Kons.Phase = 1  ' Всички тези са винаги 1 фаза
        End Select
    End Sub
    ''' <summary>
    ''' Създава списък от токови кръгове (ListTokow),
    ''' като групира консуматорите (ListKonsumator)
    ''' по комбинация от ТАБЛО и ТоковКръг.
    '''
    ''' Резултат:
    ''' За всяка уникална двойка (ТАБЛО + ТоковКръг)
    ''' се създава нов обект strTokow,
    ''' който съдържа всички консуматори към този кръг.
    ''' </summary>
    Private Sub CreateTokowList()
        If ListKonsumator Is Nothing Then Exit Sub
        ListTokow = ListKonsumator _
                    .Where(Function(k) Not String.IsNullOrEmpty(k.ТоковКръг)) _
                    .GroupBy(Function(k) New With {Key k.ТАБЛО, Key k.ТоковКръг}) _
                    .Select(Function(g) New strTokow With {
                            .Tablo = g.Key.ТАБЛО,
                            .ТоковКръг = g.Key.ТоковКръг,
                            .Konsumator = g.ToList()}).ToList()
    End Sub
    ''' <summary>
    ''' Извлича брой от стойност като "3x100" → 3, "4х18" → 4, "100" → 1
    ''' Поддържа както латиница (x), така и кирилица (х)
    ''' </summary>
    Private Function ExtractCountFromPower(powerStr As String) As Integer
        If String.IsNullOrEmpty(powerStr) Then Return 1
        ' Нормализирай - превърни в малки букви за по-лесно сравнение
        Dim normalized As String = powerStr.ToLower()
        ' Проверка за "x" на латиница ИЛИ "х" на кирилица
        If normalized.Contains("x") OrElse normalized.Contains("х") Then
            ' Разделяй и по двата вида "x"
            Dim separators() As Char = {"x"c, "X"c, "х"c, "Х"c}
            Dim parts() As String = powerStr.Split(separators)
            If parts.Length >= 1 Then
                Dim count As Integer
                ' Опитай да парснеш първата част като число
                If Integer.TryParse(parts(0).Trim(), count) AndAlso count > 0 Then
                    Return count  ' Напр. "3x100" → 3, "4х18" → 4
                End If
            End If
        End If
        Return 1
    End Function
    Private Sub InitializeBlockConfigs()
        BlockConfigs = New List(Of BlockConfig) From {
                New BlockConfig With {        ' LED ОСВЕТЛЕНИЕ
                    .BlockNames = New List(Of String) From {"LED_DENIMA", "LED_LENTA", "LED_ULTRALUX", "LED_ULTRALUX_100", "LED_ULTRALUX_НОВ",
                                                            "LED_ЛУНА", "ПЛАФОНИ", "МЕТАЛХАОГЕННА ЛАМПА", "ЛИНИЯ МХЛ - 220V", "ПОЛИЛЕЙ", "ПРОЖЕКТОР"},
                    .Category = "Lamp",
                    .DefaultPoles = 1,
                    .DefaultCable = "3" & ZnakX & "1,5",
                    .DefaultBreaker = "10",
                    .DefaultBreakerType = "EZ9 MCB",
                    .DefaultPrednaz = "Общо",
                    .DefaultPrednaz1 = "осветление",
                    .VisibilityRules = New List(Of VisRule)()
                },
                New BlockConfig With {        ' УЛИЧНО ОСВЕТЛЕНИЕ
                    .BlockNames = New List(Of String) From {"ULI4NO"},
                    .Category = "Lamp",
                    .DefaultPoles = 1,
                    .DefaultCable = "3" & ZnakX & "1,5",
                    .DefaultBreaker = "10",
                    .DefaultBreakerType = "EZ9 MCB",
                    .DefaultPrednaz = "Улично",
                    .DefaultPrednaz1 = "осветление",
                    .VisibilityRules = New List(Of VisRule)()
                },
                New BlockConfig With {        ' АВАРИЙНО ОСВЕТЛЕНИЕ
                    .BlockNames = New List(Of String) From {"АВАРИЯ", "АВАРИЯ_100"},
                    .Category = "Lamp",
                    .DefaultPoles = 1,
                    .DefaultCable = "3" & ZnakX & "1,5",
                    .DefaultBreakerType = "EZ9 MCB",
                    .DefaultBreaker = "10",
                    .DefaultPrednaz = "Аварийно",
                    .DefaultPrednaz1 = "осветление",
                    .VisibilityRules = New List(Of VisRule)()
                },
                New BlockConfig With {        ' БОЙЛЕРНО ТАБЛО
                    .BlockNames = New List(Of String) From {"БОЙЛЕРНО ТАБЛО"},
                    .Category = "Contact",
                    .DefaultPoles = 1,
                    .DefaultBreakerType = "EZ9 MCB",
                    .DefaultCable = "3" & ZnakX & "2,5",
                    .DefaultBreaker = "10",
                    .VisibilityRules = New List(Of VisRule) From {
                        New VisRule With {.VisibilityPattern = "КЛЮЧ И КОНТАКТ", .ContactCount = 1},
                        New VisRule With {.VisibilityPattern = "С ДВА КОНТАКТА", .ContactCount = 2},
                        New VisRule With {.VisibilityPattern = "С ДВА КЛЮЧА", .ContactCount = 2}
                    }
                },
                New BlockConfig With {        ' КОНТАКТИ
                    .BlockNames = New List(Of String) From {"КОНТАКТ"},
                    .Category = "Contact",
                    .DefaultPoles = 1,
                    .DefaultCable = "3" & ZnakX & "2,5",
                    .DefaultBreakerType = "EZ9 MCB",
                    .DefaultBreaker = "20",
                    .DefaultPrednaz = "Контакти",
                    .DefaultPrednaz1 = "",
                    .VisibilityRules = New List(Of VisRule) From {
                        New VisRule With {.VisibilityPattern = "ДВУГНЕЗДОВ", .Poles = 1, .ContactCount = 2},
                        New VisRule With {.VisibilityPattern = "ТРИГНЕЗДОВ", .Poles = 1, .ContactCount = 3},
                        New VisRule With {.VisibilityPattern = "ТРИФАЗЕН", .Poles = 3, .Cable = "5" & ZnakX & "2,5", .Phase = "L1,L2,L3"},
                        New VisRule With {.VisibilityPattern = "ТР+2МФ", .Poles = 3, .Cable = "5" & ZnakX & "2,5", .Phase = "L1,L2,L3", .ContactCount = 2},
                        New VisRule With {.VisibilityPattern = "ТВЪРДА ВРЪЗКА", .Poles = 1, .Cable = "3" & ZnakX & "4,0"},
                        New VisRule With {.VisibilityPattern = "УСИЛЕН", .Poles = 1, .Cable = "3" & ZnakX & "4,0"},
                        New VisRule With {.VisibilityPattern = "IP 54", .Poles = 1, .Cable = "3" & ZnakX & "2,5"},
                        New VisRule With {.VisibilityPattern = "МОНТАЖ В КАНАЛ", .Poles = 1, .Cable = "3" & ZnakX & "2,5"}
                    }
                },
                New BlockConfig With {        ' ВЕНТИЛАЦИИ, КЛИМАТИЦИ, КОНВЕКТОРИ
                    .BlockNames = New List(Of String) From {"ВЕНТИЛАЦИИ"},
                    .Category = "Device",
                    .DefaultPoles = 1,
                    .DefaultCable = "3" & ZnakX & "1,5",
                    .DefaultBreakerType = "EZ9 MCB",
                    .DefaultBreaker = "10",
                    .VisibilityRules = New List(Of VisRule) From {
                        New VisRule With {.VisibilityPattern = "3P", .Poles = 3, .Cable = "5x2,5", .Phase = "L1,L2,L3"},
                        New VisRule With {.VisibilityPattern = "КАНАЛЕН 3P", .Poles = 3, .Cable = "5x2,5", .Phase = "L1,L2,L3"},
                        New VisRule With {.VisibilityPattern = "ПРОЗОРЧЕН 3P", .Poles = 3, .Cable = "5x2,5", .Phase = "L1,L2,L3"}
                    }
                },
                New BlockConfig With {        ' БОЙЛЕРИ
                    .BlockNames = New List(Of String) From {"БОЙЛЕР"},
                    .Category = "Device",
                    .DefaultPoles = 1,
                    .DefaultCable = "3" & ZnakX & "2,5",
                    .DefaultBreakerType = "EZ9 MCB",
                    .DefaultBreaker = "20",
                    .VisibilityRules = New List(Of VisRule) From {
                        New VisRule With {.VisibilityPattern = "ИЗХОД 3P", .Poles = 3, .Cable = "5" & ZnakX & "2,5", .Phase = "L1,L2,L3"},
                        New VisRule With {.VisibilityPattern = "380V", .Poles = 3, .Cable = "5" & ZnakX & "2,5", .Phase = "L1,L2,L3"},
                        New VisRule With {.VisibilityPattern = "ПРОТОЧЕН", .Poles = 1, .Breaker = "20"},
                        New VisRule With {.VisibilityPattern = "СЕШОАР", .Poles = 1, .Breaker = "16"},
                        New VisRule With {.VisibilityPattern = "СЕШОАР С КОНТАКТ", .Poles = 1, .Breaker = "16"},
                        New VisRule With {.VisibilityPattern = "ИЗХОД ГАЗ", .Cable = "3" & ZnakX & "2,5", .Breaker = "6"}
                    }
                }
            }
    End Sub
    ''' <summary>
    ''' Обработва един консуматор спрямо конфигурацията му (BlockConfigs)
    ''' и прехвърля необходимата информация към съответния токов кръг.
    '''
    ''' Логика:
    ''' 1) Намира конфигурация по име на блок.
    ''' 2) Проверява дали има специфично правило според Visibility.
    ''' 3) Попълва кабел, прекъсвач, полюси, фаза и предназначение.
    ''' 4) Натрупва мощност и броячи (лампи/контакти).
    ''' </summary>
    Private Sub ProcessConsumerByConfig(kons As strKonsumator, ByRef tokow As strTokow)
        ' ------------------------------------------------------------
        ' 0) Подготвяме данните (унифицираме текста в UpperCase)
        ' ------------------------------------------------------------
        Dim blockName As String = kons.Name.ToUpper()
        Dim visibility As String = If(kons.Visibility IsNot Nothing,
                                  kons.Visibility.ToUpper(),
                                  "")
        ' ------------------------------------------------------------
        ' 1) Търсим основната конфигурация по име на блок
        '    Проверява дали blockName съдържа някое от имената
        '    в BlockNames списъка.
        ' ------------------------------------------------------------
        Dim config = BlockConfigs.FirstOrDefault(
                                 Function(c) c.BlockNames.Any(
                                 Function(n) blockName.Contains(n))
                                 )
        ' Ако няма намерена конфигурация → прекратяваме
        If config Is Nothing Then
            MsgBox("Блок '" & blockName & "' не е намерен в InitializeBlockConfigs!",
                   MsgBoxStyle.Critical)
            Return
        End If
        ' ------------------------------------------------------------
        ' 2) Проверяваме дали има специфично правило според Visibility
        ' ------------------------------------------------------------
        Dim visRule = config.VisibilityRules.FirstOrDefault(Function(r) visibility.Contains(r.VisibilityPattern))
        ' ------------------------------------------------------------
        ' 3) ПРЕХВЪРЛЯНЕ НА ДАННИ ОТ КОНФИГУРАЦИЯТА
        ' ------------------------------------------------------------
        '
        ' Кабел – ако има правило по Visibility → вземаме от него,
        ' иначе използваме Default стойност от конфигурацията
        tokow.Кабел_Сечение = If(visRule IsNot Nothing AndAlso
                                Not String.IsNullOrEmpty(visRule.Cable),
                                visRule.Cable,
                                config.DefaultCable)
        ' Тип кабел – фиксирана стойност
        tokow.Кабел_Тип = "СВТ"
        ' Номинален ток на прекъсвача
        Dim breakerVal As String = If(visRule IsNot Nothing AndAlso
                                    Not String.IsNullOrEmpty(visRule.Breaker),
                                    visRule.Breaker,
                                    config.DefaultBreaker)
        tokow.Breaker_Номинален_Ток = breakerVal
        ' Полюси – от правило или default
        tokow.Брой_Полюси = If(visRule IsNot Nothing AndAlso visRule.Poles <> 0,
                                     visRule.Poles,
                                     config.DefaultPoles)
        ' Числова стойност на полюсите (1 или 3)
        ' Тип апарат – от правило или default
        tokow.Breaker_Тип_Апарат = If(visRule IsNot Nothing AndAlso
                            Not String.IsNullOrEmpty(visRule.BreakerType),
                            visRule.BreakerType,
                            config.DefaultBreakerType)
        ' ------------------------------------------------------------
        ' ФАЗА
        ' ------------------------------------------------------------
        ' Ако е триполюсен → автоматично задаваме трите фази
        If tokow.Брой_Полюси = 3 Then
            tokow.Фаза = "L1,L2,L3"
        Else
            ' Ако не е 3P – запазваме съществуващата фаза
            ' или задаваме по подразбиране
            If String.IsNullOrEmpty(tokow.Фаза) Then tokow.Фаза = "L"
        End If
        ' ------------------------------------------------------------
        ' ПРЕДНАЗНАЧЕНИЕ (Default от глобалната Config)
        ' ------------------------------------------------------------
        tokow.Консуматор = config.DefaultPrednaz
        tokow.предназначение = config.DefaultPrednaz1
        ' ------------------------------------------------------------
        ' 4) МОЩНОСТ И БРОЯЧИ
        ' ------------------------------------------------------------
        ' Добавяме мощността (превръщаме W → kW)
        tokow.Мощност += kons.doubМОЩНОСТ / 1000.0
        ' Извличаме брой от текстовата мощност (ако има множител)
        Dim count As Integer = ExtractCountFromPower(kons.strМОЩНОСТ)
        ' Логика според категорията на конфигурацията
        Select Case config.Category
            Case "Lamp"
                ' Увеличаваме броя лампи
                tokow.brLamp += count
                tokow.Device = "Лампа"
            Case "Contact"
                ' Ако има специфично правило за брой контакти
                If visRule IsNot Nothing AndAlso
               visRule.ContactCount > 0 Then
                    tokow.brKontakt += visRule.ContactCount
                Else
                    tokow.brKontakt += count
                End If
                tokow.Device = "Контакт"
            Case "Device"
                ' За устройства – предназначението идва от консуматора
                tokow.Консуматор = kons.Pewdn
                tokow.предназначение = kons.PEWDN1
                ' ============================================================
                ' ПРОВЕРКА ЗА БОЙЛЕР - ТРЯБВА ЛИ ДЗТ ЗАЩИТА
                ' ============================================================
                Dim boilerTypes As String() = {
                                   "Хоризонтален",
                                   "Хоризонтален - 380V",
                                   "Вертикален",
                                   "Вертикален - 380V",
                                   "Проточен",
                                   "Проточен - 380V",
                                   "Бойлер кухня"
                }
                ' Проверяваме дали консуматорът е бойлер
                If boilerTypes.Contains(kons.Visibility) Then
                    tokow.ДТЗ_RCD = True
                    tokow.RCD_Автомат = True
                    tokow.Device = "Бойлер"
                End If
        End Select
    End Sub
    ''' <summary>
    ''' Изчислява електрическите параметри на всички токови кръгове в ListTokow.
    '''
    ''' Логика на работа:
    ''' 1) Уверява се, че конфигурацията на блоковете (BlockConfigs) е инициализирана.
    ''' 2) За всеки токов кръг:
    '''    - Нулира броячите и натрупаната мощност.
    '''    - Обработва всички консуматори в кръга чрез ProcessConsumerByConfig().
    '''    - Изчислява номиналния ток на кръга.
    '''    - Проверява дали конфигурираният прекъсвач е достатъчен.
    '''    - При нужда избира нов прекъсвач според тока.
    '''
    ''' Цел:
    ''' Да осигури коректно оразмеряване на защита (прекъсвач)
    ''' спрямо реално изчисленото натоварване на всеки токов кръг.
    ''' </summary>
    Private Sub CalculateCircuitLoads()
        ' 0) Проверка дали може да се преизчислява прекъсвача
        ' по принцип в тази процедура НЕ трбва да се влиза
        ' след първоначалното изчислиение!!!
        ' Ако се влезе тук след това се променят стойности зададени от потребителя.
        ' След първоначалния избор може да се призичлява прекъсвач
        ' само когато се избра ДЗТ и след това се премахва!!!
        ' ------------------------------------------------------------
        If Not calcBreaker Then Exit Sub
        ' ------------------------------------------------------------
        ' 1) Проверка дали конфигурацията е инициализирана.
        '    Изпълнява се само ако списъкът е празен или не е създаден.
        ' ------------------------------------------------------------
        If BlockConfigs Is Nothing OrElse BlockConfigs.Count = 0 Then InitializeBlockConfigs()
        ' ------------------------------------------------------------
        ' 2) Обработка на всеки токов кръг
        ' ------------------------------------------------------------
        For Each tokow As strTokow In ListTokow
            If tokow.Device = "Разединител" Then Continue For
            ' Нулиране на броячи и стойности преди ново изчисление
            tokow.brLamp = 0
            tokow.brKontakt = 0
            tokow.Мощност = 0
            tokow.Брой_Полюси = 1
            tokow.Device = ""
            ' --------------------------------------------------------
            ' 3) Обработка на всички консуматори в кръга
            ' --------------------------------------------------------
            For Each kons As strKonsumator In tokow.Konsumator
                ProcessConsumerByConfig(kons, tokow)
            Next
            Dim I_Def As Double = 0
            Double.TryParse(tokow.Breaker_Номинален_Ток, I_Def)
            ' --------------------------------------------------------
            ' 4) Изчисляване на номиналния ток на кръга
            '    calc_Inom() изчислява тока според мощността и полюсите
            ' --------------------------------------------------------
            tokow.Ток = calc_Inom(tokow.Мощност, tokow.Брой_Полюси)
            ' --------------------------------------------------------
            ' 5) Избор на прекъсвач
            ' --------------------------------------------------------
            CalculateBreaker(tokow)
            Dim I_Get As Double = 0
            Double.TryParse(tokow.Breaker_Номинален_Ток, I_Get)
            If I_Def > I_Get Then
                tokow.Breaker_Номинален_Ток = I_Def.ToString()
            Else
                tokow.Breaker_Номинален_Ток = I_Get.ToString()
            End If
            ' ----------------------------------------------------
            ' Избираме кабел според изчисления ток и брой полюси
            ' ----------------------------------------------------
            CalculateCable(tokow)
        Next
        ' ═══════════════════════════════════════════════════════════
        ' 7) ДОБАВИ ЗАПИС ЗА ОБЩОТО НА ТАБЛОТО
        ' ═══════════════════════════════════════════════════════════
        AddFeederRecords()
    End Sub
    ''' <summary>
    ''' Избира подходящ разединител (прекъсвач) според тока на токовия кръг.
    ''' Логиката:
    ''' 1. Определя минимален и максимален диапазон (с коефициенти)
    ''' 2. Търси най-малкия възможен апарат над минималния ток
    ''' 3. Записва резултата в обекта tokow
    ''' </summary>
    ''' <param name="tokow">Токов кръг</param>
    Private Sub CalculateDisconnector(tokow As strTokow)
        ' 1️ КОНСТАНТИ (КОЕФИЦИЕНТИ)
        ' Прекъсвачът трябва да е поне 15% над изчисления ток
        Const MIN_FACTOR As Double = 1.15
        ' Максимален коефициент (в момента не се използва във филтъра)
        Const MAX_FACTOR As Double = 1.25
        ' 2️⃣ ИЗЧИСЛЯВАНЕ НА ДИАПАЗОН
        ' Минимален допустим ток за избор
        Dim minRange As Double = tokow.Ток * MIN_FACTOR
        ' Максимален допустим ток (само информативно)
        Dim maxRange As Double = tokow.Ток * MAX_FACTOR
        ' ТЪРСЕНЕ НА ПОДХОДЯЩ АПАРАТ
        ' Търсим:
        ' - същия брой полюси
        ' - номинален ток ≥ минималния диапазон
        ' Взимаме най-малкия възможен (сортираме възходящо)
        Dim suitable = Disconnectors.Where(Function(d) d.Poles = tokow.Брой_Полюси AndAlso
                                                   d.NominalCurrent >= minRange).
                                                   OrderBy(Function(d) d.NominalCurrent).
                                                   FirstOrDefault()
        ' 4️ ПРОВЕРКА И ЗАПИС
        ' Проверка дали има намерен резултат (чрез свойството Type)
        If Not String.IsNullOrEmpty(suitable.Type) Then
            ' Записваме избрания номинален ток
            tokow.Breaker_Номинален_Ток = suitable.NominalCurrent
            ' Записваме типа на апарата
            tokow.Breaker_Тип_Апарат = suitable.Type
            ' При разединител няма характеристика (крива)
            tokow.Breaker_Крива = "-"
        Else
            ' Ако няма подходящ апарат → показваме съобщение за грешка
            MsgBox(String.Format("Грешка: Не е намерен прекъсвач за {0}А с {1} полюса.", tokow.Ток, tokow.Брой_Полюси))
        End If
    End Sub
    ''' <summary>
    ''' Връща информация за начин на монтаж на база подадена стойност (символ или текст).
    ''' </summary>
    ''' <param name="inputValue">
    ''' Входна стойност за търсене.
    ''' Може да бъде:
    ''' - символ (например "A1", "B2" и т.н.)
    ''' - текстово описание на начина на монтаж
    ''' </param>
    ''' <returns>
    ''' Връща съответстващата стойност:
    ''' - ако е подаден символ → връща текстовото описание
    ''' - ако е подаден текст → връща символа
    ''' - ако няма съвпадение → връща "Не е намерено"
    ''' </returns>
    ''' <remarks>
    ''' Функцията използва колекцията LiMountMethod, която съдържа обекти с поне две свойства:
    ''' - Simbol (символ на метода)
    ''' - Text (описание на метода)
    '''
    ''' Логика на работа:
    ''' 1. Търси първия елемент, при който:
    '''    - Simbol съвпада с inputValue
    '''    ИЛИ
    '''    - Text съвпада с inputValue
    '''
    ''' 2. Ако бъде намерен резултат:
    '''    - ако входът е символ → връща текст
    '''    - ако входът е текст → връща символ
    '''
    ''' 3. Ако няма намерен резултат:
    '''    - връща "Не е намерено"
    '''
    ''' Типично приложение:
    ''' - преобразуване между кодове и описания
    ''' - визуализация в UI (например ComboBox, DataGridView)
    ''' - валидиране на въведени данни
    '''
    ''' Потенциални особености:
    ''' - FirstOrDefault връща "празен" обект (Nothing за референтен тип или default за структура),
    '''   затова проверката result.Simbol IsNot Nothing се използва за валидност.
    ''' - Ако Simbol е String, проверката IsNot Nothing не гарантира, че е намерен реален запис
    '''   (възможно е да е празен низ "").
    ''' - Ако има дублиращи се записи, ще се върне първият срещнат.
    ''' - Сравнението е case-sensitive (зависи от настройките), което може да доведе до пропуснати съвпадения.
    ''' </remarks>
    Public Function GetMountMethodInfo(inputValue As String) As String
        ' Търсене на първия запис, който съвпада по символ или текст
        Dim result = LiMountMethod.FirstOrDefault(Function(m) m.Simbol = inputValue Or m.Text = inputValue)
        ' Проверка дали е намерен резултат
        If result.Simbol IsNot Nothing Then
            ' Ако входът съвпада със символ → връща текст
            ' Ако входът съвпада с текст → връща символ
            Return If(result.Simbol = inputValue, result.Text, result.Simbol)
        End If
        ' Ако няма съвпадение
        Return "Не е намерено"
    End Function
    ''' <summary>
    ''' Определя и задава подходящ прекъсвач за даден токов кръг.
    '''
    ''' Процедурата използва изчисления ток на кръга (tokow.Ток) и
    ''' броя на полюсите (tokow.Брой_Полюси), за да избере подходящ
    ''' прекъсвач от каталога Breakers чрез функцията SelectBreaker().
    '''
    ''' Логика на избор:
    ''' Изборът на серия прекъсвач се базира на диапазона на тока:
    '''
    ''' ≤ 63A   → MCB (Easy9, iC60N) – модулни автоматични прекъсвачи
    ''' ≤ 125A  → C120 – модулни прекъсвачи с по-висок номинал
    ''' ≤ 160A  → NSXm – компактни MCCB прекъсвачи
    ''' ≤ 630A  → NSX  – MCCB прекъсвачи за по-големи товари
    ''' > 630A  → MTZ  – въздушни прекъсвачи (ACB)
    '''
    ''' Ако се намери подходящ прекъсвач:
    ''' - параметрите на токовия кръг се обновяват със стойностите
    '''   от каталога BreakerInfo.
    '''
    ''' Ако не се намери:
    ''' - показва се информационно съобщение с данни за таблото,
    '''   кръга, мощността и изчисления ток.
    '''
    ''' Цел:
    ''' Автоматично оразмеряване на защитната апаратура спрямо
    ''' изчисленото натоварване на токовия кръг.
    ''' </summary>
    Private Sub CalculateBreaker(tokow As strTokow)
        ' 0) Проверка дали може да се преизчислява прекъсвача
        ' по принцип в тази процедура НЕ трбва да се влиза
        ' след първоначалното изчислиение!!!
        ' Ако се влезе тук след това се променят стойности зададени от потребителя.
        ' След първоначалния избор може да се призичлява прекъсвач
        ' само когато се избра ДЗТ и след това се премахва!!!
        ' ------------------------------------------------------------
        If Not calcBreaker Then Exit Sub
        ' ------------------------------------------------------------
        ' ------------------------------------------------------------
        ' Деклариране на променлива за намерения прекъсвач
        ' ------------------------------------------------------------
        ' Ако не се намери подходящ прекъсвач, стойността остава Nothing
        Dim breaker As BreakerInfo = Nothing
        ' ------------------------------------------------------------
        ' Избор на серия прекъсвач според изчисления ток
        ' ------------------------------------------------------------
        Select Case tokow.Device
            Case "Разединител"
            Case "Бойлер"
                If tokow.Ток > 17 Then
                    breaker = SelectBreaker(tokow.Ток, tokow.Брой_Полюси, "C")
                Else
                    breaker = SelectBreaker(17, tokow.Брой_Полюси, "C")
                End If
                ' За бойлери използваме по-строги критерии (крива C)
            Case "Контакт"
                ' За контакти също използваме крива C
                If tokow.Ток > 17 Then
                    breaker = SelectBreaker(tokow.Ток, tokow.Брой_Полюси, "C")
                Else
                    breaker = SelectBreaker(17, tokow.Брой_Полюси, "C")
                End If
            Case "Лампа"
                ' За лампи може да се използва по-лека крива (B), но за по-големи токове – C
                breaker = SelectBreaker(8.5, tokow.Брой_Полюси, "C")
            Case Else
                ' За други устройства – използваме крива C като универсална
                Select Case tokow.Ток
                    Case Is <= 63
                        ' MCB – модулни прекъсвачи (Easy9, iC60N)
                        ' Използва се характеристика C (подходяща за смесени товари)
                        breaker = SelectBreaker(tokow.Ток, tokow.Брой_Полюси, "C")
                    Case Is <= 125
                        ' C120 – модулни прекъсвачи за по-големи токове
                        ' Също използва крива C
                        breaker = SelectBreaker(tokow.Ток, tokow.Брой_Полюси, "C")
                    Case Is <= 160
                        ' NSXm – компактни MCCB прекъсвачи
                        ' Обикновено се използва характеристика N
                        breaker = SelectBreaker(tokow.Ток, tokow.Брой_Полюси, "N")
                    Case Is <= 630
                        ' NSX – MCCB прекъсвачи за индустриални табла
                        ' Характеристика N
                        breaker = SelectBreaker(tokow.Ток, tokow.Брой_Полюси, "N")
                    Case Else
                        ' MTZ – въздушни прекъсвачи (ACB)
                        ' Тук вместо крива се използва специална серия
                        breaker = SelectBreaker(tokow.Ток, tokow.Брой_Полюси, "MTZ")
                End Select
        End Select
        ' ------------------------------------------------------------
        ' Проверка дали е намерен подходящ прекъсвач
        ' ------------------------------------------------------------
        If breaker Is Nothing Then
            ' Ако не е намерен – показва се информативно съобщение
            ' с параметрите на токовия кръг
            Dim info As String =
                    $"Внимание: Не е намерен прекъсвач в {tokow.Tablo}!" & vbCrLf &
                    "Детайли:" & vbCrLf &
                    $"- Табло: {tokow.Tablo}" & vbCrLf &
                    $"- Кръг: {tokow.ТоковКръг}" & vbCrLf &
                    $"- Мощност: {tokow.Мощност} kW" & vbCrLf &
                    $"- Ток: {tokow.Ток} A"
            MsgBox(info, MsgBoxStyle.Exclamation, "Инфо за LayerPair")
        Else
            ' --------------------------------------------------------
            ' Обновяване на параметрите на токовия кръг
            ' със стойностите на избрания прекъсвач
            ' --------------------------------------------------------
            tokow.Breaker_Номинален_Ток = breaker.NominalCurrent.ToString() ' Номинален ток на прекъсвача
            tokow.Breaker_Тип_Апарат = breaker.Series                       ' Серия на прекъсвача (EZ9, C120, NSX, MTZ и др.)
            tokow.Breaker_Крива = breaker.Curve                             ' Характеристика на прекъсвача (B, C, D, N и др.)
            tokow.Breaker_Изкл_Възможност = breaker.Ics_kA & "kA"           ' Изключвателна способност (например 6kA, 10kA, 50kA)
            tokow.Брой_Полюси = breaker.Poles                               ' Брой полюси на прекъсвача
            tokow.Breaker_Защитен_блок = breaker.TripUnit                   ' Тип защитен блок (trip unit)
        End If
    End Sub
    ''' <summary>
    ''' Автоматично избира прекъсвач от каталога според тока и броя полюси
    ''' </summary>
    ''' <param name="calculatedCurrent">Изчислен ток (A)</param>
    ''' <param name="poles">Брой полюси (1 или 3)</param>
    ''' <param name="curve">Крива по подразбиране ("C")</param>
    ''' <returns>BreakerInfo или Nothing ако не е намерен</returns>
    Private Function SelectBreaker(calculatedCurrent As Double,
                               poles As Integer,
                               Optional curve As String = "C") As BreakerInfo
        ' Дефиниране на константи за диапазона (коефициенти)
        Const MIN_FACTOR As Double = 1.15 ' Прекъсвачът трябва да е поне 15% над изчисления ток
        Const MAX_FACTOR As Double = 1.25 ' Но не повече от 25% над него (примерно)
        Dim minRange As Double = calculatedCurrent * MIN_FACTOR
        Dim maxRange As Double = calculatedCurrent * MAX_FACTOR
        ' Филтрираме прекъсвачите, които попадат точно в този "прозорец"
        Dim suitableBreakers = Breakers.Where(Function(b) b.Poles = poles AndAlso
                            String.Equals(b.Curve, curve, StringComparison.OrdinalIgnoreCase) AndAlso
                            b.NominalCurrent >= minRange AndAlso
                            b.NominalCurrent <= maxRange
                            ).OrderBy(Function(b) b.NominalCurrent).ToList()
        ' Връщаме първия (най-малкия подходящ) от диапазона
        Dim selectedBreaker = suitableBreakers.FirstOrDefault()
        ' Ако не открием прекъсвач в този тесен диапазон, 
        ' можем да върнем първия по-голям (fallback), за да не остане празен резултат
        If selectedBreaker Is Nothing Then
            selectedBreaker = Breakers.Where(Function(b) b.Poles = poles AndAlso
                                b.NominalCurrent >= minRange
                                ).OrderBy(Function(b) b.NominalCurrent).FirstOrDefault()
        End If
        Return selectedBreaker
    End Function
    ''' <summary>
    ''' Изчислява номиналния ток за токов кръг
    ''' </summary>
    ''' <param name="Pkryg">Мощност в kW</param>
    ''' <param name="NumberPoles">Брой фази: "1P" или "3P"</param>
    ''' <param name="Motor">True за двигатели (cos φ = 0.85, КПД = 0.9)</param>
    ''' <returns>Номинален ток в Ampere</returns>
    Private Function calc_Inom(Pkryg As Double,                     ' мощност
                       NumberPoles As String,                       ' брой фази
                       Optional Motor As Boolean = False            ' Ако е двигател True - КПД и cos FI да са по 0,83
                       ) As Double                                  ' Изчислява номинален ток за товар
        Dim CosFI As Double                                         ' Декларира променлива за cos φ (фактор на мощността)
        Dim KPD As Double                                           ' Декларира променлива за КПД (коефициент на полезно действие)
        Const U380 As Double = 0.4                                  ' Дефинира константа за напрежение при 380V, преобразувано в kV (киловолти)
        Const U220 As Double = 0.23                                 ' Дефинира константа за напрежение при 220V, преобразувано в kV (киловолти)
        Dim Inom As Double = 0                                      ' Инициализира променлива за номиналния ток с начална стойност 0
        If Motor Then                                               ' Проверява дали токовият кръг е двигател
            CosFI = 0.85                                            ' Ако е двигател, задава фактор на мощността 0.83
            KPD = 0.9                                               ' Ако е двигател, задава КПД 0.83
        Else                                                        ' Ако токовият кръг не е двигател
            CosFI = 0.9                                             ' Задава фактор на мощността 0.9
            KPD = 1                                                 ' Задава КПД 1
        End If
        If NumberPoles = "3" Then                                   ' Проверява дали токовият кръг е трифазен (3 полюса)
            Inom = Pkryg / (U380 * Math.Sqrt(3) * CosFI * KPD)      ' Изчислява номиналния ток за трифазен кръг по формулата
        Else                                                        ' Ако токовият кръг е монофазен (2 полюса)
            Inom = Pkryg / (U220 * CosFI * KPD)                     ' Изчислява номиналния ток за монофазен кръг по формулата
        End If
        Return Inom                                                 ' Връща изчисления номинален ток
    End Function
    ''' <summary>
    ''' Събитие: TreeView1_AfterSelect
    ''' </summary>
    ''' <remarks>
    ''' Това събитие се извиква веднага след като потребителят избере (кликне) елемент в TreeView1.
    ''' 
    ''' Основната логика на процедурата е:
    ''' - Извикване на FillDataGridViewForPanel(), която актуализира съдържанието на DataGridView
    '''   според избрания панел или група в дървовидната структура.
    '''
    ''' Потенциални особености:
    ''' - Предполага се, че методът FillDataGridViewForPanel() използва текущо избрания елемент
    '''   (TreeView1.SelectedNode) за определяне на кои токови кръгове или табла да покаже данни.
    ''' - Ако FillDataGridViewForPanel() е тежка операция, често селектирането на различни елементи
    '''   може да забави интерфейса; в такива случаи може да се наложи оптимизация или асинхронно обновяване.
    ''' - Това е стандартен подход за синхронизация на TreeView с DataGridView в WinForms.
    ''' </remarks>
    Private Sub TreeView1_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeView1.AfterSelect
        ' Обновяване на DataGridView според избрания панел/група в TreeView
        FillDataGridViewForPanel()
    End Sub
    ''' <summary>
    ''' Попълва DataGridView1 с данни за избраното табло
    ''' </summary>
    Private Sub FillDataGridViewForPanel()
        ' Проверка дали има избран възел
        If TreeView1.SelectedNode Is Nothing Then
            MsgBox("Моля, изберете табло от дървото!", MsgBoxStyle.Exclamation, "Няма избор")
            Return
        End If
        ' Вземи името на избраното табло
        Dim selectedPanel As String = TreeView1.SelectedNode.Text
        ' Ако има "(", вземи само текста преди него
        If selectedPanel.Contains("(") Then
            selectedPanel = selectedPanel.Substring(0, selectedPanel.IndexOf("(")).Trim()
        End If
        ' Филтрирай токовите кръгове за това табло
        ' ListTokow вече е сортиран, така че просто вземи кръговете за това табло
        Dim panelCircuits = ListTokow.Where(Function(t) t.Tablo.ToUpper() = selectedPanel.ToUpper()).ToList()
        ' Проверка дали има кръгове
        If panelCircuits Is Nothing OrElse panelCircuits.Count = 0 Then
            MsgBox($"Няма намерени токови кръгове за табло '{selectedPanel}'",
                   MsgBoxStyle.Information, "Няма данни")
            Return
        End If
        GroupBox2.Text = $"Обработвам табло '{selectedPanel}'"
        ' 1. Добави колони за кръговете
        AddCircuitColumns(panelCircuits)
        ' 2. Попълни данните
        For Each row As DataGridViewRow In DataGridView1.Rows
            Dim paramName As String = row.Cells(0).Value.ToString()
            ' Пропусни разделителите и заглавията
            If paramName = "---------" OrElse
                paramName = "Прекъсвач" OrElse
                paramName = "ДТЗ" OrElse
                paramName = "Управление" OrElse
                paramName = "Кабел" Then
                Continue For
            End If
            ' Вместо For i As Integer = 0 To panelCircuits.Count - 1
            For Each circuit As strTokow In panelCircuits
                ' За индекса на колоната ще ти трябва брояч, ако държиш на него
                Dim i As Integer = panelCircuits.IndexOf(circuit)
                Dim colIndex As Integer = i + 2
                If colIndex < DataGridView1.Columns.Count Then
                    UpdateCircuitColumn(circuit, colIndex, "")
                End If
            Next
        Next
    End Sub
    ''' <summary>
    ''' Актуализира стойностите на конкретна колона в DataGridView за даден токов кръг.
    ''' </summary>
    ''' <param name="circuit">Обект от тип strTokow, съдържащ информация за токовия кръг.</param>
    ''' <param name="colIndex">Индекс на колоната в DataGridView, която ще се обнови.</param>
    ''' <remarks>
    ''' Тази процедура обхожда всички редове на DataGridView1 и за всеки ред:
    ''' - Чете първата клетка (Cells(0)), която съдържа името на параметъра.
    ''' - Според името на параметъра, присвоява съответното поле от strTokow
    '''   на клетката в колоната colIndex.
    '''
    ''' Категориите параметри са структурирани по вид:
    ''' 1. Прекъсвач:
    '''    - Тип на апарата, Номинален ток, Изкл. възможн., Крива, Защитен блок, Брой полюси
    ''' 2. ДТЗ / RCD:
    '''    - ДТЗ Нула, Вид на апарата, Клас на апарата, ДТЗ(RCD) Ном. ток, Чувствителност, Брой полюси
    ''' 3. Мощност:
    '''    - Брой лампи, Брой контакти, Инст. мощност (форматирано N3), Изчислен ток (форматирано N2)
    ''' 4. Кабел:
    '''    - Начин на монтаж, Начин на полагане, Паралелни кабели (фаза), Съседни кабели (група), Тип, Сечение
    ''' 5. Описание:
    '''    - Фаза, Консуматор, предназначение, Управление
    ''' 6. Флагове:
    '''    - Шина, Постави ДТЗ (RCD)
    '''
    ''' Потенциални особености и предупреждения:
    ''' - Ако някой ред няма стойност в Cells(0), ToString() може да хвърли изключение.
    ''' - strTokow се предполага като вече инициализиран и попълнен с всички необходими данни.
    ''' - Форматирането на стойности (N3, N2) гарантира четимост за мощност и ток.
    ''' - Ако DataGridView съдържа редове с имена, които не са включени в Select Case,
    '''   те ще останат непроменени.
    ''' - Полетата като ДТЗ_RCD и Шина са логически/флагови и служат за управление на допълнителни действия в интерфейса.
    ''' </remarks>
    Private Sub UpdateCircuitColumn(circuit As strTokow, colIndex As Integer, paramNameChe As String)
        If circuit Is Nothing Then Return
        For Each row As DataGridViewRow In DataGridView1.Rows
            If row.Cells(0).Value Is Nothing Then Continue For
            Dim paramName As String = row.Cells(0).Value.ToString()
            If paramName = paramNameChe Then Continue For
            Dim newValue As Object = Nothing
            Select Case paramName
            ' --- ПРЕКЪСВАЧ ---
                Case "Тип на апарата" : newValue = circuit.Breaker_Тип_Апарат
                Case "Номинален ток" : newValue = circuit.Breaker_Номинален_Ток
                Case "Изкл. възможн." : newValue = circuit.Breaker_Изкл_Възможност
                Case "Крива" : newValue = circuit.Breaker_Крива
                Case "Защитен блок" : newValue = circuit.Breaker_Защитен_блок
                Case "Брой полюси" : newValue = circuit.Брой_Полюси
            ' --- ДТЗ ---
                Case "ДТЗ Нула" : newValue = circuit.RCD_Нула
                Case "Вид на апарата" : newValue = circuit.RCD_Тип
                Case "Клас на апарата" : newValue = circuit.RCD_Клас
                Case "ДТЗ(RCD) Ном. ток" : newValue = circuit.RCD_Ток
                Case "Чувствителност" : newValue = circuit.RCD_Чувствителност
                Case "ДТЗ(RCD) полюси" : newValue = circuit.RCD_Полюси
            ' --- МОЩНОСТ ---
                Case "Брой лампи" : newValue = circuit.brLamp
                Case "Брой контакти" : newValue = circuit.brKontakt
                Case "Инст. мощност" : newValue = circuit.Мощност.ToString("N3")
                Case "Изчислен ток" : newValue = circuit.Ток.ToString("N2")
            ' --- КАБЕЛ ---
                Case "Начин на монтаж" : newValue = circuit.Кабел_Монтаж
                Case "Начин на полагане" : newValue = circuit.Кабел_Полагане
                Case "Паралелни кабели (фаза):" : newValue = circuit.Кабел_Брой_Фаза
                Case "Съседни кабели (група):" : newValue = circuit.Кабел_Брой_Група
                Case "Тип кабел" : newValue = circuit.Кабел_Тип
                Case "Сечение" : newValue = circuit.Кабел_Сечение
            ' --- ОПИСАНИЕ ---
                Case "Фаза" : newValue = circuit.Фаза
                Case "Консуматор" : newValue = circuit.Консуматор
                Case "предназначение" : newValue = circuit.предназначение
                Case "Управление" : newValue = circuit.Управление
            ' --- ФЛАГОВЕ ---
                Case "Шина" : newValue = circuit.Шина
                Case "Постави ДТЗ (RCD)" : newValue = circuit.ДТЗ_RCD
            End Select
            ' ✅ ЕДИНСТВЕНАТА ЗАПИСВАНЕ С ПРОВЕРКА ЗА РАЗЛИКА
            If newValue IsNot Nothing Then
                Dim currentVal As String = Convert.ToString(row.Cells(colIndex).Value)
                Dim newVal As String = Convert.ToString(newValue)
                If currentVal <> newVal Then
                    row.Cells(colIndex).Value = newValue
                End If
            End If
        Next
    End Sub
    ''' <summary>
    ''' Добавя колони за токовите кръгове на избраното табло
    ''' </summary>
    Private Sub AddCircuitColumns(panelCircuits As List(Of strTokow))
        ' 1. Изтрий старите колони за кръгове
        Dim columnsToRemove As New List(Of String)
        For Each col As DataGridViewColumn In DataGridView1.Columns
            If col.Name <> "colParameter" AndAlso
               col.Name <> "colUnit" AndAlso
               col.Name <> "colTotal" Then
                columnsToRemove.Add(col.Name)
            End If
        Next
        For Each colName As String In columnsToRemove
            DataGridView1.Columns.Remove(colName)
        Next
        ' 2. Добави нови колони за всеки кръг
        For i As Integer = 0 To panelCircuits.Count - 1
            Dim circuit As strTokow = panelCircuits(i)
            If circuit.ТоковКръг = "ОБЩО" Then Continue For
            Dim col As New DataGridViewTextBoxColumn()
            col.Name = If(circuit.ТоковКръг = "Разединител", "colDiscon", $"colCircuit{i}")
            col.HeaderText = circuit.ТоковКръг
            col.Width = 110
            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            col.SortMode = DataGridViewColumnSortMode.NotSortable
            col.Tag = circuit
            ' Добави колоната ПРЕДИ colTotal
            Dim totalIndex As Integer = DataGridView1.Columns.IndexOf(DataGridView1.Columns("colTotal"))
            DataGridView1.Columns.Insert(totalIndex, col)
        Next
        ' ✅ 3. ЗАДАЙ ТИПА КЛЕТКИ ЗА НОВИТЕ КОЛОНИ (използвайки rowData)
        Dim rowIndex As Integer = 0
        For Each row As DataGridViewRow In DataGridView1.Rows
            ' Вземи типа клетка от rowData
            Dim cellType As String = rowData(rowIndex)(2)
            ' За всяка нова колона (от индекс 2 до colTotal-1)
            For colIndex As Integer = 2 To DataGridView1.Columns.Count - 2
                'If circuit.ТоковКръг = "Разединител" Then Continue For
                Dim colName As String = DataGridView1.Columns(colIndex).Name
                ' Запази стойността от старата клетка (ако има)
                Dim oldValue As Object = Nothing
                If row.Cells(colIndex).Value IsNot Nothing Then
                    oldValue = row.Cells(colIndex).Value
                End If
                ' Създай нова клетка с правилния тип (същата логика като в SetupDataGridView)
                Dim cell As DataGridViewCell = Nothing
                Select Case cellType
                    Case "Combo"
                        cell = New DataGridViewComboBoxCell()
                        SetupComboBoxCell(cell,
                                          row.Cells(0).Value.ToString(),
                                          If(colName = "colDiscon", True, False)
                                          )
                    Case "Check"
                        cell = New DataGridViewCheckBoxCell()
                        cell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                    Case Else
                        cell = New DataGridViewTextBoxCell()
                        cell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                End Select
                ' Възстанови стойността
                If oldValue IsNot Nothing Then
                    cell.Value = oldValue
                End If
                ' Замени клетката
                row.Cells(colIndex) = cell
            Next
            rowIndex += 1
        Next
    End Sub
    ''' <summary>
    ''' Сортира ListTokow по специалния приоритет
    ''' </summary>
    Private Sub SortCircuits()
        ListTokow = ListTokow.OrderBy(
            Function(t) t.Tablo
        ).ThenBy(
            Function(t) GetCircuitSortKey(t.ТоковКръг)
        ).ToList()
    End Sub
    ''' <summary>
    ''' Връща ключ за сортиране на токов кръг със специален приоритет
    ''' Порядок: 1.ав. → 2.до. → 3.други букви → 4.числа → 5.само букви
    ''' </summary>
    Private Function GetCircuitSortKey(circuitName As String) As String
        If String.IsNullOrEmpty(circuitName) Then Return "ZZZZZZZZZZ"
        Dim name As String = circuitName.Trim().ToUpper()
        Dim priority As String = "9"  ' По подразбиране най-нисък приоритет
        Dim numberPart As String = ""
        Dim letterPart As String = ""
        ' ============================================================
        ' 1. ОПРЕДЕЛИ КАТЕГОРИЯТА (ПРИОРИТЕТ)
        ' ============================================================
        Select Case True
            ' 1. СЪЩЕСТВУВАЩИ
            Case name = "СЪЩ."
                priority = "0"
                numberPart = ExtractNumber(name)
                letterPart = "СЪЩ"
            ' 2. АВАРИЙНИ
            Case name.Contains("АВ")
                priority = "1"
                numberPart = ExtractNumber(name)
                letterPart = "АВ"
            ' 3. ДОПЪЛНИТЕЛНИ
            Case name.Contains("ДО")
                priority = "2"
                numberPart = ExtractNumber(name)
                letterPart = "ДО"
            ' 4. ЧИСТИ ЧИСЛА
            Case IsNumeric(name)
                priority = "3"
                numberPart = name
                letterPart = ""
            ' 5. ЧИСЛО + БУКВА (напр. 1а, 2б)
            Case HasNumberAndLetters(name) AndAlso Char.IsDigit(name(0))
                priority = "4"
                numberPart = ExtractNumber(name)
                letterPart = ExtractLetters(name)
            ' 6. ОБЩО (Провери го ПРЕДИ общия случай за букви)
            Case name = "ОБЩО"
                priority = "9"
                numberPart = ""
                letterPart = "ZZZZZ"
            ' 7. РЕЗЕРВА
            Case name = "РЕЗ."
                priority = "8"
                numberPart = ""
                letterPart = "РЕЗ"
            ' 8. ВСИЧКО ЗАПОЧВАЩО С БУКВА (Основни кръгове като А1, Б1 и т.н.)
            Case Not String.IsNullOrEmpty(name) AndAlso Char.IsLetter(name(0))
                priority = "5"
                numberPart = ExtractNumber(name)
                letterPart = name
                ' 9. ВСИЧКО ОСТАНАЛО
            Case Else
                priority = "8"
                numberPart = ""
                letterPart = name
        End Select
        ' ============================================================
        ' 2. СЪЗДАЙ КЛЮЧ ЗА СОРТИРАНЕ
        ' ============================================================
        ' Формат: Приоритет + Номер (с водещи нули) + Букви
        ' Пример: "10000000001АВ" за "1ав."
        If numberPart.Length > 0 Then
            ' Подравняване на числото с водещи нули (до 10 цифри)
            numberPart = numberPart.PadLeft(10, "0"c)
            Return priority & numberPart & letterPart
        Else
            ' Само букви - сортирай азбучно
            Return priority & "0000000000" & letterPart
        End If
    End Function
    ''' <summary>
    ''' Извлича числото от низ (напр. "1АВ" → "1", "А2Б" → "2")
    ''' </summary>
    Private Function ExtractNumber(text As String) As String
        Dim result As String = ""
        For Each c As Char In text
            If Char.IsDigit(c) Then
                result &= c
            End If
        Next
        Return result
    End Function
    ''' <summary>
    ''' Извлича буквите от низ (напр. "1АВ" → "АВ", "А2Б" → "АБ")
    ''' </summary>
    Private Function ExtractLetters(text As String) As String
        Dim result As String = ""
        For Each c As Char In text
            If Char.IsLetter(c) Then
                result &= c
            End If
        Next
        Return result
    End Function
    ''' <summary>
    ''' Проверява дали низът съдържа и букви и числа
    ''' </summary>
    Private Function HasNumberAndLetters(text As String) As Boolean
        Dim hasNumber As Boolean = False
        Dim hasLetter As Boolean = False
        For Each c As Char In text
            If Char.IsDigit(c) Then hasNumber = True
            If Char.IsLetter(c) Then hasLetter = True
        Next
        Return hasNumber AndAlso hasLetter
    End Function
    ''' <summary>
    ''' Проверява дали низът е само число
    ''' </summary>
    Private Function IsNumeric(text As String) As Boolean
        For Each c As Char In text
            If Not Char.IsDigit(c) Then Return False
        Next
        Return text.Length > 0
    End Function

    ''' <summary>
    ''' Събитие, което се изпълнява при промяна на стойност в клетка на DataGridView1.
    '''
    ''' Основна идея:
    ''' Таблицата се използва като редактор на параметри за даден токов кръг.
    ''' Първата колона съдържа името на параметъра (например "Тип на апарата",
    ''' "Номинален ток", "Шина", "ДТЗ (RCD)" и др.), а останалите колони съдържат
    ''' стойностите за конкретни кръгове или устройства.
    '''
    ''' Когато потребителят промени стойност:
    ''' 1. Определя се редът и колоната на промяната.
    ''' 2. От първата клетка на реда се взима името на параметъра.
    ''' 3. От текущата клетка се взима новата стойност.
    ''' 4. Чрез Select Case се определя какво действие трябва да се изпълни
    '''    според типа на параметъра.
    '''
    ''' Забележка:
    ''' Този метод служи като централизирана точка за обработка на всички
    ''' промени в таблицата. Реалната логика за всяка настройка може да се
    ''' добавя вътре в съответния Case.
    ''' </summary>
    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        If e.RowIndex < 0 OrElse e.ColumnIndex < 0 Then Return
        'If isUpdatingGrid Then Return
        Try
            'isUpdatingGrid = True
            Dim row As DataGridViewRow = DataGridView1.Rows(e.RowIndex)
            Dim col As DataGridViewColumn = DataGridView1.Columns(e.ColumnIndex)
            ' ------------------------------------------------------------
            ' 3) Името на параметъра се намира в първата колона (index 0)
            '    Например:
            '    "Тип на апарата"
            '    "Номинален ток"
            '    "Шина"
            '    "ДТЗ (RCD)"
            ' ------------------------------------------------------------
            Dim paramName As String = row.Cells(0).Value?.ToString()
            ' ------------------------------------------------------------
            ' 4) Новата стойност, въведена от потребителя в текущата клетка
            ' ------------------------------------------------------------
            Dim selectedValue As String = row.Cells(e.ColumnIndex).Value?.ToString()
            ' Първо взимаш името
            Dim currentCircuit As String = DataGridView1.Columns(e.ColumnIndex).HeaderText
            ' После го предаваш
            Dim tokow As strTokow = FindTokowByColumn(currentCircuit)
            Dim Update As Boolean = True
            If tokow.Device = "Разединител" OrElse
               tokow.Device = "Съществуващ" OrElse
               tokow.Device = "Резерва" Then Exit Sub
            If tokow IsNot Nothing AndAlso Not String.IsNullOrEmpty(selectedValue) Then
                Select Case paramName
                    Case "Тип на апарата"
                        tokow.Breaker_Тип_Апарат = selectedValue
                        Select Case tokow.Device
                            Case "Разединител"
                                Dim filteredDisco = Disconnectors.Where(Function(b) b.Type = selectedValue).ToList()
                                Dim valuesForCombo = filteredDisco _
                                                    .Select(Function(b) b.NominalCurrent.ToString()) _
                                                    .Distinct() _
                                                    .ToList()
                                UpdateComboRow("Номинален ток", valuesForCombo, e.ColumnIndex)
                            Case "Табло"
                                Dim filteredDisco = Disconnectors.Where(Function(b) b.Type = selectedValue).ToList()
                                Dim valuesForCombo = filteredDisco _
                                                    .Select(Function(b) b.NominalCurrent.ToString()) _
                                                    .Distinct() _
                                                    .ToList()
                                UpdateComboRow("Номинален ток", valuesForCombo, e.ColumnIndex)
                                tokow.Device = tokow.Device '"Табло"
                            Case Else
                                Dim filteredBreakers = Breakers.Where(Function(b) b.Series = selectedValue).ToList()
                                If filteredBreakers.Count = 0 Then Exit Select
                                tokow.Breaker_Изкл_Възможност = filteredBreakers.First().Ics_kA & "kA"
                                Dim valuesForCombo = filteredBreakers _
                                                    .Select(Function(b) b.NominalCurrent.ToString()) _
                                                    .Distinct() _
                                                    .ToList()
                                UpdateComboRow("Номинален ток", valuesForCombo, e.ColumnIndex)
                                Dim valuesCurve = filteredBreakers _
                                                    .Select(Function(b) b.Curve.ToString()) _
                                                    .Distinct() _
                                                    .ToList()
                                UpdateComboRow("Крива", valuesCurve, e.ColumnIndex)
                                Dim valuesTripUnit = filteredBreakers _
                                                    .Select(Function(b) b.TripUnit) _
                                                    .Distinct() _
                                                    .ToList()
                                UpdateComboRow("Защитен блок", valuesTripUnit, e.ColumnIndex)
                        End Select
                    Case "Постави ДТЗ (RCD)"
                        ' ✅ Първо обнови tokow от клетката!
                        tokow.ДТЗ_RCD = CBool(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value)
                        HandleRCDCheckboxChange(tokow)
                    Case "Номинален ток"
                        ' Тук може да се обработва промяна на номиналния ток
                        ' на защитния апарат (например 10A, 16A, 20A...)
                        ' 1. Първо излизаме, ако няма стойност
                        If selectedValue Is Nothing Then Exit Sub
                        ' 2. Вече сме сигурни, че имаме нещо, и правим сравнението
                        If Val(selectedValue) >= Val(tokow.Breaker_Номинален_Ток) Then
                            ' Всичко е точно, обновяваме стойността
                            tokow.Breaker_Номинален_Ток = selectedValue
                        Else
                            ' Тук се намесваме с малко "приятелски" съвет
                            Dim message As String = "Сигурен ли си в това, което правиш? " & vbCrLf &
                                   "Избраният ток е по-малък от текущия." & vbCrLf &
                                   "Честно казано, правиш простотия!" & vbCrLf &
                                   "Искаш ли наистина да продължиш към Тъмната страна?"
                            Dim result As DialogResult = MessageBox.Show(message, "Внимание: Инженерна мисъл в действие!",
                                                       MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                            If result = DialogResult.Yes Then
                                ' Потребителят е инат, записваме го
                                tokow.Breaker_Номинален_Ток = selectedValue
                            Else
                                ' Спасихме положението!
                                MessageBox.Show("Мъдро решение! Спести си един ремонт.", "Браво!", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            End If
                        End If
                        CalculateCable(tokow,
                                       Type:=tokow.Кабел_Тип,
                                       layMethod:=If(tokow.Кабел_Полагане = "Във въздух", 0, 1),
                                       mountMethod:=GetMountMethodInfo(tokow.Кабел_Монтаж),
                                       Broj_Cable:=tokow.Кабел_Брой_Група,
                                       matType:=GetCableTypeResult(tokow.Кабел_Тип)
                                       )
                    Case "Съседни кабели (група):"
                        tokow.Кабел_Брой_Група = selectedValue
                        CalculateCable(tokow,
                                       Type:=tokow.Кабел_Тип,
                                       layMethod:=If(tokow.Кабел_Полагане = "Във въздух", 0, 1),
                                       mountMethod:=GetMountMethodInfo(tokow.Кабел_Монтаж),
                                       Broj_Cable:=tokow.Кабел_Брой_Група,
                                       matType:=GetCableTypeResult(tokow.Кабел_Тип)
                                       )
                    Case "Консуматор"
                        If tokow.Device <> "Табло" Then tokow.Консуматор = selectedValue
                    Case "предназначение"
                        tokow.предназначение = selectedValue
                        Update = False
                        If tokow.Device = "Табло" Then
                            Update = True
                            ' Търсим число само след "Рпр" или "Рпр."
                            Dim pattern As String = "Рпр\.?\s*=\s*(\d+([.,]\d+)?)"
                            Dim match = System.Text.RegularExpressions.Regex.Match(tokow.предназначение, pattern)
                            Dim value As Double = -1 ' -1 = няма валидна стойност
                            ' Ако regex намери число
                            If match.Success Then
                                Dim strValue As String = match.Groups(1).Value.Replace(",", ".")
                                Double.TryParse(strValue, System.Globalization.NumberStyles.Any,
                                System.Globalization.CultureInfo.InvariantCulture,
                                value)
                            End If
                            ' Ако няма валидно число, проверяваме дали полето е просто число
                            If value < 0 Then
                                Dim onlyNumber As Double = 0
                                If Double.TryParse(tokow.предназначение.Replace(",", "."),
                                   System.Globalization.NumberStyles.Any,
                                   System.Globalization.CultureInfo.InvariantCulture,
                                   onlyNumber) Then
                                    value = onlyNumber
                                End If
                            End If
                            ' Ако има валидно число, записваме в предназначение във формат Рпр.=(число)кW
                            If value > 0 Then
                                tokow.предназначение = "Рпр.=" & value.ToString("0.##") & "кW"
                            Else
                                ' Ако няма валидно число, задаваме по подразбиране
                                tokow.предназначение = "Рпр.=15кW"
                                value = 15
                            End If
                            ' Проверка да не делим на 0
                            If tokow.Мощност <> 0 Then
                                tokow.Консуматор = "Ке=" & (value / tokow.Мощност).ToString("0.00")
                            Else
                                tokow.Консуматор = "Ке=0"
                            End If
                        End If
                    Case "Управление"
                        tokow.Управление = selectedValue
                    Case "Крива"
                        tokow.Breaker_Крива = selectedValue
                    Case "Защитен блок"
                        ' Обработка на параметър свързан със защитен модул
                        ' или допълнителна защита
                        tokow.Breaker_Защитен_блок = selectedValue
                    Case "Шина"
                        ' Шина е Boolean → True = на отделна шина, False = основна шина
                        tokow.Шина = CBool(selectedValue)
                    Case "ДТЗ (RCD)"
                            ' Управление на дефектнотокова защита (RCD) 
                            ' например включване/изключване на ДТЗ
                    Case "Начин на монтаж"
                        ' Взимаме само текстовата част за комбобокса, 
                        ' или подаваме целия списък, ако клетката е настроена за обекти
                        Dim displayValues = LiMountMethod.Select(Function(m) m.Text).ToList()
                        UpdateComboRow("Начин на монтаж", displayValues, e.ColumnIndex)
                    Case "Начин на полагане"
                        ' Правим прост списък само с двете опции
                        Dim valuesLaying As New List(Of String) From {"Във въздух", "В земя"}
                        If tokow.Кабел_Тип = "Al/R" Then
                            tokow.Кабел_Полагане = "Във въздух"
                            selectedValue = "Във въздух"
                        End If
                        CalculateCable(tokow,
                                       Type:=tokow.Кабел_Тип,
                                       layMethod:=If(selectedValue = "Във въздух", 0, 1),
                                       mountMethod:=GetMountMethodInfo(tokow.Кабел_Монтаж),
                                       Broj_Cable:=tokow.Кабел_Брой_Група,
                                       matType:=GetCableTypeResult(tokow.Кабел_Тип)
                                       )
                        ' Подаваме го към твоята процедура
                        UpdateComboRow("Начин на полагане", valuesLaying, e.ColumnIndex)
                    Case "Тип кабел"
                        ' Взимаме само уникалните имена на кабели от главния списък
                        Dim uniqueCableTypes As List(Of String) = Catalog_Cables _
                                                .Select(Function(c) c.CableType) _
                                                .Distinct() _
                                                .ToList()
                        If selectedValue = "Al/R" Then
                            tokow.Кабел_Полагане = "Във въздух"
                            CalculateCable(tokow,
                                           Type:=selectedValue,
                                           layMethod:=If(tokow.Кабел_Полагане = "Във въздух", 0, 1),
                                           mountMethod:=GetMountMethodInfo(tokow.Кабел_Монтаж),
                                           Broj_Cable:=tokow.Кабел_Брой_Група,
                                           matType:=GetCableTypeResult(selectedValue)
                                           )
                            UpdateComboRow("Тип кабел", uniqueCableTypes, e.ColumnIndex)
                            ' Правим прост списък само с двете опции
                            Dim valuesLaying As New List(Of String) From {"Във въздух", "В земя"}
                            UpdateComboRow("Начин на полагане", valuesLaying, e.ColumnIndex)
                        Else
                            ' Проверка дали стойността съществува в списъка
                            CalculateCable(tokow,
                                           Type:=selectedValue,
                                           layMethod:=If(tokow.Кабел_Полагане = "Във въздух", 0, 1),
                                           mountMethod:=GetMountMethodInfo(tokow.Кабел_Монтаж),
                                           Broj_Cable:=tokow.Кабел_Брой_Група,
                                           matType:=GetCableTypeResult(selectedValue)
                                           )
                        End If
                        ' Подаваме списъка към твоята процедура
                        UpdateComboRow("Тип кабел", uniqueCableTypes, e.ColumnIndex)
                    Case "ДТЗ Нула"
                        Dim inputValue As String = selectedValue?.ToString()
                        ' Извикай функцията за валидация
                        Dim validatedValue As String = ValidateRCDNulla(inputValue)
                        Update = False
                        ' Ако е валидно → запиши, иначе → върни старата стойност
                        If validatedValue IsNot Nothing Then
                            Update = True
                            tokow.RCD_Нула = validatedValue
                        End If
                End Select
                If Update Then UpdateCircuitColumn(tokow, col.Index, paramName)
            End If
        Finally
            'isUpdatingGrid = False
        End Try
    End Sub
    ''' <summary>
    ''' Проверява дали даден тип кабел принадлежи към предварително дефинирана група.
    ''' </summary>
    ''' <param name="cableName">Име/тип на кабела (например "САВТ", "NAYY" и др.)</param>
    ''' <returns>
    ''' Връща:
    ''' - 1 → ако кабелът принадлежи към зададената група
    ''' - 0 → ако не принадлежи
    ''' </returns>
    ''' <remarks>
    ''' Функцията сравнява подаденото име на кабел със списък от "целеви" кабели:
    '''     {"САВТ", "NA2XY", "Al/R", "NAYY"}
    '''
    ''' Използва се методът Contains(), който проверява дали има точно съвпадение.
    '''
    ''' Типично приложение:
    ''' - класификация на кабели (например алуминиеви)
    ''' - избор на коефициенти при изчисления
    ''' - условна логика при оразмеряване
    '''
    ''' Потенциални особености:
    ''' - Сравнението е case-sensitive (например "na2xy" няма да съвпадне)
    ''' - Ако входът е Nothing → ще върне 0 (без грешка)
    ''' - При нужда от разширение, списъкът може да се направи динамичен (List или база данни)
    '''
    ''' Възможно подобрение:
    ''' - Да се използва StringComparer.OrdinalIgnoreCase за нечувствително към регистъра сравнение
    ''' - Да се върне Boolean вместо Integer за по-ясна логика
    ''' </remarks>
    Public Function GetCableTypeResult(cableName As String) As Integer
        ' Списък с кабели, които попадат в конкретната група
        Dim targetCables As String() = {"САВТ", "NA2XY", "Al/R", "NAYY"}
        ' Проверка дали подаденият кабел е в списъка
        If targetCables.Contains(cableName) Then
            Return 1
        Else
            Return 0
        End If
    End Function
    ''' <summary>
    ''' Валидира и форматира стойност за RCD_Нула
    ''' </summary>
    ''' <param name="inputValue">Входна стойност (напр. "N1", "n2", "N10")</param>
    ''' <returns>Валидирана стойност или Nothing ако е невалидна</returns>
    Private Function ValidateRCDNulla(inputValue As String) As String
        ' Проверка 1: Дали е празно и дали започва с "N"
        If String.IsNullOrEmpty(inputValue) OrElse Not inputValue.ToUpper().StartsWith("N") Then Return Nothing
        ' Извлечи числото след "N"
        Dim numberPart As String = inputValue.Substring(1).Trim()
        ' Премахни всичко което НЕ е цифра
        numberPart = New String(numberPart.Where(Function(c) Char.IsDigit(c)).ToArray())
        ' Проверка 2: Дали има число
        If String.IsNullOrEmpty(numberPart) Then Return Nothing
        ' Проверка 3: Дали числото е валидно
        Dim rcdNumber As Integer
        If Not Integer.TryParse(numberPart, rcdNumber) Then Return Nothing
        ' Проверка 4: Дали числото е > 0
        If rcdNumber <= 0 Then Return Nothing
        ' ✅ Всички проверки минаха → върни валидираната стойност
        Return "N" & rcdNumber.ToString()
    End Function
    ''' <summary>
    ''' Обработва промяна на checkbox "Постави ДТЗ (RCD)"
    ''' </summary>
    Private Sub HandleRCDCheckboxChange(tokow As strTokow)
        If tokow.Device = "Контакт" Then Return
        If tokow.Device = "Разединител" Then Return
        tokow.RCD_Автомат = True
        If tokow.ДТЗ_RCD = True Then
            ' ✅ СЛАГАМЕ ДТЗ
            tokow.RCD_Автомат = True
            SetRCD(tokow)               ' Избираме ДТЗ от каталога
            ClearBreaker(tokow)         ' Изчистваме MCB данните
        Else
            ' След първоначалния избор може да се призичлява прекъсвач
            ' само когато се избра ДЗТ и след това се премахва!!!
            ' ------------------------------------------------------------
            calcBreaker = True
            ' ✅ СЛАГАМЕ ПРЕКЪСВАЧ
            CalculateBreaker(tokow)     ' Избираме прекъсвач
            ClearRCD(tokow)             ' Изчистваме ДТЗ данните
            tokow.RCD_Автомат = False
            ' След първоначалния избор може да се призичлява прекъсвач
            ' само когато се избра ДЗТ и след това се премахва!!!
            ' ------------------------------------------------------------
            calcBreaker = False
        End If
    End Sub
    ''' <summary>
    ''' Изчиства данните за прекъсвач (MCB)
    ''' </summary>
    Private Sub ClearBreaker(tokow As strTokow)
        tokow.Breaker_Тип_Апарат = ""           ' Серия апарат (EZ9, C120, NSX, MTZ)
        tokow.Breaker_Крива = ""                ' Характеристика (B, C, D)
        tokow.Breaker_Номинален_Ток = ""        ' Номинален ток (пример: "16A")
        tokow.Breaker_Изкл_Възможност = ""      ' Изключвателна способност ("6000A", "10000A")
        tokow.Breaker_Защитен_блок = ""         ' Изключвателна способност ("6000A", "10000A")
    End Sub
    ''' <summary>
    ''' Изчиства данните за ДТЗ (RCD)
    ''' </summary>
    Private Sub ClearRCD(tokow As strTokow)
        tokow.RCD_Бранд = ""
        tokow.RCD_Тип = ""
        tokow.RCD_Клас = ""
        tokow.RCD_Чувствителност = ""
        tokow.RCD_Ток = ""
        tokow.RCD_Полюси = ""
        tokow.RCD_Нула = ""
    End Sub
    ''' <summary>
    ''' Помощна процедура за обновяване на стойностите в ComboBox клетка
    ''' на определен ред и определена колона в DataGridView.
    '''
    ''' Основна идея:
    ''' DataGridView се използва като таблица с параметри, където:
    ''' - Първата колона (index 0) съдържа името на параметъра
    '''   (например "Тип на апарата", "Номинален ток", "ДТЗ (RCD)" и др.)
    ''' - Останалите колони съдържат стойности за конкретни обекти/кръгове.
    '''
    ''' Тази процедура намира реда, който съответства на даден параметър
    ''' (targetParamName) и обновява списъка със стойности на ComboBox клетката
    ''' в определена колона (columnIndex).
    '''
    ''' Типичен сценарий на използване:
    ''' Когато потребителят избере нов "Тип на апарата", трябва да се обнови
    ''' списъкът с възможни "Номинални токове" само в съответната колона.
    '''
    ''' Стъпки:
    ''' 1. Обхождаме всички редове в DataGridView.
    ''' 2. Търсим ред, чиято първа клетка съдържа името на параметъра.
    ''' 3. Взимаме клетката от конкретната колона.
    ''' 4. Проверяваме дали клетката е ComboBox.
    ''' 5. Изчистваме старите стойности и добавяме новите.
    ''' 6. Проверяваме дали текущата стойност е валидна спрямо новия списък.
    '''    Ако не е – изчистваме я.
    '''
    ''' Това позволява динамично обновяване на възможните стойности
    ''' в зависимост от други параметри в таблицата.
    ''' </summary>
    Private Sub UpdateComboRow(targetParamName As String, values As List(Of String), columnIndex As Integer)
        ' ------------------------------------------------------------
        ' 1) Обхождаме всички редове на DataGridView,
        '    за да намерим реда, който съответства на параметъра
        ' ------------------------------------------------------------
        For Each row As DataGridViewRow In DataGridView1.Rows
            ' Проверяваме дали текстът в първата колона съвпада
            ' с търсеното име на параметър
            If row.Cells(0).Value?.ToString() = targetParamName Then
                ' --------------------------------------------------------
                ' 2) Вземаме клетката от съответната колона
                ' --------------------------------------------------------
                ' Използваме TryCast, за да сме сигурни, че клетката
                ' е от тип DataGridViewComboBoxCell
                Dim comboCell = TryCast(row.Cells(columnIndex), DataGridViewComboBoxCell)
                If comboCell IsNot Nothing Then
                    ' ----------------------------------------------------
                    ' 3) Изчистваме старите стойности в ComboBox-а
                    ' ----------------------------------------------------
                    comboCell.Items.Clear()
                    ' ----------------------------------------------------
                    ' 4) Добавяме новите възможни стойности
                    ' ----------------------------------------------------
                    ' Вземаме само елементите, които НЕ са Nothing
                    Dim nonNullValues = values.Where(Function(v) v IsNot Nothing).ToArray()
                    If nonNullValues.Length > 0 Then
                        comboCell.Items.AddRange(nonNullValues)
                    End If
                    ' ----------------------------------------------------
                    ' 5) Проверяваме дали текущо избраната стойност
                    '    все още съществува в новия списък
                    ' ----------------------------------------------------
                    If comboCell.Value IsNot Nothing AndAlso
                                       Not values.Contains(comboCell.Value.ToString()) Then
                        ' Ако стойността вече не е валидна – изчистваме я
                        comboCell.Value = Nothing
                    End If
                End If
                ' --------------------------------------------------------
                ' 6) Намерили сме правилния ред – няма нужда да
                '    обхождаме останалите редове
                ' --------------------------------------------------------
                Exit For
            End If
        Next
    End Sub
    ''' <summary>
    ''' Изчислява необходимото сечение на кабел според тока и условията на полагане
    ''' Оптимизиран за сградни инсталации (90% под мазилка)
    ''' </summary>
    ''' <param name="tokow">        Токов кръг за който правим изчислението</param>
    ''' <param name="Type">         Тип на кабела: "СВТ", "САВТ", "NYY" и др. (по подразбиране "СВТ")</param>
    ''' <param name="layMethod">    Начин на полагане: 0 = във въздух (35°C), 1 = в земя (15°C) (по подразбиране 0)</param>
    ''' <param name="mountMethod">  Метод на монтаж по IEC: "A1"=гипсокартон, "B2"=под мазилка, "C"=над таван (по подразбиране "B2")</param>
    ''' <param name="Broj_Cable">   Брой кабели положени паралелно на скара (по подразбиране 1)</param>
    ''' <param name="Tipe_Cable">   Тип на проводника: 0 = кабел (3-жилен), 1 = проводник (1-жилен) (по подразбиране 0)</param>
    ''' <param name="matType">      Материал на проводника: 0 = мед (Cu), 1 = алуминий (Al) (по подразбиране 0)</param>
    ''' <param name="RetType">      Тип на връщаната стойност: 0 = само сечение, 1 = пълно означение (по подразбиране 1)</param>
    Private Sub CalculateCable(tokow As strTokow,
                                Optional Type As String = "СВТ",        ' Тип кабел (СВТ, САВТ, NYY...)
                                Optional layMethod As Integer = 0,      ' 0=въздух (35°C), 1=земя (15°C)
                                Optional mountMethod As String = "B1",  ' "A1"=гипсокартон, "B2"=под мазилка, "C"=над таван
                                Optional Broj_Cable As Integer = 1,     ' Брой паралелни кабели
                                Optional Tipe_Cable As Integer = 0,     ' 0=кабел (3-жилен), 1=проводник (1-жилен)
                                Optional matType As Integer = 0,        ' 0=мед (Cu), 1=алуминий (Al)
                                Optional RetType As Integer = 1         ' 0=само сечение, 1=пълно означение
                                )
        If tokow.Device = "Разединител" OrElse
           tokow.Device = "Съществуващ" OrElse
           tokow.Device = "Резерва" Then Exit Sub
        Dim Ibreaker As String = tokow.Breaker_Номинален_Ток
        Dim NumberPoles As String = tokow.Брой_Полюси
        ' 1. МАТЕРИАЛ И ФИЛТРИРАНЕ НА КАТАЛОГА
        Dim material As String = If(matType = 1, "Al", "Cu")
        Dim filteredCables = Catalog_Cables.Where(
                             Function(c) c.CableType = Type AndAlso
                                         c.Material = material
                             ).OrderBy(
                             Function(c) CDbl(c.PhaseSize.Replace(",", "."))
                             ).ToList()
        ' 2. КОРЕКЦИОННИ КОЕФИЦИЕНТИ
        ' K1 - брой кабели на скара
        Dim K1_Table As New Dictionary(Of Integer, Double) From {
                                        {1, 1.0},   ' 1 кабел → 100%
                                        {2, 0.88},  ' 2 кабела → 88%
                                        {3, 0.82},  ' 3 кабела → 82%
                                        {4, 0.77},  ' 4 кабела → 77%
                                        {5, 0.73},  ' 5 кабела → 73%
                                        {6, 0.7}    ' 6 кабела → 70%
                                        }
        Dim K1 As Double = If(K1_Table.ContainsKey(Broj_Cable), K1_Table(Broj_Cable), 0.7)
        ' K2 - температура
        Dim Qok As Double = If(layMethod = 1, 15, 35)  ' 15°C земя, 35°C въздух
        Const Qokdef As Double = 25
        Dim Q As Double = filteredCables(0).MaxWorkingTemp
        Dim K2 As Double = 1.0
        Dim ratio As Double = (Q - Qok) / (Q - Qokdef)
        If ratio > 0 Then K2 = Math.Sqrt(ratio)
        ' ✅ ТАБЛИЦА С КОЕФИЦИЕНТИ ЗА МОНТАЖ
        Dim MountCoefficients As New Dictionary(Of String, Double) From {
                        {"A1", 1.0},   ' Кабел в тръба в топлоизолирана стена
                        {"B1", 1.0},   ' Кабел в тръба върху стена
                        {"C", 1.0},    ' Кабел директно върху стена / кабелна скара
                        {"D1", 1.0},   ' Кабел в тръба в земята
                        {"D2", 1.0},   ' Кабел директно в земята
                        {"E", 1.0},    ' Кабел на въздух / кабелна скара
                        {"F", 1.0}     ' Кабели в пакет
                        }
        Dim K3 As Double = If(MountCoefficients.ContainsKey(mountMethod), MountCoefficients(mountMethod), 1.0)
        ' 3. ИЗБОР НА СЕЧЕНИЕ
        Dim calc As String = "######"
        Dim Inom As Double = Val(Ibreaker)
        Dim Idop As Double = Inom / (K1 * K2 * K3)
        ' ✅ ТЪРСИМ ПЪРВОТО СЕЧЕНИЕ КОЕТО ИЗДЪРЖА Idop
        For i As Integer = 0 To filteredCables.Count - 1
            Dim cable As CableInfo = filteredCables(i)
            ' ✅ ИЗБОР НА ТОК СПОРЕД layMethod
            Dim Imax As Double = 0
            If layMethod = 1 Then
                Imax = cable.MaxCurrent_Ground    ' Ток в земя
            Else
                Imax = cable.MaxCurrent_Air       ' Ток във въздух
            End If
            ' ✅ ПРОВЕРКА: ДАЛИ КАБЕЛЪТ ИЗДЪРЖА?
            If Imax >= Idop Then
                calc = cable.PhaseSize            ' Намерихме сечение!
                Exit For                          ' Излизаме от цикъла
            End If
        Next
        ' 4. ИЗВЛИЧАНЕ НА ТОКОВЕ ЗА ГОЛЕМИ СЕЧЕНИЯ (за паралелни кабели)
        Dim bestSection As String = ""
        Dim bestNum As Integer = 0
        Dim bestNeutral As String = ""
        If calc = "######" Then
            ' Променливи за съхранение на токовете
            Dim Current_120 As Double = 0
            Dim Current_150 As Double = 0
            Dim Current_185 As Double = 0
            Dim Current_240 As Double = 0
            ' Променливи за нулевите жила
            Dim Neutral_120 As String = ""
            Dim Neutral_150 As String = ""
            Dim Neutral_185 As String = ""
            Dim Neutral_240 As String = ""
            ' Търсим всяко сечение в filteredCables
            For Each cable As CableInfo In filteredCables
                Select Case cable.PhaseSize
                    Case "120"
                        Current_120 = If(layMethod = 1, cable.MaxCurrent_Ground, cable.MaxCurrent_Air)
                        Neutral_120 = cable.NeutralSize
                    Case "150"
                        Current_150 = If(layMethod = 1, cable.MaxCurrent_Ground, cable.MaxCurrent_Air)
                        Neutral_150 = cable.NeutralSize
                    Case "185"
                        Current_185 = If(layMethod = 1, cable.MaxCurrent_Ground, cable.MaxCurrent_Air)
                        Neutral_185 = cable.NeutralSize
                    Case "240"
                        Current_240 = If(layMethod = 1, cable.MaxCurrent_Ground, cable.MaxCurrent_Air)
                        Neutral_240 = cable.NeutralSize
                End Select
            Next
            Dim Idop_Adjusted As Double = Idop * 0.95

            Dim cables_120 As Integer = If(Current_120 > 0, CInt(Math.Ceiling(Idop_Adjusted / Current_120)), 0)
            Dim cables_150 As Integer = If(Current_150 > 0, CInt(Math.Ceiling(Idop_Adjusted / Current_150)), 0)
            Dim cables_185 As Integer = If(Current_185 > 0, CInt(Math.Ceiling(Idop_Adjusted / Current_185)), 0)
            Dim cables_240 As Integer = If(Current_240 > 0, CInt(Math.Ceiling(Idop_Adjusted / Current_240)), 0)

            Dim bestCables As Integer = 999

            Dim data = {
                New With {.Size = "120", .Price = 21.17 * cables_120, .Nom = cables_120, .Neutral = Neutral_120},
                New With {.Size = "150", .Price = 24.62 * cables_150, .Nom = cables_150, .Neutral = Neutral_150},
                New With {.Size = "185", .Price = 30.7 * cables_185, .Nom = cables_185, .Neutral = Neutral_185},
                New With {.Size = "240", .Price = 39.86 * cables_240, .Nom = cables_240, .Neutral = Neutral_240}
            }
            Dim bestMatch = data.Where(Function(x) x.Price > 0).OrderBy(Function(x) x.Price).FirstOrDefault()
            If bestMatch IsNot Nothing Then
                calc = bestMatch.Size
                bestSection = bestMatch.Size
                bestNum = bestMatch.Nom
                bestNeutral = bestMatch.Neutral
            End If
        End If
        ' 5. ФОРМАТИРАНЕ НА РЕЗУЛТАТА
        ' Ако RetType = 0, връщаме само сечението (напр. "2,5")
        Dim Text As String = ""
        If RetType = 0 Then
            Text = calc
        Else
            ' Определяне на броя жици според полюсите
            Dim Poles As String = If(NumberPoles = "1", "3x", "5x")
            Dim calc_N As String = ""
            ' Ако сечението е > 16mm², добавяме отделно нулево жило
            If Val(calc.Replace(",", ".")) > 16 Then
                Poles = " 4х"
                Dim index = filteredCables.FindIndex(Function(c) c.PhaseSize = calc)
                If index >= 0 Then
                    calc_N = filteredCables(index).NeutralSize
                End If
            End If
            ' Сглобяване на крайния низ
            Text = If(bestNum > 1, bestNum & "x", "")           ' Префикс за паралелни кабели
            Text += Type                                        ' Тип кабел (СВТ, САВТ...)
            Text += " "
            If Poles = "4х" AndAlso Not String.IsNullOrEmpty(calc_N) Then
                Text += "3х" & calc & "+" & calc_N              ' С нулево жило
            Else
                Text += Poles & calc                            ' Без нулево жило
            End If
            Text += "mm²"                                       ' Суфикс за единица
        End If
        tokow.Кабел_Брой_Фаза = bestNum
        tokow.Кабел_Брой_Група = Broj_Cable
        tokow.Кабел_Сечение = Text
        tokow.Кабел_Тип = Type
        tokow.Кабел_Полагане = If(layMethod = 0, "Във въздух", "В земя")
        tokow.Кабел_Монтаж = GetMountMethodInfo(mountMethod)
    End Sub
    Private Sub CalculateRCD()
        For Each tokow As strTokow In ListTokow
            If tokow.ДТЗ_RCD Then SetRCD(tokow)
        Next
    End Sub
    ''' <summary>
    ''' Определя подходяща диференциална токова защита (RCD/ДЗТ) за даден токов кръг (strTokow).
    ''' </summary>
    ''' <param name="tokow">Обект от тип strTokow, представляващ токов кръг или консуматор.</param>
    ''' <remarks>
    ''' Функцията избира RCD от каталога RCD_Catalog според следните критерии:
    ''' 1. Номинален ток >= 1.2 * ток на токовия кръг (минимум 20 A)
    ''' 2. Брой полюси (2p или 4p) спрямо фазовостта на кръга
    ''' 3. Дали устройството трябва да бъде RCBO (комбиниран с прекъсвач) или само RCCB
    '''
    ''' Стъпки на логиката:
    ''' - Определя се броят на полюсите според tokow.Брой_Полюси
    ''' - Изчислява се минималният необходим номинален ток (1.2 пъти токът на кръга или минимум 20 A)
    ''' - Филтрира се каталога RCD_Catalog по номинален ток, брой полюси и тип устройство (RCBO/RCCB)
    ''' - Ако няма съвпадение:
    '''   - Показва се предупреждение с всички търсени параметри и местоположението на токовия кръг (табло, токов кръг)
    ''' - Ако има съвпадение:
    '''   - Избира се първият подходящ RCD
    '''   - Актуализират се параметрите на tokow, включително:
    '''     Brand, DeviceType, Type, Sensitivity, NominalCurrent, Poles, Нула (N) и RCD_Автомат (Breaker)
    '''
    ''' Потенциални забележки:
    ''' - Ако RCD_Catalog е празен или няма подходящ RCD, се показва съобщение, но функцията не връща грешка програмно.
    ''' - Използването на First() предполага, че списъкът matchingRCDs е сортиран или е достатъчно добър избор първият елемент.
    ''' - Полето tokow се модифицира по стойност; ако strTokow е структура (Value Type), може да се наложи връщане на обновения обект или използване на ByRef.
    ''' - Изчислението на requiredCurrent включва коефициент 1.2; това е запас за безопасност според стандарти.
    ''' </remarks>
    Private Sub SetRCD(tokow As strTokow)
        If tokow.ТоковКръг = "ОБЩО" Then Return
        If tokow.ТоковКръг = "Разединител" Then Return
        ' Определяне на броя полюси на RCD: 4p за трифазен, 2p за еднофазен
        Dim poles As String = If(tokow.Брой_Полюси = 3, "4p", "2p")
        ' Минимален номинален ток: 1.2 * ток на кръга, но не по-малко от 20 A
        Dim requiredCurrent As Double = If(tokow.Ток * 1.2 < 20, 20, tokow.Ток * 1.2)
        ' Проверка дали е необходим RCBO (RCD с прекъсвач)
        Dim needRCBO As Boolean = tokow.RCD_Автомат
        'If poles = "4p" Then needRCBO = False
        ' Филтриране на каталога за подходящи RCD
        Dim matchingRCDs = RCD_Catalog.Where(
                           Function(r) r.NominalCurrent >= requiredCurrent AndAlso
                           r.Poles = poles AndAlso
                           r.Breaker = needRCBO
                           ).ToList()
        ' ----------------------------------------------------
        ' Ако не е намерена подходяща ДЗТ
        ' ----------------------------------------------------
        If matchingRCDs.Count = 0 Then
            Dim info As String = $"ВНИМАНИЕ: Не е намерена подходяща ДЗТ!{vbCrLf}{vbCrLf}" &
                                 $"Търсени параметри:{vbCrLf}" &
                                 $"- Мин. номинален ток: {requiredCurrent} A{vbCrLf}" &
                                 $"- Комбинирана (RCBO): {If(tokow.RCD_Автомат, "Да", "Не")}{vbCrLf}" &
                                 $"- Брой полюси: {poles}{vbCrLf}{vbCrLf}" &
                                 $"Местоположение:{vbCrLf}" &
                                 $"- Табло: {tokow.Tablo}{vbCrLf}" &
                                 $"- Токов кръг: {tokow.ТоковКръг}"
            MsgBox(info, MsgBoxStyle.Exclamation, "Липсваща апаратура в каталога")
        Else
            ' Избира се първият подходящ RCD
            Dim selectedRCD As RCDInfo = matchingRCDs.First()
            ' ------------------------------------------------
            ' Актуализиране на параметрите на токовия кръг
            ' според избраната ДЗТ
            ' ------------------------------------------------
            tokow.RCD_Бранд = selectedRCD.Brand
            tokow.RCD_Тип = selectedRCD.DeviceType
            tokow.RCD_Клас = selectedRCD.Type
            tokow.RCD_Чувствителност = selectedRCD.Sensitivity
            tokow.RCD_Ток = selectedRCD.NominalCurrent
            tokow.RCD_Полюси = selectedRCD.Poles
            tokow.RCD_Нула = "N"
            tokow.RCD_Автомат = selectedRCD.Breaker
            If tokow.RCD_Тип = "EZ9 RCBO" Then ClearBreaker(tokow)
        End If
    End Sub
    ''' <summary>
    ''' Намира токовия кръг в ListTokow базиран на:
    ''' 1. Избраното табло в TreeView1
    ''' 2. Заглавието на колоната в DataGridView (име на кръга)
    ''' </summary>
    Private Function FindTokowByColumn(circuitName As String) As strTokow
        ' 1. Взимаме таблото от TreeView
        Dim selectedTablo As String = TreeView1.SelectedNode?.Text
        If String.IsNullOrEmpty(selectedTablo) Then Return Nothing  ' Няма избрано табло
        ' 2. ИЗЧИСТВАНЕ НА ИМЕТО ОТ ДОПЪЛНИТЕЛЕН ТЕКСТ
        ' Пример: "Табло 1 (3 кръга, 5.2kW)" → "Табло 1"
        If selectedTablo.Contains("(") Then
            selectedTablo = selectedTablo.Substring(0, selectedTablo.IndexOf("(")).Trim()
        End If
        If String.IsNullOrEmpty(circuitName) Then Return Nothing  ' Няма име на кръг
        ' 3. Търсим в ListTokow по Табло + ТоковКръг
        Dim tokow As strTokow = ListTokow.FirstOrDefault(
                                Function(t) t.Tablo = selectedTablo AndAlso t.ТоковКръг = circuitName
                                )
        Return tokow  ' Може да е Nothing ако не е намерен
    End Function
    ''' <summary>
    ''' Събитие: DataGridView1_CurrentCellDirtyStateChanged
    ''' </summary>
    ''' <remarks>
    ''' Това събитие се извиква когато текущата клетка в DataGridView промени състоянието си
    ''' от "чиста" към "мръсна" (Dirty), което означава, че стойността в клетката е променена,
    ''' но все още не е записана окончателно в модела на данните.
    '''
    ''' Основната цел на тази процедура е да принуди DataGridView да запише веднага
    ''' новата стойност на клетката, когато тя е от тип:
    ''' - DataGridViewCheckBoxCell
    ''' - DataGridViewComboBoxCell
    '''
    ''' По подразбиране DataGridView записва новата стойност едва след като клетката загуби фокус.
    ''' Това може да създаде проблеми, когато логиката на програмата разчита на събитието
    ''' CellValueChanged да се изпълни веднага след промяна на стойността.
    '''
    ''' Чрез извикването на CommitEdit() стойността се записва незабавно,
    ''' което позволява на събития като:
    ''' - CellValueChanged
    ''' - CellValidated
    ''' да се задействат веднага.
    '''
    ''' Типични случаи на използване:
    ''' - когато Checkbox трябва да активира или деактивира други клетки
    ''' - когато ComboBox определя какви други параметри да се променят
    ''' - когато таблицата управлява логика на електрически параметри или избор на апаратура
    '''
    ''' Потенциални особености:
    ''' - Ако CommitEdit не се извика, CellValueChanged няма да се задейства веднага
    '''   при Checkbox и ComboBox.
    ''' - Това поведение е стандартна особеност на DataGridView в WinForms.
    ''' </remarks>
    Private Sub DataGridView1_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs) Handles DataGridView1.CurrentCellDirtyStateChanged
        ' Проверява дали текущата клетка е от тип CheckBox или ComboBox.
        ' Тези типове клетки често изискват незабавно записване на новата стойност.
        If TypeOf DataGridView1.CurrentCell Is DataGridViewCheckBoxCell OrElse
            TypeOf DataGridView1.CurrentCell Is DataGridViewComboBoxCell Then
            ' Проверява дали текущата клетка има незаписана промяна (Dirty state).
            If DataGridView1.IsCurrentCellDirty Then
                ' Принуждава DataGridView да запише новата стойност веднага.
                ' Това гарантира, че събитията за промяна на стойността ще се задействат незабавно.
                DataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit)
            End If
        End If
    End Sub
    ''' <summary>
    ''' Групира токовите кръгове с контакти по табла и създава групи за защита с ДТЗ (RCD).
    ''' </summary>
    ''' <remarks>
    ''' Процедурата анализира всички токови кръгове в колекцията ListTokow и:
    ''' - Групира кръговете по електрическо табло (Tablo).
    ''' - От всяко табло избира само кръговете, които съдържат контакти (brKontakt > 0).
    ''' - Определя броя на тези кръгове и според него създава групи за защита с ДТЗ.
    '''
    ''' Логика на групиране:
    ''' - 1 или 2 кръга → всички се поставят под една ДТЗ.
    ''' - 3 или повече кръга → извиква се процедурата GroupByThrees(),
    '''   която ги разделя на оптимални групи.
    '''
    ''' Целта е да се реализира стандартна практика при проектиране на табла,
    ''' при която няколко контактни кръга се защитават от една ДТЗ.
    '''
    ''' Потенциални особености:
    ''' - Ако няма контактни кръгове в таблото (n = 0), таблото се пропуска.
    ''' - Променливата rcdCounter служи за номериране на ДТЗ в рамките на таблото.
    ''' </remarks>
    Private Sub GroupContactsForRCD()
        ' Групиране на всички токови кръгове по табло
        Dim panels = ListTokow.GroupBy(Function(t) t.Tablo)
        For Each panelGroup In panels
            ' Избор само на кръговете, които съдържат контакти
            Dim contactCircuits = panelGroup.Where(
                            Function(t) t.brKontakt > 0 AndAlso t.Device <> "Табло"
                            ).ToList()
            ' Брой на контактните кръгове
            Dim n As Integer = contactCircuits.Count
            ' Ако няма такива кръгове – преминава към следващото табло
            If n = 0 Then Continue For
            ' Брояч за номера на ДТЗ в таблото
            Dim rcdCounter As Integer = 0
            Select Case n
            ' Един контактен кръг → една ДТЗ
                Case 1
                    rcdCounter += 1
                    CreateRCDGroup(contactCircuits, rcdCounter)
            ' Два контактни кръга → една ДТЗ
                Case 2
                    rcdCounter += 1
                    CreateRCDGroup(contactCircuits, rcdCounter)
            ' Три или повече → групиране по специална логика
                Case Is >= 3
                    GroupByThrees(contactCircuits, n, rcdCounter)
            End Select
        Next
    End Sub
    ''' <summary>
    ''' Разделя списък от токови кръгове на групи по 3 за защита с ДТЗ.
    ''' </summary>
    ''' <param name="circuits">Списък от токови кръгове.</param>
    ''' <param name="n">Общият брой кръгове.</param>
    ''' <param name="rcdCounter">Брояч на ДТЗ, предаван по референция.</param>
    ''' <remarks>
    ''' Основната цел е да се разпределят контактните кръгове в групи,
    ''' които да бъдат защитени с една ДТЗ.
    '''
    ''' Алгоритъм:
    ''' - Определя се броят на пълните групи по 3 (fullGroups).
    ''' - Определя се остатъкът (remainder).
    '''
    ''' Възможни случаи:
    ''' - remainder = 0 → всички групи са по 3 кръга.
    ''' - remainder = 1 → последните 4 кръга се групират заедно.
    ''' - remainder = 2 → последната група съдържа 2 кръга.
    '''
    ''' След създаване на групите:
    ''' - за всяка група се увеличава броячът на ДТЗ
    ''' - извиква се CreateRCDGroup() за създаване на защитата.
    '''
    ''' Потенциална особеност:
    ''' - При малък брой групи (например 4 кръга) алгоритъмът създава една група от 4,
    '''   вместо 3+1, което е по-практично при реални електрически табла.
    ''' </remarks>
    Private Sub GroupByThrees(circuits As List(Of strTokow), n As Integer, ByRef rcdCounter As Integer)
        ' Брой пълни групи по 3
        Dim fullGroups = n \ 3
        ' Остатък след групиране
        Dim remainder As Integer = n Mod 3
        ' Списък със създадените групи
        Dim groups As New List(Of List(Of strTokow))
        Select Case remainder
        ' Всички групи са по 3
            Case 0
                For i As Integer = 0 To fullGroups - 1
                    groups.Add(circuits.Skip(i * 3).Take(3).ToList())
                Next
        ' Последната група става 4
            Case 1
                For i As Integer = 0 To fullGroups - 2
                    groups.Add(circuits.Skip(i * 3).Take(3).ToList())
                Next
                groups.Add(circuits.Skip((fullGroups - 1) * 3).Take(4).ToList())
        ' Последната група е 2
            Case 2
                For i As Integer = 0 To fullGroups - 1
                    groups.Add(circuits.Skip(i * 3).Take(3).ToList())
                Next
                groups.Add(circuits.Skip(fullGroups * 3).Take(2).ToList())
        End Select
        ' Създаване на ДТЗ за всяка група
        For Each group In groups
            rcdCounter += 1
            CreateRCDGroup(group, rcdCounter)
        Next
    End Sub
    ''' <summary>
    ''' Създава група от токови кръгове, защитени от една ДТЗ.
    ''' </summary>
    ''' <param name="circuits">Списък от кръгове, които ще бъдат защитени от една ДТЗ.</param>
    ''' <param name="rcdNumber">Номер на ДТЗ в рамките на таблото.</param>
    ''' <remarks>
    ''' Процедурата извършва следните действия:
    '''
    ''' 1. Изчислява сумарния ток на всички кръгове в групата.
    ''' 2. Избира последния кръг в списъка като представителен за изчисленията.
    ''' 3. Проверява дали групата съдържа трифазен консуматор.
    ''' 4. Ако има трифазен консуматор:
    '''    - броят на полюсите се принудително задава на 3.
    ''' 5. Временно се задава сумарният ток на избрания кръг.
    ''' 6. Извиква се SetRCD(), която избира подходяща ДТЗ от каталога.
    ''' 7. На всички кръгове в групата се задава обща нула:
    '''    - "N1", "N2", "N3" и т.н.
    ''' 8. След това се възстановяват оригиналните стойности
    '''    на ток и брой полюси на последния кръг.
    '''
    ''' Потенциални особености:
    ''' - Методът използва последния кръг като временен носител на сумарния ток.
    ''' - Това е практично решение, но изисква внимателно възстановяване
    '''   на оригиналните стойности след изчислението.
    '''
    ''' Важна забележка:
    ''' - Ако структурата strTokow е Value Type (Structure),
    '''   промените върху елементите може да не се отразят в оригиналния списък,
    '''   ако не се използват по референция.
    ''' </remarks>
    Private Sub CreateRCDGroup(circuits As List(Of strTokow), rcdNumber As Integer)
        ' Сумарен ток на групата
        Dim totalCurrent As Double = circuits.Sum(Function(t) t.Ток)
        ' Последният кръг се използва като представителен за изчисленията
        Dim lastCircuit As strTokow = circuits.Last()
        ' Запазване на оригиналните параметри
        Dim originalTok As Double = lastCircuit.Ток
        Dim originalPoles As Integer = lastCircuit.Брой_Полюси
        ' Проверка дали има трифазен консуматор в групата
        Dim hasThreePhase As Boolean = circuits.Any(Function(t) t.Брой_Полюси = 3)
        ' Ако има трифазен консуматор → използва се 3-полюсна конфигурация
        If hasThreePhase Then lastCircuit.Брой_Полюси = 3
        ' Временно задаване на сумарния ток
        lastCircuit.Ток = totalCurrent
        ' Избор на подходяща ДТЗ
        SetRCD(lastCircuit)
        ' Задаване на обща нула за всички кръгове в групата
        For Each circuit In circuits
            circuit.RCD_Нула = "N" & rcdNumber.ToString()
        Next
        ' Възстановяване на оригиналните стойности
        lastCircuit.Ток = originalTok
        lastCircuit.Брой_Полюси = originalPoles
    End Sub
    Private Sub ToolStripButton_Поправи_ДЗТ_Click(sender As Object, e As EventArgs) Handles ToolStripButton_Поправи_ДЗТ.Click
        RedistributeRCDGroups()
        ' Refresh на DataGridView за да се видят новите ДЗТ настройки
        FillDataGridViewForPanel()
    End Sub
    ''' <summary>
    ''' Преразпределя ДЗТ според RCD_Нула стойностите в ListTokow
    ''' Извиква се при натискане на бутон "Поправи ДЗТ"
    ''' РАБОТИ САМО С ИЗБРАНОТО ТАБЛО В TREEVIEW!
    ''' </summary>
    Private Sub RedistributeRCDGroups()
        ' 1. ✅ Вземи избраното табло от TreeView
        Dim selectedTablo As String = TreeView1.SelectedNode?.Text
        ' Няма избрано табло
        If String.IsNullOrEmpty(selectedTablo) Then Return
        ' 2. ✅ Изчисти името от допълнителен текст (ако има)
        If selectedTablo.Contains("(") Then
            selectedTablo = selectedTablo.Substring(0, selectedTablo.IndexOf("(")).Trim()
        End If
        ' 3. ✅ Филтрирай само ТК за избраното табло
        Dim panelCircuits = ListTokow.Where(Function(t) t.Tablo = selectedTablo).ToList()
        If panelCircuits.Count = 0 Then Return
        ' 4. Филтрирай ТК с контакти
        Dim contactCircuits = panelCircuits.Where(Function(t) t.brKontakt > 0).ToList()
        If contactCircuits.Count = 0 Then Return
        ' 5. ✅ Изчистване на ТК без RCD_Нула
        For Each circuit In contactCircuits
            ClearRCDFields(circuit)
        Next
        ' 6. ✅ Групиране по RCD_Нула (само тези които имат стойност)
        Dim rcdGroups = contactCircuits.Where(Function(t) Not String.IsNullOrEmpty(t.RCD_Нула)).GroupBy(Function(t) t.RCD_Нула)
        ' 7. За всяка група → пресметни ДЗТ
        For Each rcdGroup In rcdGroups
            Dim circuitsInGroup = rcdGroup.ToList()
            ProcessRCDGroup(circuitsInGroup)
        Next
    End Sub
    Private Sub ClearRCDFields(circuit As strTokow)
        circuit.RCD_Бранд = ""
        circuit.RCD_Клас = ""
        circuit.RCD_Тип = ""
        circuit.RCD_Чувствителност = ""
        circuit.RCD_Ток = ""
        circuit.RCD_Полюси = ""
        circuit.RCD_Автомат = False
    End Sub
    ''' <summary>
    ''' Обработва една група ТК с еднакво RCD_Нула
    ''' Подобно на CreateRCDGroup() но без rcdNumber параметър
    ''' </summary>
    Private Sub ProcessRCDGroup(circuits As List(Of strTokow))
        ' 1. Сумирай токовете
        Dim totalCurrent As Double = circuits.Sum(Function(t) t.Ток)
        ' 2. Проверка за 3-фазни кръгове
        Dim hasThreePhase As Boolean = circuits.Any(Function(t) t.Брой_Полюси = 3 OrElse t.Фаза = "L1,L2,L3")
        ' 3. Вземи последния ТК
        Dim lastCircuit As strTokow = circuits.Last()
        ' 4. Запази оригиналните полюси и фаза
        Dim originalPoles As Integer = lastCircuit.Брой_Полюси
        Dim originalFaza As String = lastCircuit.Фаза
        Dim originalNula As String = lastCircuit.RCD_Нула
        ' 5. Ако има 3-фазен → временно промени
        If hasThreePhase Then lastCircuit.Брой_Полюси = 3
        ' 6. Извикай SetRCD()
        Dim originalTok As Double = lastCircuit.Ток
        lastCircuit.Ток = totalCurrent
        SetRCD(lastCircuit)
        lastCircuit.Ток = originalTok
        ' 7. Върни оригиналните полюси и фаза
        lastCircuit.Брой_Полюси = originalPoles
        lastCircuit.RCD_Нула = originalNula
    End Sub
    Private Sub ToolStripButton_Балансирай_фазите_Click(sender As Object, e As EventArgs) Handles ToolStripButton_Балансирай_фазите.Click
        ' =====================================================
        ' 1. ВЗЕМИ ИЗБРАНОТО ТАБЛО
        ' =====================================================
        Dim selectedTablo As String = TreeView1.SelectedNode?.Text
        ' Ако няма избран възел → прекратяване
        If String.IsNullOrEmpty(selectedTablo) Then Return
        ' Премахване на допълнителен текст (например "(...)" )
        If selectedTablo.Contains("(") Then
            selectedTablo = selectedTablo.Substring(0, selectedTablo.IndexOf("(")).Trim()
        End If
        BalancePhases(selectedTablo)
        FillDataGridViewForPanel()
    End Sub
    ''' <summary>
    ''' Балансира фазите (L1, L2, L3) за дадено табло и изчислява резултатните токове.
    ''' </summary>
    ''' <param name="selectedTablo">Име на таблото, за което ще се извърши балансиране.</param>
    ''' <remarks>
    ''' Процедурата извършва пълно балансиране на токовите кръгове и изчислява
    ''' крайното натоварване по фази, което записва в реда "ОБЩО".
    '''
    ''' Основна логика:
    ''' 1. Извлича всички кръгове за таблото (без ред "ОБЩО")
    ''' 2. Проверява за наличие на трифазни консуматори
    '''    - ако няма → пита потребителя дали да продължи
    ''' 3. Намира реда "ОБЩО" и го маркира като трифазен
    ''' 4. Създава групи за балансиране (Bus, RCD, Normal)
    ''' 5. Инициализира токовете по фази
    ''' 6. Добавя трифазните товари към всички фази
    ''' 7. Разпределя групите към най-слабо натоварената фаза (greedy алгоритъм)
    ''' 8. Преизчислява реалните токове по фази след разпределението
    ''' 9. Записва резултатите в ред "ОБЩО"
    ''' 10. Определя максималния фазов ток
    '''
    ''' Важни особености:
    ''' - Редът "ОБЩО" се използва като обобщение на резултатите
    ''' - Трифазните консуматори се разпределят равномерно към всички фази
    ''' - Еднофазните се разпределят чрез групиране
    '''
    ''' Потенциални рискове:
    ''' - Ако няма ред "ОБЩО" → ще възникне грешка (NullReference)
    ''' - Използва CDbl → зависи от регионалните настройки (десетичен разделител)
    ''' - Split(">") предполага винаги валиден формат на текста
    '''
    ''' Възможни подобрения:
    ''' - Проверка за Nothing при totalRow
    ''' - Използване на числови стойности вместо парсване на текст
    ''' - Сортиране на групите по ток преди балансиране
    ''' </remarks>
    Private Sub BalancePhases(selectedTablo As String)
        ' =====================================================
        ' 1. ВЗЕМИ КРЪГОВЕТЕ (БЕЗ "ОБЩО")
        ' =====================================================
        Dim panelCircuits = ListTokow.Where(Function(t)
                                                Return t.Tablo = selectedTablo AndAlso
                                                   t.ТоковКръг <> "ОБЩО"
                                            End Function).ToList()
        ' Ако няма кръгове → прекратяване
        If panelCircuits.Count = 0 Then Return
        ' =====================================================
        ' 2. ПРОВЕРКА ЗА ТРИФАЗНИ КОНСУМАТОРИ
        ' =====================================================
        Dim hasThreePhase As Boolean = panelCircuits.Any(
                                       Function(t) t.Брой_Полюси = 3 OrElse t.Фаза = "L1,L2,L3"
                                       )
        ' Ако няма → пита потребителя
        If Not hasThreePhase Then
            Dim result As MsgBoxResult = MessageBox.Show(
            "Няма трифазни консуматори в това табло." & vbCrLf & vbCrLf &
            "Искате ли да балансирате таблото?",
            "Балансиране на фазите",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question
        )
            If result = MsgBoxResult.No Then Return
        End If
        ' =====================================================
        ' 3. НАМИРАНЕ НА РЕД "ОБЩО"
        ' =====================================================
        Dim totalRow = ListTokow.FirstOrDefault(Function(t)
                                                    Return t.Tablo = selectedTablo AndAlso
                                                       t.ТоковКръг = "ОБЩО"
                                                End Function)

        ' Маркиране като трифазен
        totalRow.Брой_Полюси = 3
        totalRow.Фаза = "L1,L2,L3"
        ' =====================================================
        ' 4. СЪЗДАВАНЕ НА ГРУПИ
        ' =====================================================
        Dim balanceGroups As List(Of BalanceGroup) = CreateBalanceGroups(panelCircuits)
        ' =====================================================
        ' 5. ИНИЦИАЛИЗАЦИЯ НА ФАЗИТЕ
        ' =====================================================
        Dim phaseCurrents As New Dictionary(Of String, Double) From {
                                {"L1", 0},
                                {"L2", 0},
                                {"L3", 0}
        }
        ' =====================================================
        ' 6. ДОБАВЯНЕ НА ТРИФАЗНИ ТОВАРИ
        ' =====================================================
        Dim threePhaseCircuits = panelCircuits.Where(
                                 Function(t) t.Брой_Полюси = 3 OrElse t.Фаза = "L1,L2,L3"
                                 ).ToList()
        For Each circuit In threePhaseCircuits
            phaseCurrents("L1") += circuit.Ток
            phaseCurrents("L2") += circuit.Ток
            phaseCurrents("L3") += circuit.Ток
        Next
        ' =====================================================
        ' 7. БАЛАНСИРАНЕ (GREEDY)
        ' =====================================================
        For Each group In balanceGroups
            ' Най-слабо натоварена фаза
            Dim minPhase As String = phaseCurrents.Keys.
                                     OrderBy(Function(p) phaseCurrents(p)).
                                     First()
            ' Присвояване
            group.AssignedPhase = minPhase
            ' Запис в кръговете
            For Each circuit In group.Circuits
                circuit.Фаза = group.AssignedPhase
            Next
            ' Добавяне на товар
            phaseCurrents(minPhase) += group.TotalCurrent
        Next
        ' =====================================================
        ' 8. ПРЕИЗЧИСЛЕНИЕ НА ФАЗИТЕ
        ' =====================================================
        phaseCurrents("L1") = 0
        phaseCurrents("L2") = 0
        phaseCurrents("L3") = 0
        For Each circuit In panelCircuits
            ' Трифазен → към всички
            If circuit.Брой_Полюси = 3 OrElse circuit.Фаза = "L1,L2,L3" Then
                phaseCurrents("L1") += circuit.Ток
                phaseCurrents("L2") += circuit.Ток
                phaseCurrents("L3") += circuit.Ток

            Else
                ' Еднофазен → към конкретната фаза
                Dim p As String = circuit.Фаза.Trim().ToUpper()
                If phaseCurrents.ContainsKey(p) Then
                    phaseCurrents(p) += circuit.Ток
                End If
            End If
        Next
        ' =====================================================
        ' 9. ЗАПИС В "ОБЩО"
        ' =====================================================
        totalRow.RCD_Тип = "Ток фази"
        totalRow.RCD_Клас = "Фаза L1->" & phaseCurrents("L1").ToString("N2")
        totalRow.RCD_Ток = "Фаза L2->" & phaseCurrents("L2").ToString("N2")
        totalRow.RCD_Чувствителност = "Фаза L3->" & phaseCurrents("L3").ToString("N2")
        ' =====================================================
        ' 10. ОПРЕДЕЛЯНЕ НА МАКСИМАЛЕН ТОК
        ' =====================================================
        Dim valL1 As Double = CDbl(totalRow.RCD_Клас.Split(">"c)(1))
        Dim valL2 As Double = CDbl(totalRow.RCD_Ток.Split(">"c)(1))
        Dim valL3 As Double = CDbl(totalRow.RCD_Чувствителност.Split(">"c)(1))
        totalRow.Ток = Math.Max(valL1, Math.Max(valL2, valL3))
    End Sub
    ''' <summary>
    ''' Създава групи от токови кръгове за целите на балансиране на фазите.
    ''' </summary>
    ''' <param name="panelCircuits">Списък от токови кръгове (strTokow), принадлежащи към едно табло.</param>
    ''' <returns>Списък от групи (BalanceGroup), използвани за по-нататъшно разпределение по фази.</returns>
    ''' <remarks>
    ''' Функцията разделя всички токови кръгове в три основни типа групи:
    '''
    ''' 1. Шинни групи (Bus):
    '''    - Включва кръгове, които са маркирани с Шина = True и са еднофазни.
    '''    - Изчислява се процентното участие на шинните консуматори спрямо общата мощност.
    '''    - В зависимост от процента:
    '''          под 10% → "SmallBus"
    '''          над 10% → "LargeBus"
    '''
    ''' 2. Групи по ДТЗ (RCD):
    '''    - Включва еднофазни кръгове, които имат зададена RCD_Нула (N1, N2, ...)
    '''    - Изключва кръговете, които вече са част от шинна група.
    '''    - Групира се по стойността на RCD_Нула.
    '''
    ''' 3. Нормални групи (Normal):
    '''    - Включва всички останали еднофазни кръгове:
    '''         - без ДТЗ
    '''         - не са част от шинна група
    '''
    ''' За всяка група се изчислява:
    ''' - списък с кръгове
    ''' - общ ток (TotalCurrent)
    ''' - тип на групата (GroupType)
    ''' - ключ (GroupKey), използван за идентификация
    '''
    ''' Функцията връща списък от BalanceGroup, които могат да се използват за:
    ''' - балансиране на фазите
    ''' - оптимално разпределение на товарите
    ''' - анализ на натоварването
    '''
    ''' Потенциални особености:
    ''' - Само еднофазни кръгове (Брой_Полюси = 1) се включват в логиката за балансиране.
    ''' - Трифазните консуматори не се обработват тук (вероятно се третират отделно).
    ''' - Ако общата мощност е 0, се избягва деление на нула при изчисляване на процента.
    ''' - Използването на IIf може да доведе до изпълнение и на двата клона (VB особеност),
    '''   но в този контекст няма странични ефекти.
    ''' - Debug.Print се използва за диагностика и проследяване на създадените групи.
    ''' </remarks>
    Private Function CreateBalanceGroups(panelCircuits As List(Of strTokow)) As List(Of BalanceGroup)
        ' Списък с резултатните групи
        Dim groups As New List(Of BalanceGroup)
        ' ----------------------------------------------------
        ' 1. ШИННИ ГРУПИ (Bus)
        ' ----------------------------------------------------
        Dim busCircuits = panelCircuits.Where(
                                Function(t) t.Шина = True AndAlso t.Брой_Полюси = 1
                                ).ToList()
        If busCircuits.Count > 0 Then
            ' Обща мощност на таблото
            Dim totalPower As Double = panelCircuits.Sum(Function(t) t.Мощност)
            ' Мощност на шинните консуматори
            Dim busPower As Double = busCircuits.Sum(Function(t) t.Мощност)
            ' Процентно участие на шината
            Dim busPowerPercent As Double = 0
            If totalPower > 0 Then
                busPowerPercent = (busPower / totalPower) * 100
            End If
            ' Създаване на група за шината
            Dim busGroup As New BalanceGroup With {
                                .GroupType = IIf(busPowerPercent < 10, "SmallBus", "LargeBus"),
                                .GroupKey = "Bus",
                                .Circuits = busCircuits,
                                .TotalCurrent = busCircuits.Sum(Function(t) t.Ток)
            }

            groups.Add(busGroup)
        End If
        ' ----------------------------------------------------
        ' 2. ГРУПИ ПО ДТЗ (RCD)
        ' ----------------------------------------------------
        Dim rcdGroups = panelCircuits.Where(
        Function(t) t.Брой_Полюси = 1 AndAlso
                            Not String.IsNullOrEmpty(t.RCD_Нула) AndAlso
                            t.Шина = False   ' Изключва вече включените в шинна група
                            ).GroupBy(Function(t) t.RCD_Нула)
        For Each rcdGroup In rcdGroups
            Dim balanceGroup As New BalanceGroup With {
                            .GroupType = "RCD",
                            .GroupKey = rcdGroup.Key,  ' Например: N1, N2, N3
                            .Circuits = rcdGroup.ToList(),
                            .TotalCurrent = rcdGroup.Sum(Function(t) t.Ток)
                        }
            groups.Add(balanceGroup)
        Next
        ' ----------------------------------------------------
        ' 3. НОРМАЛНИ ГРУПИ (без ДТЗ и без шина)
        ' ----------------------------------------------------
        Dim normalCircuits = panelCircuits.Where(
                                Function(t) t.Брой_Полюси = 1 AndAlso
                                String.IsNullOrEmpty(t.RCD_Нула) AndAlso
                                t.Шина = False
                                ).ToList()

        For Each circuit In normalCircuits
            Dim normalGroup As New BalanceGroup With {
                        .GroupType = "Normal",
                        .GroupKey = Nothing,
                        .Circuits = New List(Of strTokow) From {circuit},
                        .TotalCurrent = circuit.Ток
                    }
            groups.Add(normalGroup)
        Next
        groups = groups.OrderByDescending(Function(g) g.TotalCurrent).ToList()
        Return groups
    End Function
    Private Sub SaveToolStripButton_Click(sender As Object, e As EventArgs) Handles SaveToolStripButton.Click
        Try
            ' Вземаме пътя на текущо отворения DWG
            Dim dwgPath As String = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Name
            Dim dwgFolder As String = IO.Path.GetDirectoryName(dwgPath)
            Dim dwgName As String = IO.Path.GetFileNameWithoutExtension(dwgPath)
            ' Създаваме име на JSON файла
            Dim savePath As String = IO.Path.Combine(dwgFolder, dwgName & "_Tokowi.json")
            ' Сериализиране на ListTokow
            Dim json As String = JsonConvert.SerializeObject(ListTokow, Formatting.Indented)
            ' Записване във файл
            IO.File.WriteAllText(savePath, json)
            MessageBox.Show($"Файлът е записан успешно: {savePath}", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Грешка при запис: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub OpenToolStripButton_Click(sender As Object, e As EventArgs) Handles OpenToolStripButton.Click
        Try
            ' Вземаме пътя на текущо отворения DWG
            Dim dwgPath As String = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Name
            Dim dwgFolder As String = IO.Path.GetDirectoryName(dwgPath)
            Dim dwgName As String = IO.Path.GetFileNameWithoutExtension(dwgPath)
            ' Създаваме име на JSON файла
            Dim openPath As String = IO.Path.Combine(dwgFolder, dwgName & "_Tokowi.json")
            ' Проверка дали файлът съществува
            If Not IO.File.Exists(openPath) Then
                MessageBox.Show("Файлът не е намерен: " & openPath, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            ' Четене и десериализиране
            Dim json As String = IO.File.ReadAllText(openPath)
            Dim loadedList As List(Of strTokow) = JsonConvert.DeserializeObject(Of List(Of strTokow))(json)
            If loadedList IsNot Nothing Then
                ' 🔹 Изтриваме текущото съдържание
                ListTokow.Clear()
                ' 🔹 Записваме съдържанието от файла
                ListTokow.AddRange(loadedList)
            Else
                ListTokow = New List(Of strTokow)
            End If
            ' 👉 Възстановяване на ObjectId от Handle_Block
            Dim doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db = doc.Database
            For Each t In ListTokow
                For Each k In t.Konsumator
                    If Not String.IsNullOrEmpty(k.Handle_Block) Then
                        Try
                            Dim h As New Handle(Convert.ToInt64(k.Handle_Block, 16))
                            k.ID_Block = db.GetObjectId(False, h, 0)
                        Catch
                            k.ID_Block = ObjectId.Null ' блокът вече не съществува
                        End Try
                    End If
                Next
            Next
            MessageBox.Show("Файлът е зареден успешно: " & openPath, "OK", MessageBoxButtons.OK, MessageBoxIcon.Information)
            ' 👉 Ако имаш DataGridView или UI, можеш да го обновиш:
            ' RefreshGrid()
        Catch ex As Exception
            MessageBox.Show("Грешка при четене на файла: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        BuildTreeViewFromKonsumatori()
        SetupDataGridView()
    End Sub
    Private Sub NewToolStripButton_Click(sender As Object, e As EventArgs) Handles NewToolStripButton.Click
        Try
            ' Потвърждение към потребителя (по желание)
            If MessageBox.Show("Сигурни ли сте, че искате да създадете нов проект? Всички незапазени данни ще бъдат загубени.",
                           "Ново", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.No Then
                Return
            End If
            ' 🔹 Изчистваме текущия ListTokow
            ListTokow.Clear()
            MessageBox.Show("Нов проект е готов. Всички стари данни са изчистени.", "New", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Грешка при създаване на нов проект: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        BuildTreeViewFromKonsumatori()
        SetupDataGridView()
    End Sub
    ''' <summary>
    ''' Копира стойността на избрана клетка от DataGridView в Clipboard (буфера).
    ''' </summary>
    ''' <param name="sender">Източник на събитието (бутон или shortcut).</param>
    ''' <param name="e">Аргументи на събитието.</param>
    ''' <remarks>
    ''' Процедурата реализира базова функционалност за копиране (Copy):
    '''
    ''' Логика на работа:
    ''' 1. Проверява дали има избрани клетки в DataGridView
    ''' 2. Взема първата избрана клетка
    ''' 3. Проверява дали стойността ѝ не е Nothing
    ''' 4. Преобразува стойността в текст (ToString)
    ''' 5. Записва текста в системния Clipboard
    '''
    ''' Поведение:
    ''' - Копира се само първата избрана клетка (дори да са избрани повече)
    ''' - Ако клетката е празна → не се прави нищо
    '''
    ''' Ограничения:
    ''' - Не поддържа копиране на множество клетки (табличен формат)
    ''' - Не запазва структура (редове/колони), както в Excel
    '''
    ''' Възможни подобрения:
    ''' - Поддръжка на multi-cell copy (TAB/NEWLINE формат)
    ''' - Копиране на цели редове или колони
    ''' - Проверка за тип на клетката (напр. ComboBox → SelectedValue/Text)
    '''
    ''' Забележка:
    ''' - Clipboard работи със стрингове, затова всички стойности се конвертират чрез ToString()
    ''' </remarks>
    Private Sub CopyToolStripButton_Click(sender As Object, e As EventArgs) Handles CopyToolStripButton.Click
        ' Проверка дали има избрани клетки
        If DataGridView1.SelectedCells.Count > 0 Then
            ' Взема първата избрана клетка
            Dim cellValue As Object = DataGridView1.SelectedCells(0).Value
            ' Проверка дали клетката съдържа стойност
            If cellValue IsNot Nothing Then
                ' Запис на стойността в Clipboard като текст
                My.Computer.Clipboard.SetText(cellValue.ToString())
            End If
        End If
    End Sub
    ''' <summary>
    ''' Поставя съдържание от Clipboard (буфера) в избраните клетки на DataGridView.
    ''' </summary>
    ''' <param name="sender">Източник на събитието (бутон или shortcut).</param>
    ''' <param name="e">Аргументи на събитието.</param>
    ''' <remarks>
    ''' Процедурата реализира функционалност за поставяне (Paste), подобна на Excel:
    '''
    ''' Логика на работа:
    ''' 1. Проверява дали Clipboard съдържа текст
    ''' 2. Ако има текст:
    '''    - извлича съдържанието
    '''    - обхожда всички маркирани клетки в DataGridView
    ''' 3. За всяка клетка:
    '''    - проверява дали не е ReadOnly
    '''    - опитва да зададе стойността
    '''    - при грешка (несъвместим тип) прескача клетката
    ''' 4. Обновява грида чрез Refresh()
    '''
    ''' Поведение:
    ''' - Един и същ текст се поставя във всички избрани клетки
    ''' - Ако клетката не позволява запис → пропуска се
    ''' - Ако типът не съвпада (напр. текст в числова клетка) → грешката се игнорира
    '''
    ''' Ограничения:
    ''' - Не поддържа multi-cell paste (таблични данни с редове/колони)
    ''' - Не обработва специални формати (напр. табулации, нови редове)
    '''
    ''' Възможни подобрения:
    ''' - Поддръжка на paste от Excel (разделяне по TAB и NewLine)
    ''' - Проверка за тип на клетката (ComboBox, CheckBox и др.)
    ''' - Валидация на входните данни преди запис
    ''' - Частично обновяване вместо пълен Refresh()
    '''
    ''' Забележка:
    ''' - My.Computer.Clipboard работи само в STA нишка (валидно за WinForms)
    ''' </remarks>
    Private Sub PasteToolStripButton_Click(sender As Object, e As EventArgs) Handles PasteToolStripButton.Click
        ' Проверка дали Clipboard съдържа текст
        If My.Computer.Clipboard.ContainsText() Then
            ' Вземане на текста от буфера
            Dim textToPaste As String = My.Computer.Clipboard.GetText()
            ' Обхождане на всички избрани клетки
            For Each cell As DataGridViewCell In DataGridView1.SelectedCells
                ' Проверка дали клетката може да се редактира
                If Not cell.ReadOnly Then
                    Try
                        ' Задаване на стойността
                        cell.Value = textToPaste
                    Catch ex As Exception
                        ' При грешка (напр. невалиден тип) → пропуска клетката
                        Continue For
                    End Try
                End If
            Next
            ' Обновяване на визуализацията
            DataGridView1.Refresh()
        Else
            ' Съобщение при празен или невалиден Clipboard
            MessageBox.Show("Буферът е празен или не съдържа текст!")
        End If
    End Sub
    ''' <summary>
    ''' Обработва натискане на клавиши в DataGridView за реализиране на бързи команди (shortcuts).
    ''' </summary>
    ''' <param name="sender">Източник на събитието (DataGridView).</param>
    ''' <param name="e">Информация за натиснатия клавиш.</param>
    ''' <remarks>
    ''' Процедурата прихваща клавишни комбинации с Ctrl и изпълнява съответните действия:
    '''
    ''' Поддържани комбинации:
    ''' - Ctrl + C → копиране (извиква CopyToolStripButton_Click)
    ''' - Ctrl + V → поставяне (извиква PasteToolStripButton_Click)
    '''
    ''' Логика:
    ''' 1. Проверява дали е натиснат Ctrl
    ''' 2. Проверява кой точно клавиш е натиснат (C, V и т.н.)
    ''' 3. Извиква съответната процедура
    ''' 4. Задава e.Handled = True, за да предотврати стандартното поведение на DataGridView
    '''
    ''' Предимства:
    ''' - Осигурява Excel-подобно поведение
    ''' - Централизира управление на shortcut-и
    '''
    ''' Разширяемост:
    ''' - Лесно могат да се добавят нови комбинации:
    '''     Ctrl + X → изрязване
    '''     Ctrl + A → селектиране на всичко
    '''     Ctrl + D → дублиране на ред и др.
    '''
    ''' Забележки:
    ''' - Ако не се зададе e.Handled = True, DataGridView може да изпълни и своето стандартно действие
    ''' - Copy/Paste логиката трябва да бъде реализирана в съответните методи
    ''' </remarks>
    Private Sub DataGridView1_KeyDown(sender As Object, e As KeyEventArgs) Handles DataGridView1.KeyDown
        ' Проверка дали е натиснат Ctrl
        If e.Control Then
            ' Проверка кой клавиш е натиснат
            Select Case e.KeyCode
            ' Ctrl + C → Копиране
                Case Keys.C
                    CopyToolStripButton_Click(sender, e)
                    e.Handled = True   ' Спира стандартното поведение
            ' Ctrl + V → Поставяне
                Case Keys.V
                    PasteToolStripButton_Click(sender, e)
                    e.Handled = True   ' Спира стандартното поведение

                    ' Пример за разширение:
                    ' Ctrl + X → Изрязване
                    'Case Keys.X
                    '    CutToolStripButton_Click(sender, e)
                    '    e.Handled = True

            End Select
        End If
    End Sub
    Private Sub ToolStripButton_Вмъни_Autocad_Click(sender As Object, e As EventArgs) Handles ToolStripButton_Вмъкни_Autocad.Click
        Try
            ' 1. ВЗЕМИ ИЗБРАНОТО ТАБЛО ОТ TREEVIEW
            Dim selectedTablo As String = TreeView1.SelectedNode?.Text
            If String.IsNullOrEmpty(selectedTablo) Then
                MsgBox("Моля, изберете табло от дървото!", MsgBoxStyle.Exclamation)
                Exit Sub
            End If
            ' Премахване на допълнителен текст (например "(...)" )
            If selectedTablo.Contains("(") Then
                selectedTablo = selectedTablo.Substring(0, selectedTablo.IndexOf("(")).Trim()
            End If
            '' ФИЛТРИРАЙ КРЪГОВЕТЕ ЗА ТОВА ТАБЛО
            Dim panelCircuits = ListTokow.Where(Function(t)
                                                    Return t.Tablo = selectedTablo
                                                End Function).ToList()
            If panelCircuits.Count = 0 Then
                MsgBox("Няма намерени кръгове за табло: " & selectedTablo, MsgBoxStyle.Exclamation)
                Exit Sub
            End If
            ' Вземане на текущия AutoCAD документ, редактор и база
            Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim edt As Editor = acDoc.Editor
            Dim acCurDb As Database = acDoc.Database
            ' ВЗЕМИ БАЗОВА ТОЧКА ОТ ПОТРЕБИТЕЛЯ
            Dim ptBasePointRes As PromptPointResult
            Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
            pPtOpts.Message = vbLf & "Изберете долен ляв ъгъл на таблото: "
            ptBasePointRes = acDoc.Editor.GetPoint(pPtOpts)
            If ptBasePointRes.Status = PromptStatus.Cancel Then Exit Sub
            Dim ptBasePoint As Point3d = ptBasePointRes.Value
            Me.Visible = False
            ' Проверяваме дали има кръгове на отделна шина
            twoBus = panelCircuits.Any(Function(c) c.Шина)
            hasDisconnector = panelCircuits.Any(Function(c) c.Device = "Разединител")
            If twoBus Then
                ' Проверяваме дали НЯМА нито един елемент с Device = "Разединител"
                If Not hasDisconnector Then
                    ' Извеждаме съобщение и прекратяваме процедурата
                    MessageBox.Show("Две шини – добре. Разединител – няма. Софтуерът изпада в депресия!")
                    Return
                End If
            End If
            ' 6. СТАРТИРАЙ ЧЕРТАНЕТО В ТРАНЗАКЦИЯ
            Using trans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Try
                    ' ПРЕДИЗЧИСЛЯВАНЕ НА ПАРАМЕТРИТЕ
                    ' Тук ще извикваме процедурите за чертане една по една
                    DrawPanelFrame(acDoc, acCurDb, ptBasePoint, panelCircuits, selectedTablo)   ' Тук чертаем рамката на таблото
                    DrawBusbars(acDoc, acCurDb, ptBasePoint, panelCircuits)                     ' Тук чертаем шините
                    DrawCircuits(acDoc, acCurDb, ptBasePoint, panelCircuits)                    ' Тук чертаем всеки токов кръг (прекъсвачи, текстове, линии)
                    DrawRCDBusbar(acDoc, acCurDb, ptBasePoint, panelCircuits)                   ' Тук чертаем ДЗТ за токовите кръгове (прекъсвачи, текстове, линии)




                    DrawMainSwitch(acDoc, acCurDb, ptBasePoint, panelCircuits)
                    DrawGrounding(acDoc, acCurDb, ptBasePoint.X, ptBasePoint, selectedTablo)   ' Чертaем заземление само за главно разпределително табло
                    DrawAnnotations(ptBasePoint, panelCircuits)                                ' Процедурата създава текстови анотации
                Catch ex As Exception
                    trans.Abort()
                    MsgBox("Възникна грешка при чертане: " & vbCrLf & ex.Message & vbCrLf & vbCrLf & ex.StackTrace, MsgBoxStyle.Critical)
                Finally
                    trans.Commit()
                End Try
            End Using
        Catch ex As Exception
            ' Извличане на информация за реда
            Dim st As New StackTrace(ex, True)
            Dim frame As StackFrame = st.GetFrame(0) ' Взема последния фрейм, където е гръмнало
            Dim line As Integer = frame.GetFileLineNumber()
            Dim fileName As String = frame.GetFileName()

            Dim errorMsg As String = String.Format(
                "Грешка: {0}" & vbCrLf &
                "Файл: {1}" & vbCrLf &
                "Ред: {2}" & vbCrLf & vbCrLf &
                "StackTrace: {3}",
                ex.Message, fileName, line, ex.StackTrace)
            MsgBox(errorMsg, MsgBoxStyle.Critical)
        Finally
            Me.Visible = True
        End Try
        FillDataGridViewForPanel()
    End Sub
    ''' <summary>
    ''' Чертaе RCD (ДТЗ) групи и разпределителни линии в таблото.
    ''' Логиката:
    ''' 1. Изчислява позициите на шините
    ''' 2. Обхожда всички токови кръгове
    ''' 3. Групира кръговете по RCD_Нула
    ''' 4. При смяна на групата затваря предишната
    ''' 5. Изчертава обща линия и RCD блок за всяка група
    ''' </summary>
    Private Sub DrawRCDBusbar(acDoc As Document, acCurDb As Database,
                          basePoint As Point3d,
                          circuits As List(Of strTokow))
        ' Изчисляване на началната X координата
        Dim X_Start As Double =
        basePoint.X + widthText + widthTextDim
        ' Y координата на главната шина
        Dim Y_Shina As Double = basePoint.Y + Y_Шина
        ' Y координата на RCD шината
        Dim Y_RCD As Double = Y_Shina - 118
        ' Начална колона на текущата RCD група
        Dim rcdGroupStart As Integer = 0
        ' RCD_Нула от предишния токов кръг
        Dim previousRCD_Null As String = ""
        ' Флаг дали в момента сме в активна RCD група
        Dim inRCDGroup As Boolean = False
        ' Текущ индекс на колоната
        Dim colIndex As Integer = 0
        ' Запазва последния токов кръг в активната група
        Dim currentGroupCircuit As strTokow = Nothing
        Try
            ' Обхождаме всички токови кръгове
            For Each circuit As strTokow In circuits
                ' Пропускаме специалните кръгове
                If circuit.ТоковКръг = "Разединител" OrElse
               circuit.ТоковКръг = "ОБЩО" Then Continue For
                ' Преминаваме към следващата колона
                colIndex += 1
                ' Проверка дали кръгът има ДТЗ
                Dim hasRCD As Boolean =
                Not String.IsNullOrEmpty(circuit.RCD_Нула) AndAlso
                circuit.RCD_Нула.Trim().ToUpper() <> "N"
                If hasRCD Then
                    ' Ако няма активна група → започваме нова
                    If Not inRCDGroup Then
                        rcdGroupStart = colIndex
                        previousRCD_Null = circuit.RCD_Нула.Trim().ToUpper()
                        currentGroupCircuit = circuit
                        inRCDGroup = True
                        ' Ако RCD_Нула е различно → затваряме старата група
                    ElseIf circuit.RCD_Нула.Trim().ToUpper() <> previousRCD_Null Then
                        DrawRCDGroupLine(
                                        acDoc,
                                        acCurDb,
                                        X_Start,
                                        Y_RCD,
                                        Y_Shina,
                                        rcdGroupStart,
                                        colIndex - 1,
                                        currentGroupCircuit
                                        )
                        ' Стартираме нова група
                        rcdGroupStart = colIndex
                        previousRCD_Null = circuit.RCD_Нула.Trim().ToUpper()
                        currentGroupCircuit = circuit
                    Else
                        ' Същото RCD_Нула → групата продължава
                        currentGroupCircuit = circuit
                    End If
                Else
                    ' Ако няма ДТЗ и има активна група → затваряме я
                    If inRCDGroup Then
                        DrawRCDGroupLine(
                                acDoc,
                                acCurDb,
                                X_Start,
                                Y_RCD,
                                Y_Shina,
                                rcdGroupStart,
                                colIndex - 1,
                                currentGroupCircuit)
                        inRCDGroup = False
                        previousRCD_Null = ""
                        currentGroupCircuit = Nothing
                    End If
                End If
            Next
            ' Ако последната група е останала отворена → затваряме я
            If inRCDGroup Then
                DrawRCDGroupLine(
                        acDoc,
                        acCurDb,
                        X_Start,
                        Y_RCD,
                        Y_Shina,
                        rcdGroupStart,
                        colIndex,
                        currentGroupCircuit)

            End If
        Catch ex As Exception
            MsgBox(
            "Възникна грешка: " &
            vbCrLf &
            ex.Message &
            vbCrLf &
            vbCrLf &
            ex.StackTrace,
            MsgBoxStyle.Critical
        )
        End Try
    End Sub
    ''' <summary>
    ''' Чертaе група с ДТЗ (RCD):
    ''' - хоризонтална шина
    ''' - блок на RCD в центъра
    ''' - текст с фази над шината
    ''' - попълва атрибутите на блока
    ''' </summary>
    ''' <param name="acDoc">Текущ документ</param>
    ''' <param name="acCurDb">Текуща база данни</param>
    ''' <param name="X_Start">Начална X позиция</param>
    ''' <param name="Y_RCD">Y позиция на шината (хоризонталната линия)</param>
    ''' <param name="Y_Shina">Y позиция за поставяне на RCD блока</param>
    ''' <param name="groupStart">Начална колона на групата</param>
    ''' <param name="groupEnd">Крайна колона на групата</param>
    ''' <param name="circuits">Данни за токовия кръг (RCD параметри)</param>
    Private Sub DrawRCDGroupLine(acDoc As Document, acCurDb As Database,
                         X_Start As Double, Y_RCD As Double, Y_Shina As Double,
                         groupStart As Integer, groupEnd As Integer,
                         circuits As strTokow
                         )
        Try
            ' 1 ИЗЧИСЛЯВАНЕ НА X ПОЗИЦИИТЕ
            ' Лява граница на групата
            Dim X_First As Double = X_Start + (groupStart - 1) * widthColom + widthColom / 4
            ' Дясна граница на групата
            Dim X_Last As Double = X_Start + (groupEnd) * widthColom - widthColom / 4
            ' Център на групата (за позициониране на RCD блока)
            Dim X_Center As Double = (X_First + X_Last) / 2
            ' 2️ ЧЕРТАНЕ НА ХОРИЗОНТАЛНА ЛИНИЯ (ШИНА)
            ' Чертaе шината между първата и последната позиция
            cu.DrowLine(New Point3d(X_First, Y_RCD, 0),
                New Point3d(X_Last, Y_RCD, 0),
                "EL_ТАБЛА",
                Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight070,
                "ByLayer")
            ' 3️ ВМЪКВАНЕ НА RCD БЛОК
            ' Поставя блока в центъра на групата
            Dim rcdBlockId As ObjectId = cu.InsertBlock("s_id_res_circ_break",
                                                 New Point3d(X_Center, Y_Shina, 0),
                                                 "EL_ТАБЛА",
                                                 New Scale3d(5, 5, 5))
            ' 4️⃣ ДОБАВЯНЕ НА ТЕКСТ НАД ШИНАТА
            ' Y позиция на текста (малко над линията)
            Dim textY As Double = Y_RCD + 15
            ' Текст с фази + нула + защитен проводник
            Dim phaseText As String = circuits.Фаза & "," & circuits.RCD_Нула & "," & "PE"
            ' Вмъкване на текста
            cu.InsertText(phaseText,
                  New Point3d(X_First, textY, 0),
                  "EL__DIM",
                  10,
                  TextHorizontalMode.TextLeft,
                  TextVerticalMode.TextBase)
            ' 5️ ПОПЪЛВАНЕ НА АТРИБУТИ НА БЛОКА
            ' Проверка дали блокът е създаден успешно
            If Not rcdBlockId.IsNull Then
                ' Стартираме транзакция за редакция
                Using trans As Transaction = acCurDb.TransactionManager.StartTransaction()
                    ' Вземаме референция към блока
                    Dim acBlkRef As BlockReference =
                                    DirectCast(trans.GetObject(rcdBlockId, OpenMode.ForWrite), BlockReference)
                    ' Обхождаме всички атрибути на блока
                    For Each objID As ObjectId In acBlkRef.AttributeCollection
                        ' Вземаме конкретен атрибут
                        Dim acAttRef As AttributeReference =
                                        DirectCast(trans.GetObject(objID, OpenMode.ForWrite), AttributeReference)
                        ' Попълваме според TAG-а
                        Select Case acAttRef.Tag
                            Case "1" : acAttRef.TextString = circuits.RCD_Клас
                            Case "2" : acAttRef.TextString = circuits.RCD_Полюси
                            Case "3" : acAttRef.TextString = circuits.RCD_Ток & "А"
                            Case "4" : acAttRef.TextString = "Мигновена"
                            Case "5" : acAttRef.TextString = circuits.RCD_Чувствителност & "mА"
                            Case "SHORTNAME" : acAttRef.TextString = circuits.RCD_Тип
                            Case "REFNB" : acAttRef.TextString = circuits.Tablo
                            Case "DESIGNATION" : acAttRef.TextString = ""
                        End Select
                    Next
                    ' Записваме промените
                    trans.Commit()
                End Using
            End If
        Catch ex As Exception
            ' Обработка на грешка – показва съобщение с детайли
            MsgBox("Възникна грешка: " & vbCrLf &
               ex.Message & vbCrLf & vbCrLf &
               ex.StackTrace,
               MsgBoxStyle.Critical)
        End Try
    End Sub
    Private Sub AddLine(lines As List(Of LineDefinition),
                    startPoint As Point3d, endPoint As Point3d,
                    Optional layer As String = "EL_ТАБЛА",
                    Optional lineWeight As Autodesk.AutoCAD.DatabaseServices.LineWeight = Autodesk.AutoCAD.DatabaseServices.LineWeight.ByLayer,
                    Optional lineType As String = "ByLayer",
                    Optional colorIndex As Integer = -1)
        lines.Add(New LineDefinition(startPoint, endPoint, layer, lineWeight, lineType, colorIndex))
    End Sub
    ''' <summary>
    ''' Изчертава рамката на електрическо табло в AutoCAD.
    ''' Включва позициониране спрямо базова точка и използва данни от подадените токови кръгове.
    ''' </summary>
    ''' <param name="acDoc">Активният AutoCAD документ.</param>
    ''' <param name="acCurDb">Текущата база данни на чертежа.</param>
    ''' <param name="basePoint">Базова точка за позициониране на рамката.</param>
    ''' <param name="circuits">Списък с токови кръгове, използван за определяне на размерите и съдържанието.</param>
    ''' <param name="selectedTablo">Име на таблото, за което се чертае рамката.</param>
    ''' <remarks>
    ''' Процедурата изгражда графичната рамка на таблото, като използва
    ''' геометрични зависимости и данни от токовите кръгове.
    ''' Използва помощни функции за чертане на линии и текст.
    ''' </remarks>
    Private Sub DrawPanelFrame(acDoc As Document, acCurDb As Database, basePoint As Point3d,
                               circuits As List(Of strTokow), selectedTablo As String)
        Try
            ' =====================================================
            ' 1️ ИЗЧИСЛЯВАНЕ НА ОСНОВНИТЕ РАЗМЕРИ
            ' =====================================================
            Dim brColums As Integer = circuits.Count - If(twoBus, 1, 0)
            Dim tableWidth As Double = basePoint.X + widthText + widthTextDim + (brColums) * widthColom
            Dim tableHeight As Double = 10 * heightRow
            ' =====================================================
            ' 2️ СЪЗДАВАНЕ НА СПИСЪК С ЛИНИИТЕ
            ' =====================================================
            Dim lines As New List(Of LineDefinition)
            ' --- Хоризонтални линии на таблицата ---
            ' Долна линия (ред 0)
            AddLine(lines, New Point3d(basePoint.X, basePoint.Y, 0),
                       New Point3d(tableWidth, basePoint.Y, 0))
            ' Хоризонтални линии за редове 3-10
            For row As Integer = 3 To 10
                AddLine(lines, New Point3d(basePoint.X, basePoint.Y + row * heightRow, 0),
                           New Point3d(tableWidth, basePoint.Y + row * heightRow, 0))
            Next
            ' --- Вертикални линии на таблицата ---
            ' Ляв край
            AddLine(lines, New Point3d(basePoint.X, basePoint.Y, 0),
                       New Point3d(basePoint.X, basePoint.Y + tableHeight, 0))
            ' След "Токов кръг"
            AddLine(lines, New Point3d(basePoint.X + widthText, basePoint.Y, 0),
                       New Point3d(basePoint.X + widthText, basePoint.Y + tableHeight, 0))
            ' След "№"
            AddLine(lines, New Point3d(basePoint.X + widthText + widthTextDim, basePoint.Y, 0),
                       New Point3d(basePoint.X + widthText + widthTextDim, basePoint.Y + tableHeight, 0))
            ' Вертикални линии за всеки токов кръг
            For col As Integer = 1 To brColums
                Dim xLine As Double = basePoint.X + widthText + widthTextDim + col * widthColom
                AddLine(lines, New Point3d(xLine, basePoint.Y, 0),
                               New Point3d(xLine, basePoint.Y + tableHeight, 0))
            Next
            ' --- Рамка на блока с информация за шината ---
            Dim blockStartY As Double = basePoint.Y + tableHeight + lengthProw
            Dim blockEndY As Double = blockStartY + widthTablo
            ' Лява страна (CENTER тип)
            AddLine(lines, New Point3d(basePoint.X + widthText, blockStartY, 0),
                       New Point3d(basePoint.X + widthText, blockEndY, 0),
                       "EL_ТАБЛА", Autodesk.AutoCAD.DatabaseServices.LineWeight.ByLayer, "CENTER")
            ' Долна страна
            AddLine(lines, New Point3d(basePoint.X + widthText, blockStartY, 0),
                       New Point3d(tableWidth, blockStartY, 0),
                       "EL_ТАБЛА", Autodesk.AutoCAD.DatabaseServices.LineWeight.ByLayer, "CENTER")
            ' Дясна страна
            AddLine(lines, New Point3d(tableWidth, blockStartY, 0),
                       New Point3d(tableWidth, blockEndY, 0),
                       "EL_ТАБЛА", Autodesk.AutoCAD.DatabaseServices.LineWeight.ByLayer, "CENTER")
            ' Горна страна
            AddLine(lines, New Point3d(basePoint.X + widthText, blockEndY, 0),
                       New Point3d(tableWidth, blockEndY, 0),
                       "EL_ТАБЛА", Autodesk.AutoCAD.DatabaseServices.LineWeight.ByLayer, "CENTER")
            ' --- Червен кръст за маркировка (Defpoints) ---
            Dim crossCenterX As Double = basePoint.X + widthText + 18
            Dim crossCenterY As Double = blockEndY - 18
            ' Вертикална линия на кръста
            AddLine(lines, New Point3d(crossCenterX, blockEndY, 0),
                           New Point3d(crossCenterX, blockEndY - 36, 0),
                           "Defpoints", Autodesk.AutoCAD.DatabaseServices.LineWeight.ByLayer, "ByLayer", 1)
            ' Хоризонтална линия на кръста
            AddLine(lines, New Point3d(basePoint.X + widthText, crossCenterY, 0),
                       New Point3d(basePoint.X + widthText + 36, crossCenterY, 0),
                       "Defpoints", Autodesk.AutoCAD.DatabaseServices.LineWeight.ByLayer, "ByLayer", 1)
            ' =====================================================
            ' 3️ ЧЕРТАЕНЕ НА ВСИЧКИ ЛИНИИ ОТ СПИСЪКА
            ' =====================================================
            For Each line As LineDefinition In lines
                If line.ColorIndex = -1 Then
                    cu.DrowLine(line.StartPoint, line.EndPoint, line.Layer, line.LineWeightValue, line.LineType)
                Else
                    cu.DrowLine(line.StartPoint, line.EndPoint, line.Layer, line.LineWeightValue, line.LineType, line.ColorIndex)
                End If
            Next
            ' =====================================================
            ' 4️ ТЕКСТОВЕ - ПЪРВА КОЛОНА (ЗАГЛАВКИ)
            ' =====================================================
            Dim textX As Double = basePoint.X + padingText
            Dim textY As Double = basePoint.Y + (heightRow - heightText) / 2
            cu.InsertText("Токов кръг", New Point3d(textX, textY + 9 * heightRow, 0),
                      "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("Брой лампи", New Point3d(textX, textY + 8 * heightRow, 0),
                      "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("Брой контакти", New Point3d(textX, textY + 7 * heightRow, 0),
                      "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("Инстал. мощност", New Point3d(textX, textY + 6 * heightRow, 0),
                      "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("Тип кабел", New Point3d(textX, textY + 5 * heightRow, 0),
                      "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("Сечение кабел", New Point3d(textX, textY + 4 * heightRow, 0),
                      "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("Фаза", New Point3d(textX, textY + 3 * heightRow, 0),
                      "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("Консуматор", New Point3d(textX, textY + 2 * heightRow, 0),
                      "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            ' =====================================================
            ' 5️ ТЕКСТОВЕ - ВТОРА КОЛОНА (МЕРНИ ЕДИНИЦИ)
            ' =====================================================
            textX = textX + widthText
            cu.InsertText("№", New Point3d(textX, textY + 9 * heightRow, 0),
                      "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("бр.", New Point3d(textX, textY + 8 * heightRow, 0),
                      "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("бр.", New Point3d(textX, textY + 7 * heightRow, 0),
                      "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("kW", New Point3d(textX, textY + 6 * heightRow, 0),
                      "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("---", New Point3d(textX, textY + 5 * heightRow, 0),
                      "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("mm²", New Point3d(textX, textY + 4 * heightRow, 0),
                      "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("---", New Point3d(textX, textY + 3 * heightRow, 0),
                      "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("---", New Point3d(textX, textY + 2 * heightRow, 0),
                      "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)

            Dim X = basePoint.X + widthText + widthTextDim
            cu.InsertText(selectedTablo,
                          New Point3d(X + (brColums - 1) * widthColom,
                                      basePoint.Y + Y_Шина + 95,
                                      0),
                          "EL__DIM", heightText + 5, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
        Catch ex As Exception
            ' --------------------------------------------------------
            ' Обработка на грешки
            ' --------------------------------------------------------
            MsgBox("Възникна грешка: " &
                   ex.Message &
                   vbCrLf & vbCrLf &
                   ex.StackTrace.ToString)
        End Try
    End Sub
    ''' <summary>
    ''' Процедурата DrawBusbars отговаря за изчертаването на шините (busbars)
    ''' в електрическо табло в AutoCAD.
    ''' 
    ''' В зависимост от конфигурацията (една или две шини), процедурата:
    ''' - изчислява геометрията (позиции и дължини)
    ''' - определя фазите
    ''' - изчертава линии за шините
    ''' - добавя текстови надписи за фазите
    ''' - при две шини – визуализира връзката между тях
    ''' 
    ''' Използва данни от списък с токови кръгове (strTokow).
    ''' </summary>
    ''' <param name="acDoc">AutoCAD документ (не се използва директно, но е част от контекста).</param>
    ''' <param name="acCurDb">AutoCAD база данни (не се използва директно тук).</param>
    ''' <param name="basePoint">Базова точка за позициониране на всички елементи.</param>
    ''' <param name="circuits">Списък от токови кръгове, използван за изчисления и логика.</param>
    Private Sub DrawBusbars(acDoc As Document, acCurDb As Database, basePoint As Point3d, circuits As List(Of strTokow))
        Try
            ' 1️ ИЗЧИСЛЯВАНЕ НА ОСНОВНИТЕ РАЗМЕРИ
            Dim brColums As Integer = circuits.Count - 1
            Dim X_Start As Double = basePoint.X + widthText + widthTextDim
            Dim X_End As Double = basePoint.X + widthText + widthTextDim + brColums * widthColom + widthColom / 2
            Dim Y_Shina As Double = basePoint.Y + Y_Шина
            ' Брой токови кръгове, които принадлежат към първата шина.
            brTokKrygoweNa6ina = circuits.Where(Function(c) c.Шина = True).Count()
            Dim Faza_Първа_шина = circuits.Any(Function(c) c.Фаза = "L1" Or c.Фаза = "L2" Or c.Фаза = "L3")
            Dim circuitOBSTO = circuits.FirstOrDefault(Function(c) c.ТоковКръг = "ОБЩО")
            Dim Faza_Втора_шина = circuitOBSTO.Фаза
            ' 3️ ТЕКСТ ЗА ФАЗИТЕ НА ШИНАТА
            Dim phaseText As String = IIf(Faza_Първа_шина, "L1,L2,L3,N,PE", "L,N,PE")
            ' 4️ ЧЕРТАЕНЕ НА ШИНИТЕ
            Dim X_Split As Double = 0
            Dim X_SecondStart As Double = 0
            Dim X_SecondEnd As Double = 0
            If Not twoBus Then
                ' ----- ЕДНА ШИНА -----
                cu.DrowLine(New Point3d(X_Start, Y_Shina, 0),
                        New Point3d(X_End, Y_Shina, 0),
                        "EL_ТАБЛА", Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight070, "ByLayer")
                cu.InsertText(Faza_Втора_шина & ",N,PE",
                          New Point3d(X_Start, Y_Shina + 2 * padingText, 0),
                          "EL__DIM", heightText,
                          TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
                X_SecondStart = X_Start
                X_SecondEnd = X_End
            Else
                ' ----- ДВЕ ШИНИ -----
                X_Split = X_Start + brTokKrygoweNa6ina * widthColom - widthColom / 2
                ' Чертае първата (лява) шина.
                cu.DrowLine(New Point3d(X_Start, Y_Shina, 0),
                        New Point3d(X_Split, Y_Shina, 0),
                        "EL_ТАБЛА", Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight070, "ByLayer")
                ' Начална позиция на втората (дясна) шина.
                X_SecondStart = X_Start + brTokKrygoweNa6ina * widthColom + widthColom / 2
                ' Край на втората шина (същият като X_End).
                X_SecondEnd = basePoint.X + widthText + widthTextDim + (brColums - 1) * widthColom + widthColom / 2
                ' Чертае втората шина.
                cu.DrowLine(New Point3d(X_SecondStart, Y_Shina, 0),
                        New Point3d(X_SecondEnd, Y_Shina, 0),
                        "EL_ТАБЛА", Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight070, "ByLayer")
                ' Надпис за първата шина (зависи от наличието на трифазни товари).
                cu.InsertText(phaseText,
                          New Point3d(X_Start, Y_Shina + 2 * padingText, 0),
                          "EL__DIM", heightText,
                          TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
                ' Надпис за втората шина (взет от "ОБЩО").
                cu.InsertText(Faza_Втора_шина & ",N,PE",
                          New Point3d(X_SecondStart, Y_Shina + 2 * padingText, 0),
                          "EL__DIM", heightText,
                          TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
                ' Чертае връзка между двете шини (между разединители).
                ' Позиционирана е над основните шини (+95 по Y).
                Dim X_6ina1 As Double = (X_SecondStart + X_End) / 2
                Dim X_6ina2 As Double = (X_Start + X_Split) / 2
                cu.DrowLine(New Point3d(X_6ina1, Y_Shina + 95, 0),
                        New Point3d(X_6ina2, Y_Shina + 95, 0),
                        "EL_ТАБЛА", Autodesk.AutoCAD.DatabaseServices.LineWeight.ByLayer, "ByLayer")
                ' Вмъква разединител в средата на шината и попълва атрибутите
                Dim circuit = circuits.FirstOrDefault(Function(c) c.Device = "Разединител")
                If circuit Is Nothing Then Return

                Dim X_disconn As Double = (X_Start + X_Split) / 2
                Dim Y_disconn As Double = Y_Shina + 95
                Dim blkRecId As ObjectId = cu.InsertBlock("s_i_ng_switch_disconn",
                                               New Point3d(X_disconn, Y_disconn, 0),
                                               "EL_ТАБЛА",
                                               New Scale3d(5, 5, 5))
                If Not blkRecId.IsNull Then
                    Using trans As Transaction = acCurDb.TransactionManager.StartTransaction()
                        Dim acBlkRef As BlockReference = DirectCast(trans.GetObject(blkRecId, OpenMode.ForWrite), BlockReference)
                        For Each objID As ObjectId In acBlkRef.AttributeCollection
                            Dim acAttRef As AttributeReference = DirectCast(trans.GetObject(objID, OpenMode.ForWrite), AttributeReference)
                            Select Case acAttRef.Tag
                                Case "1" : acAttRef.TextString = ""
                                Case "2" : acAttRef.TextString = circuit.Брой_Полюси & "p"
                                Case "3" : acAttRef.TextString = circuit.Breaker_Номинален_Ток & "A"
                                Case "4" : acAttRef.TextString = ""
                                Case "5" : acAttRef.TextString = ""
                                Case "SHORTNAME" : acAttRef.TextString = circuit.Breaker_Тип_Апарат
                                Case "REFNB" : acAttRef.TextString = circuit.Tablo
                                Case "DESIGNATION" : acAttRef.TextString = ""
                            End Select
                        Next
                        trans.Commit()
                    End Using
                End If
            End If
            Dim circuit_Общо = circuits.FirstOrDefault(Function(c) c.Device = "Табло")
            Dim X_Общо As Double = (X_SecondStart + X_SecondEnd) / 2
            Dim Y_Общо As Double = Y_Shina + 95
            Dim blkRecId_Общо As ObjectId = cu.InsertBlock("s_i_ng_switch_disconn",
                               New Point3d(X_Общо, Y_Общо, 0),
                               "EL_ТАБЛА",
                               New Scale3d(5, 5, 5))
            If Not blkRecId_Общо.IsNull Then
                Using trans As Transaction = acCurDb.TransactionManager.StartTransaction()
                    Dim acBlkRef As BlockReference = DirectCast(trans.GetObject(blkRecId_Общо, OpenMode.ForWrite), BlockReference)
                    For Each objID As ObjectId In acBlkRef.AttributeCollection
                        Dim acAttRef As AttributeReference = DirectCast(trans.GetObject(objID, OpenMode.ForWrite), AttributeReference)
                        Select Case acAttRef.Tag
                            Case "1" : acAttRef.TextString = ""
                            Case "2" : acAttRef.TextString = circuit_Общо.Брой_Полюси & "p"
                            Case "3" : acAttRef.TextString = circuit_Общо.Breaker_Номинален_Ток & "A"
                            Case "4" : acAttRef.TextString = ""
                            Case "5" : acAttRef.TextString = ""
                            Case "SHORTNAME" : acAttRef.TextString = circuit_Общо.Breaker_Тип_Апарат
                            Case "REFNB" : acAttRef.TextString = circuit_Общо.Tablo
                            Case "DESIGNATION" : acAttRef.TextString = ""
                        End Select
                    Next
                    trans.Commit()
                End Using
                ' Чертае вертикална линия над прекъсвача.
                cu.DrowLine(New Point3d(X_Общо, Y_Общо, 0),
                        New Point3d(X_Общо, Y_Общо + 125, 0),
                        "EL_ТАБЛА", Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight070, "ByLayer")

                Dim blkRecId_Текст As ObjectId = cu.InsertBlock("Кабел",
                                   New Point3d(X_Общо, Y_Общо + 90, 0),
                                   "EL__DIM",
                                   New Scale3d(1, 1, 1))
                Using trans As Transaction = acCurDb.TransactionManager.StartTransaction()
                    Dim acBlkRef As BlockReference = DirectCast(trans.GetObject(blkRecId_Текст, OpenMode.ForWrite), BlockReference)
                    For Each objID As ObjectId In acBlkRef.AttributeCollection
                        Dim acAttRef As AttributeReference = DirectCast(trans.GetObject(objID, OpenMode.ForWrite), AttributeReference)
                        Select Case acAttRef.Tag
                            Case "NA4IN_0" : acAttRef.TextString = circuit_Общо.Кабел_Сечение
                            Case "NA4IN_1" : acAttRef.TextString = "от табло " & circuit_Общо.Табло_Родител
                            Case "NA4IN_2" : acAttRef.TextString = ""
                            Case "NA4IN_3" : acAttRef.TextString = ""
                            Case "NA4IN_4" : acAttRef.TextString = ""
                            Case "NA4IN_5" : acAttRef.TextString = ""
                            Case "NA4IN_6" : acAttRef.TextString = ""
                            Case "NA4IN_7" : acAttRef.TextString = ""
                            Case "NA4IN_8" : acAttRef.TextString = ""
                            Case "NA4IN_9" : acAttRef.TextString = ""
                            Case "NA4IN_10" : acAttRef.TextString = ""
                        End Select
                    Next
                    Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection
                    For Each prop As DynamicBlockReferenceProperty In props
                        If prop.PropertyName = "Visibility" Then prop.Value = "Точка"
                    Next
                    trans.Commit()
                End Using
                cu.EditDynamicBlockReferenceKabel(blkRecId_Текст)
            End If
        Catch ex As Exception
            ' Показва съобщение при възникване на грешка,
            ' включително текста на грешката и stack trace.
            ' Полезно при дебъг, но не е подходящо за production среда.
            MsgBox("Възникна грешка:  " &
               ex.Message &
               vbCrLf & vbCrLf &
               ex.StackTrace.ToString)
        End Try
    End Sub
    ''' <summary>
    ''' Чертaе всички токови кръгове в таблото.
    ''' Логиката:
    ''' 1. Изчислява началните координати
    ''' 2. Обхожда всички токови кръгове
    ''' 3. Изчертава текстовете за всеки кръг
    ''' 4. Вмъква прекъсвач и управляващи елементи
    ''' 5. Чертaе свързващите линии
    ''' </summary>
    Private Sub DrawCircuits(acDoc As Document, acCurDb As Database, basePoint As Point3d, circuits As List(Of strTokow))
        ' Изчислява общия брой колони.
        ' Ако има двойна шина (twoBus=True), една колона се резервира и не участва.
        Dim brColums As Integer = circuits.Count - If(twoBus, 1, 0)
        ' Начална X координата след текстовата зона
        Dim X_Start As Double = basePoint.X + widthText + widthTextDim
        ' Y координата на шината
        Dim Y_Shina As Double = basePoint.Y + Y_Шина
        Try
            ' Индекс на текущата колона
            Dim colIndex As Integer = 0
            ' Обхождаме всички токови кръгове
            For Each circuit As strTokow In circuits
                ' Пропускаме специалните кръгове тип "Разединител"
                If circuit.Device = "Разединител" Then Continue For
                ' Изчисляване на X позицията за текущия кръг
                Dim X As Double =
                X_Start + colIndex * widthColom + widthColom / 2
                ' Чертaе текстовата информация за токовия кръг
                DrawCircuitTexts(acDoc, acCurDb, basePoint, circuit, X)
                ' Ако е запис "Табло" → пишем само текстовете
                ' Не се чертаят прекъсвачи и линии
                If circuit.Device = "Табло" Then Continue For
                ' Вмъква блок за прекъсвач
                DrawBreakerBlock(acDoc, acCurDb, basePoint, circuit, X, Y_Shina)
                ' Чертaе управляващо устройство (ако има)
                DrawControlDevice(acDoc, acCurDb, circuit, X, Y_Shina)
                ' Чертaе вертикалните линии за кръга
                DrawCircuitLines(X, circuit, Y_Shina)
                ' Преминава към следващата колона
                colIndex += 1
            Next
        Catch ex As Exception
            MsgBox("Възникна грешка: " &
               vbCrLf &
               ex.Message &
               vbCrLf &
               vbCrLf &
               ex.StackTrace,
               MsgBoxStyle.Critical)
        End Try
    End Sub
    ''' <summary>
    ''' Функцията Calculate_GV2 избира подходящ моторен прекъсвач (тип GV2)
    ''' на база вече изчислен ток.
    ''' 
    ''' Логиката включва:
    ''' 1. Преобразуване на входния ток от текст към число
    ''' 2. Търсене на съвпадение в база данни (GV_Database)
    ''' 3. Връщане на конкретна информация според параметъра "Връща"
    ''' 
    ''' Функцията НЕ изчислява ток – очаква той да е подаден отвън.
    ''' Това я прави по-гъвкава и независима от начина на изчисление.
    ''' </summary>
    ''' <param name="Ток">
    ''' Ток като текст (например "10", "10.5", "10,5").
    ''' Допуска се използване на запетая или точка като десетичен разделител.
    ''' </param>
    ''' <param name="Връща">
    ''' Определя какъв резултат да бъде върнат:
    ''' 1 → Тип на защитата (например GV2-ME)
    ''' 2 → Мощност по каталог (при 400V)
    ''' 3 → Диапазон на настройка
    ''' </param>
    ''' <returns>
    ''' Връща String със съответния резултат или съобщение:
    ''' - "N/A" при невалиден ток
    ''' - "Out of range (...A)" ако няма подходящ апарат
    ''' - "Грешен параметър" при невалиден вход за "Връща"
    ''' </returns>
    Private Function Calculate_GV2(Ток As String, Връща As Integer) As String
        ' =====================================================
        ' 1️ ПРЕОБРАЗУВАНЕ НА ВХОДНИЯ ТОК
        ' =====================================================
        ' Замяна на запетая с точка, за да се осигури коректно
        ' преобразуване към числов тип.
        Dim I_val As String = Ток.Replace(",", ".")
        ' Преобразуване на текстовата стойност към Double.
        ' Val извлича числото от началото на низа.
        Dim I_double As Double = Val(I_val)
        ' Проверка за невалиден или нулев ток.
        If I_double <= 0 Then Return "N/A"
        ' =====================================================
        ' 2️ ТЪРСЕНЕ В БАЗАТА ДАННИ
        ' =====================================================
        ' Търсене на първия запис в GV_Database,
        ' при който токът попада в диапазона:
        ' MinCurrent ≤ I_double ≤ MaxCurrent
        Dim match = GV_Database.FirstOrDefault(Function(x) I_double >= x.MinCurrent And I_double <= x.MaxCurrent)
        ' Ако няма намерен подходящ апарат,
        ' връщаме информация за тока.
        If match Is Nothing Then Return "Out of range (" & I_double.ToString("F2") & "A)"
        ' =====================================================
        ' 3️ ВРЪЩАНЕ НА РЕЗУЛТАТ
        ' =====================================================
        ' В зависимост от параметъра "Връща",
        Select Case Връща
            Case 1 : Return match.Type               ' Връща типа на апарата (например GV2-ME).
            Case 2 : Return match.MotorPower         ' Връща мощността по каталог (при 400V).
            Case 3 : Return match.SettingRange       ' Връща диапазона на настройка на тока.
            Case Else : Return "Грешен параметър"    ' Невалиден параметър "Връща".
        End Select
    End Function
    ''' <summary>
    ''' Чертaе вертикална линия за токов кръг в таблото.
    ''' Позицията на линията зависи от:
    ''' - наличието на управление
    ''' - типа на устройството (напр. "Контакт")
    ''' </summary>
    ''' <param name="X">Х координата на линията</param>
    ''' <param name="circuit">Обект с данни за токовия кръг</param>
    ''' <param name="Y_Shina">Y координата на шината (референтна точка)</param>
    Public Sub DrawCircuitLines(X As Double, circuit As strTokow, Y_Shina As Double)
        ' Ако резерва не чертаем линия
        If circuit.Device = "Резерва" Then Return
        ' 1️ ДЕФИНИРАНЕ НА НАЧАЛНИ КООРДИНАТИ
        ' Начална позиция по Y (горна точка на линията)
        Dim startY As Double = Y_Shina - 135
        ' Крайна позиция по Y (долна точка на линията)
        Dim endY As Double = Y_Shina - 370
        ' 2️ ПРОВЕРКА ЗА УПРАВЛЕНИЕ
        ' Ако има управление (различно от Nothing, празно или "Няма"),
        ' линията се измества надолу (правим място за управление)
        If Not String.IsNullOrEmpty(circuit.Управление) AndAlso circuit.Управление <> "Няма" Then startY -= 95
        ' 3️⃣ ПРОВЕРКА ЗА ТИП УСТРОЙСТВО
        ' Ако устройството е "Контакт",
        ' използваме фиксирана позиция (override на startY)
        If circuit.Device = "Контакт" Then startY = Y_Shina - 253
        ' 4️ СЪЗДАВАНЕ НА ТОЧКИ
        ' Начална точка на линията
        Dim startPt As New Point3d(X, startY, 0)
        ' Крайна точка на линията
        Dim endPt As New Point3d(X, endY, 0)
        ' 5️ ЧЕРТАНЕ НА ЛИНИЯТА
        ' Чертaе линията в слой "EL_ТАБЛА"
        ' с настройки по слой (ByLayer)
        cu.DrowLine(startPt, endPt,
                "EL_ТАБЛА",
                DatabaseServices.LineWeight.ByLayer,
                "ByLayer")
    End Sub
    ''' <summary>
    ''' Чертае управляващо устройство под прекъсвача
    ''' </summary>
    ''' <param name="acDoc">AutoCAD документ</param>
    ''' <param name="acCurDb">AutoCAD база данни</param>
    ''' <param name="circuit">Токов кръг</param>
    ''' <param name="X">X координата (център на колоната)</param>
    ''' <param name="breakerY">Y позиция на прекъсвача</param>
    Private Sub DrawControlDevice(acDoc As Document, acCurDb As Database,
                                  circuit As strTokow, X As Double, breakerY As Double)
        Try
            ' =====================================================
            ' 1️ ПРОВЕРКА ДАЛИ ИМА УПРАВЛЕНИЕ
            ' =====================================================
            If String.IsNullOrEmpty(circuit.Управление) Then Return
            If circuit.Управление = "Няма" Then Return
            ' =====================================================
            ' 2️ НАМИРАНЕ НА БЛОКА ОТ РЕЧНИКА
            ' =====================================================
            If Not ControlBlockMap.ContainsKey(circuit.Управление) Then Return
            Dim blockName As String = ControlBlockMap(circuit.Управление)
            If String.IsNullOrEmpty(blockName) Then Return
            ' =====================================================
            ' 3️ ИЗЧИСЛЯВАНЕ НА ПОЗИЦИЯТА (ВИНАГИ под прекъсвача)
            ' =====================================================
            Dim controlY As Double = breakerY - 135
            Dim insertPoint As New Point3d(X, controlY, 0)
            Dim blockScale As New Scale3d(5, 5, 5)
            ' =====================================================
            ' 4️ ПОЛУЧАВАНЕ НА ПАРАМЕТРИТЕ ЗА ТОЗИ ТИП
            ' =====================================================
            Dim config As ControlDeviceConfig = GetControlDeviceConfig(circuit)
            ' =====================================================
            ' 5️ ВМЪКВАНЕ НА БЛОКА
            ' =====================================================
            Dim blkRecId As ObjectId = cu.InsertBlock(blockName, insertPoint, "EL_ТАБЛА", blockScale)
            ' =====================================================
            ' 6️ ПОПЪЛВАНЕ НА АТРИБУТИТЕ
            ' =====================================================
            If Not blkRecId.IsNull Then
                Using trans As Transaction = acCurDb.TransactionManager.StartTransaction()
                    Dim acBlkRef As BlockReference = DirectCast(trans.GetObject(blkRecId, OpenMode.ForWrite), BlockReference)
                    For Each objID As ObjectId In acBlkRef.AttributeCollection
                        Dim acAttRef As AttributeReference = DirectCast(trans.GetObject(objID, OpenMode.ForWrite), AttributeReference)
                        Select Case acAttRef.Tag
                            Case "1" : acAttRef.TextString = config.Str_1
                            Case "2" : acAttRef.TextString = config.Str_2
                            Case "3" : acAttRef.TextString = config.Str_3
                            Case "4" : acAttRef.TextString = config.Str_4
                            Case "5" : acAttRef.TextString = ""
                            Case "SHORTNAME" : acAttRef.TextString = config.ShortName
                            Case "REFNB" : acAttRef.TextString = circuit.Tablo
                            Case "DESIGNATION" : acAttRef.TextString = ""
                        End Select
                    Next
                    Dim kvadrat As Boolean = True
                    If kvadrat Then
                        Dim Y_kvadrat As Double = controlY - 195
                        cu.InsertBlock("Ключ_квадрат",
                                       New Point3d(X - 32, Y_kvadrat, 0),
                                       "EL_ТАБЛА",
                                       New Scale3d(1, 1, 1))
                        cu.DrowLine(New Point3d(X - 32,
                                                Y_kvadrat + 25,
                                                0),
                                    New Point3d(X - 32,
                                                Y_kvadrat + 133,
                                                0),
                                    "EL_ТАБЛА",
                                    DatabaseServices.LineWeight.ByLayer,
                                    "ByLayer")
                    End If
                    trans.Commit()
                End Using
            End If
        Catch ex As Exception
            MsgBox("Възникна грешка: " & vbCrLf & ex.Message & vbCrLf & vbCrLf & ex.StackTrace, MsgBoxStyle.Critical)
        End Try
    End Sub
    ''' <summary>
    ''' Връща конфигурацията за даден тип управление
    ''' </summary>
    Private Function GetControlDeviceConfig(circuit As strTokow) As ControlDeviceConfig
        ' New(str_1 , str_2 , str_3 , str_4 , str_5 , shortName)
        Select Case circuit.Управление
            Case "Импулсно реле"
                Return New ControlDeviceConfig("", "1p", If(circuit.Ток * 1.1 > 16, "32A", "16A"), "220VAC", "", "iTL")
            Case "Контактор"
                Return New ControlDeviceConfig(circuit.Брой_Полюси.ToString & "НО", "", If(circuit.Ток * 1.1 > 16, "25A", "16A"), "220VAC", "", "iCT")
                'Case "Моторна защита"
                '    Return New ControlDeviceConfig("3NO", "1NO+1NC", "9A", "230VAC", "LC1D")
                'Case "Моторен механизъм"
                '    Return New ControlDeviceConfig("", "", "", "", "NS100")
                'Case "Честотен регулатор"
                '    Return New ControlDeviceConfig("", "", "", "", "ATV")
            Case "Стълбищен автомат"
                Return New ControlDeviceConfig("", "", "0.5-20min",
                                               If(circuit.Ток * 1.1 > 16, "---", "16A"),
                                               "", "MINp")
                'Case "Електромер"
                '    Return New ControlDeviceConfig("", "", "", "", "kWh")
            Case "Фото реле"
                Return New ControlDeviceConfig("", "2-100 Lx",
                                               If(circuit.Ток * 1.1 > 10, "---", "10A"),
                                               "", "", "IC100")
            Case Else
                Return New ControlDeviceConfig("", "", "", "", "", "")
        End Select
    End Function
    ''' <summary>
    ''' Процедурата DrawBreakerBlock добавя визуално представяне (блок) на автоматичен прекъсвач или RCD в AutoCAD документ.
    ''' Използва се в контекста на проектиране на електрически табла и схеми на токови кръгове.
    ''' </summary>
    ''' <param name="acDoc">Документът на AutoCAD, в който се добавя блокът.</param>
    ''' <param name="acCurDb">Текущата база данни на AutoCAD, за извършване на транзакции и операции върху блокове.</param>
    ''' <param name="basePoint">Базова точка (не се използва директно в тази процедура, но може да е за разширения).</param>
    ''' <param name="circuit">Обект от тип strTokow, който съдържа данните за конкретния токов кръг, като тип апарат, RCD информация, брой полюси и др.</param>
    ''' <param name="X">X координата за позициониране на блока.</param>
    ''' <param name="Y_Shina">Y координата на шината, върху която се поставя блокът.</param>
    Private Sub DrawBreakerBlock(acDoc As Document, acCurDb As Database, basePoint As Point3d,
                                circuit As strTokow, X As Double, Y_Shina As Double)
        ' Име на блока по подразбиране – стандартен прекъсвач C60
        Dim blockName As String = "s_c60_circ_break"
        ' Масштаб на блока – зададен като 5х5х5.
        Dim blockScale As New Scale3d(5, 5, 5)
        ' Начална позиция за поставяне на блока
        Dim insertPoint As New Point3d(X, Y_Shina, 0)
        ' Флаг, който указва какво представлява блокът - за специално попълване на атрибути
        Dim rcd_Yes As String = ""
        ' Ако има RCD_Нула и тя не е "N", местим блока надолу по Y с 117.5 единици
        If Not String.IsNullOrEmpty(circuit.RCD_Нула) AndAlso
                    circuit.RCD_Нула.Trim().ToUpper() <> "N" Then
            insertPoint = New Point3d(X, Y_Shina - 117.5, 0)
        End If
        ' Избор на блок според типа апарат
        ' 1. Първоначални настройки
        rcd_Yes = "Прекъсвач"
        blockName = "s_c60_circ_break"
        ' 2. Логика за избор на тип апарат и блок
        Select Case True
    ' Първи приоритет: Резерви
            Case circuit.Device = "Резерва"
                rcd_Yes = "Резерва"
                blockName = "s_c60_circ_break"
    ' Първи приоритет: Резерви
            Case circuit.Device = "Съществуващ"
                rcd_Yes = "Съществуващ"
                blockName = "s_c60_circ_break"
    ' Втори приоритет: Моторна защита
            Case circuit.Управление = "Моторна защита"
                rcd_Yes = "Моторна защита"
                blockName = "s_GV2"
    ' Трети приоритет: Проверка за RCD (ако има попълнен тип)
            Case Not String.IsNullOrWhiteSpace(circuit.RCD_Тип)
                rcd_Yes = "RCD"
                blockName = "s_dpnn_vigi_circ_break"
                ' Всичко останало (Default)
            Case Else
                rcd_Yes = "Прекъсвач"
                blockName = "s_c60_circ_break"
        End Select
        ' Вмъкване на блока в AutoCAD с помощта на функция InsertBlock
        ' (предполага се, че cu е помощен модул/клас за CAD операции)
        Dim blkRecId As ObjectId = cu.InsertBlock(blockName, insertPoint, "EL_ТАБЛА", blockScale)
        ' Ако вмъкването е успешно (ObjectId не е Null)
        If Not blkRecId.IsNull Then
            ' Стартиране на транзакция за промяна на атрибутите на блока
            Try
                Using trans As Transaction = acCurDb.TransactionManager.StartTransaction()
                    ' Получаваме референция към блока за писане
                    Dim acBlkRef As BlockReference = DirectCast(trans.GetObject(blkRecId, OpenMode.ForWrite), BlockReference)
                    ' Обхождане на всички атрибути на блока
                    For Each objID As ObjectId In acBlkRef.AttributeCollection
                        Dim acAttRef As AttributeReference = DirectCast(trans.GetObject(objID, OpenMode.ForWrite), AttributeReference)
                        ' Попълване на атрибутите в зависимост дали е RCD или обикновен прекъсвач
                        If acAttRef.Tag = "DESIGNATION" Then acAttRef.TextString = ""
                        If acAttRef.Tag = "REFNB" Then acAttRef.TextString = circuit.Tablo
                        Select Case rcd_Yes
                            Case "Съществуващ"
                                Select Case acAttRef.Tag
                                    Case "1" : acAttRef.TextString = ""
                                    Case "2" : acAttRef.TextString = ""
                                    Case "3" : acAttRef.TextString = ""
                                    Case "4" : acAttRef.TextString = circuit.Breaker_Номинален_Ток
                                    Case "5" : acAttRef.TextString = ""
                                    Case "SHORTNAME" : acAttRef.TextString = ""
                                End Select
                            Case "Резерва"
                                Select Case acAttRef.Tag
                                    Case "1" : acAttRef.TextString = ""
                                    Case "2" : acAttRef.TextString = ""
                                    Case "3" : acAttRef.TextString = ""
                                    Case "4" : acAttRef.TextString = circuit.Breaker_Номинален_Ток
                                    Case "5" : acAttRef.TextString = ""
                                    Case "SHORTNAME" : acAttRef.TextString = ""
                                End Select
                            Case "Моторна защита"
                                Select Case acAttRef.Tag
                                    Case "1" : acAttRef.TextString = Calculate_GV2(circuit.Ток, 3)
                                    Case "2" : acAttRef.TextString = "3P"
                                    Case "3" : acAttRef.TextString = Calculate_GV2(circuit.Ток, 2)
                                    Case "4" : acAttRef.TextString = ""
                                    Case "5" : acAttRef.TextString = ""
                                    Case "SHORTNAME" : acAttRef.TextString = Calculate_GV2(circuit.Ток, 1)
                                End Select
                            Case "RCD"
                                ' Атрибути за RCD
                                Select Case acAttRef.Tag
                                    Case "SHORTNAME" : acAttRef.TextString = circuit.RCD_Тип
                                    Case "1" : acAttRef.TextString = circuit.RCD_Клас
                                    Case "2" : acAttRef.TextString = circuit.Брой_Полюси & "p"
                                    Case "3" : acAttRef.TextString = "C"
                                    Case "4" : acAttRef.TextString = circuit.RCD_Ток & "A"
                                    Case "5" : acAttRef.TextString = circuit.RCD_Чувствителност & "mA"
                                End Select
                            Case "Прекъсвач"
                                ' Атрибути за прекъсвач
                                Select Case acAttRef.Tag
                                    Case "SHORTNAME" : acAttRef.TextString = circuit.Breaker_Тип_Апарат
                                    Case "2" : acAttRef.TextString = circuit.Breaker_Крива
                                    Case "3" : acAttRef.TextString = circuit.Брой_Полюси & "p"
                                    Case "4" : acAttRef.TextString = circuit.Breaker_Номинален_Ток & "A"
                                End Select
                        End Select
                    Next
                    ' Потвърждаваме промяната на атрибутите
                    trans.Commit()
                End Using
            Catch ex As Exception

            End Try
        End If
    End Sub
    ''' <summary>
    ''' Чертaе текстовата информация за един токов кръг в таблицата на таблото.
    ''' </summary>
    ''' <param name="acDoc">Текущият AutoCAD документ.</param>
    ''' <param name="acCurDb">Текущата база данни.</param>
    ''' <param name="basePoint">Начална точка на таблицата.</param>
    ''' <param name="circuit">Обект с данни за токовия кръг.</param>
    ''' <param name="X">X координата на колоната за съответния кръг.</param>
    ''' <remarks>
    ''' Процедурата позиционира и изчертава всички текстове за даден токов кръг
    ''' в съответната колона на таблицата.
    '''
    ''' Особености:
    ''' - Използва центрирано подравняване за числови и кратки стойности
    ''' - Използва ляво подравняване за текстови описания
    ''' - При нулеви стойности (лампи/контакти) показва "----"
    ''' - Всички координати се изчисляват спрямо basePoint
    '''
    ''' Възможни подобрения:
    ''' - Проверка за празни/null стойности (Nothing)
    ''' - Форматиране на текста спрямо дължината (truncate/auto-scale)
    ''' - Унифициране на височината на текста (в момента има 12 и heightText)
    ''' </remarks>
    Private Sub DrawCircuitTexts(acDoc As Document, acCurDb As Database, basePoint As Point3d,
                             circuit As strTokow, X As Double)
        Dim Y_Base As Double = basePoint.Y
        Dim textLayer As String = "EL__DIM"
        ' Токов кръг (ред 1)
        cu.InsertText(circuit.ТоковКръг,
                  New Point3d(X + padingText, Y_Base + 9 * heightRow + heightRow / 2, 0),
                  textLayer, heightText, TextHorizontalMode.TextMid, TextVerticalMode.TextBase)
        ' Брой лампи (ред 2)
        cu.InsertText(IIf(circuit.brLamp = 0, "----", circuit.brLamp.ToString()),
                  New Point3d(X + padingText, Y_Base + 8 * heightRow + heightRow / 2, 0),
                  textLayer, heightText, TextHorizontalMode.TextMid, TextVerticalMode.TextBase)
        ' Брой контакти (ред 3)
        cu.InsertText(IIf(circuit.brKontakt = 0, "----", circuit.brKontakt.ToString()),
                  New Point3d(X + padingText, Y_Base + 7 * heightRow + heightRow / 2, 0),
                  textLayer, heightText, TextHorizontalMode.TextMid, TextVerticalMode.TextBase)
        ' Мощност (ред 4)
        cu.InsertText(circuit.Мощност.ToString("0.000"),
                  New Point3d(X + padingText, Y_Base + 6 * heightRow + heightRow / 2, 0),
                  textLayer, heightText, TextHorizontalMode.TextMid, TextVerticalMode.TextBase)
        ' Тип кабел (ред 5)
        cu.InsertText(circuit.Кабел_Тип,
                  New Point3d(X + padingText, Y_Base + 5 * heightRow + heightRow / 2, 0),
                  textLayer, heightText, TextHorizontalMode.TextMid, TextVerticalMode.TextBase)
        ' Сечение кабел (ред 6)
        cu.InsertText(circuit.Кабел_Сечение,
                  New Point3d(X + padingText, Y_Base + 4 * heightRow + heightRow / 2, 0),
                  textLayer, heightText, TextHorizontalMode.TextMid, TextVerticalMode.TextBase)
        ' Фаза (ред 7)
        cu.InsertText(circuit.Фаза,
                  New Point3d(X + padingText, Y_Base + 3 * heightRow + heightRow / 2, 0),
                  textLayer, heightText, TextHorizontalMode.TextMid, TextVerticalMode.TextBase)
        ' Консуматор (ред 8) - ляво подравнен
        cu.InsertText(circuit.Консуматор,
                  New Point3d(X - widthColom / 2 + padingText, Y_Base + 2 * heightRow + (heightRow - heightText) / 2, 0),
                  textLayer, 12, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
        ' Предназначение (ред 9) - ляво подравнен
        cu.InsertText(circuit.предназначение,
                  New Point3d(X - widthColom / 2 + padingText, Y_Base + 1 * heightRow + (heightRow - heightText) / 2, 0),
                  textLayer, 12, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
    End Sub
    Private Sub DrawMainSwitch(acDoc As Document, acCurDb As Database, basePoint As Point3d, circuits As List(Of strTokow))
        ' Тук ще чертаем главния прекъсвач/разединител
    End Sub
    ''' <summary>
    ''' Чертaе заземителната схема към главното табло.
    ''' Логиката:
    ''' 1. Проверява дали таблото е Гл.Р.Т.
    ''' 2. Чертaе връзката към заземителя
    ''' 3. Добавя текст за съпротивление на заземяване
    ''' 4. Вмъква динамичен блок "Заземление"
    ''' 5. Настройва параметрите и атрибутите на блока
    ''' 6. Добавя означение за PE проводник
    ''' </summary>
    Private Sub DrawGrounding(acDoc As Document, acCurDb As Database, X As Double, ptbasePoint As Point3d, panelName As String)
        ' Чертaем заземление само за главно разпределително табло
        If panelName <> "Гл.Р.Т." AndAlso panelName <> "ГлРТ" Then Return
        ' Хоризонтална линия към заземителя
        cu.DrowLine(
                    New Point3d(X, ptbasePoint.Y + Y_Шина, 0),
                    New Point3d(X - widthColom, ptbasePoint.Y + Y_Шина, 0),
                    "EL_ТАБЛА",
                    Autodesk.AutoCAD.DatabaseServices.LineWeight.ByLayer,
                    "ByLayer"
                    )
        ' Текст за съпротивление на заземяване
        cu.InsertText(
                "R<30Ω",
                New Point3d(X - widthColom,
                            ptbasePoint.Y + Y_Шина + 2 * padingText,
                            0),
                "EL__DIM",
                heightText,
                TextHorizontalMode.TextLeft,
                TextVerticalMode.TextBase
                )
        ' Вмъкване на блока "Заземление"
        Dim blkRecId =
        cu.InsertBlock(
            "Заземление",
            New Point3d(X - widthColom,
                         ptbasePoint.Y + Y_Шина,
                         0),
            "EL_ТАБЛА",
            New Scale3d(0.21, 0.21, 0.21)
        )
        ' Взимаме активния документ
        Dim doc As Document =
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Using trans As Transaction = doc.TransactionManager.StartTransaction()
            ' Взимаме BlockTable
            Dim acBlkTbl As BlockTable =
            trans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            ' Взимаме BlockReference на вмъкнатия блок
            Dim acBlkRef As BlockReference =
            DirectCast(
                trans.GetObject(blkRecId, OpenMode.ForWrite),
                BlockReference
            )
            ' Достъп до динамичните параметри на блока
            Dim props As DynamicBlockReferencePropertyCollection =
            acBlkRef.DynamicBlockReferencePropertyCollection
            ' Настройване на параметрите на динамичния блок
            For Each prop As DynamicBlockReferenceProperty In props
                Select Case prop.PropertyName
                    Case "Visibility"
                        ' Настройка на Visibility State
                        prop.Value = "Заземител-БЕЗ контролна клема"
                    Case "Position1 X"
                        ' Настройка на X позиция
                        prop.Value = -10.0
                    Case "Position1 Y"
                        ' Настройка на Y позиция
                        prop.Value = -80.0
                    Case "Angle1"
                        ' Настройка на ъгъл
                        prop.Value = 0.0
                End Select
            Next
            ' Достъп до атрибутите на блока
            Dim attCol As AttributeCollection =
            acBlkRef.AttributeCollection
            ' Попълване на атрибути
            For Each objID As ObjectId In attCol
                Dim dbObj As DBObject =
                trans.GetObject(objID, OpenMode.ForWrite)
                Dim acAttRef As AttributeReference = dbObj
                ' Попълване на атрибут "ТАБЛО"
                If acAttRef.Tag = "ТАБЛО" Then acAttRef.TextString = "2к"
            Next
            ' Запис на промените
            trans.Commit()
        End Using
        ' Добавяне на означение за защитен проводник
        cu.InsertText(
                "PE",
                New Point3d(
                    X - widthColom + 3 * padingText,
                    ptbasePoint.Y + Y_Шина - heightText - padingText,
                    0
                ),
                "EL__DIM",
                heightText,
                TextHorizontalMode.TextLeft,
                TextVerticalMode.TextBase
                )
    End Sub
    ''' <summary>
    ''' Процедурата DrawAnnotations създава текстови анотации (бележки) в AutoCAD чертеж,
    ''' свързани с електрическо табло, както и изчислява и визуализира общия брой полюси
    ''' на всички токови кръгове.
    ''' 
    ''' Използва се в контекста на автоматизирано чертане на табла, където освен графиката
    ''' е необходимо да се добавят и нормативни указания и обобщена информация.
    ''' </summary>
    ''' <param name="basePoint">
    ''' Базова точка за позициониране на текстовете. Всички анотации се разполагат
    ''' относително спрямо тази точка.
    ''' </param>
    ''' <param name="circuits">
    ''' Списък от токови кръгове (strTokow), използван за изчисляване на общия брой полюси.
    ''' </param>
    Private Sub DrawAnnotations(basePoint As Point3d, circuits As List(Of strTokow))
        Dim Zabelevka As String = "1. Таблото да се изпълни в съответствие с изискванията на БДС EN 61439-1."
        ' Добавяне на нови редове с допълнителни изисквания към таблото
        Zabelevka += vbCrLf & "2. Aпаратурата и тоководящите части да бъдат монтирани зад защитни капаци. "
        Zabelevka += vbCrLf & "3. Достъпа до палците и ръкохватките на комутационните апарати се осигурява посредством отвори в защитните капаци."
        Zabelevka += vbCrLf & "4. Апаратурата е избрана по каталога на SCHNEIDER ELECTRIC."
        Zabelevka += vbCrLf & "5. Изборът на автоматичните прекъсвачи е съобразен с токовете на к.с., спазени са изискванията за селективност."
        Zabelevka += vbCrLf & "6. При замяна типа на апаратурата да се преизчисли схемата."
        Zabelevka += vbCrLf & "7. При замяна номиналният ток на апаратурата да се преизчисли сечението на кабелите."
        cu.InsertMText("ЗАБЕЛЕЖКИ:",
                       New Point3d(basePoint.X,
                                   basePoint.Y - 20, 0),
                       "EL__DIM", 10)
        cu.InsertMText(Zabelevka,
                       New Point3d(basePoint.X + 30,
                                   basePoint.Y - 20 - heightRow, 0),
                       "EL__DIM", 10)
        Dim pol As Integer = 0
        For Each circuit As strTokow In circuits
            pol += circuit.Брой_Полюси
            Select Case circuit.Управление
                Case "Няма", "", "Електромер",
                 "Честотен регулатор",
                 "Моторен механизъм"
                    ' При тези типове управление не се добавят допълнителни полюси.
                Case "Стълбищен автомат", "Импулсно реле"
                    pol += 1
                Case "Фото реле"
                    pol += 3
                Case "Контактор", "Моторна защита"
                    pol += circuit.Брой_Полюси
            End Select
            Select Case circuit.RCD_Полюси
                Case "2p"
                    pol += 2
                Case "4p"
                    pol += 4
            End Select
        Next
        cu.InsertMText("Полюси -> " & pol.ToString(0),
                       New Point3d(basePoint.X + 160,
                                   basePoint.Y + 900, 0),
                       "Defpoints", 20, 1)
    End Sub
    Private Sub ToolStripButton_ШИНА_Click(sender As Object, e As EventArgs) Handles ToolStripButton_ШИНА.Click
        ' 1. Вземи избраното табло
        Dim selectedTablo As String = TreeView1.SelectedNode?.Text
        If String.IsNullOrEmpty(selectedTablo) Then Return
        If selectedTablo.Contains("(") Then
            selectedTablo = selectedTablo.Substring(0, selectedTablo.IndexOf("(")).Trim()
        End If
        ' 2. Извикай процедурата за избор на разединител на шина
        UpdateDisconnectorRecord(selectedTablo)
        SetupDataGridView_Total()
        ' 3. Refresh на DataGridView
        FillDataGridViewForPanel()
    End Sub
    ''' <summary>
    ''' Добавя обобщен запис "ОБЩО" за всяко табло.
    ''' Логиката:
    ''' 1. Намира всички уникални табла
    ''' 2. За всяко табло събира всички кръгове
    ''' 3. Изчислява общи стойности (мощност, брой консуматори)
    ''' 4. Определя фаза и полюси
    ''' 5. Създава нов запис "ОБЩО"
    ''' 6. Изчислява ток, прекъсвач и кабел
    ''' </summary>
    Private Sub AddFeederRecords()
        ' 1️ НАМИРАНЕ НА ВСИЧКИ УНИКАЛНИ ТАБЛА
        ' Извлича всички уникални стойности на Tablo от ListTokow
        Dim allTablos As List(Of String) =
        ListTokow.Select(Function(t) t.Tablo).Distinct().ToList()
        ' 2️ ОБХОЖДАНЕ НА ВСЯКО ТАБЛО
        For Each tabloName As String In allTablos
            BuildPanelSummaryRecord(tabloName)
        Next
    End Sub
    ''' <summary>
    ''' Изгражда обобщен запис "ОБЩО" за дадено табло.
    ''' Логиката включва:
    ''' - събиране на всички кръгове
    ''' - изчисляване на мощности и консуматори
    ''' - определяне на фази и полюси
    ''' - намиране или създаване на запис "ОБЩО"
    ''' - изчисляване на ток и избор на апаратура
    ''' </summary>
    Private Sub BuildPanelSummaryRecord(tabloName As String)
        ' Взима всички кръгове за текущото табло без вече съществуващите "ОБЩО"
        Dim panelCircuits As List(Of strTokow) = ListTokow.Where(Function(t)
                                                                     Return t.Tablo = tabloName AndAlso
                                                                  t.ТоковКръг <> "ОБЩО"
                                                                 End Function).ToList()
        ' Ако няма кръгове, прекратяваме обработката
        If panelCircuits.Count = 0 Then Exit Sub
        ' Обща мощност на таблото
        Dim totalPower As Double = panelCircuits.Sum(Function(c) c.Мощност)
        ' Общ брой осветителни тела
        Dim totalLamps As Integer = panelCircuits.Sum(Function(c) c.brLamp)
        ' Общ брой контакти
        Dim totalContacts As Integer = panelCircuits.Sum(Function(c) c.brKontakt)
        ' Проверка дали има трифазни консуматори
        Dim hasThreePhase As Boolean = panelCircuits.Any(Function(c) c.Брой_Полюси = 3)
        ' Определяне на брой полюси (3 ако има трифазни, иначе 1)
        Dim mostCommonPoles As Integer = If(panelCircuits.Any(Function(c) c.Брой_Полюси = 3), 3, 1)
        ' Определяне на фазово обозначение
        Dim totalPhase As String = If(hasThreePhase, "L1,L2,L3", "L")
        ' Търсене на съществуващ запис "ОБЩО"
        Dim totalTokow = ListTokow.FirstOrDefault(Function(t)
                                                      Return t.Tablo = tabloName AndAlso
                                                      t.ТоковКръг = "ОБЩО"
                                                  End Function)
        ' Ако не съществува, създаваме нов запис
        If totalTokow Is Nothing Then
            totalTokow = New strTokow With {
                             .Tablo = tabloName,
                             .ТоковКръг = "ОБЩО"
            }
            ListTokow.Add(totalTokow)
        End If
        ' Попълване на обобщените данни
        With totalTokow
            .Device = "Табло"
            .Tablo = tabloName
            .ТоковКръг = "ОБЩО"
            .Брой_Полюси = mostCommonPoles
            .Мощност = totalPower
            .Фаза = totalPhase
            .brLamp = totalLamps
            .brKontakt = totalContacts
            .Табло_Родител = ""
            If .Tablo = ROOT_NODE_TEXT Then
                .Консуматор = "Ке="
                .предназначение = "Рпр.=15кW"
            Else
                .Консуматор = ""
                .предназначение = ""
            End If
        End With
        ' Изчисляване на ток
        If hasThreePhase Then
            ' Балансиране на фазите
            BalancePhases(tabloName)
            ' Извличане на стойности от текстови полета (формат "X>стойност")
            Dim valL1 As Double = CDbl(totalTokow.RCD_Клас.Split(">"c)(1))
            Dim valL2 As Double = CDbl(totalTokow.RCD_Ток.Split(">"c)(1))
            Dim valL3 As Double = CDbl(totalTokow.RCD_Чувствителност.Split(">"c)(1))
            ' Изчисляване на максимален ток
            totalTokow.Ток = Math.Max(valL1, Math.Max(valL2, valL3))
        Else
            ' Изчисляване на еднофазен ток
            totalTokow.Ток = calc_Inom(totalTokow.Мощност, totalTokow.Брой_Полюси)
        End If
        ' Избор на прекъсвач/разединител
        CalculateDisconnector(totalTokow)
        ' Избор на кабел
        CalculateCable(totalTokow)
    End Sub
    ''' <summary>
    ''' Актуализира (или създава) запис за разединител в дадено табло.
    ''' Логиката:
    ''' 1. Взема всички токови кръгове за таблото
    ''' 2. Премахва стария разединител (ако има)
    ''' 3. Изчислява обща мощност и ток от шините
    ''' 4. Създава нов разединител
    ''' 5. Вмъква го след последната шина
    ''' </summary>
    ''' <param name="tabloName">Име на таблото</param>
    Private Sub UpdateDisconnectorRecord(tabloName As String)
        ' 1️ ВЗЕМАМЕ ВСИЧКИ КРЪГОВЕ ОТ ТАБЛОТО
        Dim circuitsInTablo As List(Of strTokow) =
        ListTokow.Where(Function(t) t.Tablo = tabloName).ToList()
        If circuitsInTablo.Count = 0 Then Exit Sub
        ' 2️ ПРЕМАХВАМЕ СТАР РАЗЕДИНИТЕЛ (АКО ИМА)
        ListTokow.RemoveAll(Function(t) t.Tablo = tabloName AndAlso t.Device = "Разединител")
        ' 3️ ВЗЕМАМЕ САМО ШИНИТЕ
        Dim busCircuits As List(Of strTokow) =
        circuitsInTablo.Where(Function(t) t.Шина = True).ToList()
        If busCircuits.Count = 0 Then Exit Sub
        ' 4️ ИЗЧИСЛЕНИЯ
        ' Обща мощност на всички кръгове към шината
        Dim totalPower As Double = busCircuits.Sum(Function(c) c.Мощност)
        ' Проверка дали има трифазни консуматори
        Dim hasThreePhase As Boolean =
            busCircuits.Any(Function(c) c.Брой_Полюси = 3)
        ' Определяне на брой полюси и фази
        Dim poles As Integer = If(hasThreePhase, 3, 1)
        Dim phases As String = If(hasThreePhase, "L1,L2,L3", "L")
        ' Изчисляване на общ ток
        Dim totalCurrent As Double
        If hasThreePhase Then
            totalCurrent = (totalPower * 1000) / (Math.Sqrt(3) * 400)
        Else
            totalCurrent = (totalPower * 1000) / 230
        End If
        ' 5️ СЪЗДАВАНЕ НА НОВ РАЗЕДИНИТЕЛ
        Dim disconnector As New strTokow With {
                                .Device = "Разединител",
                                .ТоковКръг = "Разединител",
                                .Tablo = tabloName,
                                .Брой_Полюси = poles,
                                .Мощност = totalPower,
                                .Ток = totalCurrent,
                                .Фаза = phases,
                                .Консуматор = "Разединител",
                                .предназначение = "за I шина"
        }
        ' 6️ НАМИРАНЕ НА ПОЗИЦИЯ (СЛЕД ПОСЛЕДНАТА ШИНА)
        Dim lastBusIndex As Integer = -1
        For i As Integer = 0 To ListTokow.Count - 1
            If ListTokow(i).Tablo = tabloName AndAlso ListTokow(i).Шина = True Then
                lastBusIndex = i
            End If
        Next
        ' 7️ ВМЪКВАНЕ В СПИСЪКА
        ListTokow.Insert(lastBusIndex + 1, disconnector)
        ' Определя конкретен тип разединител според тока
        CalculateDisconnector(disconnector)
    End Sub
    ''' <summary>
    ''' Отваря форма за масово добавяне на резервни/нови кръгове към избраното табло.
    ''' Взима текущо избраното табло от TreeView, почиства името му и подава списъка ListTokow към подчинена форма.
    ''' След успешно добавяне обновява визуализацията чрез сортиране и презареждане на DataGridView.
    ''' </summary>
    Private Sub ToolStripButton_Добави_резерва_Click(sender As Object, e As EventArgs) Handles ToolStripButton_Добави_резерва.Click
        ' 1. Взимаме текущо избраното табло от TreeView
        Dim selectedTablo As String = TreeView1.SelectedNode?.Text
        ' 2. Проверка дали има избрано табло
        If String.IsNullOrEmpty(selectedTablo) Then
            MessageBox.Show("Моля, първо изберете табло от списъка вляво.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        ' 3. Почистваме текста (махаме допълнителна информация в скоби)
        If selectedTablo.Contains("(") Then
            selectedTablo = selectedTablo.Substring(0, selectedTablo.IndexOf("(")).Trim()
        End If
        ' 4. Отваряме форма за масово добавяне на резервни/нови кръгове
        ' Подаваме текущия списък и избраното табло като контекст
        Using frm As New Form_BatchAddCircuits(ListTokow, selectedTablo)
            If frm.ShowDialog(Me) = DialogResult.OK Then
                ' 5. Ако потребителят е потвърдил → обновяваме визуализацията
                SortCircuits()
                FillDataGridViewForPanel()
            End If
        End Using
    End Sub
End Class

#Region "Клас за добавяне на резервни и съществуващи кръгове"
Public Class Form_BatchAddCircuits
    Inherits Form
    ' Данни, подадени от основната форма
    Private _targetList As List(Of Form_Tablo_new.strTokow)
    Private _tabloName As String
    ' Контроли, до които ще имаме достъп по-късно
    Private numExist As NumericUpDown
    Private numReserve As NumericUpDown
    Private btnOk As Button
    Private btnCancel As Button
    Private lblInfo As Label
    ''' <summary>
    ''' Инициализира формата, приема входните данни и извиква процедурите за изграждане.
    ''' </summary>
    Public Sub New(targetList As List(Of Form_Tablo_new.strTokow), tabloName As String)
        ' 1. Записваме подадените данни в локални полета
        _targetList = targetList
        _tabloName = tabloName
        ' 2. Извикваме процедурите в строго определен ред
        ConfigureFormSettings()   ' Настройки на самата форма
        BuildUserInterface()      ' Създаване и подреждане на визуалните елементи
        SetupEventHandlers()      ' Свързване на събития и клавишни преки пътища
    End Sub
    ''' <summary>
    ''' Задава базовите свойства на прозореца (размер, стил, позиция, цветове).
    ''' </summary>
    Private Sub ConfigureFormSettings()
        Me.Text = "Добавяне на кръгове за Същеструващи/Резерви"
        Me.Size = New Size(400, 220)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.StartPosition = FormStartPosition.CenterParent
        Me.BackColor = Color.White
        Me.MaximizeBox = False
        Me.MinimizeBox = False
    End Sub
    ''' <summary>
    ''' Създава, конфигурира и добавя всички контроли към формата.
    ''' Отговаря само за визуалната структура.
    ''' </summary>
    Private Sub BuildUserInterface()
        ' --- Хедър ---
        lblInfo = New Label With {
            .Text = "Табло: " & _tabloName,
            .Dock = DockStyle.Top,
            .Height = 35,
            .TextAlign = ContentAlignment.MiddleCenter,
            .BackColor = Color.FromArgb(0, 102, 204),
            .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 14, FontStyle.Bold)
        }
        Me.Controls.Add(lblInfo)
        ' --- GroupBox контейнер ---
        Dim grp As New GroupBox With {
            .Text = " Брой кръгове за добавяне ",
            .Location = New Point(15, 50),
            .Size = New Size(355, 70),
            .Font = New Font("Segoe UI", 12, FontStyle.Bold)
        }
        Me.Controls.Add(grp)
        ' --- TableLayoutPanel за подредба (Етикет - Поле - Етикет - Поле) ---
        Dim tbl As New TableLayoutPanel With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 4,
            .RowCount = 1,
            .Padding = New Padding(5)
        }
        tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 45))
        tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 20))
        tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 45))
        tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 20))
        grp.Controls.Add(tbl)
        ' --- Лейбъл за Съществуващи ---
        Dim lblE As New Label With {
                .Text = "Същеструващи:",
                .AutoSize = False,
                .TextAlign = ContentAlignment.MiddleRight,
                .Dock = DockStyle.Fill,
                .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }

        ' --- Лейбъл за Резерви ---
        Dim lblR As New Label With {
                .Text = "Резерви:",
                .AutoSize = False,
                .TextAlign = ContentAlignment.MiddleRight,
                .Dock = DockStyle.Fill,
                .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }
        numExist = New NumericUpDown With {
                .Minimum = 0,
                .Maximum = 100,
                .Value = 1,
                .Width = 60,
                .TextAlign = HorizontalAlignment.Center,
                .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }
        numReserve = New NumericUpDown With {
            .Minimum = 0,
            .Maximum = 100,
            .Value = 1,
            .Width = 60,
            .TextAlign = HorizontalAlignment.Center,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }
        tbl.Controls.Add(lblE, 0, 0)
        tbl.Controls.Add(numExist, 1, 0)
        tbl.Controls.Add(lblR, 2, 0)
        tbl.Controls.Add(numReserve, 3, 0)
        ' --- Панел за бутоните (долу вдясно) ---
        Dim pnlBtns As New FlowLayoutPanel With {
            .FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft,
            .Location = New Point(15, 130),
            .Size = New Size(355, 45),
            .WrapContents = False,
            .Padding = New Padding(15, 5, 0, 0)
        }
        Me.Controls.Add(pnlBtns)
        ' --- Бутон ГЕНЕРИРАЙ ---
        btnOk = New Button With {
            .Text = "ГЕНЕРИРАЙ",
            .Size = New Size(95, 32),
            .BackColor = Color.FromArgb(0, 102, 204),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.System,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .DialogResult = DialogResult.OK,
            .Margin = New Padding(0, 0, 10, 0)
        }
        btnOk.FlatAppearance.BorderSize = 0
        btnOk.FlatAppearance.MouseOverBackColor = Color.FromArgb(0, 120, 240)
        ' --- Бутон ОТКАЗ ---
        btnCancel = New Button With {
            .Text = "ОТКАЗ",
            .Size = New Size(95, 32),
            .BackColor = Color.FromArgb(240, 240, 240),
            .ForeColor = Color.FromArgb(60, 60, 60),
            .FlatStyle = FlatStyle.System,
            .Font = New Font("Segoe UI", 10, FontStyle.Regular),
            .DialogResult = DialogResult.Cancel, .Margin = New Padding(0)
        }
        btnCancel.FlatAppearance.BorderSize = 0
        btnCancel.FlatAppearance.MouseOverBackColor = Color.FromArgb(220, 220, 220)

        pnlBtns.Controls.Add(btnOk)
        pnlBtns.Controls.Add(btnCancel)
    End Sub
    ''' <summary>
    ''' Дефинира кои бутони затварят формата и какви действия изпълняват.
    ''' </summary>
    Private Sub SetupEventHandlers()
        Me.AcceptButton = btnOk
        Me.CancelButton = btnCancel

        ' Свързваме логиката с бутона. 
        ' (DialogResult затваря формата автоматично след изпълнение на този handler)
        AddHandler btnOk.Click, AddressOf ProcessGeneration
    End Sub
    ''' <summary>
    ''' Прочита въведените стойности, създава обектите и ги добавя към списъка.
    ''' Отговаря САМО за логиката и данните, не за визуалната част.
    ''' </summary>
    Private Sub ProcessGeneration(sender As Object, e As EventArgs)
        ' Бърза валидация: да не се добавя нищо при нули
        If numExist.Value = 0 AndAlso numReserve.Value = 0 Then
            MessageBox.Show("Моля, въведете поне 1 за добавяне.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.DialogResult = DialogResult.None ' Спира автоматичното затваряне
            Return
        End If
        ' 1. Генериране на "съществуващи" кръгове
        For i As Integer = 1 To CInt(numExist.Value)
            _targetList.Add(New Form_Tablo_new.strTokow With {
                .Tablo = _tabloName, .Device = "Съществуващ", .ТоковКръг = "същ.",
                .Консуматор = "Съществуващ", .предназначение = "не се променя",
                .Breaker_Номинален_Ток = "Същ.", .Мощност = 0, .Ток = 0,
                .Брой_Полюси = 0, .Фаза = "---"
            })
        Next
        ' 2. Генериране на резервни кръгове
        For i As Integer = 1 To CInt(numReserve.Value)
            _targetList.Add(New Form_Tablo_new.strTokow With {
                .Tablo = _tabloName, .Device = "Резерва", .ТоковКръг = "рез.",
                .Консуматор = "Резерв", .предназначение = "",
                .Breaker_Номинален_Ток = "Същ.", .Breaker_Тип_Апарат = "EZ9 MCB",
                .Брой_Полюси = 1, .Фаза = "---", .Мощност = 0, .Ток = 0
            })
        Next
    End Sub
End Class
#End Region