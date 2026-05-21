' =========================================================================
' ФАЙЛ: Form_Tablo_CatalogManager.vb
' ОТГОВОРНОСТ: Капсулиране на всички каталожни данни (кабели, прекъсвачи)
'              и инженерната логика за техния избор.
' =========================================================================

Imports System.Linq

''' <summary>
''' Главен мениджър за управление на каталозите с ел. апаратура и материали.
''' </summary>
Public Class Form_Tablo_CatalogManager

    ' Достъп до отделните специализирани каталози
    Public ReadOnly Property Breakers As BreakerCatalog
    Public ReadOnly Property Cables As CableCatalog
    Public ReadOnly Property Disconnectors As DisconnectorCatalog

    Public Sub New()
        ' Инициализиране на подкласовете
        _Breakers = New BreakerCatalog()
        _Cables = New CableCatalog()
        _Disconnectors = New DisconnectorCatalog()
    End Sub

    ''' <summary>
    ''' Зарежда абсолютно всички каталози наведнъж (извиква се в Form_Load).
    ''' </summary>
    Public Sub SetCatalog()
        Cables.LoadCatalog()
        Breakers.LoadCatalog()
        Disconnectors.LoadCatalog()
        ' Тук ще извикаме и другите, ако се наложи
    End Sub

End Class

' =========================================================================
' 1. МОДУЛ: АВТОМАТИЧНИ ПРЕКЪСВАЧИ
' =========================================================================
Public Class BreakerCatalog
    ' Преместваме структурата (или класа) вътре, за да не тежи в основната форма
    Public Structure BreakerInfo
        Public Brand As String
        Public Series As String
        Public NominalCurrent As Double
        Public Characteristic As String
        Public Poles As Integer
        ' Добави тук и другите полета, които структурата ти има в момента
    End Structure

    ' Списъкът с данни (склада)
    Public Property DataList As List(Of BreakerInfo)

    Public Sub New()
        DataList = New List(Of BreakerInfo)()
    End Sub

    ''' <summary>
    ''' Зарежда данните за автоматичните прекъсвачи (старото ти FillBreakers).
    ''' </summary>
    Public Sub LoadCatalog()
        ' TODO: Тук ще копираме кода от твоето FillBreakers()
    End Sub

    ''' <summary>
    ''' Автоматичен избор на най-подходящия прекъсвач по изчислен ток и характеристика.
    ''' </summary>
    Public Function SelectBestBreaker(calculatedCurrent As Double, characteristic As String) As BreakerInfo
        ' Намира първия прекъсвач, чийто номинален ток е по-голям или равен на изчисления, 
        ' и съвпада по търсената характеристика (напр. "C", "B")
        Dim selected = DataList.
            Where(Function(b) b.NominalCurrent >= calculatedCurrent AndAlso b.Characteristic.ToUpper() = characteristic.ToUpper()).
            OrderBy(Function(b) b.NominalCurrent).
            FirstOrDefault()

        Return selected
    End Function
End Class

' =========================================================================
' 2. МОДУЛ: КАБЕЛИ
' =========================================================================
Public Class CableCatalog
    ' Тук ще дефинираме структурата CableInfo, когато ми я изпратиш
    Public Property DataList As List(Of Object)

    Public Sub New()
        DataList = New List(Of Object)()
    End Sub

    ''' <summary>
    ''' Зарежда речника/списъка с кабели (старото ти FillCables).
    ''' </summary>
    Public Sub LoadCatalog()
        ' TODO: Тук ще преместим кода от FillCables()
    End Sub

    ''' <summary>
    ''' Автоматичен избор на сечение на кабела.
    ''' </summary>
    Public Function SelectCable(breakerCurrent As Double, installationMethod As String) As Object
        ' TODO: Тук ще напишем чистата инженерна логика за избор на кабел
        Return Nothing
    End Function
End Class

' =========================================================================
' 3. МОДУЛ: РАЗЕДИНИТЕЛИ (ТОВАРИ ПРЕКЪСВАЧИ)
' =========================================================================
Public Class DisconnectorCatalog
    Public Structure DisconnectorInfo
        Public NominalCurrent As Double
        Public Model As String
        ' Още полета според твоя код...
    End Structure

    Public Property DataList As List(Of DisconnectorInfo)

    Public Sub New()
        DataList = New List(Of DisconnectorInfo)()
    End Sub

    ''' <summary>
    ''' Зарежда списъка с разединители.
    ''' </summary>
    Public Sub LoadCatalog()
        ' TODO: Тук ще дойде Disconnectors = New List(Of DisconnectorInfo) From { ... }
    End Sub
End Class