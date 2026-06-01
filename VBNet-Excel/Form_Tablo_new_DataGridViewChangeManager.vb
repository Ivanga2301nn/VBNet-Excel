Imports System.Reflection

Public Class DataGridViewChangeManager
    ' Пазим препратки към шестте компонента
    Private _cableCatalog As CableCatalog
    Private _breakerCatalog As BreakerCatalog
    Private _disconnectorCatalog As DisconnectorCatalog
    Private _rcdCatalog As RCDCatalog
    Private _calculationEngine As ElectricalCalculationEngine

    ''' <summary>
    ''' Конструкторът вече приема точно шестте компонента от формата
    ''' </summary>
    Public Sub New(ByVal breakerCat As BreakerCatalog,
                   ByVal disconnectorCat As DisconnectorCatalog,
                   ByVal rcdCat As RCDCatalog,
                   ByVal cableCat As CableCatalog,
                   ByVal calcEngine As ElectricalCalculationEngine)
        Me._breakerCatalog = breakerCat
        Me._disconnectorCatalog = disconnectorCat
        Me._rcdCatalog = rcdCat
        Me._cableCatalog = cableCat
        Me._calculationEngine = calcEngine
    End Sub
    ''' <summary>
    ''' Главната входна точка. Взема името на процедурата от формата (Индекс 4) 
    ''' и я извиква динамично чрез Reflection.
    ''' </summary>
    Public Sub UpdateCircuitProperty(ByVal circuit As clsTokow, ByVal procedureToExecute As String, ByVal newValue As String)
        ' Тъй като вече си подсигурил защитите във формата, тук директно търсим метода
        Try
            ' Намира публичния метод със съответното име в този клас
            Dim method As MethodInfo = Me.GetType().GetMethod(procedureToExecute)
            If method IsNot Nothing Then
                ' Изпълнява намерената процедура, подавайки circuit и newValue като аргументи
                method.Invoke(Me, New Object() {circuit, newValue})
            Else
                ' Полезно съобщение в конзолата, ако името в шаблона е сгрешено или методът още липсва
                Debug.WriteLine($"[Грешка] Методът '{procedureToExecute}' не е дефиниран в DataGridViewChangeManager.")
            End If
        Catch ex As TargetInvocationException
            ' Улавяме грешка, възникнала вътре в самата инженерна процедура (напр. в изчислителния модул)
            Throw New Exception($"Грешка в процедура {procedureToExecute}: {ex.InnerException?.Message}", ex)
        End Try
    End Sub
    ''' <summary>
    ''' Извиква се динамично, когато потребителят промени типа на прекъсвача в Grid-а.
    ''' </summary>
    Public Sub HandleBreakerTypeChange(ByVal circuit As clsTokow, ByVal value As String)
        ' 1. Защита: Проверяваме дали подаденият обект е валиден
        If circuit Is Nothing Then Exit Sub

        ' 2. Записваме новия тип апарат в обекта на токовия кръг
        ' (Например: "Автоматичен прекъсвач", "Моторен прекъсвач", "Товаров прекъсвач")
        circuit.Breaker_Тип_Апарат = value

    End Sub

End Class