Imports System.Reflection
Imports System.Windows.Forms
Imports iTextSharp.text.pdf.qrcode.Version

Public Class DataGridViewChangeManager
    ' Пазим препратки към шестте компонента
    Private _cableCatalog As CableCatalog
    Private _breakerCatalog As BreakerCatalog
    Private _disconnectorCatalog As DisconnectorCatalog
    Private _rcdCatalog As RCDCatalog
    Private _calculationEngine As ElectricalCalculationEngine
    Private _allCircuits As List(Of clsTokow)
    Private _boardManager As BoardStructureManager
    ''' <summary>
    ''' Конструкторът вече приема точно шестте компонента от формата
    ''' </summary>
    Public Sub New()
        Me._breakerCatalog = AppSettings.BreakerCatalog
        Me._disconnectorCatalog = AppSettings.DisconnectorCatalog
        Me._rcdCatalog = AppSettings.RcdCatalog
        Me._cableCatalog = AppSettings.CableCatalog
        Me._calculationEngine = AppSettings.ElectricalCalculationEngine
        Me._allCircuits = AppSettings.ListTokow
        Me._boardManager = AppSettings.BoardStructureManager
    End Sub
    ''' <summary>
    ''' Главната входна точка. Взема името на процедурата от формата (Индекс 4) 
    ''' и я извиква динамично чрез Reflection.
    ''' </summary>
    Public Sub UpdateCircuitProperty(ByVal tokow As clsTokow, ByVal procedureToExecute As String, ByVal newValue As String)
        ' ------------------------------------------------------------
        ' 🛡️ ОБЩА ПРОВЕРКА ЗА КОРЕКТНОСТ (За трите параметъра)
        ' ------------------------------------------------------------
        If tokow Is Nothing OrElse
           String.IsNullOrEmpty(procedureToExecute) Then
            ' Показваме съобщение за грешни/непълни данни
            MessageBox.Show(
                "Операцията е прекратена! Подадени са некоректни или празни данни за токовия кръг, процедурата или стойността.",
                "Невалидни данни",
                MessageBoxButtons.OK,
                MessageBoxIcon.Exclamation
            )
            ' Изход от процедурата
            Exit Sub
        End If
        If tokow.Device = "Разединител" OrElse
           tokow.Device = "Съществуващ" OrElse
           tokow.Device = "Резерва" Then Exit Sub
        ' Тъй като вече си подсигурил защитите във формата, тук директно търсим метода
        Try
            ' Намира публичния метод със съответното име в този клас
            Dim method As MethodInfo = Me.GetType().GetMethod(procedureToExecute)
            If method IsNot Nothing Then
                ' Изпълнява намерената процедура, подавайки tokow и newValue като аргументи
                method.Invoke(Me, New Object() {tokow, newValue})
            End If
        Catch ex As TargetInvocationException
            ' Улавяме грешка, възникнала вътре в самата инженерна процедура (напр. в изчислителния модул)
            ' Извличаме реалната инженерна грешка, която е възникнала вътре в извикания метод
            Dim realException As Exception = ex.InnerException
            Dim errorMessage As String = If(realException IsNot Nothing, realException.Message, ex.Message)
            ' Показваме елегантно съобщение на потребителя, вместо да чупим AutoCAD
            MessageBox.Show(
                $"Възникна грешка при обработка на промяната ({procedureToExecute}):{Environment.NewLine}{errorMessage}",
                "Инженерен изчислителен модул",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning)
        Finally
            ' Тук може да добавиш код, който винаги трябва да се изпълни след опита за извикване на процедурата
            ' Например, ако искаш да обновиш някакъв статус в UI-то или да логнеш действието

        End Try
    End Sub
    ' =================================================================
    ' === СЪЩИНСКИ ПРОЦЕДУРИ (Пренесени от стария Select Case) ===
    ' =================================================================
    ''' <summary>
    ''' Обработва промяната в групата на ДТЗ Нула, преизчислява логиката на таблото
    ''' и автоматично обновява съответните редове и колони в потребителския интерфейс.
    ''' </summary>
    Public Sub HandleRcdZeroChange(ByVal panelCircuits As clsTokow, ByVal value As String)
        ' 1. Пускаме текста през санитарния филтър
        Dim validatedValue As String = ValidateRCDNulla(value)
        ' 2. Ако филтърът върне валиден резултат → записваме го в обекта
        If validatedValue IsNot Nothing Then
            panelCircuits.RCD_Нула = validatedValue
        End If
        ' 3. Филтрираме всички токови кръгове за текущата сграда и табло
        Dim panels = AppSettings.ListTokow.Where(Function(t)
                                                     Return t.BuildingName = panelCircuits.BuildingName AndAlso
                                                            t.Tablo = panelCircuits.Tablo
                                                 End Function).ToList()
        ' 4. КРИТИЧНА ЗАЩИТА: Сваляме флага на True, за да блокираме CellValueChanged събитията на Grid-а
        AppSettings.IsGridLoading = True
        Try
            ' 5. Извикваме мениджъра, за да пренареди и преизчисли ДТЗ групите на таблото
            _boardManager.ProcessPanelRCDLogic(panels)
            ' 6. Извикваме втората процедура, която софтуерно ще попълни новите данни в Grid-а
            AppSettings.IsGridLoading = False
            AppSettings.DataGridViewManager.UpdateRcdGridValues(panels)
        Finally
            ' 7. Вдигаме флага обратно на False,
            ' за да разрешим отново нормалната работа на потребителя с Grid-а
            AppSettings.IsGridLoading = False
        End Try
    End Sub
    ' =================================================================
    ' === ПОМОЩНИ ФУНКЦИИ (Валидации и санитарни филтри) ===
    ' =================================================================
    ''' <summary>
    ''' Изчиства текста и проверява дали форматът отговаря на "N" + число (напр. N1, N2)
    ''' </summary>
    Private Function ValidateRCDNulla(ByVal inputValue As String) As String
        ' Проверка 1: Дали е празно и дали започва с "N"
        If String.IsNullOrEmpty(inputValue) OrElse Not inputValue.ToUpper().StartsWith("N") Then Return Nothing
        ' Извлечи числото след "N"
        Dim numberPart As String = inputValue.Substring(1).Trim()
        ' Премахни всичко, което НЕ е цифра
        numberPart = New String(numberPart.Where(Function(c) Char.IsDigit(c)).ToArray())
        ' Проверка 2: Дали има останали числа след чистенето
        If String.IsNullOrEmpty(numberPart) Then Return Nothing
        ' Проверка 3: Дали числото е валидно
        Dim rcdNumber As Integer
        If Not Integer.TryParse(numberPart, rcdNumber) Then Return Nothing
        ' Проверка 4: Дали числото е строго по-голямо от 0
        If rcdNumber <= 0 Then Return Nothing
        ' ✅ Всички проверки минаха → връщаме стандартизирания текст с главна буква
        Return "N" & rcdNumber.ToString()
    End Function
    ''' <summary>
    ''' Сменя: "Тип на апарата". 
    ''' Филтрира каталозите според вида на устройството и връща новите списъци за ComboBox клетките.
    ''' </summary>
    Public Function HandleBreakerTypeChange(ByVal circuit As clsTokow, ByVal value As String) As Dictionary(Of String, List(Of String))
        ' Създаваме речник, в който ще запишем филтрираните списъци за различните редове
        Dim resultLists As New Dictionary(Of String, List(Of String))()
        If circuit Is Nothing OrElse String.IsNullOrEmpty(value) Then Return resultLists
        ' 1. Записваме новия избран тип апарат в обекта
        circuit.Breaker_Тип_Апарат = value
        ' 2. Разделяме логиката според типа на устройството (Device)
        Select Case circuit.Device
            Case "Разединител", "Табло"
                ' Използваме инжектирания каталог за разединители
                If _disconnectorCatalog IsNot Nothing Then
                    ' Филтрираме разединителите по избрания тип (selectedValue)
                    Dim filteredDisco = _disconnectorCatalog.Disconnectors.Where(Function(b) b.Type = value).ToList()

                    ' Извличаме уникалните номинални токове
                    Dim valuesForCombo = filteredDisco _
                                    .Select(Function(b) b.NominalCurrent.ToString()) _
                                    .Distinct() _
                                    .ToList()
                    ' Записваме в речника, че за ред "Номинален ток" имаме нов списък
                    resultLists.Add("Номинален ток", valuesForCombo)
                End If
            Case Else
                ' За автоматични прекъсвачи – използваме инжектирания каталог _breakerCatalog
                If _breakerCatalog IsNot Nothing Then
                    ' Филтрираме прекъсвачите по избраната серия (Series)
                    Dim filteredBreakers = _breakerCatalog.Breakers.Where(Function(b) b.Series = value).ToList()
                    If filteredBreakers.Count = 0 Then Return resultLists
                    ' 3. Записваме автоматично изключвателната възможност (Ics) в обекта
                    circuit.Breaker_Изкл_Възможност = filteredBreakers.First().Ics_kA & "kA"
                    ' 4. Генерираме уникалните списъци за ComboBox-овете в Grid-а
                    ' Списък за "Номинален ток"
                    Dim valuesNominal = filteredBreakers.Select(Function(b) b.NominalCurrent.ToString()).Distinct().ToList()
                    resultLists.Add("Номинален ток", valuesNominal)
                    ' Списък за "Крива"
                    Dim valuesCurve = filteredBreakers.Select(Function(b) b.Curve.ToString()).Distinct().ToList()
                    resultLists.Add("Крива", valuesCurve)
                    ' Списък за "Защитен блок"
                    Dim valuesTripUnit = filteredBreakers.Select(Function(b) b.TripUnit).Distinct().ToList()
                    resultLists.Add("Защитен блок", valuesTripUnit)
                End If
        End Select
        ' Връщаме събраните списъци обратно към формата
        Return resultLists
    End Function
    ''' <summary>
    ''' Сменя: "Номинален ток"
    ''' </summary>
    Public Sub HandleNominalCurrentChange(ByVal circuit As clsTokow, ByVal value As String)
        If String.IsNullOrEmpty(value) Then Exit Sub
        If Val(value) >= Val(circuit.Breaker_Номинален_Ток) Then
            circuit.Breaker_Номинален_Ток = value
        Else
            Dim msg As String = "Сигурен ли си в това, което правиш? " & vbCrLf &
                               "Избраният ток е по-малък от текущия." & vbCrLf &
                               "Искаш ли наистина да продължиш към Тъмната страна?"
            Dim result As DialogResult = MessageBox.Show(
                msg, "Внимание: Инженерна мисъл в действие!",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
            If result = DialogResult.Yes Then
                circuit.Breaker_Номинален_Ток = value
            Else
                MessageBox.Show("Мъдро решение! Спести си един ремонт.",
                                "Браво!",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information)
                Exit Sub
            End If
        End If
        _cableCatalog.CalculateCable(circuit)
    End Sub
    ''' <summary>
    ''' Обработва промяна на настройката за 
    ''' наличие на ДТЗ за конкретен токов кръг.
    '''
    ''' При валидна входна стойност обновява
    ''' свойството ДТЗ_RCD на подадения обект.
    '''
    ''' За обекти от тип "Табло" използването на ДТЗ
    ''' не е разрешено и стойността винаги се задава като False.
    ''' </summary>
    ''' <param name="circuit">
    ''' Обектът clsTokow, за който се променя настройката за ДТЗ.
    ''' </param>
    ''' <param name="value">
    ''' Новата стойност, подадена като текст, която се преобразува до Boolean.
    ''' </param>
    Public Sub HandleRcdToggleChange(ByVal circuit As clsTokow, ByVal value As String)
        If circuit.Device = "Табло" Then
            circuit.ДТЗ_RCD = False
            Exit Sub
        End If
        Dim result As Boolean
        If Boolean.TryParse(value, result) Then circuit.ДТЗ_RCD = result
        circuit.RCD_Автомат = True
        _breakerCatalog.ClearBreaker(circuit)
        _rcdCatalog.ClearRCD(circuit)
        If circuit.ДТЗ_RCD Then
            circuit.RCD_Нула = "N0"
            _rcdCatalog.SetRCD(circuit)
            circuit.RCD_Нула = "N"
        Else
            _breakerCatalog.CalculateBreaker(circuit)
        End If
        _cableCatalog.CalculateCable(circuit)
    End Sub
    Public Sub HandleCableTypeChange(ByVal circuit As clsTokow, ByVal value As String)
        _cableCatalog.CalculateCable(circuit, Cable_Type:=value)
    End Sub
    Public Sub HandleCableRoutingChange(ByVal circuit As clsTokow, ByVal value As String)
        _cableCatalog.CalculateCable(circuit, layMethod:=value)
    End Sub
    Public Sub HandleGroupCablesChange(ByVal circuit As clsTokow, ByVal value As String)
        _cableCatalog.CalculateCable(circuit, Broj_Cable:=value)
    End Sub
End Class