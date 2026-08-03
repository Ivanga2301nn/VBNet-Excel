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
        If AppSettings.IsProcessingRcd Then
            Dim validatedSilent = ValidateRCDNulla(value)
            If validatedSilent IsNot Nothing Then
                AppSettings.EnterGridUpdate()
                Try
                    panelCircuits.RCD_Нула = validatedSilent
                Finally
                    AppSettings.ExitGridUpdate()
                End Try
            End If
            Return
        End If
        Dim validatedValue As String = ValidateRCDNulla(value)
        If validatedValue IsNot Nothing Then panelCircuits.RCD_Нула = validatedValue
        Dim panels = AppSettings.ListTokow.Where(Function(t) t.BuildingName = panelCircuits.BuildingName AndAlso t.Tablo = panelCircuits.Tablo).ToList()
        AppSettings.EnterGridUpdate()
        Try
            AppSettings.BoardStructureManager.ProcessPanelRCDLogic(panels)
            AppSettings.DataGridViewManager.UpdateRcdGridValues(panels)
        Finally
            AppSettings.ExitGridUpdate()
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
    ' =============================================================
    ' Процедура: HandleRcdToggleChange
    ' =============================================================
    ''' <summary>
    ''' Обработва промяна на настройката за използване
    ''' на ДТЗ (RCD) за конкретен токов кръг.
    '''
    ''' Процедурата променя състоянието на ДТЗ,
    ''' нулира предварително избраните защитни апарати,
    ''' при необходимост избира подходяща ДТЗ
    ''' и преизчислява свързаните параметри.
    '''
    ''' При промяна на защитата се извършва
    ''' ново изчисление на прекъсвача и кабела,
    ''' тъй като изборът на ДТЗ влияе
    ''' върху крайната конфигурация на кръга.
    '''
    ''' За главни записи от тип "Табло"
    ''' използването на ДТЗ не се допуска.
    ''' </summary>
    ''' <param name="circuit">
    ''' Токовият кръг, за който се променя
    ''' настройката на ДТЗ.
    ''' </param>
    ''' <param name="value">
    ''' Нова стойност на настройката,
    ''' подадена като текст,
    ''' която се преобразува до Boolean.
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
    ' =============================================================
    ' Процедура: HandleCableTypeChange
    ' =============================================================
    ''' <summary>
    ''' Обработва промяна на типа кабел
    ''' за даден токов кръг.
    '''
    ''' След промяна на типа кабел
    ''' извиква преизчисляване на кабелните параметри,
    ''' като използва новия тип кабел.
    ''' </summary>
    ''' <param name="circuit">
    ''' Токовият кръг,
    ''' за който се променя кабелът.
    ''' </param>
    ''' <param name="value">
    ''' Новият тип кабел,
    ''' избран от потребителя.
    ''' </param>
    Public Sub HandleCableTypeChange(ByVal circuit As clsTokow, ByVal value As String)
        _cableCatalog.CalculateCable(circuit, Cable_Type:=value)
    End Sub
    ' =============================================================
    ' Процедура: HandleCableRoutingChange
    ' =============================================================
    ''' <summary>
    ''' Обработва промяна на начина
    ''' на полагане на кабела.
    '''
    ''' При промяна на начина на монтаж
    ''' преизчислява параметрите на кабела,
    ''' тъй като условията на полагане
    ''' влияят върху допустимия ток.
    ''' </summary>
    ''' <param name="circuit">
    ''' Токовият кръг,
    ''' за който се променя
    ''' начинът на полагане.
    ''' </param>
    ''' <param name="value">
    ''' Новият начин на полагане,
    ''' избран от потребителя.
    ''' </param>
    Public Sub HandleCableRoutingChange(ByVal circuit As clsTokow, ByVal value As String)
        _cableCatalog.CalculateCable(circuit, layMethod:=value)
    End Sub
    ' =============================================================
    ' Процедура: HandleGroupCablesChange
    ' =============================================================
    ''' <summary>
    ''' Обработва промяна на броя
    ''' паралелно положени кабели
    ''' за даден токов кръг.
    '''
    ''' След промяна извършва ново изчисление
    ''' на кабелните параметри,
    ''' като отчита новия брой кабели.
    ''' </summary>
    ''' <param name="circuit">
    ''' Токовият кръг,
    ''' за който се променя
    ''' броят кабели.
    ''' </param>
    ''' <param name="value">
    ''' Новият брой кабели,
    ''' въведен или избран от потребителя.
    ''' </param>
    Public Sub HandleGroupCablesChange(ByVal circuit As clsTokow, ByVal value As String)
        _cableCatalog.CalculateCable(circuit, Broj_Cable:=value)
    End Sub
    ' =============================================================
    ' Процедура: HandleBusbarChange
    ' =============================================================
    ''' <summary>
    ''' Обработва промяна на настройката за използване на шина.
    ''' Записва избраната стойност директно в свойството на токовия кръг.
    ''' Използва се при управление на параметрите на таблата и техните свързващи елементи.
    ''' </summary>
    ''' <param name="circuit">
    ''' Токовият кръг или елементът, към който принадлежи шината.
    ''' </param>
    ''' <param name="value">
    ''' Нова стойност на параметъра за шината.
    ''' </param>
    Public Sub HandleBusbarChange(ByVal circuit As clsTokow, ByVal value As String)
        If circuit Is Nothing Then Exit Sub
        ' Записваме новата стойност на тикчето
        Boolean.TryParse(value, circuit.Шина)
        ' Взимаме директно само маркираните консуматори за това табло
        Dim busbarCircuits = AppSettings.ListTokow.Where(Function(c) _
                             c.BuildingName = circuit.BuildingName AndAlso
                             c.Tablo = circuit.Tablo AndAlso
                             c.ТоковКръг <> "ОБЩО" AndAlso
                             c.Device <> "Разединител" AndAlso
                             c.Шина = True).ToList()
        ' Търсим дали ВЕЧЕ съществува такъв апарат за това табло
        Dim existingDisconnector = AppSettings.ListTokow.FirstOrDefault(Function(c) _
                                   c.BuildingName = circuit.BuildingName AndAlso
                                   c.Tablo = circuit.Tablo AndAlso
                                   c.Device = "Разединител")
        ' ПРОВЕРКА: Ако няма нито едно тикче (Шина = 0)
        If busbarCircuits.Count = 0 Then
            If existingDisconnector IsNot Nothing Then
                AppSettings.ListTokow.Remove(existingDisconnector)
            End If
            Exit Sub
        End If
        ' АКО ИМА ТИКЧЕТА: Намираме или създаваме Разединителя
        If existingDisconnector Is Nothing Then
            existingDisconnector = New clsTokow() With {
            .BuildingName = circuit.BuildingName,
            .Tablo = circuit.Tablo,
            .ТоковКръг = "Шина 1",
            .Консуматор = "Разединител Шина",
            .Device = "Разединител"}
            AppSettings.ListTokow.Add(existingDisconnector)
        End If
        ' -------------------------------------------------------------
        ' 6. ИЗЧИСЛЕНИЯ САМО ЗА МАРКИРАНИТЕ В ШИНАТА
        ' -------------------------------------------------------------
        Dim totalPower As Double = busbarCircuits.Sum(Function(c) c.Мощност)
        Dim totalLamps As Integer = busbarCircuits.Sum(Function(c) c.brLamp)
        Dim totalContacts As Integer = busbarCircuits.Sum(Function(c) c.brKontakt)
        Dim hasThreePhase As Boolean = busbarCircuits.Any(Function(c) c.Брой_Полюси = 3)
        Dim mostCommonPoles As Integer = If(hasThreePhase, 3, 1)
        Dim totalPhase As String = If(hasThreePhase, "L1,L2,L3", "L")

        Dim phaseCurrents As New Dictionary(Of String, Double) From {
                                 {"L1", 0}, {"L2", 0}, {"L3", 0}
                                 }

        ' Сумираме токовете по фази САМО за консуматорите на шината
        For Each busCircuit In busbarCircuits
            If busCircuit.Брой_Полюси = 3 OrElse busCircuit.Фаза = "L1,L2,L3" Then
                phaseCurrents("L1") += busCircuit.Ток
                phaseCurrents("L2") += busCircuit.Ток
                phaseCurrents("L3") += busCircuit.Ток
            Else
                Dim p As String = busCircuit.Фаза.Trim().ToUpper()
                If phaseCurrents.ContainsKey(p) Then
                    phaseCurrents(p) += busCircuit.Ток
                End If
            End If
        Next
        ' Актуализираме данните на разединителя
        With existingDisconnector
            .Табло_Родител = circuit.Tablo
            .Брой_Полюси = mostCommonPoles
            .Мощност = totalPower
            .Фаза = totalPhase
            .brLamp = totalLamps
            .brKontakt = totalContacts
            ' Записваме фазовите токове в служебните полета
            .RCD_Тип = "Ток фази"
            .RCD_Клас = "L1->" & phaseCurrents("L1").ToString("N2")
            .RCD_Ток = "L2->" & phaseCurrents("L2").ToString("N2")
            .RCD_Чувствителност = "L3->" & phaseCurrents("L3").ToString("N2")
        End With
        ' -------------------------------------------------------------
        ' 7. ИЗЧИСЛЯВАНЕ НА ТОКА И ИЗБОР ОТ КАТАЛОЗИТЕ
        ' -------------------------------------------------------------
        If hasThreePhase Then
            ' Можеш да ползваш своя calculation engine или баланс
            existingDisconnector.Ток = _calculationEngine.calc_Inom(totalPower, 3)
        Else
            existingDisconnector.Ток = _calculationEngine.calc_Inom(totalPower, 1)
        End If
        ' Автоматичен избор на апарат и кабел
        _disconnectorCatalog.CalculateDisconnector(existingDisconnector)
        _cableCatalog.CalculateCable(existingDisconnector)
    End Sub
End Class