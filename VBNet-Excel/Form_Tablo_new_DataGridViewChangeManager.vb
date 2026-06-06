Imports System.Reflection
Imports System.Windows.Forms



Public Class DataGridViewChangeManager
    ' Пазим препратки към шестте компонента
    Private _cableCatalog As CableCatalog
    Private _breakerCatalog As BreakerCatalog
    Private _disconnectorCatalog As DisconnectorCatalog
    Private _rcdCatalog As RCDCatalog
    Private _calculationEngine As ElectricalCalculationEngine
    Private _allCircuits As List(Of clsTokow)
    Dim _boardManager As New BoardStructureManager(_rcdCatalog)
    ''' <summary>
    ''' Конструкторът вече приема точно шестте компонента от формата
    ''' </summary>
    Public Sub New(ByVal breakerCat As BreakerCatalog,
                   ByVal disconnectorCat As DisconnectorCatalog,
                   ByVal rcdCat As RCDCatalog,
                   ByVal cableCat As CableCatalog,
                   ByVal calcEngine As ElectricalCalculationEngine) ' <-- Новите попълнения

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
    Public Sub UpdateCircuitProperty(ByVal tokow As clsTokow, ByVal procedureToExecute As String, ByVal newValue As String)
        ' ------------------------------------------------------------
        ' 🛡️ ОБЩА ПРОВЕРКА ЗА КОРЕКТНОСТ (За трите параметъра)
        ' ------------------------------------------------------------
        If tokow Is Nothing OrElse
           String.IsNullOrEmpty(procedureToExecute) OrElse
           String.IsNullOrEmpty(newValue) Then
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
        End Try
    End Sub
    ' =================================================================
    ' === СЪЩИНСКИ ПРОЦЕДУРИ (Пренесени от стария Select Case) ===
    ' =================================================================
    ''' <summary>
    ''' Сменя: "ДТЗ Нула" (Валидира текста и го записва в обекта)
    ''' Извиква се при: "HandleRcdZeroChange"
    ''' </summary>
    Public Sub HandleRcdZeroChange(ByVal panelCircuits As clsTokow, ByVal value As String)
        ' 1. Пускаме текста през санитарния филтър
        Dim validatedValue As String = ValidateRCDNulla(value)
        ' 2. Ако филтърът върне валиден резултат → записваме го в обекта
        If validatedValue IsNot Nothing Then panelCircuits.RCD_Нула = validatedValue
        Dim panels = AppSettings.ListTokow.Where(Function(t) New With {Key t.BuildingName, Key t.Tablo} = currentPanelKey) _
            .ToList()

        ' Викаме публичния метод от BoardStructureManager, за да пренареди ДТЗ-тата на таблото
        _boardManager.ProcessPanelRCDLogic(panels)
    End Sub
    ' =================================================================
    ' === ПОМОЩНИ ФУНКЦИИ (Валидации и санитарни филтри) ===
    ' =================================================================
    ''' <summary>
    ''' Санитарен филтър: Изчиства текста и проверява дали форматът отговаря на "N" + число (напр. N1, N2)
    ''' </summary>
    Private Function ValidateRCDNulla(ByVal inputValue As String) As String
        ' Проверка 1: Дали е празно и дали започва с "N"
        If String.IsNullOrEmpty(inputValue) OrElse Not inputValue.ToUpper().StartsWith("N") Then
            Return Nothing
        End If
        ' Извлечи числото след "N"
        Dim numberPart As String = inputValue.Substring(1).Trim()
        ' Премахни всичко, което НЕ е цифра
        numberPart = New String(numberPart.Where(Function(c) Char.IsDigit(c)).ToArray())
        ' Проверка 2: Дали има останали числа след чистенето
        If String.IsNullOrEmpty(numberPart) Then
            Return Nothing
        End If
        ' Проверка 3: Дали числото е валидно
        Dim rcdNumber As Integer
        If Not Integer.TryParse(numberPart, rcdNumber) Then
            Return Nothing
        End If
        ' Проверка 4: Дали числото е строго по-голямо од 0
        If rcdNumber <= 0 Then
            Return Nothing
        End If
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
    End Sub
End Class

'''' <summary>
'''' Събитие, което се изпълнява при промяна на стойност в клетка на DataGridView1.
''''
'''' Основна идея:
'''' Таблицата се използва като редактор на параметри за даден токов кръг.
'''' Първата колона съдържа името на параметъра (например "Тип на апарата",
'''' "Номинален ток", "Шина", "ДТЗ (RCD)" и др.), а останалите колони съдържат
'''' стойностите за конкретни кръгове или устройства.
''''
'''' Когато потребителят промени стойност:
'''' 1. Определя се редът и колоната на промяната.
'''' 2. От първата клетка на реда се взима името на параметъра.
'''' 3. От текущата клетка се взима новата стойност.
'''' 4. Чрез Select Case се определя какво действие трябва да се изпълни
''''    според типа на параметъра.
''''
'''' Забележка:
'''' Този метод служи като централизирана точка за обработка на всички
'''' промени в таблицата. Реалната логика за всяка настройка може да се
'''' добавя вътре в съответния Case.
'''' </summary>
''Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
''    If e.RowIndex < 0 OrElse e.ColumnIndex < 0 Then Return
''    If isUpdatingGrid Then Return
''    Try
''        isUpdatingGrid = True
''        Dim row As DataGridViewRow = DataGridView1.Rows(e.RowIndex)
''        Dim col As DataGridViewColumn = DataGridView1.Columns(e.ColumnIndex)
''         ------------------------------------------------------------
''         3) Името на параметъра се намира в първата колона (index 0)
''            Например:
''            "Тип на апарата"
''            "Номинален ток"
''            "Шина"
''            "ДТЗ (RCD)"
''         ------------------------------------------------------------
''        Dim paramName As String = row.Cells(0).Value?.ToString()
''         ------------------------------------------------------------
''         4) Новата стойност, въведена от потребителя в текущата клетка
''         ------------------------------------------------------------
''        Dim selectedValue As String = row.Cells(e.ColumnIndex).Value?.ToString()
''        Първо взимаш името
''        Dim currentCircuit As String = DataGridView1.Columns(e.ColumnIndex).HeaderText
''        После го предаваш
''        Dim tokow As strTokow = FindTokowByColumn(currentCircuit)
''        Dim Update As Boolean = True
''        If tokow.Device = "Разединител" OrElse
''               tokow.Device = "Съществуващ" OrElse
''               tokow.Device = "Резерва" Then Exit Sub
''        If tokow IsNot Nothing AndAlso Not String.IsNullOrEmpty(selectedValue) Then
''            Select Case paramName
''                Case "Тип на апарата"
''                    tokow.Breaker_Тип_Апарат = selectedValue
''                    Select Case tokow.Device
''                        Case "Разединител"
''                            Dim filteredDisco = Disconnectors.Where(Function(b) b.Type = selectedValue).ToList()
''                            Dim valuesForCombo = filteredDisco _
''                                                    .Select(Function(b) b.NominalCurrent.ToString()) _
''                                                    .Distinct() _
''                                                    .ToList()
''                            UpdateComboRow("Номинален ток", valuesForCombo, e.ColumnIndex)
''                        Case "Табло"
''                            Dim filteredDisco = Disconnectors.Where(Function(b) b.Type = selectedValue).ToList()
''                            Dim valuesForCombo = filteredDisco _
''                                                    .Select(Function(b) b.NominalCurrent.ToString()) _
''                                                    .Distinct() _
''                                                    .ToList()
''                            UpdateComboRow("Номинален ток", valuesForCombo, e.ColumnIndex)
''                            tokow.Device = tokow.Device '"Табло"
''                        Case Else
''                            Dim filteredBreakers = NewBreakers.Breakers.Where(Function(b) b.Series = selectedValue).ToList()
''                            If filteredBreakers.Count = 0 Then Exit Select
''                            tokow.Breaker_Изкл_Възможност = filteredBreakers.First().Ics_kA & "kA"
''                            Dim valuesForCombo = filteredBreakers _
''                                                    .Select(Function(b) b.NominalCurrent.ToString()) _
''                                                    .Distinct() _
''                                                    .ToList()
''                            UpdateComboRow("Номинален ток", valuesForCombo, e.ColumnIndex)
''                            Dim valuesCurve = filteredBreakers _
''                                                    .Select(Function(b) b.Curve.ToString()) _
''                                                    .Distinct() _
''                                                    .ToList()
''                            UpdateComboRow("Крива", valuesCurve, e.ColumnIndex)
''                            Dim valuesTripUnit = filteredBreakers _
''                                                    .Select(Function(b) b.TripUnit) _
''                                                    .Distinct() _
''                                                    .ToList()
''                            UpdateComboRow("Защитен блок", valuesTripUnit, e.ColumnIndex)
''                    End Select
''                Case "Постави ДТЗ (RCD)"
''                     ✅ Първо обнови tokow от клетката!
''                    tokow.ДТЗ_RCD = CBool(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value)
''                    HandleRCDCheckboxChange(tokow)
''                Case "Номинален ток"
''                    Тук може да се обработва промяна на номиналния ток
''                     на защитния апарат (например 10A, 16A, 20A...)
''                     1. Първо излизаме, ако няма стойност
''                    If selectedValue Is Nothing Then Exit Sub
''2. Вече сме сигурни, че имаме нещо, и правим сравнението
''                    If Val(selectedValue) >= Val(tokow.Breaker_Номинален_Ток) Then
''                        Всичко е точно, обновяваме стойността
''                        tokow.Breaker_Номинален_Ток = selectedValue
''                    Else
''                        Тук се намесваме с малко "приятелски" съвет
''                        Dim message As String = "Сигурен ли си в това, което правиш? " & vbCrLf &
''                                   "Избраният ток е по-малък от текущия." & vbCrLf &
''                                   "Честно казано, правиш простотия!" & vbCrLf &
''                                   "Искаш ли наистина да продължиш към Тъмната страна?"
''                        Dim result As DialogResult = MessageBox.Show(message, "Внимание: Инженерна мисъл в действие!",
''                                                       MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
''                        If result = DialogResult.Yes Then
''                            Потребителят е инат, записваме го
''                            tokow.Breaker_Номинален_Ток = selectedValue
''                        Else
''                            Спасихме положението!
''                            MessageBox.Show("Мъдро решение! Спести си един ремонт.", "Браво!", MessageBoxButtons.OK, MessageBoxIcon.Information)
''                        End If
''                    End If
''                    NewCables.CalculateCable(tokow,
''                                       Type:=tokow.Кабел_Тип,
''                                       layMethod:=If(tokow.Кабел_Полагане = "Във въздух", 0, 1),
''                                       mountMethod:=NewCables.GetMountMethodInfo(tokow.Кабел_Монтаж),
''                                       Broj_Cable:=tokow.Кабел_Брой_Група,
''                                       matType:=NewCables.GetCableTypeResult(tokow.Кабел_Тип)
''                                       )
''                Case "Съседни кабели (група):"
''                    tokow.Кабел_Брой_Група = selectedValue
''                    NewCables.CalculateCable(tokow,
''                                       Type:=tokow.Кабел_Тип,
''                                       layMethod:=If(tokow.Кабел_Полагане = "Във въздух", 0, 1),
''                                       mountMethod:=NewCables.GetMountMethodInfo(tokow.Кабел_Монтаж),
''                                       Broj_Cable:=tokow.Кабел_Брой_Група,
''                                       matType:=NewCables.GetCableTypeResult(tokow.Кабел_Тип)
''                                       )
''                Case "Консуматор"
''                    If tokow.Device <> "Табло" Then tokow.Консуматор = selectedValue
''                Case "предназначение"
''                    tokow.предназначение = selectedValue
''                    Update = False
''                    If tokow.Device = "Табло" Then
''                        Update = True
''                        Търсим число само след "Рпр" или "Рпр."
''                        Dim pattern As String = "Рпр\.?\s*=\s*(\d+([.,]\d+)?)"
''                        Dim match = System.Text.RegularExpressions.Regex.Match(tokow.предназначение, pattern)
''                        Dim value As Double = -1 ' -1 = няма валидна стойност
''                        Ако Regex намери число
''                        If match.Success Then
''                            Dim strValue As String = match.Groups(1).Value.Replace(",", ".")
''                            Double.TryParse(strValue, System.Globalization.NumberStyles.Any,
''                                System.Globalization.CultureInfo.InvariantCulture,
''                                value)
''                        End If
''                        Ако няма валидно число, проверяваме дали полето е просто число
''                        If value < 0 Then
''                            Dim onlyNumber As Double = 0
''                            If Double.TryParse(tokow.предназначение.Replace(",", "."),
''                                   System.Globalization.NumberStyles.Any,
''                                   System.Globalization.CultureInfo.InvariantCulture,
''                                   onlyNumber) Then
''                                value = onlyNumber
''                            End If
''                        End If
''                        Ако има валидно число, записваме в предназначение във формат Рпр.=(число)кW
''                        If value > 0 Then
''                            tokow.предназначение = "Рпр.=" & value.ToString("0.##") & "кW"
''                        Else
''                            Ако няма валидно число, задаваме по подразбиране
''                            tokow.предназначение = "Рпр.=15кW"
''                            value = 15
''                        End If
''                        Проверка да не делим на 0
''                        If tokow.Мощност <> 0 Then
''                            tokow.Консуматор = "Ке=" & (value / tokow.Мощност).ToString("0.00")
''                        Else
''                            tokow.Консуматор = "Ке=0"
''                        End If
''                    End If
''                Case "Управление"
''                    tokow.Управление = selectedValue
''                Case "Крива"
''                    tokow.Breaker_Крива = selectedValue
''                Case "Защитен блок"
''                    Обработка на параметър свързан със защитен модул
''                     или допълнителна защита
''                    tokow.Breaker_Защитен_блок = selectedValue
''                Case "Шина"
''                    Шина е Boolean → True = на отделна шина, False = основна шина
''                    tokow.Шина = CBool(selectedValue)
''                Case "ДТЗ (RCD)"
''                    Управление на дефектнотокова защита (RCD) 
''                             например включване / изключване на ДТЗ
''                Case "Начин на монтаж"
''                    Взимаме само текстовата част за комбобокса, 
''                     или подаваме целия списък, ако клетката е настроена за обекти
''                    Dim displayValues = NewCables.LiMountMethod.Select(Function(m) m.Text).ToList()
''                    UpdateComboRow("Начин на монтаж", displayValues, e.ColumnIndex)
''                Case "Начин на полагане"
''                    Правим прост списък само с двете опции
''                    Dim valuesLaying As New List(Of String) From {"Във въздух", "В земя"}
''                    If tokow.Кабел_Тип = "Al/R" Then
''                        tokow.Кабел_Полагане = "Във въздух"
''                        selectedValue = "Във въздух"
''                    End If
''                    NewCables.CalculateCable(tokow,
''                                       Type:=tokow.Кабел_Тип,
''                                       layMethod:=If(selectedValue = "Във въздух", 0, 1),
''                                       mountMethod:=NewCables.GetMountMethodInfo(tokow.Кабел_Монтаж),
''                                       Broj_Cable:=tokow.Кабел_Брой_Група,
''                                       matType:=NewCables.GetCableTypeResult(tokow.Кабел_Тип)
''                                       )
''                    Подаваме го към твоята процедура
''                    UpdateComboRow("Начин на полагане", valuesLaying, e.ColumnIndex)
''                Case "Тип кабел"
''                    Взимаме само уникалните имена на кабели от главния списък
''                    Dim uniqueCableTypes As List(Of String) = NewCables.CableTypesForCombo
''                    If selectedValue = "Al/R" Then
''                        tokow.Кабел_Полагане = "Във въздух"
''                        NewCables.CalculateCable(tokow,
''                                           Type:=selectedValue,
''                                           layMethod:=If(tokow.Кабел_Полагане = "Във въздух", 0, 1),
''                                           mountMethod:=NewCables.GetMountMethodInfo(tokow.Кабел_Монтаж),
''                                           Broj_Cable:=tokow.Кабел_Брой_Група,
''                                           matType:=NewCables.GetCableTypeResult(selectedValue)
''                                           )
''                        UpdateComboRow("Тип кабел", uniqueCableTypes, e.ColumnIndex)
''                        Правим прост списък само с двете опции
''                        Dim valuesLaying As New List(Of String) From {"Във въздух", "В земя"}
''                        UpdateComboRow("Начин на полагане", valuesLaying, e.ColumnIndex)
''                    Else
''                        Проверка дали стойността съществува в списъка
''                        NewCables.CalculateCable(tokow,
''                                           Type:=selectedValue,
''                                           layMethod:=If(tokow.Кабел_Полагане = "Във въздух", 0, 1),
''                                           mountMethod:=NewCables.GetMountMethodInfo(tokow.Кабел_Монтаж),
''                                           Broj_Cable:=tokow.Кабел_Брой_Група,
''                                           matType:=NewCables.GetCableTypeResult(selectedValue)
''                                           )
''                    End If
''                    Подаваме списъка към твоята процедура
''                    UpdateComboRow("Тип кабел", uniqueCableTypes, e.ColumnIndex)
''                Case "ДТЗ Нула"
''                    Dim inputValue As String = selectedValue?.ToString()
''                    Извикай функцията за валидация
''                    Dim validatedValue As String = ValidateRCDNulla(inputValue)
''                    Update = False
''                    Ако е валидно → запиши, иначе → върни старата стойност
''                    If validatedValue IsNot Nothing Then
''                        Update = True
''                        tokow.RCD_Нула = validatedValue
''                    End If
''            End Select
''            If Update Then UpdateCircuitColumn(tokow, col.Index, paramName)
''        End If
''    Finally
''        isUpdatingGrid = False
''    End Try
''End Sub
