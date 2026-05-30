Imports System.Drawing
Imports System.Windows.Forms

Public Class DataGridViewManager
    ' Всички технически каталози
    Private ReadOnly _disconnectorCatalog As DisconnectorCatalog
    Private ReadOnly _breakerCatalog As BreakerCatalog
    Private ReadOnly _cableCatalog As CableCatalog
    Private ReadOnly _rcdCatalog As RCDCatalog
    ' --- Специфични списъци за зареждане на ComboBox клетките в таблицата ---
    ' Вътрешен списък с опции за ComboBox-а на ред "Управление"
    Private ReadOnly _ComboItems_control As String() = {
                     "Фото реле",
                     "Стълбищен автомат",
                     "Импулсно реле",
                     "Контактор",
                     "Моторна защита",
                     "Моторен механизъм",
                     "Честотен регулатор",
                     "Електромер"
                     }
    ' Динамични списъци (зареждат се в паметта от каталога за прекъсвачи BreakerCatalog)
    Private _ComboItems_breakerType As New List(Of String)        ' Списък за ред "Тип на апарата" (сериите: iC60, EZ9 и др.)
    Private _ComboItems_breakerIn As New List(Of String)          ' Списък за ред "Номинален ток" (амперажите: 10А, 16А, 25А и др.)
    Private _ComboItems_breakerCurve As New List(Of String)       ' Списък за ред "Крива" (характеристиките: B, C, D)
    Private _ComboItems_breakerUnit As New List(Of String)        ' Списък за ред "Защитен блок" (Vigi модули и електронни блокове)

    ' Динамични списъци (зареждат се в паметта от каталога за кабели CableCatalog)
    Private _ComboItems_cableType As New List(Of String)          ' Списък за ред "Тип кабел" (материалите: СВТ, NYM, ПВ-А1 и др.)
    Private _ComboItems_cableInstallation As New List(Of String)  ' Списък за ред "Начин на монтаж" (кодовете по стандарт: A1, B1, C, E и др.)
    Private ReadOnly _ComboItems_cableEnvironment As String() = {
                     "Във въздух", "В земя"}                        ' Списък за ред "Начин на полагане" (Основна среда за охлаждане на кабела)

    ' Динамични списъци (зареждат се в паметта от каталога за разединители DisconnectorCatalog)
    Private _ComboItems_disconType As New List(Of String)       ' Списък за ред "Тип на апарата" при разединители (напр. Interpact INS, iSW)
    Private _ComboItems_disconIn As New List(Of String)         ' Списък за ред "Номинален ток" за разединители (напр. 40А, 63А, 100А...)


    ' Пазим референция към контролата, за да може целият клас да си я знае
    Private ReadOnly _dgv As DataGridView
    ''' <summary>
    ''' Конструктор на мениджъра за DataGridView.
    ''' </summary>
    Public Sub New(ByVal dgv As DataGridView,
                   ByVal disconnectorCat As DisconnectorCatalog,
                   ByVal breakerCat As BreakerCatalog,
                   ByVal cableCat As CableCatalog,
                   ByVal rcdCat As RCDCatalog)
        ' Запомняме таблицата веднъж завинаги в този клас
        Me._dgv = dgv
        ' Записваме референциите към данните и каталозите
        Me._disconnectorCatalog = disconnectorCat
        Me._breakerCatalog = breakerCat
        Me._cableCatalog = cableCat
        Me._rcdCatalog = rcdCat

        ' Зареждаме динамичните списъци за ComboBox клетките от съответните каталози    
        _ComboItems_cableType = cableCat.GetUniqueCableTypes()                      ' Взима уникалните типове кабели от каталога
        _ComboItems_breakerType = breakerCat.GetUniqueBreakerTypes("63А", "1p")     ' Взима уникалните типове прекъсвачи от каталога
        _ComboItems_breakerIn = breakerCat.GetUniqueBreakerCurrents("NSXm", "1p")   ' Взима уникалните амперажи от каталога
        _ComboItems_breakerCurve = breakerCat.GetUniqueBreakerCurves("NSXm", "1p")  ' Взима уникалните криви от каталога
        _ComboItems_breakerUnit = breakerCat.GetUniqueBreakerUnits("NSXm", "1p")    ' Взима уникалните защитни блокове от каталога

        _ComboItems_disconType = disconnectorCat.GetUniqueDisconnectorTypes("63А", "1p")       ' Взима уникалните типове разединители от каталога 
        _ComboItems_disconIn = disconnectorCat.GetUniqueDisconnectorCurrents("iSW", "1p")      ' Взима уникалните амперажи за разединители от каталога
    End Sub
    Public ReadOnly Property rowTemplate As List(Of Object())
        Get
            Return New List(Of Object()) From {
            New Object() {"Прекъсвач", "", "Text", Function(c As clsTokow) c.Breaker_Тип_Апарат},
            New Object() {"Изчислен ток", "A", "Text", Function(c As clsTokow) c.Ток.ToString("F2")},
            New Object() {"Тип на апарата", "", "Combo", Function(c As clsTokow) c.Breaker_Тип_Апарат},
            New Object() {"Номинален ток", "A", "Combo", Function(c As clsTokow) c.Breaker_Номинален_Ток},
            New Object() {"Изкл. възможн.", "kA", "Text", Function(c As clsTokow) c.Breaker_Изкл_Възможност},
            New Object() {"Крива", "", "Combo", Function(c As clsTokow) c.Breaker_Крива},
            New Object() {"Защитен блок", "", "Combo", Function(c As clsTokow) c.Breaker_Защитен_блок},
            New Object() {"Брой полюси", "бр.", "Text", Function(c As clsTokow) c.Брой_Полюси.ToString() & "p"},
            New Object() {"ДТЗ (RCD)", "", "Text", Function(c As clsTokow) c.RCD_Тип},
            New Object() {"ДТЗ Нула", "", "Text", Function(c As clsTokow) c.RCD_Нула},
            New Object() {"Вид на апарата", "", "Text", Function(c As clsTokow) c.RCD_Бранд},
            New Object() {"Клас на апарата", "", "Text", Function(c As clsTokow) c.RCD_Клас},
            New Object() {"ДТЗ(RCD) Ном. ток", "A", "Text", Function(c As clsTokow) c.RCD_Ток},
            New Object() {"Чувствителност", "mA", "Text", Function(c As clsTokow) c.RCD_Чувствителност},
            New Object() {"ДТЗ(RCD) полюси", "бр.", "Text", Function(c As clsTokow) c.RCD_Полюси},
            New Object() {"---------", "", "Text", Function(c As clsTokow) ""},
            New Object() {"Брой лампи", "бр.", "Text", Function(c As clsTokow) c.brLamp.ToString()},
            New Object() {"Брой контакти", "бр.", "Text", Function(c As clsTokow) c.brKontakt.ToString()},
            New Object() {"Инст. мощност", "kW", "Text", Function(c As clsTokow) c.Мощност.ToString("F2")},
            New Object() {"---------", "", "Text", Function(c As clsTokow) ""},
            New Object() {"Кабел", "", "Text", Function(c As clsTokow) c.Кабел_Тип},
            New Object() {"Начин на монтаж", "--", "Combo", Function(c As clsTokow) c.Кабел_Монтаж},
            New Object() {"Начин на полагане", "--", "Combo", Function(c As clsTokow) c.Кабел_Полагане},
            New Object() {"Паралелни кабели (фаза): ", "бр.", "Text", Function(c As clsTokow) c.Кабел_Брой_Фаза},
            New Object() {"Съседни кабели (група):", "бр.", "Text", Function(c As clsTokow) c.Кабел_Брой_Група},
            New Object() {"Тип кабел", "---", "Combo", Function(c As clsTokow) c.Кабел_Тип},
            New Object() {"Сечение", "mm²", "Text", Function(c As clsTokow) c.Кабел_Сечение},
            New Object() {"---------", "", "Text", Function(c As clsTokow) ""},
            New Object() {"Фаза", "", "Text", Function(c As clsTokow) c.Фаза},
            New Object() {"Консуматор", "---", "Text", Function(c As clsTokow) c.Консуматор},
            New Object() {"предназначение", "---", "Text", Function(c As clsTokow) c.предназначение},
            New Object() {"Управление", "---", "Combo", Function(c As clsTokow) c.Управление},
            New Object() {"---------", "", "Text", Function(c As clsTokow) ""},
            New Object() {"Шина", "---", "Check", Function(c As clsTokow) c.Шина},
            New Object() {"Постави ДТЗ (RCD)", "---", "Check", Function(c As clsTokow) c.ДТЗ_RCD}
        }
        End Get
    End Property
    ''' <summary>
    ''' Създава и конфигурира основната структура на DataGridView за показване на електрически табла и кръгове.
    ''' Изграждат се фиксирани колони (Параметър, Мерна единица, ОБЩО) и динамични колони според rowData.
    ''' Клетките се генерират според тип (ComboBox, CheckBox, TextBox), след което се прилага визуално форматиране.
    ''' </summary>
    Public Sub InitializeGridStructure()
        ' Изчистване на старата структура
        _dgv.Columns.Clear()
        _dgv.Rows.Clear()
        _dgv.RowHeadersVisible = False
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
        _dgv.Columns.Add(colParam)
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
        _dgv.Columns.Add(colUnit)
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
        _dgv.Columns.Add(colTotal)
        ' =====================================================
        ' 4. РЕДОВЕ: попълване от rowData шаблона
        ' =====================================================
        For Each row As Object() In rowTemplate
            Dim dgvRow As New DataGridViewRow()
            dgvRow.CreateCells(_dgv)
            ' Параметър
            dgvRow.Cells(0).Value = row(0)
            ' Мерна единица
            dgvRow.Cells(1).Value = row(1)
            ' Тип на клетките за останалите колони
            Dim cellType As String = row(2)
            ' Генериране на клетки за динамичните колони
            For colIndex As Integer = 2 To _dgv.Columns.Count - 2
                Dim cell As DataGridViewCell = Nothing
                Select Case cellType
                    Case "Combo"
                        cell = New DataGridViewComboBoxCell()
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
            _dgv.Rows.Add(dgvRow)
        Next
        ' =====================================================
        ' 5. НАСТРОЙКИ
        ' =====================================================
        _dgv.AllowUserToAddRows = False                                    ' Забранява на потребителя да добавя празен нов ред в края на таблицата
        _dgv.AllowUserToDeleteRows = False                                 ' Забранява на потребителя да изтрива редове с натискане на Delete
        _dgv.ReadOnly = False                                              ' Позволява редакция на клетките (важно за ComboBox и CheckBox)
        _dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None    ' Изключва автоматичното оразмеряване (разчита на зададен Width)
        _dgv.ColumnHeadersDefaultCellStyle.Font = New Font("Arial", 10, FontStyle.Bold) ' Задава шрифт Bold за заглавния ред
        _dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter ' Центрира текста в заглавията на колоните
        _dgv.ColumnHeadersHeight = 25                                      ' Фиксира височината на заглавната лента на 170 пиксела
        _dgv.RowTemplate.Height = 25                                       ' Задава стандартна височина на всеки нов ред с данни
        _dgv.BackgroundColor = Color.White                                 ' Променя цвета на фона на самата контрола (зад редовете) на бял
        _dgv.GridColor = Color.Gray                                        ' Задава сив цвят за линиите на мрежата между клетките
        _dgv.BorderStyle = BorderStyle.Fixed3D                             ' Прави рамката на цялата таблица да изглежда обемна (3D)
        _dgv.CellBorderStyle = DataGridViewCellBorderStyle.Single          ' Задава единична тънка линия за граница между отделните клетки
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
    Public Sub SetupDataGridView_Total()
        ' Списък с индекси на целевите колони, които ще се обработват
        Dim targetColumns As New List(Of Integer)
        ' Имена на колоните, които търсим в DataGridView
        Dim colNames() As String = {"colTotal", "colDiscon"}
        ' Проверка дали колоните съществуват в грида и взимане на индексите им
        For Each colName In colNames
            If _dgv.Columns.Contains(colName) Then
                targetColumns.Add(_dgv.Columns(colName).Index)
            End If
        Next
        ' Обхождане на всички редове в DataGridView
        For i As Integer = 0 To _dgv.Rows.Count - 1
            Dim dgvRow As DataGridViewRow = _dgv.Rows(i)
            ' Защита от несъответствие между визуалните редове и източника на данни
            If i >= rowTemplate.Count Then Continue For
            ' Взима съответния ред от източника на данни
            Dim data As Object() = rowTemplate(i)
            ' Тип на реда (определя какви клетки ще се създадат)
            Dim cellType As String = data(2).ToString()
            ' Обхождане на целевите колони (Total / Disconnector)
            For Each colIndex In targetColumns
                Dim specialCell As DataGridViewCell = Nothing
                ' Определяне на типа клетка според cellType
                Select Case cellType
                    Case "Combo"
                        ' Клетка тип ComboBox
                        Dim comboCell As New DataGridViewComboBoxCell()
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
    ''' Основна процедура за визуализиране на цялостната структура на избраното табло.
    ''' Извиква се при клик в TreeViewManager.
    ''' </summary>
    ''' <param name="dgv">Контролата DataGridView, която ще се управлява.</param>
    ''' <param name="selectedObject">Обектът от тип clsTokow, изпратен от TreeView.</param>
    Public Sub DisplayBoardStructure(ByVal selectedObject As clsTokow)
        ' --- ЗА ТЕСТВАНЕ НА ВРЪЗКАТА ---
        ' За момента я оставяме празна, за да тестваме само извикването и предаването на обекта.
        ' Когато му дойде времето, тук ще напишем логиката, която взема таблото, 
        ' намира кръговете му и чертае колоните (включително "ОБЩО" / Главния разединител).
        ' 2. Накрая викаш новата функция, за да изсипе данните в "ОБЩО"
        FillTotalColumnWithData(selectedObject)
    End Sub
    ''' <summary>
    ''' Обхожда редовете на таблицата и попълва колона "ОБЩО" с реалните стойности от обекта, 
    ''' изпълнявайки ламбда функциите, записани директно в rowTemplate.
    ''' </summary>
    Public Sub FillTotalColumnWithData(ByVal circuit As clsTokow)
        ' 1. Намираме индекса на колоната "ОБЩО" по име, за да сме сигурни къде пишем
        Dim totalColIndex As Integer = -1
        If _dgv.Columns.Contains("colTotal") Then
            totalColIndex = _dgv.Columns("colTotal").Index
        Else
            Exit Sub ' Ако колоната я няма, спираме, за да не гръмне кода
        End If
        ' Защита: Ако не ни е подаден обект, няма какво да наливаме
        If circuit Is Nothing Then Exit Sub
        ' 2. Въртим цикъл по всички редове на грида
        For rowIndex As Integer = 0 To _dgv.Rows.Count - 1
            ' Защита: да не излезем извън границите на шаблона
            If rowIndex >= rowTemplate.Count Then Continue For
            Dim rowData As Object() = rowTemplate(rowIndex)
            Dim cellType As String = rowData(2).ToString()
            Dim targetCell As DataGridViewCell = _dgv.Rows(rowIndex).Cells(totalColIndex)
            Dim parameterName As String = rowData(0).ToString()
            ' 3. МАГИЯТА: Вземаме ламбда функцията от индекс 3 и я изпълняваме, подавайки circuit
            ' 1. Кастваме към базовия Delegate клас в .NET
            Dim resolver As [Delegate] = DirectCast(rowData(3), [Delegate])
            ' 2. Изпълняваме я динамично, като подаваме circuit в масив от обекти
            Dim rawValue As Object = resolver.DynamicInvoke(circuit)
            ' 4. Записваме стойността в клетката съобразно нейния тип
            If cellType = "Combo" Then
                Dim comboCell = DirectCast(targetCell, DataGridViewComboBoxCell)
                ' ЕТО ГО СЪКРАЩЕНИЕТО: Викаме новата процедура да напълни списъка
                PopulateComboBoxItems(comboCell, parameterName, circuit)
                Dim valStr As String = If(rawValue IsNot Nothing, rawValue.ToString(), "")
                ' Проверяваме дали стойността от обекта съществува в списъка на ComboBox-а
                If comboCell.Items.Contains(valStr) Then
                    comboCell.Value = valStr
                ElseIf comboCell.Items.Count > 0 Then
                    comboCell.Value = comboCell.Items(0) ' Падащ вариант по подразбиране (напр. "---")
                End If
            Else
                ' За стандартен TextBox (String) или CheckBox (Boolean) - директно наливаме обекта
                targetCell.Value = rawValue
            End If
        Next
    End Sub
    ''' <summary>
    ''' Пълни Items на конкретна ComboBox клетка с филтрирани данни от каталозите според текущия токов кръг.
    ''' Автоматично превключва между каталог за Разединители (за главното) и Прекъсвачи (за токови кръгове).
    ''' </summary>
    Private Sub PopulateComboBoxItems(ByVal comboCell As DataGridViewComboBoxCell,
                                  ByVal parameterName As String,
                                  ByVal circuit As clsTokow)
        comboCell.Items.Clear()
        comboCell.Items.Add("---") ' Опция по подразбиране

        Dim currentPoles As String = circuit.Брой_Полюси.ToString() & "p"
        Dim currentBreakerType As String = circuit.Breaker_Тип_Апарат
        Dim currentIn As String = circuit.Breaker_Номинален_Ток

        ' ЛОГИКА ЗА АВТОМАТИЧНО РАЗПОЗНАВАНЕ:
        ' Ако името на токовия кръг съдържа "главен", "ввод" или устройството е дефинирано като главно -> ползваме Разединител
        ' 1. Първо си дефинираш променливата
        Dim isDisconnector As Boolean = False

        ' 2. Правиш правилната проверка
        If circuit.Device IsNot Nothing AndAlso
           (circuit.Device.ToLower().Contains("табло") _
           OrElse circuit.Device = "Разединител") Then
            isDisconnector = True
        End If
        Select Case parameterName
            Case "Тип на апарата"
                If isDisconnector Then
                    comboCell.Items.AddRange(_disconnectorCatalog.GetUniqueDisconnectorTypes(currentIn, currentPoles).ToArray())
                Else
                    comboCell.Items.AddRange(_breakerCatalog.GetUniqueBreakerTypes(currentIn, currentPoles).ToArray())
                End If
            Case "Номинален ток"
                If isDisconnector Then
                    comboCell.Items.AddRange(_disconnectorCatalog.GetUniqueDisconnectorCurrents(currentBreakerType, currentPoles).ToArray())
                Else
                    comboCell.Items.AddRange(_breakerCatalog.GetUniqueBreakerCurrents(currentBreakerType, currentPoles).ToArray())
                End If
            Case "Крива"
                ' Разединителите нямат крива на изключване, зареждаме само за прекъсвачи
                If Not isDisconnector Then
                    comboCell.Items.AddRange(_breakerCatalog.GetUniqueBreakerCurves(currentBreakerType, currentPoles).ToArray())
                End If
            Case "Защитен блок"
                ' Разединителите нямат защитен блок, зареждаме само за прекъсвачи
                If Not isDisconnector Then
                    comboCell.Items.AddRange(_breakerCatalog.GetUniqueBreakerUnits(currentBreakerType, currentPoles).ToArray())
                End If
            Case "Тип кабел"
                comboCell.Items.AddRange(_ComboItems_cableType.ToArray())
            Case "Начин на монтаж"
                comboCell.Items.AddRange(_ComboItems_cableInstallation.ToArray())
            Case "Начин на полагане"
                comboCell.Items.AddRange(_ComboItems_cableEnvironment.ToArray())
            Case "Управление"
                comboCell.Items.AddRange(_ComboItems_control.ToArray())
        End Select
    End Sub
End Class