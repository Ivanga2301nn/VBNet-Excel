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
                     "Няма",
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
    Public ReadOnly Property rowTemplate As List(Of String())
        Get
            Return New List(Of String()) From {
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
        For Each row As String() In rowTemplate
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
            Dim data As String() = rowTemplate(i)
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
    Private Sub SetupComboBoxCell(ByVal cell As DataGridViewCell, ByVal parameterName As String, ByVal isDisconnector As Boolean)
        ' Безопасно преобразуване - ако клетката не е ComboBox, методът просто ще излезе без грешка
        Dim comboCell = TryCast(cell, DataGridViewComboBoxCell)
        If comboCell Is Nothing Then Exit Sub
        comboCell.Items.Clear()
        ' Пълним клетката със съответния нов именуван списък, който вече имаме в класа
        Select Case parameterName
            Case "Тип на апарата"
                If isDisconnector Then
                    ' Тук ще викаме списъка от разединителите
                Else
                    comboCell.Items.AddRange(_ComboItems_breakerType.ToArray())
                End If
            Case "Номинален ток"
                If isDisconnector Then
                    ' Списък от разединители
                Else
                    comboCell.Items.AddRange(_ComboItems_breakerIn.ToArray())
                End If
            Case "Крива"
                If isDisconnector Then
                    comboCell.Items.Add("-")
                Else
                    comboCell.Items.AddRange(_ComboItems_breakerCurve.ToArray())
                End If
            Case "Управление"
                If isDisconnector Then
                    comboCell.Items.Add("Няма")
                Else
                    ' Ползваме нашето ново име!
                    comboCell.Items.AddRange(_ComboItems_control)
                End If
            Case "Тип кабел"
                comboCell.Items.AddRange(_ComboItems_cableType.ToArray())
            Case "Начин на монтаж"
                comboCell.Items.AddRange(_ComboItems_cableInstallation.ToArray())
            Case "Начин на полагане"
                comboCell.Items.AddRange(_ComboItems_cableEnvironment.ToArray())
        End Select
        ' Задаваме начална стойност по подразбиране
        If comboCell.Items.Count > 0 Then
            comboCell.Value = comboCell.Items(0)
        End If
        ' Задаваме визуално клетката да изглежда като истински ComboBox през цялото време
        comboCell.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox
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
    End Sub
End Class