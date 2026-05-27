Imports System.Drawing
Imports System.Windows.Forms

Public Module DataGridViewConfig
    ' =====================================================
    ' РЕДОВЕ: Параметри с мерни единици и типове клетки
    ' =====================================================
    ' Структура: {Параметър, Мерна единица, Тип клетка}
    ' Тип клетка: "Text", "Combo", "Check"
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
    Public Sub InitializeGridStructure(ByVal dgv As DataGridView)
        ' Изчистване на старата структура
        dgv.Columns.Clear()
        dgv.Rows.Clear()
        dgv.RowHeadersVisible = False
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
        dgv.Columns.Add(colParam)
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
        dgv.Columns.Add(colUnit)
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
        dgv.Columns.Add(colTotal)
        ' =====================================================
        ' 4. РЕДОВЕ: попълване от rowData шаблона
        ' =====================================================
        For Each row As String() In rowTemplate
            Dim dgvRow As New DataGridViewRow()
            dgvRow.CreateCells(dgv)
            ' Параметър
            dgvRow.Cells(0).Value = row(0)
            ' Мерна единица
            dgvRow.Cells(1).Value = row(1)
            ' Тип на клетките за останалите колони
            Dim cellType As String = row(2)
            ' Генериране на клетки за динамичните колони
            For colIndex As Integer = 2 To dgv.Columns.Count - 2
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
            dgv.Rows.Add(dgvRow)
        Next
        ' =====================================================
        ' 5. НАСТРОЙКИ
        ' =====================================================
        dgv.AllowUserToAddRows = False                                    ' Забранява на потребителя да добавя празен нов ред в края на таблицата
        dgv.AllowUserToDeleteRows = False                                 ' Забранява на потребителя да изтрива редове с натискане на Delete
        dgv.ReadOnly = False                                              ' Позволява редакция на клетките (важно за ComboBox и CheckBox)
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None    ' Изключва автоматичното оразмеряване (разчита на зададен Width)
        dgv.ColumnHeadersDefaultCellStyle.Font = New Font("Arial", 10, FontStyle.Bold) ' Задава шрифт Bold за заглавния ред
        dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter ' Центрира текста в заглавията на колоните
        dgv.ColumnHeadersHeight = 25                                      ' Фиксира височината на заглавната лента на 170 пиксела
        dgv.RowTemplate.Height = 25                                       ' Задава стандартна височина на всеки нов ред с данни
        dgv.BackgroundColor = Color.White                                 ' Променя цвета на фона на самата контрола (зад редовете) на бял
        dgv.GridColor = Color.Gray                                        ' Задава сив цвят за линиите на мрежата между клетките
        dgv.BorderStyle = BorderStyle.Fixed3D                             ' Прави рамката на цялата таблица да изглежда обемна (3D)
        dgv.CellBorderStyle = DataGridViewCellBorderStyle.Single          ' Задава единична тънка линия за граница между отделните клетки
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
    Public Sub SetupDataGridView_Total(ByVal dgv As DataGridView)
        ' Списък с индекси на целевите колони, които ще се обработват
        Dim targetColumns As New List(Of Integer)
        ' Имена на колоните, които търсим в DataGridView
        Dim colNames() As String = {"colTotal", "colDiscon"}
        ' Проверка дали колоните съществуват в грида и взимане на индексите им
        For Each colName In colNames
            If dgv.Columns.Contains(colName) Then
                targetColumns.Add(dgv.Columns(colName).Index)
            End If
        Next
        ' Обхождане на всички редове в DataGridView
        For i As Integer = 0 To dgv.Rows.Count - 1
            Dim dgvRow As DataGridViewRow = dgv.Rows(i)
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
End Module


' Допълнителна логика според първия елемент на реда (data(0))
'Select Case data(0).ToString()
'    Case "Управление"
'        ' Специална инициализация за управление
'        SetupComboBoxCell(comboCell, data(0), True)
'    Case "Тип на апарата"
'        ' Попълване от списък с прекъсвачи
'        comboCell.Items.Clear()
'        comboCell.Items.AddRange(Disconnectors_For_combo.ToArray())
'    Case "Номинален ток"
'        ' Попълване на номинални токове
'        comboCell.Items.Clear()
'        comboCell.Items.AddRange(Discon_Tok_For_combo.ToArray())
'    Case "Тип кабел"
'        ' Попълване на кабели
'        comboCell.Items.Clear()
'        comboCell.Items.AddRange(NewCables.CableTypesForCombo.ToArray())
'    Case "Начин на монтаж"
'        ' Попълване от дефиниран списък с монтажи
'        comboCell.Items.Clear()
'        comboCell.Items.AddRange(NewCables.LiMountMethod.Select(Function(m) m.Text).ToArray())
'    Case "Начин на полагане"
'        ' Фиксиран списък за полагане
'        Dim valuesLaying As New List(Of String) From {"Във въздух", "В земя"}
'        comboCell.Items.Clear()
'        comboCell.Items.AddRange(valuesLaying.ToArray())
'    Case Else
'        ' Няма дефинирана логика за този случай
'End Select



'' =============================================================
'' Процедура: SetupComboBoxCell
'' =============================================================
'' <summary>
'' Настройва DataGridViewComboBoxCell
'' според подадения параметър.
''
'' Процедурата:
'' - изчиства старите елементи
'' - зарежда нови стойности
'' - задава начална стойност
'' - настройва визуалния режим на ComboBox
''
'' Използва се при динамично изграждане
'' на DataGridView редове и клетки.
'' </summary>
''
'' <param name="cell">
'' Клетка от DataGridView,
'' която се преобразува в ComboBoxCell.
'' </param>
''
'' <param name="parameter">
'' Определя какъв тип данни
'' да бъдат заредени в ComboBox.
'' </param>
''
'' <param name="Discon">
'' Флаг указващ дали:
'' - се работи с разединители
'' - или със стандартни прекъсвачи
'' </param>
'Private Sub SetupComboBoxCell(cell As DataGridViewCell,
'                              parameter As String,
'                              Discon As Boolean)
'    ' Преобразува стандартната клетка
'    ' към DataGridViewComboBoxCell
'    Dim comboCell As DataGridViewComboBoxCell =
'        CType(cell, DataGridViewComboBoxCell)
'    ' Изчиства старите елементи,
'    ' за да не се натрупват дублирания
'    comboCell.Items.Clear()
'    ' Зареждане на различни стойности
'    ' според типа параметър
'    Select Case parameter
'        Case "Тип на апарата"
'            ' Зарежда:
'            ' - разединители
'            ' - или прекъсвачи
'            If Discon Then
'                comboCell.Items.AddRange(Disconnectors_For_combo.ToArray())
'            Else
'                comboCell.Items.AddRange(NewBreakers.Breakers_For_combo.ToArray())
'            End If
'        Case "Номинален ток"
'            ' Зарежда стандартни номинални токове
'            If Discon Then
'                comboCell.Items.AddRange(Discon_Tok_For_combo.ToArray())
'            Else
'                comboCell.Items.AddRange("6", "10", "16", "20", "25", "32", "40", "50", "63")
'            End If
'        Case "Крива"
'            ' Зарежда токови характеристики
'            ' на прекъсвачите
'            If Discon Then
'                comboCell.Items.AddRange("-")
'            Else
'                comboCell.Items.AddRange("C", "B", "D")
'            End If
'        Case "Управление"
'            ' Зарежда типове управление
'            ' и допълнителни устройства
'            If Discon Then
'                comboCell.Items.AddRange("Няма")
'            Else
'                comboCell.Items.AddRange(
'                    "Няма",
'                    "Фото реле",
'                    "Стълбищен автомат",
'                    "Импулсно реле",
'                    "Контактор",
'                    "Моторна защита",
'                    "Моторен механизъм",
'                    "Честотен регулатор",
'                    "Електромер"
'                )
'            End If
'        Case "Тип кабел"
'            ' Зарежда наличните типове кабели
'            comboCell.Items.AddRange(_cableCatalog.CableTypesForCombo.ToArray())
'            ' Възможност за бъдещо разширяване
'            ' с допълнителни типове кабели
'    End Select
'    ' Ако има заредени елементи:
'    ' задава първия като начална стойност
'    '
'    ' Това предотвратява празни клетки
'    If comboCell.Items.Count > 0 Then
'        comboCell.Value = comboCell.Items(0)
'    End If
'    ' Настройка на режима на ComboBox
'    '
'    ' DropDown:
'    ' - позволява избор
'    ' - позволява и ръчно въвеждане
'    comboCell.DisplayStyle = ComboBoxStyle.DropDown
'End Sub
