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
            New Object() {"ДТЗ (RCD)", "", "Text", Function(c As clsTokow) ""},
            New Object() {"ДТЗ Нула", "", "Text", Function(c As clsTokow) c.RCD_Нула},
            New Object() {"Вид на апарата", "", "Text", Function(c As clsTokow) c.RCD_Тип},
            New Object() {"Клас на апарата", "", "Text", Function(c As clsTokow) c.RCD_Клас},
            New Object() {"ДТЗ(RCD) Ном. ток", "A", "Text", Function(c As clsTokow) c.RCD_Ток},
            New Object() {"Чувствителност", "mA", "Text", Function(c As clsTokow) c.RCD_Чувствителност},
            New Object() {"ДТЗ(RCD) полюси", "бр.", "Text", Function(c As clsTokow) c.RCD_Полюси},
            New Object() {"---------", "", "Text", Function(c As clsTokow) ""},
            New Object() {"Брой лампи", "бр.", "Text", Function(c As clsTokow) c.brLamp.ToString()},
            New Object() {"Брой контакти", "бр.", "Text", Function(c As clsTokow) c.brKontakt.ToString()},
            New Object() {"Инст. мощност", "kW", "Text", Function(c As clsTokow) c.Мощност.ToString("F2")},
            New Object() {"---------", "", "Text", Function(c As clsTokow) ""},
            New Object() {"Кабел", "", "Text", Function(c As clsTokow) ""},
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
    Public Sub DisplayBoardStructure(ByVal selectedObject As clsTokow)
        ' 1. Защита: Ако няма избран обект, няма какво да изобразяваме
        If selectedObject Is Nothing Then Exit Sub
        ' 2. МЕТЛАТА: Изчистваме абсолютно всичко след първите две колони (Параметър и Мярка)
        ' Така излитат и старите кръгове, и старата колона ОБЩО на един замах
        For i As Integer = _dgv.Columns.Count - 1 To 2 Step -1
            _dgv.Columns.RemoveAt(i)
        Next
        ' 3. Взимаме токовите кръгове за това табло
        Dim circuitsList As List(Of clsTokow) = selectedObject.GetMyCircuits()
        ' 4. ПЪРВИ ЦИКЪЛ: Добавяме и пълним колоните за всеки токов кръг (ако има такива)
        If circuitsList IsNot Nothing Then
            For Each circuit As clsTokow In circuitsList
                ' Извикваме споделения майстор
                CreateAndFillColumn(circuit)
            Next
        End If
        ' 5. ФИНАЛЕН АКОРД: Добавяме и пълним колоната за ОБЩО, 
        ' като използваме СЪЩАТА процедура, но подаваме главното табло
        CreateAndFillColumn(selectedObject)
    End Sub
    ''' <summary>
    ''' Пълни Items на конкретна ComboBox клетка по нейните индекси с филтрирани данни от каталозите.
    ''' </summary>
    Private Sub PopulateComboBoxItems(ByVal colIndex As Integer,
                                      ByVal rowIndex As Integer,
                                      ByVal parameterName As String,
                                      ByVal circuit As clsTokow)
        ' Извличаме клетката директно от грида и я кастваме към ComboBoxCell
        Dim comboCell As DataGridViewComboBoxCell = DirectCast(_dgv.Rows(rowIndex).Cells(colIndex), DataGridViewComboBoxCell)
        comboCell.Items.Clear()
        comboCell.Items.Add("---") ' Опция по подразбиране
        ' 1. Вземане на текущите филтри от обекта
        Dim currentPoles As String = If(circuit.Брой_Полюси > 0, circuit.Брой_Полюси.ToString() & "p", "1p")
        Dim currentBreakerType As String = circuit.Breaker_Тип_Апарат
        Dim currentIn As String = circuit.Breaker_Номинален_Ток
        ' 2. Логика за автоматично разпознаване на апарата (Главен или Кръг)
        Dim isDisconnector As Boolean = False
        If circuit.Device.ToLower().Contains("табло") OrElse
       circuit.Device = "Разединител" OrElse
       String.IsNullOrEmpty(circuit.ТоковКръг) Then
            isDisconnector = True
        End If
        ' 3. Наливане на опциите според параметъра
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
                If Not isDisconnector Then
                    comboCell.Items.AddRange(_breakerCatalog.GetUniqueBreakerCurves(currentBreakerType, currentPoles).ToArray())
                End If
            Case "Защитен блок"
                If Not isDisconnector Then
                    comboCell.Items.AddRange(_breakerCatalog.GetUniqueBreakerUnits(currentBreakerType, currentPoles).ToArray())
                End If
        ' --- Кабели и Управление ---
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
    ''' <summary>
    ''' Споделена процедура, която създава колона в края на грида и я конфигурира/пълни 1:1 по rowTemplate.
    ''' </summary>
    Private Sub CreateAndFillColumn(ByVal circuit As clsTokow)
        If circuit Is Nothing Then Exit Sub
        ' 1. Определяме името и заглавието на колоната спрямо обекта
        Dim columnName As String = ""
        Dim columnHeader As String = ""
        ' Ако свойството "ТоковКръг" е празно или нищо, значи това е главното табло (ОБЩО)
        If String.IsNullOrEmpty(circuit.ТоковКръг) Then
            columnName = "colTotal"
            columnHeader = "ОБЩО"
        Else
            ' Ако има име (напр. "Кръг 1"), правим динамично име на колоната
            columnName = "col_" & circuit.ТоковКръг
            columnHeader = circuit.ТоковКръг
        End If
        ' 2. Създаваме самата колона в края на Grid-a
        Dim newCol As New DataGridViewColumn()
        With newCol
            .Name = columnName
            .HeaderText = columnHeader
            .CellTemplate = New DataGridViewTextBoxCell() ' Базово раждане като TextBox
            .Width = 90
            .SortMode = DataGridViewColumnSortMode.NotSortable
        End With
        Dim colIndex As Integer = _dgv.Columns.Add(newCol)
        ' 3. Структура: Конфигуриране на типовете клетки (Combo, Check, Text)
        SetupDataGridView_ColumnStructure(colIndex)
        ' 4. Каталози: Наливане на опциите в ComboBox клетките
        FillColumnComboBoxOptions(colIndex, circuit)
        ' 5. Данни: Наливане на реалните стойности от AutoCAD обекта
        FillColumnValues(colIndex, circuit)
    End Sub
    Public Sub SetupDataGridView_ColumnStructure(ByVal colIndex As Integer)
        If colIndex < 0 OrElse colIndex >= _dgv.Columns.Count Then Exit Sub
        For i As Integer = 0 To _dgv.Rows.Count - 1
            Dim dgvRow As DataGridViewRow = _dgv.Rows(i)
            If i >= rowTemplate.Count Then Continue For
            Dim data As Object() = rowTemplate(i)
            Dim parameterName As String = data(0).ToString()
            Dim cellType As String = data(2).ToString()
            ' 1. СЪЗДАВАНЕ НА ПРАВИЛНАТА КЛЕТКА
            Dim specialCell As DataGridViewCell = Nothing
            Select Case cellType
                Case "Combo"
                    Dim comboCell As New DataGridViewComboBoxCell()
                    ' Показва стрелката винаги, за да личи че е падащо меню
                    comboCell.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox
                    ' === ТОВА ПРЕМАХВА СИВИТЕ БОРДЕРИ И РАМКИ НА БУТОНА ===
                    comboCell.FlatStyle = FlatStyle.Flat
                    ' Настройваме стрелката и бордера да се държат като модерна уеб форма:
                    ' Бордерът ще е тънък и сив, а стрелката ще се появява дискретно при посочване/редакция
                    comboCell.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox
                    specialCell = comboCell
                Case "Check"
                    specialCell = New DataGridViewCheckBoxCell()
                    specialCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                    specialCell.Value = False
                Case Else
                    specialCell = New DataGridViewTextBoxCell()
                    specialCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            End Select
            ' Записваме новата клетка в колоната
            dgvRow.Cells(colIndex) = specialCell
            ' 2. ВИЗУАЛНО ФОРМАТИРАНЕ НА НОВАТА КЛЕТКА
            Select Case parameterName
                Case "---------"
                    specialCell.Style.BackColor = Color.FromArgb(220, 220, 220)
                    specialCell.ReadOnly = True
                    specialCell.Value = "" ' Чистим чертичките в клетките с данни
                Case "Прекъсвач", "ДТЗ (RCD)", "Кабел", "Фаза"
                    '' Боядисваме клетката в заглавния син цвят
                    'specialCell.Style.BackColor = Color.FromArgb(180, 200, 255)
                    'specialCell.Style.Font = New Font("Arial", 10, FontStyle.Bold)
                    specialCell.ReadOnly = True
                Case Else
                    ' Специфичен цвят за колона ОБЩО, ако редът не е заглавен
                    If _dgv.Columns(colIndex).Name = "colTotal" Then
                    End If
            End Select
        Next
    End Sub
    ''' <summary>
    ''' Обхожда редовете на конкретна колона и зарежда опциите в ComboBox клетките от каталозите.
    ''' </summary>
    Private Sub FillColumnComboBoxOptions(ByVal colIndex As Integer, ByVal circuit As clsTokow)
        For rowIndex As Integer = 0 To _dgv.Rows.Count - 1
            If rowIndex >= rowTemplate.Count Then Continue For
            Dim data As Object() = rowTemplate(rowIndex)
            Dim parameterName As String = data(0).ToString()
            Dim cellType As String = data(2).ToString()
            ' Пълним списъците само ако клетката е Combo
            If cellType = "Combo" Then
                PopulateComboBoxItems(colIndex, rowIndex, parameterName, circuit)
            End If
        Next
    End Sub
    ''' <summary>
    ''' Обхожда редовете на конкретна колона и налива реалните стойности от обекта circuit.
    ''' </summary>
    Private Sub FillColumnValues(ByVal colIndex As Integer, ByVal circuit As clsTokow)
        For rowIndex As Integer = 0 To _dgv.Rows.Count - 1
            If rowIndex >= rowTemplate.Count Then Continue For
            Dim data As Object() = rowTemplate(rowIndex)
            Dim cellType As String = data(2).ToString()
            Dim resolver As [Delegate] = DirectCast(data(3), [Delegate])
            Dim targetCell As DataGridViewCell = _dgv.Rows(rowIndex).Cells(colIndex)
            Dim rawValue As Object = resolver.DynamicInvoke(circuit)
            If cellType = "Combo" Then
                Dim comboCell = DirectCast(targetCell, DataGridViewComboBoxCell)
                Dim valStr As String = If(rawValue IsNot Nothing, rawValue.ToString(), "")
                ' Избираме стойността, ако я има в заредения каталог
                If comboCell.Items.Contains(valStr) Then
                    comboCell.Value = valStr
                Else
                    comboCell.Value = "---"
                End If
            ElseIf cellType = "Check" Then
                If TypeOf rawValue Is Boolean Then
                    targetCell.Value = DirectCast(rawValue, Boolean)
                Else
                    targetCell.Value = False
                End If
            Else
                targetCell.Value = rawValue
            End If
        Next
    End Sub
End Class