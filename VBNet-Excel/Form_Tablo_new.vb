Imports System.Collections.Generic
Imports System.Drawing
Imports System.Security.Cryptography
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.DatabaseServices.Filters
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Internal.DatabaseServices
Imports Autodesk.AutoCAD.PlottingServices
Imports Autodesk.AutoCAD.Runtime
Imports VBNet_Excel.Form_Tablo_new

'Imports System.IO
'Imports System.Windows.Forms

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
        Me.Height = 850
        Me.Width = 1600
        SetCatalog()
        GetKonsumatori()
        CreateTokowList()
        InitializeBlockConfigs()
        CalculateCircuitLoads()


        BuildTreeViewFromKonsumatori()
        SetupDataGridView()
    End Sub
    Dim PI As Double = 3.1415926535897931
    Dim cu As CommonUtil = New CommonUtil()
    Private ListKonsumator As New List(Of strKonsumator)
    ' Списък за токовите кръгове
    Dim ListTokow As New List(Of strTokow)
    ' ============================================================
    ' КАТАЛОЖНИ ПРОМЕНЛИВИ (на ниво форма)
    ' ============================================================
    Private BlockConfigs As New List(Of BlockConfig)
    Private Breakers As New List(Of BreakerInfo)
    Private Cables As New List(Of CableInfo)
    Private Busbars_Cu As New List(Of BusbarInfo)
    Private Busbars_Al As New List(Of BusbarInfo)
    Private RCD_Catalog As New List(Of RCDInfo)
    Private IcableDict As New Dictionary(Of String, Integer())
    Private Kable_Size_L As String()
    Private Kable_Size_N As String()
    Dim Disconnectors As New List(Of DisconnectorInfo)
    ' ============================================================
    ' КАТАЛОЖНИ СТРУКТУРИ
    ' ============================================================
    Public Structure BreakerInfo
        Dim NominalCurrent As Integer        ' 6, 10, 16, 20...
        Dim Type As String                   ' "EZ9", "C120", "NSX", "MTZ"
        Dim Brand As String                  ' "Schneider"
        Dim Poles As Integer                 ' 1, 2, 3, 4
        Dim Curve As String                  ' "B", "C", "D"
        Dim BreakingCapacity As Integer      ' 6000, 10000, 25000... (A)
    End Structure
    Public Structure DisconnectorInfo
        Dim NominalCurrent As Integer    ' 20, 32, 40...
        Dim Type As String               ' "iSW", "INS", "IN"
        Dim Brand As String              ' "Acti9", "Easy9"
        Dim Poles As Integer             ' 2, 3, 4
    End Structure
    Public Structure CableInfo
        Dim Section As String            ' "3x2.5", "5x16"
        Dim Material As String           ' "Cu", "Al"
        Dim Conductors As Integer        ' 2, 3, 4, 5
        Dim CurrentCapacity As Integer   ' Допустим ток
        Dim InstallationMethod As String ' "air", "ground"
    End Structure
    Public Structure BusbarInfo
        Dim CurrentCapacity As Integer   ' Допустим ток
        Dim Section As String            ' "30x4", "50x5"
        Dim Material As String           ' "Cu", "Al"
    End Structure
    Public Structure RCDInfo
        Dim NominalCurrent As Integer    ' 25, 40, 63...
        Dim Type As String               ' "AC", "A", "F"
        Dim Poles As String              ' "2p", "4p"
        Dim Sensitivity As Integer       ' 10, 30, 100, 300, 500 (mA)
        Dim DeviceType As String         ' "RCCB", "RCBO", "iID"
    End Structure
    Public Structure strKonsumator
        Dim Name As String              ' Име на блока
        Dim ID_Block As ObjectId        ' Връзка към AutoCAD
        Dim ТоковКръг As String         ' Токов кръг
        Dim strМОЩНОСТ As String        ' Мощност като текст от атрибут
        Dim doubМОЩНОСТ As Double       ' Мощност като число
        Dim ТАБЛО As String             ' Табло
        Dim Pewdn As String             ' Предназначение
        Dim PEWDN1 As String            ' Доп. предназначение
        Dim Dylvina_Led As Double       ' За LED ленти
        Dim Visibility As String        ' За динамични блокове
        Dim Phase As Integer            ' Брой фази (1, 3)
    End Structure
    Public Class strTokow
        ' ============================================================
        ' ИДЕНТИФИКАЦИЯ (Български имена за DataGridView)
        ' ============================================================
        Public Tablo As String                 ' Родителско табло
        Public ТоковКръг As String             ' Име/номер на кръга
        Public БройПолюси As Integer           ' 1 или 3 (ще ни трябва за избора на прекъсвач)
        ' ============================================================
        ' БРОЯЧИ
        ' ============================================================
        Public brLamp As Integer               ' Брой лампи
        Public brKontakt As Integer            ' Брой контакти
        ' ============================================================
        ' МОЩНОСТ И ТОК
        ' ============================================================
        Public Мощност As Double               ' kW (обща мощност)
        Public Ток As Double                   ' A (изчислен ток I = P/U)
        Public Фаза As String                  ' "1P", "3P", "L1", "L2", "L3"
        ' ============================================================
        ' КАБЕЛ
        ' ============================================================
        Public Кабел_Сечение As String         ' "3x2.5", "5x4"
        Public Кабел_Тип As String             ' "NYM", "YJV", "CBT"
        ' ============================================================
        ' ЗАЩИТА (ПРЕКЪСВАЧ)
        ' ============================================================
        Public Тип_Апарат As String            ' "EZ9", "C120", "NSX", "MTZ"
        Public Брой_Полюси As String           ' "1p", "3p"
        Public Крива As String                 ' "B", "C", "D"
        Public Номинален_Ток As String         ' "10A", "16A", "20A"...
        Public Изкл_Възможност As String       ' "6000A", "10000A"...
        ' ============================================================
        ' ДТЗ (RCD) - ОПЦИОНАЛНО
        ' ============================================================
        Public RCD_Тип As String               ' "AC", "A", "F"
        Public RCD_Чувствителност As String    ' "30mA", "100mA", "300mA"
        Public RCD_Ток As String               ' "25A", "40A", "63A"
        Dim RCD_Полюси As String            ' "2p", "4p"
        ' ============================================================
        ' КОНСУМАТОРИ В КРЪГА
        ' ============================================================
        Public Konsumator As List(Of strKonsumator)
    End Class
    Public Structure strTablo
        Dim countTablo As Integer
        Dim Name As String              ' "АП-1"
        Dim prevTablo As String         ' "Гл.Р.Т."
        Dim countTokKryg As Integer
        ' За TreeView групиране - ДОБАВЕНО:
        Dim Floor As String             ' "Първи етаж", "Подземен"
        Dim Building As String          ' "Сграда А" (по желание)
        Dim Tokowkryg As List(Of strTokow)  ' ПРОМЯНА: масив → List
        Dim TabloType As String
        ' Изчислени (за показване в TreeView)
        Dim TotalPower As Double        ' Сума от кръговете
        Dim SupplyCable As String       ' "NYM 5x16"
        ' Допълнителни за таблото (по желание)
        Dim Width As Integer
        Dim Height As Integer
        Dim IP_Rating As String
    End Structure
    ''' <summary>
    ''' Конфигурация за всеки тип блок
    ''' </summary>
    Public Class BlockConfig
        Public BlockNames As List(Of String)      ' Възможни имена на блока
        Public Category As String                 ' "Lamp", "Contact", "Device", "Panel"
        Public DefaultPoles As String             ' "1p" или "3p"
        Public DefaultCable As String             ' "3x1.5", "3x2.5", "5x2.5"
        Public DefaultBreaker As String           ' "10", "16", "20"
        Public VisibilityRules As List(Of VisRule) ' Правила за visibility
    End Class
    ''' <summary>
    ''' Правило за конкретна visibility стойност
    ''' </summary>
    Public Class VisRule
        Public VisibilityPattern As String        ' "3P", "Двугнездов", "Проточен"
        Public Poles As String                    ' "1p" или "3p"
        Public Cable As String                    ' "3x2.5", "5x4"
        Public Breaker As String                  ' "16", "25", "32"
        Public Phase As String                    ' "L" или "L1,L2,L3"
        Public ContactCount As Integer            ' Колко контакта добавя (1, 2, 3)
    End Class


    Private Sub GetKonsumatori()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim edt As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database
        Dim SelectedSet = cu.GetObjects("INSERT", "Изберете блок")
        If SelectedSet Is Nothing Then
            MsgBox("НЕ Е маркиран нито един блок.")
            Exit Sub
        End If
        Dim blkRecId As ObjectId = ObjectId.Null
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                Dim acBlkTbl As BlockTable
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                ToolStripProgressBar1.Maximum = SelectedSet.Count
                ToolStripProgressBar1.Value = 0
                For Each sObj As SelectedObject In SelectedSet
                    blkRecId = sObj.ObjectId
                    Dim Kons As New strKonsumator
                    Dim acBlkRef As BlockReference = DirectCast(acTrans.GetObject(blkRecId, OpenMode.ForRead), BlockReference)
                    Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                    Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection
                    Kons.Name = (CType(acBlkRef.DynamicBlockTableRecord.GetObject(OpenMode.ForRead), BlockTableRecord)).Name
                    Kons.ID_Block = blkRecId
                    For Each attId As ObjectId In attCol
                        Dim dbObj As DBObject = acTrans.GetObject(attId, OpenMode.ForRead)
                        ' Преобразува обекта в AttributeReference, за да работи с атрибутите.
                        Dim acAttRef As AttributeReference = dbObj
                        If acAttRef.Tag = "ТАБЛО" Then Kons.ТАБЛО = acAttRef.TextString
                        If acAttRef.Tag = "КРЪГ" Then Kons.ТоковКръг = acAttRef.TextString
                        If acAttRef.Tag = "Pewdn" Then Kons.Pewdn = acAttRef.TextString
                        If acAttRef.Tag = "PEWDN1" Then Kons.PEWDN1 = acAttRef.TextString
                        If acAttRef.Tag = "LED" Then Kons.strМОЩНОСТ = acAttRef.TextString
                        If acAttRef.Tag = "МОЩНОСТ" Then Kons.strМОЩНОСТ = acAttRef.TextString
                    Next

                    Dim Visibility As String = ""
                    For Each prop As DynamicBlockReferenceProperty In props
                        If prop.PropertyName = "Visibility1" Then Kons.Visibility = prop.Value
                        If prop.PropertyName = "Visibility" Then Kons.Visibility = prop.Value
                        If prop.PropertyName = "Дължина" Then Kons.Dylvina_Led = prop.Value
                    Next

                    Kons.doubМОЩНОСТ = CalcPower(Kons.strМОЩНОСТ, Kons.Dylvina_Led)
                    ProcessBlockByType(Kons, Kons.Name, Kons.Visibility)

                    If Kons.doubМОЩНОСТ > 0 Then ListKonsumator.Add(Kons)

                    ToolStripProgressBar1.Value += 1
                Next
                acTrans.Commit()
            Catch ex As Exception
                MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
                acTrans.Abort()
            End Try
        End Using
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
    ''' Групира консуматорите по табла и изгражда TreeView
    ''' Структура: Етаж → Табло
    ''' </summary>
    Private Sub BuildTreeViewFromKonsumatori()
        ' 1. Изчисти старото дърво
        TreeView.Nodes.Clear()

        ' 2. Групирай консуматорите по ТАБЛО
        Dim panels = ListKonsumator.GroupBy(Function(k) k.ТАБЛО).ToList()

        ' 3. За всяко табло създай възел
        For Each panelGroup In panels
            Dim panelName As String = panelGroup.Key

            ' Пропусни ако няма име на табло
            If String.IsNullOrEmpty(panelName) Then
                panelName = "Без табло"
            End If

            ' Брой кръгове в това табло (уникални ТоковКръг стойности)
            Dim circuitCount As Integer = panelGroup.Select(Function(k) k.ТоковКръг).Distinct().Count()

            ' Обща мощност (сума от всички консуматори)
            Dim totalPower As Double = panelGroup.Sum(Function(k) k.doubМОЩНОСТ)

            ' Създай възел за таблото
            Dim panelNode As New TreeNode()
            panelNode.Text = GetPanelNodeText(panelName, circuitCount, totalPower)
            panelNode.Tag = panelGroup.ToList()  ' Запази консуматорите за по-късно

            ' Добави възела в TreeView
            TreeView.Nodes.Add(panelNode)
        Next

        ' 4. Разгъни дървото
        TreeView.ExpandAll()

    End Sub
    ''' <summary>
    ''' Форматира текста за възела на таблото
    ''' </summary>
    Private Function GetPanelNodeText(panelName As String,
                                  circuitCount As Integer,
                                  totalPower As Double) As String
        Dim powerkW As Double = totalPower / 1000.0
        Dim circuitText As String = If(circuitCount = 1, "кръг", "кръга")
        Return $"{panelName} ({circuitCount} кръга, {powerkW:F3}kW)"
    End Function
    Private Sub SetupDataGridView()
        DataGridView1.Columns.Clear()
        DataGridView1.Rows.Clear()
        ' =====================================================
        ' 1. ПЪРВА КОЛОНА: Параметри
        ' =====================================================
        Dim colParam As New DataGridViewTextBoxColumn()
        colParam.Name = "colParameter"
        colParam.HeaderText = "Параметър"
        colParam.Width = 150
        colParam.Frozen = True
        colParam.DefaultCellStyle.Font = New Drawing.Font("Arial", 10, FontStyle.Bold)
        colParam.DefaultCellStyle.BackColor = Color.FromArgb(200, 220, 255)
        DataGridView1.Columns.Add(colParam)
        ' =====================================================
        ' 2. ВТОРА КОЛОНА: Мерни единици (дименсии)
        ' =====================================================
        Dim colUnit As New DataGridViewTextBoxColumn()
        colUnit.Name = "colUnit"
        colUnit.HeaderText = ""
        colUnit.Width = 50
        colUnit.Frozen = True
        colUnit.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        colUnit.DefaultCellStyle.Font = New Drawing.Font("Arial", 10, FontStyle.Italic)
        colUnit.DefaultCellStyle.ForeColor = Color.Gray
        DataGridView1.Columns.Add(colUnit)
        ' =====================================================
        ' 3. КОЛОНИ ЗА ТОКОВИ КРЪГОВЕ
        ' =====================================================
        For i As Integer = 1 To 20
            Dim col As New DataGridViewTextBoxColumn()
            col.Name = $"colCircuit{i}"
            col.HeaderText = i.ToString()
            col.Width = 70
            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns.Add(col)
        Next
        ' Колона ОБЩО
        Dim colTotal As New DataGridViewTextBoxColumn()
        colTotal.Name = "colTotal"
        colTotal.HeaderText = "ОБЩО"
        colTotal.Width = 90
        colTotal.DefaultCellStyle.Font = New Drawing.Font("Arial", 10, FontStyle.Bold)
        colTotal.DefaultCellStyle.BackColor = Color.FromArgb(230, 240, 255)
        colTotal.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns.Add(colTotal)
        ' =====================================================
        ' 4. РЕДОВЕ: Параметри с мерни единици и типове клетки
        ' =====================================================
        ' Структура: {Параметър, Мерна единица, Тип клетка}
        ' Тип клетка: "Text", "Combo", "Check"
        Dim rowData As String()() = {
        New String() {"---------", "", "Text"},
        New String() {"Прекъсвач", "", "Text"},
        New String() {"---------", "", "Text"},
        New String() {"Изчислен ток", "A", "Text"},
        New String() {"Тип на апарата", "", "Combo"},
        New String() {"Номинален ток", "A", "Combo"},
        New String() {"Изкл. възможн.", "", "Text"},
        New String() {"Крива", "", "Text"},
        New String() {"Брой полюси", "бр.", "Text"},
        New String() {"---------", "", "Text"},
        New String() {"ДТЗ", "", "Text"},
        New String() {"Вид на апарата", "", "Text"},
        New String() {"Клас на апарата", "", "Text"},
        New String() {"Номинален ток", "A", "Text"},
        New String() {"Изкл. възможн.", "mA", "Text"},
        New String() {"Брой полюси", "бр.", "Text"},
        New String() {"---------", "", "Text"},
        New String() {"Брой лампи", "бр.", "Text"},
        New String() {"Брой контакти", "бр.", "Text"},
        New String() {"Инст. мощност", "kW", "Text"},
        New String() {"Тип кабел", "---", "Text"},
        New String() {"Сечение", "---", "Text"},
        New String() {"Фаза", "---", "Text"},
        New String() {"Консуматор", "---", "Text"},
        New String() {"---------", "", "Text"},
        New String() {"Управление", "---", "Combo"},
        New String() {"---------", "", "Text"},
        New String() {"Шина", "---", "Check"},
        New String() {"ДТЗ (RCD)", "---", "Check"}
    }
        For Each row As String() In rowData
            Dim dgvRow As New DataGridViewRow()
            dgvRow.CreateCells(DataGridView1)
            ' Колона 0: Параметър
            dgvRow.Cells(0).Value = row(0)
            ' Колона 1: Мерна единица
            dgvRow.Cells(1).Value = row(1)
            ' Определи типа на клетката
            Dim cellType As String = row(2)
            ' За колони 2+ (кръгове), създай подходящ тип клетка
            For colIndex As Integer = 2 To DataGridView1.Columns.Count - 1
                Dim cell As DataGridViewCell = Nothing
                Select Case cellType
                    Case "Combo"
                        cell = New DataGridViewComboBoxCell()
                        SetupComboBoxCell(cell, row(0))
                    Case "Check"
                        cell = New DataGridViewCheckBoxCell()
                        cell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                    Case Else
                        cell = New DataGridViewTextBoxCell()
                        cell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                End Select

                dgvRow.Cells(colIndex) = cell
            Next
            ' Оцветяване
            If row(0) = "---------" Then
                dgvRow.DefaultCellStyle.BackColor = Color.FromArgb(220, 220, 220)
            ElseIf row(0) = "Прекъсвач" OrElse row(0) = "ДТЗ" Then
                dgvRow.DefaultCellStyle.BackColor = Color.FromArgb(180, 200, 255)
                dgvRow.DefaultCellStyle.Font = New Drawing.Font("Arial", 9, FontStyle.Bold)
            End If

            DataGridView1.Rows.Add(dgvRow)
        Next
        ' =====================================================
        ' 5. НАСТРОЙКИ
        ' =====================================================
        DataGridView1.AllowUserToAddRows = False
        DataGridView1.AllowUserToDeleteRows = False
        DataGridView1.ReadOnly = False  ' Позволи редакция за ComboBox и CheckBox
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        DataGridView1.ColumnHeadersDefaultCellStyle.Font = New Drawing.Font("Arial", 10, FontStyle.Bold)
        DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.ColumnHeadersHeight = 40
        DataGridView1.RowTemplate.Height = 25
        DataGridView1.BackgroundColor = Color.White
        DataGridView1.GridColor = Color.Gray
        DataGridView1.BorderStyle = BorderStyle.Fixed3D
        DataGridView1.CellBorderStyle = DataGridViewCellBorderStyle.Single
    End Sub
    Private Sub SetupComboBoxCell(cell As DataGridViewCell, parameter As String)
        Dim comboCell As DataGridViewComboBoxCell = CType(cell, DataGridViewComboBoxCell)
        Select Case parameter
            Case "Тип на апарата"
                comboCell.Items.AddRange("EZ9 MCB", "EZ9 RCCB", "EZ9 RCBO", "iSW", "A9 MCB")
            Case "Номинален ток"
                comboCell.Items.AddRange("6A", "10A", "16A", "20A", "25A", "32A", "40A", "50A", "63A")
            Case "Управление"
                comboCell.Items.AddRange("Ръчно", "Автоматично", "Дистанционно")
        End Select
        comboCell.DisplayStyle = ComboBoxStyle.DropDownList
    End Sub
    ' Добави това след SetupDataGridView()
    Private Sub DataGridView1_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
        ' Игнорирай грешките при форматиране
        e.ThrowException = False
        e.Cancel = True
    End Sub
    ' ============================================================
    ' ФУНКЦИЯ ЗА ЗАРЕЖДАНЕ НА КАТАЛОЗИТЕ
    ' ============================================================
    Private Sub SetCatalog()
        'Допустими токови натоварвания на кабели и проводници
        IcableDict = New Dictionary(Of String, Integer()) From {
    {"0_0_0", {20, 27, 36, 45, 63, 82, 113, 138, 168, 210, 262, 307, 352, 405, 482}},   ' Меден 1 жило положен във въздух
    {"0_0_1", {19, 25, 34, 43, 59, 79, 105, 126, 157, 199, 246, 285, 326, 374, 445}},   ' Меден 3 жилен положен във въздух
    {"0_1_0", {0, 0, 28, 38, 48, 63, 85, 105, 127, 165, 205, 235, 270, 315, 375}},      ' Алуминиев 1 жило положен във въздух
    {"0_1_1", {0, 20, 26, 34, 43, 64, 82, 100, 119, 152, 185, 215, 245, 285, 338}},     ' Алуминиев 3 жилен положен във въздух
    {"1_0_0", {29, 38, 49, 62, 83, 104, 136, 162, 192, 236, 285, 322, 363, 410, 475}},  ' Меден 1 жило положен във земя
    {"1_0_1", {25, 34, 45, 55, 76, 96, 126, 151, 178, 225, 270, 306, 346, 390, 458}},   ' Меден 3 жилен положен във земя
    {"1_1_0", {0, 0, 38, 52, 63, 82, 106, 128, 150, 186, 220, 250, 282, 320, 375}},     ' Алуминиев 1 жило положен във земя
    {"1_1_1", {0, 25, 32, 42, 53, 75, 92, 110, 134, 170, 210, 245, 274, 310, 360}}      ' Алуминиев 3 жилен положен във земя
    }
        ' Общ масив за всички сечения ФАЗОВОТО ЖИЛО
        Kable_Size_L = {"1,5", "2,5", "4,0", "6,0", "10", "16", "25", "35", "50", "70", "95", "120", "150", "185", "240"}
        ' Общ масив за всички сечения НУЛЕВОТО ЖИЛО
        Kable_Size_N = {"0", "0", "0", "0", "0", "0", "16", "16", "25", "35", "50", "70", "70", "95", "120"}
        ' Речник за всички автоматични прекъсвачи
        Breakers = New List(Of BreakerInfo) From {
New BreakerInfo With {.NominalCurrent = 6, .Type = "EZ9", .Brand = "Schneider", .Poles = 1, .Curve = "B", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 10, .Type = "EZ9", .Brand = "Schneider", .Poles = 1, .Curve = "B", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 16, .Type = "EZ9", .Brand = "Schneider", .Poles = 1, .Curve = "B", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 20, .Type = "EZ9", .Brand = "Schneider", .Poles = 1, .Curve = "B", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 25, .Type = "EZ9", .Brand = "Schneider", .Poles = 1, .Curve = "B", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 32, .Type = "EZ9", .Brand = "Schneider", .Poles = 1, .Curve = "B", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 40, .Type = "EZ9", .Brand = "Schneider", .Poles = 1, .Curve = "B", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 50, .Type = "EZ9", .Brand = "Schneider", .Poles = 1, .Curve = "B", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 63, .Type = "EZ9", .Brand = "Schneider", .Poles = 1, .Curve = "B", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 6, .Type = "EZ9", .Brand = "Schneider", .Poles = 1, .Curve = "C", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 10, .Type = "EZ9", .Brand = "Schneider", .Poles = 1, .Curve = "C", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 16, .Type = "EZ9", .Brand = "Schneider", .Poles = 1, .Curve = "C", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 20, .Type = "EZ9", .Brand = "Schneider", .Poles = 1, .Curve = "C", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 25, .Type = "EZ9", .Brand = "Schneider", .Poles = 1, .Curve = "C", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 32, .Type = "EZ9", .Brand = "Schneider", .Poles = 1, .Curve = "C", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 40, .Type = "EZ9", .Brand = "Schneider", .Poles = 1, .Curve = "C", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 50, .Type = "EZ9", .Brand = "Schneider", .Poles = 1, .Curve = "C", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 63, .Type = "EZ9", .Brand = "Schneider", .Poles = 1, .Curve = "C", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 6, .Type = "EZ9", .Brand = "Schneider", .Poles = 1, .Curve = "D", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 10, .Type = "EZ9", .Brand = "Schneider", .Poles = 1, .Curve = "D", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 16, .Type = "EZ9", .Brand = "Schneider", .Poles = 1, .Curve = "D", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 20, .Type = "EZ9", .Brand = "Schneider", .Poles = 1, .Curve = "D", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 25, .Type = "EZ9", .Brand = "Schneider", .Poles = 1, .Curve = "D", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 32, .Type = "EZ9", .Brand = "Schneider", .Poles = 1, .Curve = "D", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 40, .Type = "EZ9", .Brand = "Schneider", .Poles = 1, .Curve = "D", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 50, .Type = "EZ9", .Brand = "Schneider", .Poles = 1, .Curve = "D", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 63, .Type = "EZ9", .Brand = "Schneider", .Poles = 1, .Curve = "D", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 6, .Type = "EZ9", .Brand = "Schneider", .Poles = 3, .Curve = "B", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 10, .Type = "EZ9", .Brand = "Schneider", .Poles = 3, .Curve = "B", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 16, .Type = "EZ9", .Brand = "Schneider", .Poles = 3, .Curve = "B", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 20, .Type = "EZ9", .Brand = "Schneider", .Poles = 3, .Curve = "B", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 25, .Type = "EZ9", .Brand = "Schneider", .Poles = 3, .Curve = "B", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 32, .Type = "EZ9", .Brand = "Schneider", .Poles = 3, .Curve = "B", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 40, .Type = "EZ9", .Brand = "Schneider", .Poles = 3, .Curve = "B", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 50, .Type = "EZ9", .Brand = "Schneider", .Poles = 3, .Curve = "B", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 6, .Type = "EZ9", .Brand = "Schneider", .Poles = 3, .Curve = "C", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 10, .Type = "EZ9", .Brand = "Schneider", .Poles = 3, .Curve = "C", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 16, .Type = "EZ9", .Brand = "Schneider", .Poles = 3, .Curve = "C", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 20, .Type = "EZ9", .Brand = "Schneider", .Poles = 3, .Curve = "C", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 25, .Type = "EZ9", .Brand = "Schneider", .Poles = 3, .Curve = "C", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 32, .Type = "EZ9", .Brand = "Schneider", .Poles = 3, .Curve = "C", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 40, .Type = "EZ9", .Brand = "Schneider", .Poles = 3, .Curve = "C", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 50, .Type = "EZ9", .Brand = "Schneider", .Poles = 3, .Curve = "C", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 63, .Type = "EZ9", .Brand = "Schneider", .Poles = 3, .Curve = "C", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 6, .Type = "EZ9", .Brand = "Schneider", .Poles = 3, .Curve = "D", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 10, .Type = "EZ9", .Brand = "Schneider", .Poles = 3, .Curve = "D", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 16, .Type = "EZ9", .Brand = "Schneider", .Poles = 3, .Curve = "D", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 20, .Type = "EZ9", .Brand = "Schneider", .Poles = 3, .Curve = "D", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 25, .Type = "EZ9", .Brand = "Schneider", .Poles = 3, .Curve = "D", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 32, .Type = "EZ9", .Brand = "Schneider", .Poles = 3, .Curve = "D", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 40, .Type = "EZ9", .Brand = "Schneider", .Poles = 3, .Curve = "D", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 50, .Type = "EZ9", .Brand = "Schneider", .Poles = 3, .Curve = "D", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 63, .Type = "EZ9", .Brand = "Schneider", .Poles = 3, .Curve = "D", .BreakingCapacity = 6000},
New BreakerInfo With {.NominalCurrent = 80, .Type = "C120", .Brand = "Schneider", .Poles = 1, .Curve = "C", .BreakingCapacity = 10000},
New BreakerInfo With {.NominalCurrent = 100, .Type = "C120", .Brand = "Schneider", .Poles = 1, .Curve = "C", .BreakingCapacity = 10000},
New BreakerInfo With {.NominalCurrent = 125, .Type = "C120", .Brand = "Schneider", .Poles = 1, .Curve = "C", .BreakingCapacity = 10000},
New BreakerInfo With {.NominalCurrent = 80, .Type = "C120", .Brand = "Schneider", .Poles = 1, .Curve = "D", .BreakingCapacity = 10000},
New BreakerInfo With {.NominalCurrent = 100, .Type = "C120", .Brand = "Schneider", .Poles = 1, .Curve = "D", .BreakingCapacity = 10000},
New BreakerInfo With {.NominalCurrent = 125, .Type = "C120", .Brand = "Schneider", .Poles = 1, .Curve = "D", .BreakingCapacity = 10000},
New BreakerInfo With {.NominalCurrent = 80, .Type = "C120", .Brand = "Schneider", .Poles = 3, .Curve = "C", .BreakingCapacity = 10000},
New BreakerInfo With {.NominalCurrent = 100, .Type = "C120", .Brand = "Schneider", .Poles = 3, .Curve = "C", .BreakingCapacity = 10000},
New BreakerInfo With {.NominalCurrent = 125, .Type = "C120", .Brand = "Schneider", .Poles = 3, .Curve = "C", .BreakingCapacity = 10000},
New BreakerInfo With {.NominalCurrent = 80, .Type = "C120", .Brand = "Schneider", .Poles = 3, .Curve = "D", .BreakingCapacity = 10000},
New BreakerInfo With {.NominalCurrent = 100, .Type = "C120", .Brand = "Schneider", .Poles = 3, .Curve = "D", .BreakingCapacity = 10000},
New BreakerInfo With {.NominalCurrent = 125, .Type = "C120", .Brand = "Schneider", .Poles = 3, .Curve = "D", .BreakingCapacity = 10000},
New BreakerInfo With {.NominalCurrent = 100, .Type = "NSX", .Brand = "Schneider", .Poles = 3, .Curve = "C", .BreakingCapacity = 25000},
New BreakerInfo With {.NominalCurrent = 125, .Type = "NSX", .Brand = "Schneider", .Poles = 3, .Curve = "C", .BreakingCapacity = 25000},
New BreakerInfo With {.NominalCurrent = 160, .Type = "NSX", .Brand = "Schneider", .Poles = 3, .Curve = "C", .BreakingCapacity = 25000},
New BreakerInfo With {.NominalCurrent = 200, .Type = "NSX", .Brand = "Schneider", .Poles = 3, .Curve = "C", .BreakingCapacity = 25000},
New BreakerInfo With {.NominalCurrent = 250, .Type = "NSX", .Brand = "Schneider", .Poles = 3, .Curve = "C", .BreakingCapacity = 25000},
New BreakerInfo With {.NominalCurrent = 320, .Type = "NSX", .Brand = "Schneider", .Poles = 3, .Curve = "C", .BreakingCapacity = 25000},
New BreakerInfo With {.NominalCurrent = 400, .Type = "NSX", .Brand = "Schneider", .Poles = 3, .Curve = "C", .BreakingCapacity = 25000},
New BreakerInfo With {.NominalCurrent = 500, .Type = "NSX", .Brand = "Schneider", .Poles = 3, .Curve = "C", .BreakingCapacity = 25000},
New BreakerInfo With {.NominalCurrent = 630, .Type = "NSX", .Brand = "Schneider", .Poles = 3, .Curve = "C", .BreakingCapacity = 25000},
New BreakerInfo With {.NominalCurrent = 100, .Type = "NSX", .Brand = "Schneider", .Poles = 3, .Curve = "D", .BreakingCapacity = 25000},
New BreakerInfo With {.NominalCurrent = 125, .Type = "NSX", .Brand = "Schneider", .Poles = 3, .Curve = "D", .BreakingCapacity = 25000},
New BreakerInfo With {.NominalCurrent = 160, .Type = "NSX", .Brand = "Schneider", .Poles = 3, .Curve = "D", .BreakingCapacity = 25000},
New BreakerInfo With {.NominalCurrent = 200, .Type = "NSX", .Brand = "Schneider", .Poles = 3, .Curve = "D", .BreakingCapacity = 25000},
New BreakerInfo With {.NominalCurrent = 250, .Type = "NSX", .Brand = "Schneider", .Poles = 3, .Curve = "D", .BreakingCapacity = 25000},
New BreakerInfo With {.NominalCurrent = 320, .Type = "NSX", .Brand = "Schneider", .Poles = 3, .Curve = "D", .BreakingCapacity = 25000},
New BreakerInfo With {.NominalCurrent = 400, .Type = "NSX", .Brand = "Schneider", .Poles = 3, .Curve = "D", .BreakingCapacity = 25000},
New BreakerInfo With {.NominalCurrent = 500, .Type = "NSX", .Brand = "Schneider", .Poles = 3, .Curve = "D", .BreakingCapacity = 25000},
New BreakerInfo With {.NominalCurrent = 630, .Type = "NSX", .Brand = "Schneider", .Poles = 3, .Curve = "D", .BreakingCapacity = 25000},
New BreakerInfo With {.NominalCurrent = 800, .Type = "MTZ", .Brand = "Schneider", .Poles = 3, .Curve = "C", .BreakingCapacity = 50000},
New BreakerInfo With {.NominalCurrent = 1000, .Type = "MTZ", .Brand = "Schneider", .Poles = 3, .Curve = "C", .BreakingCapacity = 50000},
New BreakerInfo With {.NominalCurrent = 1250, .Type = "MTZ", .Brand = "Schneider", .Poles = 3, .Curve = "C", .BreakingCapacity = 50000},
New BreakerInfo With {.NominalCurrent = 1600, .Type = "MTZ", .Brand = "Schneider", .Poles = 3, .Curve = "C", .BreakingCapacity = 50000},
New BreakerInfo With {.NominalCurrent = 2000, .Type = "MTZ", .Brand = "Schneider", .Poles = 3, .Curve = "C", .BreakingCapacity = 50000},
New BreakerInfo With {.NominalCurrent = 2500, .Type = "MTZ", .Brand = "Schneider", .Poles = 3, .Curve = "C", .BreakingCapacity = 50000},
New BreakerInfo With {.NominalCurrent = 3200, .Type = "MTZ", .Brand = "Schneider", .Poles = 3, .Curve = "C", .BreakingCapacity = 50000},
New BreakerInfo With {.NominalCurrent = 4000, .Type = "MTZ", .Brand = "Schneider", .Poles = 3, .Curve = "C", .BreakingCapacity = 50000},
New BreakerInfo With {.NominalCurrent = 5000, .Type = "MTZ", .Brand = "Schneider", .Poles = 3, .Curve = "C", .BreakingCapacity = 50000},
New BreakerInfo With {.NominalCurrent = 6300, .Type = "MTZ", .Brand = "Schneider", .Poles = 3, .Curve = "C", .BreakingCapacity = 50000},
New BreakerInfo With {.NominalCurrent = 800, .Type = "MTZ", .Brand = "Schneider", .Poles = 3, .Curve = "D", .BreakingCapacity = 50000},
New BreakerInfo With {.NominalCurrent = 1000, .Type = "MTZ", .Brand = "Schneider", .Poles = 3, .Curve = "D", .BreakingCapacity = 50000},
New BreakerInfo With {.NominalCurrent = 1250, .Type = "MTZ", .Brand = "Schneider", .Poles = 3, .Curve = "D", .BreakingCapacity = 50000},
New BreakerInfo With {.NominalCurrent = 1600, .Type = "MTZ", .Brand = "Schneider", .Poles = 3, .Curve = "D", .BreakingCapacity = 50000},
New BreakerInfo With {.NominalCurrent = 2000, .Type = "MTZ", .Brand = "Schneider", .Poles = 3, .Curve = "D", .BreakingCapacity = 50000},
New BreakerInfo With {.NominalCurrent = 2500, .Type = "MTZ", .Brand = "Schneider", .Poles = 3, .Curve = "D", .BreakingCapacity = 50000},
New BreakerInfo With {.NominalCurrent = 3200, .Type = "MTZ", .Brand = "Schneider", .Poles = 3, .Curve = "D", .BreakingCapacity = 50000},
New BreakerInfo With {.NominalCurrent = 4000, .Type = "MTZ", .Brand = "Schneider", .Poles = 3, .Curve = "D", .BreakingCapacity = 50000},
New BreakerInfo With {.NominalCurrent = 5000, .Type = "MTZ", .Brand = "Schneider", .Poles = 3, .Curve = "D", .BreakingCapacity = 50000},
New BreakerInfo With {.NominalCurrent = 6300, .Type = "MTZ", .Brand = "Schneider", .Poles = 3, .Curve = "D", .BreakingCapacity = 50000}
}
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
        New RCDInfo With {.NominalCurrent = 25, .Type = "AC", .Poles = "2p", .Sensitivity = 30, .DeviceType = "RCCB"},
        New RCDInfo With {.NominalCurrent = 25, .Type = "AC", .Poles = "4p", .Sensitivity = 30, .DeviceType = "RCCB"},
        New RCDInfo With {.NominalCurrent = 40, .Type = "AC", .Poles = "2p", .Sensitivity = 30, .DeviceType = "RCCB"},
        New RCDInfo With {.NominalCurrent = 40, .Type = "AC", .Poles = "4p", .Sensitivity = 30, .DeviceType = "RCCB"},
        New RCDInfo With {.NominalCurrent = 63, .Type = "AC", .Poles = "2p", .Sensitivity = 30, .DeviceType = "RCCB"},
        New RCDInfo With {.NominalCurrent = 63, .Type = "AC", .Poles = "4p", .Sensitivity = 30, .DeviceType = "RCCB"},
        New RCDInfo With {.NominalCurrent = 16, .Type = "AC", .Poles = "2p", .Sensitivity = 30, .DeviceType = "RCBO"},
        New RCDInfo With {.NominalCurrent = 20, .Type = "AC", .Poles = "2p", .Sensitivity = 30, .DeviceType = "RCBO"},
        New RCDInfo With {.NominalCurrent = 25, .Type = "AC", .Poles = "2p", .Sensitivity = 30, .DeviceType = "RCBO"},
        New RCDInfo With {.NominalCurrent = 32, .Type = "AC", .Poles = "2p", .Sensitivity = 30, .DeviceType = "RCBO"},
        New RCDInfo With {.NominalCurrent = 40, .Type = "AC", .Poles = "2p", .Sensitivity = 30, .DeviceType = "RCBO"},
        New RCDInfo With {.NominalCurrent = 25, .Type = "AC", .Poles = "2p", .Sensitivity = 300, .DeviceType = "iID"},
        New RCDInfo With {.NominalCurrent = 40, .Type = "AC", .Poles = "2p", .Sensitivity = 300, .DeviceType = "iID"},
        New RCDInfo With {.NominalCurrent = 63, .Type = "AC", .Poles = "2p", .Sensitivity = 300, .DeviceType = "iID"},
        New RCDInfo With {.NominalCurrent = 25, .Type = "AC", .Poles = "4p", .Sensitivity = 300, .DeviceType = "iID"},
        New RCDInfo With {.NominalCurrent = 40, .Type = "AC", .Poles = "4p", .Sensitivity = 300, .DeviceType = "iID"},
        New RCDInfo With {.NominalCurrent = 63, .Type = "AC", .Poles = "4p", .Sensitivity = 300, .DeviceType = "iID"}
    }
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
            Case name.Contains("LED_DENIMA"), name.Contains("LED_LENTA"), name.Contains("LED_ULTRALUX"), name.Contains("LED_ЛУНА"),
         name.Contains("АВАРИЯ"), name.Contains("БОЙЛЕРНО ТАБЛО"), name.Contains("ЛАМПИ_СПАЛНЯ"), name.Contains("ЛИНИЯ МХЛ"),
         name.Contains("ЛУМИНЕСЦЕНТНА"), name.Contains("МЕТАЛХАЛОГЕННА"), name.Contains("ПЛАФОНИ"),
         name.Contains("АПЛИК"), name.Contains("ПЕНДЕЛ"), name.Contains("ЛАМПИОН"),
         name.Contains("НАСТОЛНА ЛАМПА"), name.Contains("ФАСАДНО"), name.Contains("БАНСКИ АПЛИК"), name.Contains("ДАТЧИК"),
         name.Contains("ФОТОДАТЧИК"), name.Contains("ПОЛИЛЕЙ"), name.Contains("ПРОЖЕКТОР")

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
            .Konsumator = g.ToList()
        }).ToList()
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
            .BlockNames = New List(Of String) From {"LED_DENIMA", "LED_LENTA", "LED_ULTRALUX", "LED_ULTRALUX_100", "LED_ULTRALUX_НОВ", "LED_ЛУНА"},
            .Category = "Lamp",
            .DefaultPoles = "1p",
            .DefaultCable = "3x1.5",
            .DefaultBreaker = "10",
                .VisibilityRules = New List(Of VisRule)()
        },
        New BlockConfig With {        ' АВАРИЙНО ОСВЕТЛЕНИЕ
            .BlockNames = New List(Of String) From {"АВАРИЯ", "АВАРИЯ_100"},
            .Category = "Lamp",
            .DefaultPoles = "1p",
            .DefaultCable = "3x1.5",
            .DefaultBreaker = "10",
                .VisibilityRules = New List(Of VisRule)()
        },
        New BlockConfig With {        ' БОЙЛЕРИ
            .BlockNames = New List(Of String) From {"БОЙЛЕР"},
            .Category = "Device",
            .DefaultPoles = "1p",
            .DefaultCable = "3x2.5",
            .DefaultBreaker = "16",
            .VisibilityRules = New List(Of VisRule) From {
                New VisRule With {.VisibilityPattern = "ИЗХОД 3P", .Poles = "3p", .Cable = "5x2.5", .Phase = "L1,L2,L3"},
                New VisRule With {.VisibilityPattern = "380V", .Poles = "3p", .Cable = "5x2.5", .Phase = "L1,L2,L3"},
                New VisRule With {.VisibilityPattern = "ПРОТОЧЕН", .Breaker = "20"},
                New VisRule With {.VisibilityPattern = "СЕШОАР", .Breaker = "16"},
                New VisRule With {.VisibilityPattern = "СЕШОАР С КОНТАКТ", .Breaker = "16"},
                New VisRule With {.VisibilityPattern = "ИЗХОД ГАЗ", .Cable = "3x1.5", .Breaker = "6"}
            }
        },
        New BlockConfig With {        ' БОЙЛЕРНО ТАБЛО
            .BlockNames = New List(Of String) From {"БОЙЛЕРНО ТАБЛО"},
            .Category = "Contact",
            .DefaultPoles = "1p",
            .DefaultCable = "3x1.5",
            .DefaultBreaker = "10",
            .VisibilityRules = New List(Of VisRule) From {
                New VisRule With {.VisibilityPattern = "КЛЮЧ И КОНТАКТ", .ContactCount = 1},
                New VisRule With {.VisibilityPattern = "С ДВА КОНТАКТА", .ContactCount = 2},
                New VisRule With {.VisibilityPattern = "С ДВА КЛЮЧА", .ContactCount = 2}
            }
        },
        New BlockConfig With {        ' ВЕНТИЛАЦИИ, КЛИМАТИЦИ, КОНВЕКТОРИ
            .BlockNames = New List(Of String) From {"ВЕНТИЛАЦИИ", "ВЕНТИЛАТОР", "КЛИМАТИК", "КОНВЕКТОР", "ГОРЕЛКА", "НАГРЕВАТЕЛ", "ЕЛ. ЛИРА"},
            .Category = "Device",
            .DefaultPoles = "1p",
            .DefaultCable = "3x1.5",
            .DefaultBreaker = "10",
            .VisibilityRules = New List(Of VisRule) From {
                New VisRule With {.VisibilityPattern = "3P", .Poles = "3p", .Cable = "5x1.5", .Phase = "L1,L2,L3"},
                New VisRule With {.VisibilityPattern = "КАНАЛЕН 3P", .Poles = "3p", .Cable = "5x1.5", .Phase = "L1,L2,L3"},
                New VisRule With {.VisibilityPattern = "ПРОЗОРЧЕН 3P", .Poles = "3p", .Cable = "5x1.5", .Phase = "L1,L2,L3"}
            }
        },
        New BlockConfig With {        ' КОНТАКТИ
            .BlockNames = New List(Of String) From {"КОНТАКТ"},
            .Category = "Contact",
            .DefaultPoles = "1p",
            .DefaultCable = "3x2.5",
            .DefaultBreaker = "16",
            .VisibilityRules = New List(Of VisRule) From {
                New VisRule With {.VisibilityPattern = "ДВУГНЕЗДОВ", .ContactCount = 1},
                New VisRule With {.VisibilityPattern = "ТРИГНЕЗДОВ", .ContactCount = 2},
                New VisRule With {.VisibilityPattern = "ТРИФАЗЕН", .Poles = "3p", .Cable = "5x2.5", .Phase = "L1,L2,L3"},
                New VisRule With {.VisibilityPattern = "ТР+2МФ", .Poles = "3p", .Cable = "5x2.5", .Phase = "L1,L2,L3", .ContactCount = 2},
                New VisRule With {.VisibilityPattern = "ТВЪРДА ВРЪЗКА", .Cable = "3x4.0"},
                New VisRule With {.VisibilityPattern = "УСИЛЕН", .Cable = "3x4.0"},
                New VisRule With {.VisibilityPattern = "IP 54", .Cable = "3x2.5"},
                New VisRule With {.VisibilityPattern = "МОНТАЖ В КАНАЛ", .Cable = "3x2.5"}
            }
        },
        New BlockConfig With {        ' ЛИНИЯ МХЛ
            .BlockNames = New List(Of String) From {"ЛИНИЯ МХЛ"},
            .Category = "Lamp",
            .DefaultPoles = "1p",
            .DefaultCable = "3x1.5",
            .DefaultBreaker = "10",
            .VisibilityRules = New List(Of VisRule)()
        },
        New BlockConfig With {        ' ЛУМИНЕСЦЕНТНИ ЛАМПИ
            .BlockNames = New List(Of String) From {"ЛУМИНЕСЦЕНТНА"},
            .Category = "Lamp",
            .DefaultPoles = "1p",
            .DefaultCable = "3x1.5",
            .DefaultBreaker = "10",
            .VisibilityRules = New List(Of VisRule)()
        },
        New BlockConfig With {        ' МЕТАЛОХАЛОГЕННИ ЛАМПИ
            .BlockNames = New List(Of String) From {"МЕТАЛХАОГЕННА", "МЕТАЛХАЛОГЕННА"},
            .Category = "Lamp",
            .DefaultPoles = "1p",
            .DefaultCable = "3x1.5",
            .DefaultBreaker = "10",
            .VisibilityRules = New List(Of VisRule)()
        },
        New BlockConfig With {        ' ПЛАФОНИ, АПЛИЦИ, ПЕНДЕЛИ, ЛАМПИОНИ, ДАТЧИЦИ
            .BlockNames = New List(Of String) From {"ПЛАФОНИ",
                                        "АПЛИК", "ПЕНДЕЛ", "ЛАМПИОН",
                                        "НАСТОЛНА ЛАМПА", "ФАСАДНО", "БАНСКИ АПЛИК",
                                        "ДАТЧИК", "ФОТОДАТЧИК"},
            .Category = "Lamp",
            .DefaultPoles = "1p",
            .DefaultCable = "3x1.5",
            .DefaultBreaker = "10",
            .VisibilityRules = New List(Of VisRule)()
        },
        New BlockConfig With {        ' ПОЛИЛЕИ
            .BlockNames = New List(Of String) From {"ПОЛИЛЕЙ"},
            .Category = "Lamp",
            .DefaultPoles = "1p",
            .DefaultCable = "3x1.5",
            .DefaultBreaker = "10",
            .VisibilityRules = New List(Of VisRule)()
        },
        New BlockConfig With {        ' ПРОЖЕКТОРИ
            .BlockNames = New List(Of String) From {"ПРОЖЕКТОР"},
            .Category = "Lamp",
            .DefaultPoles = "1p",
            .DefaultCable = "3x1.5",
            .DefaultBreaker = "10",
            .VisibilityRules = New List(Of VisRule)()
        }
    }
    End Sub
    Private Sub ProcessConsumerByConfig(kons As strKonsumator, ByRef tokow As strTokow)
        Dim blockName As String = kons.Name.ToUpper()
        ' ✅ БЕЗОПАСНО вземане на Visibility
        Dim visibility As String = ""
        If kons.Visibility IsNot Nothing Then visibility = kons.Visibility.ToUpper()
        ' Намери конфигурацията за този блок
        Dim config = BlockConfigs.FirstOrDefault(
        Function(c) c.BlockNames.Any(Function(n) blockName.Contains(n)))
        If config Is Nothing Then
            ' Непознат блок - използвай настройки по подразбиране
            tokow.brLamp += 1  ' По подразбиране брои като лампа
            Return
        End If
        ' Извличане на брой от мощност (напр. "3x100" → 3)
        Dim count As Integer = ExtractCountFromPower(kons.strМОЩНОСТ)
        ' ============================================================
        ' БРОЕНЕ НА ЛАМПИ ИЛИ КОНТАКТИ
        ' ============================================================
        Select Case config.Category
            Case "Lamp"
                tokow.brLamp += count
            Case "Contact"
                tokow.brKontakt += count
                ' Провери за специални правила за контакти
                If Not String.IsNullOrEmpty(visibility) Then
                    Dim visRule = config.VisibilityRules.FirstOrDefault(
                    Function(r) visibility.Contains(r.VisibilityPattern))
                    If visRule IsNot Nothing AndAlso visRule.ContactCount > 0 Then
                        tokow.brKontakt += (visRule.ContactCount - 1)
                    End If
                End If
            Case "Device"



                ' Уреди не се броят като лампи/контакти
        End Select
        ' ============================================================
        ' ПРОВЕРКА ЗА ТРИФАЗНОСТ
        ' ============================================================
        If kons.Phase = 3 Then
            tokow.БройПолюси = 3
            tokow.Фаза = "3P"
        Else
            ' Провери visibility правила за 3 фази
            If Not String.IsNullOrEmpty(visibility) Then
                Dim visRule = config.VisibilityRules.FirstOrDefault(
                Function(r) visibility.Contains(r.VisibilityPattern) AndAlso r.Poles = "3p"
            )
                If visRule IsNot Nothing Then
                    tokow.БройПолюси = 3
                    tokow.Фаза = "3P"
                End If
            End If
        End If
        ' ============================================================
        ' МОЩНОСТ
        ' ============================================================
        tokow.Мощност += kons.doubМОЩНОСТ / 1000.0  ' W → kW
    End Sub
    Private Sub CalculateCircuitLoads()
        ' Инициализирай конфигурацията (само веднъж)
        If BlockConfigs Is Nothing OrElse BlockConfigs.Count = 0 Then
            InitializeBlockConfigs()
        End If
        For Each tokow As strTokow In ListTokow
            ' Нулирай броячите
            tokow.brLamp = 0
            tokow.brKontakt = 0
            tokow.Мощност = 0
            tokow.БройПолюси = 1
            tokow.Фаза = "1P"
            ' Обработи всеки консуматор
            For Each kons As strKonsumator In tokow.Konsumator
                ProcessConsumerByConfig(kons, tokow)
            Next
            ' Изчисли тока
            If tokow.БройПолюси = 3 Then
                tokow.Ток = (tokow.Мощност * 1000) / (Math.Sqrt(3) * 400)
            Else
                tokow.Ток = (tokow.Мощност * 1000) / 230
            End If
        Next
    End Sub
End Class