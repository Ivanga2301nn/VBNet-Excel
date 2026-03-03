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
Imports Org.BouncyCastle.Math.EC.ECCurve
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
        SortCircuits()




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
    Private Kable_Type As String()
    Dim Disconnectors As New List(Of DisconnectorInfo)
    ' ============================================================
    ' КАТАЛОЖНИ СТРУКТУРИ
    ' ============================================================
    ' =====================================================
    ' 4. РЕДОВЕ: Параметри с мерни единици и типове клетки
    ' =====================================================
    ' Структура: {Параметър, Мерна единица, Тип клетка}
    ' Тип клетка: "Text", "Combo", "Check"
    Dim rowData As String()() = {
        New String() {"Прекъсвач", "", "Text"},
        New String() {"Изчислен ток", "A", "Text"},
        New String() {"Тип на апарата", "", "Combo"},
        New String() {"Номинален ток", "A", "Combo"},
        New String() {"Изкл. възможн.", "A", "Text"},
        New String() {"Крива", "", "Text"},
        New String() {"Защитен блок", "", "Combo"},
        New String() {"Брой полюси", "бр.", "Text"},
        New String() {"ДТЗ", "", "Text"},
        New String() {"Вид на апарата", "", "Text"},
        New String() {"Клас на апарата", "", "Text"},
        New String() {"Номинален ток", "A", "Text"},
        New String() {"Чувствителност", "mA", "Text"},
        New String() {"Брой полюси", "бр.", "Text"},
        New String() {"Брой лампи", "бр.", "Text"},
        New String() {"Брой контакти", "бр.", "Text"},
        New String() {"Инст. мощност", "kW", "Text"},
        New String() {"Кабел", "", "Text"},
        New String() {"Тип", "---", "Combo"},
        New String() {"Сечение", "mm²", "Combo"},
        New String() {"Фаза", "mm²", "Text"},
        New String() {"---------", "", "Text"},
        New String() {"Консуматор", "---", "Text"},
        New String() {"предназначение", "---", "Text"},
        New String() {"Управление", "---", "Combo"},
        New String() {"---------", "", "Text"},
        New String() {"Шина", "---", "Check"},
        New String() {"ДТЗ (RCD)", "---", "Check"}
    }

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
    ''' <summary>
    ''' Структура токов кръг в електрическо табло.
    ''' Съдържа идентификация, мощност, ток, кабел,
    ''' защитна апаратура (прекъсвач и ДТЗ),
    ''' както и списък с консуматори в кръга.
    ''' Използва се като логическа структура за обработка
    ''' и визуализация (напр. в DataGridView).
    ''' </summary>
    Public Class strTokow
        ' ============================================================
        ' ИДЕНТИФИКАЦИЯ
        ' ============================================================
        Public Tablo As String                 ' Табло към което принадлежи кръгът
        Public ТоковКръг As String             ' Име или номер на токовия кръг
        Public БройПолюси As Integer           ' 1 или 3 – използва се при избор на апарат
        ' ============================================================
        ' БРОЯЧИ
        ' ============================================================
        Public brLamp As Integer               ' Брой лампи в кръга
        Public brKontakt As Integer            ' Брой контакти в кръга
        ' ============================================================
        ' МОЩНОСТ И ТОК
        ' ============================================================
        Public Мощност As Double               ' Обща мощност на кръга (kW)
        Public Ток As Double                   ' Изчислен ток (A)
        Public Фаза As String                  ' Фаза: "1P", "3P", "L1", "L2", "L3"
        ' ============================================================
        ' КАБЕЛ
        ' ============================================================
        Public Кабел_Сечение As String         ' Сечение на кабела (пример: "3x2.5")
        Public Кабел_Тип As String             ' Тип кабел (NYM, YJV, CBT и др.)
        ' ============================================================
        ' ЗАЩИТА (ПРЕКЪСВАЧ)
        ' ============================================================
        Public Тип_Апарат As String            ' Серия апарат (EZ9, C120, NSX, MTZ)
        Public Брой_Полюси As String           ' Брой полюси на прекъсвача ("1p", "3p")
        Public Крива As String                 ' Характеристика (B, C, D)
        Public Номинален_Ток As String         ' Номинален ток (пример: "16A")
        Public Изкл_Възможност As String       ' Изключвателна способност ("6000A", "10000A")
        Public Защитен_блок As String          ' Изключвателна способност ("6000A", "10000A")
        ' ============================================================
        ' ДТЗ (RCD)
        ' ============================================================
        Public RCD_Тип As String               ' Тип ДТЗ (AC, A, F)
        Public RCD_Чувствителност As String    ' Чувствителност ("30mA", "100mA", "300mA")
        Public RCD_Ток As String               ' Номинален ток на ДТЗ ("25A", "40A", "63A")
        Public RCD_Полюси As String            ' Полюси на ДТЗ ("2p", "4p")
        ' ============================================================
        ' ОПИСАНИЕ / ТЕКСТОВЕ
        ' ============================================================
        Public Консуматор As String            ' Обобщен текст за консуматора
        Public предназначение As String        ' Предназначение на кръга
        ' ============================================================
        ' ДОПЪЛНИТЕЛНИ ФЛАГОВЕ
        ' ============================================================
        Public Управление As String            ' Тип управление (ако има)
        Public Шина As Boolean                 ' Дали кръгът е на шинена
        Public ДТЗ_RCD As Boolean              ' Дали има задължително трявба да има ДТЗ
        ' ============================================================
        ' КОНСУМАТОРИ В КРЪГА
        ' ============================================================
        Public Konsumator As List(Of strKonsumator)
        ' Списък с всички реални консуматори,
        ' принадлежащи към този токов кръг.
    End Class
    ''' <summary>
    ''' Представя автоматичен прекъсвач – MCB, MCCB или ACB.
    ''' Може да се използва за избор на прекъсвач за генераторни табла,
    ''' както и за по-сложни сценарии с селективност и късо съединение.
    ''' </summary>
    Public Class BreakerInfo
        Public Brand As String              ' Производител на прекъсвача (например "Schneider").
        Public Series As String             ' Серия или модел на прекъсвача (например "EZ9", "C120", "NSX", "MTZ").
        ''' <summary>
        ''' Категория на прекъсвача:
        ''' - "MCB" – миниатюрен автоматичен прекъсвач
        ''' - "MCCB" – корпусен прекъсвач
        ''' - "ACB" – въздушен прекъсвач
        ''' </summary>
        Public Category As String
        Public NominalCurrent As Integer         ' Номинален ток на прекъсвача в ампери.
        Public Poles As Integer ' Брой полюси (1P, 2P, 3P или 4P).
        ''' <summary>
        ''' Работна прекъсвателна способност (Ics) в kA.
        ''' Това е стойността, до която прекъсвачът може да изключва многократно.
        ''' </summary>
        Public Ics_kA As Decimal
        ''' <summary>
        ''' Крива на MCB (B, C или D). 
        ''' Само за миниатюрни автоматични прекъсвачи.  
        ''' Определя характеристиката на изключване при късо съединение.
        ''' MCCB и ACB не използват това поле.
        ''' </summary>
        Public Curve As String
        ''' <summary>
        ''' Тип на защитната единица (Trip Unit) – TM-D, Micrologic и т.н.
        ''' Само за MCCB и ACB.  
        ''' Определя електронната или термомагнитната защита.
        ''' </summary>
        Public TripUnit As String
        ''' <summary>
        ''' Дали прекъсвачът има регулируеми настройки на Ir, Ii, Isd и други параметри.
        ''' True – може да се настройва (MCCB, ACB).  
        ''' False – фиксирани характеристики (MCB).
        ''' </summary>
        Public IsAdjustable As Boolean
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
        Public BlockNames As List(Of String)        ' Възможни имена на блока
        Public Category As String                   ' "Lamp", "Contact", "Device", "Panel"
        Public DefaultPoles As String               ' "1p" или "3p"
        Public DefaultCable As String               ' "3x1.5", "3x2.5", "5x2.5"
        Public DefaultBreaker As String             ' "10", "16", "20"
        Public DefaultPrednaz As String             ' Предназначение 
        Public DefaultPrednaz1 As String            ' Предназначение 
        Public VisibilityRules As List(Of VisRule)  ' Правила за visibility
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
                MsgBox("Възникна грешка:  " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
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
        TreeView1.Nodes.Clear()
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
            TreeView1.Nodes.Add(panelNode)
        Next
        ' 4. Разгъни дървото
        TreeView1.ExpandAll()
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
        colParam.SortMode = DataGridViewColumnSortMode.NotSortable
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
        colUnit.SortMode = DataGridViewColumnSortMode.NotSortable
        DataGridView1.Columns.Add(colUnit)
        ' Колона ОБЩО
        Dim colTotal As New DataGridViewTextBoxColumn()
        colTotal.Name = "colTotal"
        colTotal.HeaderText = "ОБЩО"
        colTotal.Width = 90
        colTotal.DefaultCellStyle.Font = New Drawing.Font("Arial", 10, FontStyle.Bold)
        colTotal.DefaultCellStyle.BackColor = Color.FromArgb(230, 240, 255)
        colTotal.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        colTotal.SortMode = DataGridViewColumnSortMode.NotSortable
        DataGridView1.Columns.Add(colTotal)
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
            Select Case row(0).ToString()
                Case "---------"
                    dgvRow.DefaultCellStyle.BackColor = Color.FromArgb(220, 220, 220)
                Case "Прекъсвач", "ДТЗ", "Кабел"
                    dgvRow.DefaultCellStyle.BackColor = Color.FromArgb(180, 200, 255)
                    dgvRow.DefaultCellStyle.Font = New Drawing.Font("Arial", 10, FontStyle.Bold)
                Case Else
                    ' Тук можеш да сложиш форматиране по подразбиране, ако е необходимо
            End Select
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
                comboCell.Items.AddRange("6", "10", "16", "20", "25", "32", "40", "50", "63")
            Case "Управление"
                comboCell.Items.AddRange("Няма",
                                         "Импулсно реле",
                                         "Моторна защита",
                                         "Контактор",
                                         "Моторен механизъм",
                                         "Честотен регулатор",
                                         "Стълбищен автомат",
                                         "Електромер",
                                         "Фото реле")
            Case "Сечение"
                comboCell.Items.AddRange(Kable_Size_L)
            Case "Тип"
                comboCell.Items.AddRange(Kable_Type)
        End Select
        ' ✅ ЗАДАЙ ПЪРВИЯ ЕЛЕМЕНТ КАТО СТОЙНОСТ
        If comboCell.Items.Count > 0 Then comboCell.Value = comboCell.Items(0)
        'comboCell.DisplayStyle = ComboBoxStyle.Simple
        comboCell.DisplayStyle = ComboBoxStyle.DropDown
        'comboCell.DisplayStyle = ComboBoxStyle.DropDownList
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
        ' Речник за типове кабели
        Kable_Type = {"СВТ", "САВТ", "Al/R", "Al/R+СВТ", "Al/R+САВТ"}
        ' Речник за всички автоматични прекъсвачи
        ' Инициализиране на списъка
        Breakers = New List(Of BreakerInfo)
        ' Попълване на всички прекъсвачи чрез отделната процедура
        FillBreakers()
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
    ''' Процедура за добавяне на всички прекъсвачи.
    ''' Тук се генерират MCB, MCCB и ACB.
    ''' </summary>
    Private Sub FillBreakers()
        ' ==========================
        ' MCB – EZ9
        ' ==========================
        Dim EZ9_Currents = {6, 10, 16, 20, 25, 32, 40, 50, 63}
        Dim EZ9_Curves = {"B", "C", "D"}
        Dim EZ9_Poles = {1, 3}
        For Each Inom In EZ9_Currents
            For Each curve In EZ9_Curves
                For Each poles In EZ9_Poles
                    Breakers.Add(New BreakerInfo With {
                    .Brand = "Schneider",
                    .Series = "Easy9",
                    .Category = "MCB",
                    .NominalCurrent = Inom,
                    .Poles = poles,
                    .Curve = curve,
                    .Ics_kA = 6,
                    .TripUnit = Nothing,
                    .IsAdjustable = False
                })
                Next
            Next
        Next
        ' ==========================
        ' MCB – Acti9 iC60N (6kA / 10kA)
        ' ==========================
        ' iC60N предлага изключително малки токове за защита на контролни вериги
        Dim iC60_Currents = {0.5, 1, 2, 3, 4, 6, 10, 13, 16, 20, 25, 32, 40, 50, 63}
        Dim iC60_Curves = {"B", "C", "D"}
        Dim iC60_Poles = {1, 3} ' Добавяме 2P и 4P, които са стандарт тук

        For Each Inom In iC60_Currents
            For Each curve In iC60_Curves
                For Each poles In iC60_Poles
                    Breakers.Add(New BreakerInfo With {
                .Brand = "Schneider",
                .Series = "Acti9 iC60N",
                .Category = "MCB",
                .NominalCurrent = Inom,
                .Poles = poles,
                .Curve = curve,
                .Ics_kA = 6,
                .TripUnit = Nothing,
                .IsAdjustable = False
            })
                Next
            Next
        Next
        ' ==========================
        ' MCCB – ComPacT NSXm (16A до 160A)
        ' ==========================
        Dim NSXm_Currents = {16, 25, 32, 40, 50, 63, 80, 100, 125, 160}
        Dim NSXm_Poles = {3, 4} ' Предлага се основно в 3P и 4P
        ' Дефинираме нивата на изключвателна способност (Icu @ 415V)
        ' E=16kA, B=25kA, F=36kA, N=50kA, H=70kA
        Dim NSXm_Levels = New Dictionary(Of String, Integer) From {
            {"E", 16}, {"B", 25}, {"F", 36}, {"N", 50}, {"H", 70}
        }

        For Each Inom In NSXm_Currents
            For Each level In NSXm_Levels
                For Each poles In NSXm_Poles
                    Breakers.Add(New BreakerInfo With {
                .Brand = "Schneider",
                .Series = "ComPacT NSXm",
                .Category = "MCCB",
                .NominalCurrent = Inom,
                .Poles = poles,
                .TripUnit = "TM-D",
                .Ics_kA = level.Value,
                .Curve = level.Key,
                .IsAdjustable = True
            })
                Next
            Next
        Next
        ' ==========================
        ' MCB – C120
        ' ==========================
        Dim C120_Currents = {80, 100, 125}
        Dim C120_Curves = {"C", "D"}
        Dim C120_Poles = {1, 3}
        For Each Inom In C120_Currents
            For Each curve In C120_Curves
                For Each poles In C120_Poles
                    Breakers.Add(New BreakerInfo With {
                    .Brand = "Schneider",
                    .Series = "Acti9 C120",
                    .Category = "MCB",
                    .NominalCurrent = Inom,
                    .Poles = poles,
                    .Curve = curve,
                    .Ics_kA = 10,
                    .TripUnit = Nothing,
                    .IsAdjustable = False
                })
                Next
            Next
        Next
        ' NSX100 – TM‑D, TM‑DC
        Dim NSX100_Currents = {16, 25, 32, 40, 63, 80, 100}
        Dim NSX100_Curves = {"B", "F", "N", "H", "S", "L"}
        Dim NSX100_TripUnits = {"TM-D", "TM-DC"}

        For Each Inom In NSX100_Currents
            For Each curve In NSX100_Curves
                For Each trip In NSX100_TripUnits
                    Breakers.Add(New BreakerInfo With {
                        .Brand = "Schneider",
                        .Series = "ComPacT NSX100",
                        .Category = "MCCB",
                        .NominalCurrent = Inom,
                        .Poles = 3,
                        .TripUnit = trip,
                        .Ics_kA = 25,
                        .Curve = curve,
                        .IsAdjustable = True
                    })
                Next
            Next
        Next

        ' NSX160 – TM‑D, TM‑DC
        Dim NSX160_Currents = {80, 100, 125, 160}
        Dim NSX160_Curves = {"B", "F", "N", "H", "S", "L"}
        Dim NSX160_TripUnits = {"TM-D", "TM-DC"}

        For Each Inom In NSX160_Currents
            For Each curve In NSX160_Curves
                For Each trip In NSX160_TripUnits
                    Breakers.Add(New BreakerInfo With {
                        .Brand = "Schneider",
                        .Series = "ComPacT NSX160",
                        .Category = "MCCB",
                        .NominalCurrent = Inom,
                        .Poles = 3,
                        .TripUnit = trip,
                        .Ics_kA = 36,
                        .Curve = curve,
                        .IsAdjustable = True
                    })
                Next
            Next
        Next

        ' NSX250 – Micrologic (по‑големи токове обикновено с електронна защита)
        Dim NSX250_Currents = {125, 160, 200, 250}
        Dim NSX250_Curves = {"B", "F", "N", "H", "S", "L"}
        Dim NSX250_TripUnits = {"Micrologic 2.0", "Micrologic 5.0"}

        For Each Inom In NSX250_Currents
            For Each curve In NSX250_Curves
                For Each trip In NSX250_TripUnits
                    Breakers.Add(New BreakerInfo With {
                        .Brand = "Schneider",
                        .Series = "ComPacT NSX250",
                        .Category = "MCCB",
                        .NominalCurrent = Inom,
                        .Poles = 3,
                        .TripUnit = trip,
                        .Ics_kA = 50,
                        .Curve = curve,
                        .IsAdjustable = True
                    })
                Next
            Next
        Next
        ' NSX400/NSX630 – Micrologic
        Dim NSX400_Currents = {250, 320, 400}
        Dim NSX400_Curves = {"F", "N", "H", "S", "L"}
        Dim NSX_High_TripUnits = {"Micrologic 2.3"}
        For Each Inom In NSX400_Currents
            For Each curve In NSX400_Curves
                For Each trip In NSX_High_TripUnits
                    Breakers.Add(New BreakerInfo With {
                        .Brand = "Schneider",
                        .Series = "ComPacT NSX400",
                        .Category = "MCCB",
                        .NominalCurrent = Inom,
                        .Poles = 3,
                        .TripUnit = trip,
                        .Ics_kA = 70,
                        .Curve = curve,
                        .IsAdjustable = True
                    })
                Next
            Next
        Next
        Dim NSX630_Currents = {400, 500, 630}
        For Each Inom In NSX630_Currents
            For Each curve In NSX400_Curves
                For Each trip In NSX_High_TripUnits
                    Breakers.Add(New BreakerInfo With {
                        .Brand = "Schneider",
                        .Series = "ComPacT NSX630",
                        .Category = "MCCB",
                        .NominalCurrent = Inom,
                        .Poles = 3,
                        .TripUnit = trip,
                        .Ics_kA = 100,
                        .Curve = curve,
                        .IsAdjustable = True
                    })
                Next
            Next
        Next
        ' ==========================
        ' ACB – MTZ
        ' ==========================
        Dim MTZ_Currents = {800, 1000, 1250, 1600, 2000, 2500, 3200, 4000, 5000, 6300}
        Dim MTZ_Icu = {42, 65, 100}
        Dim MTZ_Poles = {3, 4}
        For Each Inom In MTZ_Currents
            For Each icuValue In MTZ_Icu
                For Each poles In MTZ_Poles
                    Breakers.Add(New BreakerInfo With {
                    .Brand = "Schneider",
                    .Series = "Masterpact MTZ",
                    .Category = "ACB",
                    .NominalCurrent = Inom,
                    .Poles = poles,
                    .TripUnit = "Micrologic 6.0",
                    .Ics_kA = icuValue,
                    .Curve = Nothing,
                    .IsAdjustable = True
                })
                Next
            Next
        Next
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
            .BlockNames = New List(Of String) From {"LED_DENIMA", "LED_LENTA", "LED_ULTRALUX", "LED_ULTRALUX_100", "LED_ULTRALUX_НОВ",
                                                    "LED_ЛУНА", "ПЛАФОНИ", "МЕТАЛХАОГЕННА ЛАМПА", "ЛИНИЯ МХЛ - 220V", "ПОЛИЛЕЙ", "ПРОЖЕКТОР"},
            .Category = "Lamp",
            .DefaultPoles = "1p",
            .DefaultCable = "3x1,5",
            .DefaultBreaker = "10",
            .DefaultPrednaz = "Общо",
            .DefaultPrednaz1 = "осветление",
            .VisibilityRules = New List(Of VisRule)()
        },
        New BlockConfig With {        ' УЛИЧНО ОСВЕТЛЕНИЕ
            .BlockNames = New List(Of String) From {"ULI4NO"},
            .Category = "Lamp",
            .DefaultPoles = "1p",
            .DefaultCable = "3x1,5",
            .DefaultBreaker = "10",
            .DefaultPrednaz = "Улично",
            .DefaultPrednaz1 = "осветление",
            .VisibilityRules = New List(Of VisRule)()
        },
        New BlockConfig With {        ' АВАРИЙНО ОСВЕТЛЕНИЕ
            .BlockNames = New List(Of String) From {"АВАРИЯ", "АВАРИЯ_100"},
            .Category = "Lamp",
            .DefaultPoles = "1p",
            .DefaultCable = "3x1,5",
            .DefaultBreaker = "10",
            .DefaultPrednaz = "Аварийно",
            .DefaultPrednaz1 = "осветление",
            .VisibilityRules = New List(Of VisRule)()
        },
        New BlockConfig With {        ' БОЙЛЕРИ
            .BlockNames = New List(Of String) From {"БОЙЛЕР"},
            .Category = "Device",
            .DefaultPoles = "1p",
            .DefaultCable = "3x2,5",
            .DefaultBreaker = "16",
            .VisibilityRules = New List(Of VisRule) From {
                New VisRule With {.VisibilityPattern = "ИЗХОД 3P", .Poles = "3p", .Cable = "5x2,5", .Phase = "L1,L2,L3"},
                New VisRule With {.VisibilityPattern = "380V", .Poles = "3p", .Cable = "5x2,5", .Phase = "L1,L2,L3"},
                New VisRule With {.VisibilityPattern = "ПРОТОЧЕН", .Breaker = "20"},
                New VisRule With {.VisibilityPattern = "СЕШОАР", .Breaker = "16"},
                New VisRule With {.VisibilityPattern = "СЕШОАР С КОНТАКТ", .Breaker = "16"},
                New VisRule With {.VisibilityPattern = "ИЗХОД ГАЗ", .Cable = "3x1,5", .Breaker = "6"}
            }
        },
        New BlockConfig With {        ' БОЙЛЕРНО ТАБЛО
            .BlockNames = New List(Of String) From {"БОЙЛЕРНО ТАБЛО"},
            .Category = "Contact",
            .DefaultPoles = "1p",
            .DefaultCable = "3x2,5",
            .DefaultBreaker = "10",
            .VisibilityRules = New List(Of VisRule) From {
                New VisRule With {.VisibilityPattern = "КЛЮЧ И КОНТАКТ", .ContactCount = 1},
                New VisRule With {.VisibilityPattern = "С ДВА КОНТАКТА", .ContactCount = 2},
                New VisRule With {.VisibilityPattern = "С ДВА КЛЮЧА", .ContactCount = 2}
            }
        },
        New BlockConfig With {        ' ВЕНТИЛАЦИИ, КЛИМАТИЦИ, КОНВЕКТОРИ
            .BlockNames = New List(Of String) From {"ВЕНТИЛАЦИИ"},
            .Category = "Device",
            .DefaultPoles = "1p",
            .DefaultCable = "3x1,5",
            .DefaultBreaker = "10",
            .VisibilityRules = New List(Of VisRule) From {
                New VisRule With {.VisibilityPattern = "3P", .Poles = "3p", .Cable = "5x2,5", .Phase = "L1,L2,L3"},
                New VisRule With {.VisibilityPattern = "КАНАЛЕН 3P", .Poles = "3p", .Cable = "5x2,5", .Phase = "L1,L2,L3"},
                New VisRule With {.VisibilityPattern = "ПРОЗОРЧЕН 3P", .Poles = "3p", .Cable = "5x2,5", .Phase = "L1,L2,L3"}
            }
        },
        New BlockConfig With {        ' КОНТАКТИ
            .BlockNames = New List(Of String) From {"КОНТАКТ"},
            .Category = "Contact",
            .DefaultPoles = "1p",
            .DefaultCable = "3x2,5",
            .DefaultBreaker = "16",
            .DefaultPrednaz = "Контакти",
            .DefaultPrednaz1 = "",
            .VisibilityRules = New List(Of VisRule) From {
                New VisRule With {.VisibilityPattern = "ДВУГНЕЗДОВ", .ContactCount = 1},
                New VisRule With {.VisibilityPattern = "ТРИГНЕЗДОВ", .ContactCount = 2},
                New VisRule With {.VisibilityPattern = "ТРИФАЗЕН", .Poles = "3p", .Cable = "5x2,5", .Phase = "L1,L2,L3"},
                New VisRule With {.VisibilityPattern = "ТР+2МФ", .Poles = "3p", .Cable = "5x2,5", .Phase = "L1,L2,L3", .ContactCount = 2},
                New VisRule With {.VisibilityPattern = "ТВЪРДА ВРЪЗКА", .Cable = "3x4,0"},
                New VisRule With {.VisibilityPattern = "УСИЛЕН", .Cable = "3x4,0"},
                New VisRule With {.VisibilityPattern = "IP 54", .Cable = "3x2,5"},
                New VisRule With {.VisibilityPattern = "МОНТАЖ В КАНАЛ", .Cable = "3x2,5"}
            }
        }
    }
    End Sub
    Private Sub ProcessConsumerByConfig(kons As strKonsumator, ByRef tokow As strTokow)
        Dim blockName As String = kons.Name.ToUpper()
        Dim visibility As String = If(kons.Visibility IsNot Nothing, kons.Visibility.ToUpper(), "")
        ' 1. Намиране на основната конфигурация за блока
        Dim config = BlockConfigs.FirstOrDefault(Function(c) c.BlockNames.Any(Function(n) blockName.Contains(n)))
        If config Is Nothing Then
            MsgBox("Блок '" & blockName & "' не е намерен в InitializeBlockConfigs!", MsgBoxStyle.Critical)
            Return
        End If
        ' 2. Намиране на специфично правило по Visibility (ако съществува)
        Dim visRule = config.VisibilityRules.FirstOrDefault(Function(r) visibility.Contains(r.VisibilityPattern))
        ' 3. ПРЕХВЪРЛЯНЕ НА ИНФОРМАЦИЯТА (Спрямо Config)
        ' Кабел и Предпазител: Вземаме от правилото, ако има такова, иначе от Default на конфигурацията
        tokow.Кабел_Сечение = If(visRule IsNot Nothing AndAlso Not String.IsNullOrEmpty(visRule.Cable), visRule.Cable, config.DefaultCable)
        Dim breakerVal As String = If(visRule IsNot Nothing AndAlso Not String.IsNullOrEmpty(visRule.Breaker), visRule.Breaker, config.DefaultBreaker)
        tokow.Номинален_Ток = breakerVal & "A"
        ' Полюси
        tokow.Брой_Полюси = If(visRule IsNot Nothing AndAlso Not String.IsNullOrEmpty(visRule.Poles), visRule.Poles, config.DefaultPoles)
        tokow.БройПолюси = If(tokow.Брой_Полюси.ToLower() = "3p", 3, 1)
        ' Фаза
        If tokow.БройПолюси = 3 Then
            tokow.Фаза = "L1,L2,L3"
        Else
            ' Ако не е 3P, запазваме съществуващата фаза (L1/L2/L3) или слагаме L1 по подразбиране
            If String.IsNullOrEmpty(tokow.Фаза) Then tokow.Фаза = "L"
        End If
        ' Предназначение
        tokow.Консуматор = config.DefaultPrednaz
        tokow.предназначение = config.DefaultPrednaz1
        ' 4. МОЩНОСТ И БРОЯЧИ
        tokow.Мощност += kons.doubМОЩНОСТ / 1000.0 ' Превръщаме W в kW
        Dim count As Integer = ExtractCountFromPower(kons.strМОЩНОСТ)
        Select Case config.Category
            Case "Lamp"
                tokow.brLamp += count
            Case "Contact"
                ' Ако в правилото има специфичен брой контакти (напр. за двугнездов), ползваме него
                If visRule IsNot Nothing AndAlso visRule.ContactCount > 0 Then
                    tokow.brKontakt += visRule.ContactCount
                Else
                    tokow.brKontakt += count
                End If
                ' Автоматично маркираме, че за контакти се изисква ДТЗ
                tokow.ДТЗ_RCD = True
        End Select
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
            ' Обработи всеки консуматор
            For Each kons As strKonsumator In tokow.Konsumator
                ProcessConsumerByConfig(kons, tokow)
            Next
            ' Изчисли тока
            tokow.Ток = calc_Inom(tokow.Мощност, tokow.Брой_Полюси)
            ' ✅ ИЗБЕРИ ПРЕКЪСВАЧ
            Dim poles As Integer = If(tokow.БройПолюси = 3, 3, 1)
            Dim breaker As BreakerInfo = SelectBreaker(tokow.Ток, poles, "C")

            'tokow.Тип_Апарат = breaker.Type
            tokow.Номинален_Ток = breaker.NominalCurrent.ToString()
            tokow.Крива = breaker.Curve
            'tokow.Изкл_Възможност = breaker.BreakingCapacity.ToString() & "A"
            tokow.Брой_Полюси = breaker.Poles & "P"
        Next
    End Sub
    ''' <summary>
    ''' Автоматично избира прекъсвач от каталога според тока и броя полюси
    ''' </summary>
    ''' <param name="calculatedCurrent">Изчислен ток (A)</param>
    ''' <param name="poles">Брой полюси (1 или 3)</param>
    ''' <param name="curve">Крива по подразбиране ("C")</param>
    ''' <returns>BreakerInfo или Nothing ако не е намерен</returns>
    Private Function SelectBreaker(calculatedCurrent As Double,
                                poles As Integer,
                                Optional curve As String = "C") As BreakerInfo

        ' 1. Коефициент на оразмеряване (1.25 за резерв)
        Dim designCurrent As Double = calculatedCurrent * 1.25
        ' 2. Филтрирай прекъсвачите по полюси и крива
        Dim suitableBreakers = Breakers.Where(
                               Function(b) b.Poles = poles AndAlso
                               String.Equals(b.Curve, curve, StringComparison.OrdinalIgnoreCase)
                               ).OrderBy(Function(b) b.NominalCurrent).ToList()
        ' 3. Намери първия прекъсвач с номинален ток >= designCurrent
        Dim selectedBreaker = suitableBreakers.FirstOrDefault(
                              Function(b) b.NominalCurrent >= designCurrent
                              )
        Return selectedBreaker
    End Function
    ''' <summary>
    ''' Изчислява номиналния ток за токов кръг
    ''' </summary>
    ''' <param name="Pkryg">Мощност в kW</param>
    ''' <param name="NumberPoles">Брой фази: "1P" или "3P"</param>
    ''' <param name="Motor">True за двигатели (cos φ = 0.85, КПД = 0.9)</param>
    ''' <returns>Номинален ток в Ampere</returns>
    Private Function calc_Inom(Pkryg As Double,                     ' мощност
                       NumberPoles As String,                       ' брой фази
                       Optional Motor As Boolean = False            ' Ако е двигател True - КПД и cos FI да са по 0,83
                       ) As Double                                  ' Изчислява номинален ток за товар
        Dim CosFI As Double                                         ' Декларира променлива за cos φ (фактор на мощността)
        Dim KPD As Double                                           ' Декларира променлива за КПД (коефициент на полезно действие)
        Const U380 As Double = 0.4                                  ' Дефинира константа за напрежение при 380V, преобразувано в kV (киловолти)
        Const U220 As Double = 0.23                                 ' Дефинира константа за напрежение при 220V, преобразувано в kV (киловолти)
        Dim Inom As Double = 0                                      ' Инициализира променлива за номиналния ток с начална стойност 0
        If Motor Then                                               ' Проверява дали токовият кръг е двигател
            CosFI = 0.85                                            ' Ако е двигател, задава фактор на мощността 0.83
            KPD = 0.9                                               ' Ако е двигател, задава КПД 0.83
        Else                                                        ' Ако токовият кръг не е двигател
            CosFI = 0.9                                             ' Задава фактор на мощността 0.9
            KPD = 1                                                 ' Задава КПД 1
        End If
        If NumberPoles = "3p" Then                                  ' Проверява дали токовият кръг е трифазен (3 полюса)
            Inom = Pkryg / (U380 * Math.Sqrt(3) * CosFI * KPD)      ' Изчислява номиналния ток за трифазен кръг по формулата
        Else                                                        ' Ако токовият кръг е монофазен (2 полюса)
            Inom = Pkryg / (U220 * CosFI * KPD)                     ' Изчислява номиналния ток за монофазен кръг по формулата
        End If
        Return Inom                                                 ' Връща изчисления номинален ток
    End Function
    Private Sub TreeView1_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeView1.AfterSelect
        FillDataGridViewForPanel()
    End Sub
    ''' <summary>
    ''' Попълва DataGridView1 с данни за избраното табло
    ''' </summary>
    Private Sub FillDataGridViewForPanel()
        ' Проверка дали има избран възел
        If TreeView1.SelectedNode Is Nothing Then
            MsgBox("Моля, изберете табло от дървото!", MsgBoxStyle.Exclamation, "Няма избор")
            Return
        End If
        ' Вземи името на избраното табло
        Dim selectedPanel As String = TreeView1.SelectedNode.Text
        ' Ако има "(", вземи само текста преди него
        If selectedPanel.Contains("(") Then
            selectedPanel = selectedPanel.Substring(0, selectedPanel.IndexOf("(")).Trim()
        End If
        ' Филтрирай токовите кръгове за това табло
        ' ListTokow вече е сортиран, така че просто вземи кръговете за това табло
        Dim panelCircuits = ListTokow.Where(Function(t) t.Tablo.ToUpper() = selectedPanel.ToUpper()).ToList()
        ' Проверка дали има кръгове
        If panelCircuits Is Nothing OrElse panelCircuits.Count = 0 Then
            MsgBox($"Няма намерени токови кръгове за табло '{selectedPanel}'",
                   MsgBoxStyle.Information, "Няма данни")
            Return
        End If
        GroupBox2.Text = $"Обработвам табло '{selectedPanel}'"
        ' 1. Добави колони за кръговете
        AddCircuitColumns(panelCircuits)
        ' ============================================================
        ' 1. ИЗЧИСЛИ ВСИЧКИ ОБЩИ СТОЙНОСТИ (САМО ВЕДНЪЖ!)
        ' ============================================================
        Dim totalLamps As Integer = panelCircuits.Sum(Function(c) c.brLamp)
        Dim totalContacts As Integer = panelCircuits.Sum(Function(c) c.brKontakt)
        Dim totalPower As Double = panelCircuits.Sum(Function(c) c.Мощност)

        Dim hasThreePhase As Boolean = panelCircuits.Any(Function(c) c.БройПолюси = 3)
        Dim totalCurrent As Double = If(hasThreePhase,
                                        (totalPower * 1000) / (Math.Sqrt(3) * 400),
                                        (totalPower * 1000) / 230)

        Dim mostCommonPoles As String = panelCircuits.GroupBy(Function(c) c.Брой_Полюси) _
                                             .OrderByDescending(Function(g) g.Count()) _
                                             .FirstOrDefault()?.Key
        If mostCommonPoles Is Nothing Then mostCommonPoles = "1p"

        Dim totalPhase As String = If(hasThreePhase, "3P", "1P")
        ' 2. Попълни данните
        For Each row As DataGridViewRow In DataGridView1.Rows
            Dim paramName As String = row.Cells(0).Value.ToString()
            ' Пропусни разделителите и заглавията
            If paramName = "---------" OrElse
                paramName = "Прекъсвач" OrElse
                paramName = "ДТЗ" OrElse
                paramName = "Управление" OrElse
                paramName = "Кабел" Then

                Continue For
            End If
            ' Попълни клетките за всеки кръг
            For i As Integer = 0 To panelCircuits.Count - 1
                Dim circuit As strTokow = panelCircuits(i)
                Dim colIndex As Integer = i + 2
                If colIndex < DataGridView1.Columns.Count - 1 Then
                    Select Case paramName
                        ' --- ЗАЩИТА (ПРЕКЪСВАЧ) ---
                        Case "Тип на апарата" : row.Cells(colIndex).Value = panelCircuits(i).Тип_Апарат
                        Case "Номинален ток" : row.Cells(colIndex).Value = panelCircuits(i).Номинален_Ток
                        Case "Изкл. възможн." : row.Cells(colIndex).Value = panelCircuits(i).Изкл_Възможност
                        Case "Крива" : row.Cells(colIndex).Value = panelCircuits(i).Крива
                        Case "Защитен блок" : row.Cells(colIndex).Value = panelCircuits(i).Защитен_блок
                        Case "Брой полюси" : row.Cells(colIndex).Value = panelCircuits(i).Брой_Полюси
                        ' --- ДТЗ (RCD) ---
                        Case "Вид на апарата" : row.Cells(colIndex).Value = panelCircuits(i).RCD_Тип
                        Case "Номинален ток" : row.Cells(colIndex).Value = panelCircuits(i).RCD_Ток
                        Case "Чувствителност" : row.Cells(colIndex).Value = panelCircuits(i).RCD_Чувствителност
                        Case "Брой полюси" : row.Cells(colIndex).Value = panelCircuits(i).RCD_Полюси
                        ' --- БРОЯЧИ И МОЩНОСТ ---
                        Case "Брой лампи" : row.Cells(colIndex).Value = panelCircuits(i).brLamp
                        Case "Брой контакти" : row.Cells(colIndex).Value = panelCircuits(i).brKontakt
                        Case "Инст. мощност" : row.Cells(colIndex).Value = panelCircuits(i).Мощност.ToString("N3")
                        Case "Изчислен ток" : row.Cells(colIndex).Value = panelCircuits(i).Ток.ToString("N2")
                        ' --- КАБЕЛ И ФАЗА ---
                        Case "Тип" : row.Cells(colIndex).Value = panelCircuits(i).Кабел_Тип
                        Case "Сечение" : row.Cells(colIndex).Value = panelCircuits(i).Кабел_Сечение
                        Case "Фаза" : row.Cells(colIndex).Value = panelCircuits(i).Фаза
                        ' --- ОПИСАНИЯ ---
                        Case "Консуматор" : row.Cells(colIndex).Value = panelCircuits(i).Консуматор
                        Case "предназначение" : row.Cells(colIndex).Value = panelCircuits(i).предназначение
                        Case "Управление" : row.Cells(colIndex).Value = panelCircuits(i).Управление
                        ' --- ФЛАГОВЕ ---
                        Case "Шина" : row.Cells(colIndex).Value = panelCircuits(i).Шина
                        Case "ДТЗ (RCD)" : row.Cells(colIndex).Value = panelCircuits(i).ДТЗ_RCD
                    End Select
                End If
            Next
            ' 3. ОБЩО (последна колона)
            Dim totalColIndex As Integer = DataGridView1.Columns.Count - 1
            Select Case paramName
                Case "Брой лампи"
                    row.Cells(totalColIndex).Value = panelCircuits.Sum(Function(c) c.brLamp)
                Case "Брой контакти"
                    row.Cells(totalColIndex).Value = panelCircuits.Sum(Function(c) c.brKontakt)
                Case "Инст. мощност"
                    row.Cells(totalColIndex).Value = panelCircuits.Sum(Function(c) c.Мощност).ToString("0.00")
                Case "Изчислен ток"
                    If hasThreePhase Then
                        totalCurrent = (totalPower * 1000) / (Math.Sqrt(3) * 400)
                    Else
                        totalCurrent = (totalPower * 1000) / 230
                    End If
                    row.Cells(totalColIndex).Value = totalCurrent.ToString("0.00")
            End Select
        Next
    End Sub
    ''' <summary>
    ''' Добавя колони за токовите кръгове на избраното табло
    ''' </summary>
    Private Sub AddCircuitColumns(panelCircuits As List(Of strTokow))
        ' 1. Изтрий старите колони за кръгове
        Dim columnsToRemove As New List(Of String)
        For Each col As DataGridViewColumn In DataGridView1.Columns
            If col.Name <> "colParameter" AndAlso
           col.Name <> "colUnit" AndAlso
           col.Name <> "colTotal" Then
                columnsToRemove.Add(col.Name)
            End If
        Next
        For Each colName As String In columnsToRemove
            DataGridView1.Columns.Remove(colName)
        Next
        ' 2. Добави нови колони за всеки кръг
        For i As Integer = 0 To panelCircuits.Count - 1
            Dim circuit As strTokow = panelCircuits(i)
            Dim col As New DataGridViewTextBoxColumn()
            col.Name = $"colCircuit{i}"
            col.HeaderText = circuit.ТоковКръг
            col.Width = 110
            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            col.SortMode = DataGridViewColumnSortMode.NotSortable
            col.Tag = circuit
            ' Добави колоната ПРЕДИ colTotal
            Dim totalIndex As Integer = DataGridView1.Columns.IndexOf(DataGridView1.Columns("colTotal"))
            DataGridView1.Columns.Insert(totalIndex, col)
        Next
        ' ✅ 3. ЗАДАЙ ТИПА КЛЕТКИ ЗА НОВИТЕ КОЛОНИ (използвайки rowData)
        Dim rowIndex As Integer = 0
        For Each row As DataGridViewRow In DataGridView1.Rows
            ' Вземи типа клетка от rowData
            Dim cellType As String = rowData(rowIndex)(2)
            ' За всяка нова колона (от индекс 2 до colTotal-1)
            For colIndex As Integer = 2 To DataGridView1.Columns.Count - 2
                Dim colName As String = DataGridView1.Columns(colIndex).Name
                ' Пропусни ако не е колона за кръг
                If Not colName.StartsWith("colCircuit") Then Continue For
                ' Запази стойността от старата клетка (ако има)
                Dim oldValue As Object = Nothing
                If row.Cells(colIndex).Value IsNot Nothing Then
                    oldValue = row.Cells(colIndex).Value
                End If
                ' Създай нова клетка с правилния тип (същата логика като в SetupDataGridView)
                Dim cell As DataGridViewCell = Nothing
                Select Case cellType
                    Case "Combo"
                        cell = New DataGridViewComboBoxCell()
                        SetupComboBoxCell(cell, row.Cells(0).Value.ToString())
                    Case "Check"
                        cell = New DataGridViewCheckBoxCell()
                        cell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                    Case Else
                        cell = New DataGridViewTextBoxCell()
                        cell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                End Select
                ' Възстанови стойността
                If oldValue IsNot Nothing Then
                    cell.Value = oldValue
                End If
                ' Замени клетката
                row.Cells(colIndex) = cell
            Next
            rowIndex += 1
        Next
    End Sub
    ''' <summary>
    ''' Връща ключ за сортиране на токов кръг със специален приоритет
    ''' Порядок: 1.ав. → 2.до. → 3.други букви → 4.числа → 5.само букви
    ''' </summary>
    Private Function GetCircuitSortKey(circuitName As String) As String
        If String.IsNullOrEmpty(circuitName) Then Return "ZZZZZZZZZZ"
        Dim name As String = circuitName.Trim().ToUpper()
        Dim priority As String = "9"  ' По подразбиране най-нисък приоритет
        Dim numberPart As String = ""
        Dim letterPart As String = ""
        ' ============================================================
        ' 1. ОПРЕДЕЛИ КАТЕГОРИЯТА (ПРИОРИТЕТ)
        ' ============================================================
        ' Проверка за "АВ." (авариен?)
        If name.Contains("АВ.") OrElse name.EndsWith("АВ") Then
            priority = "1"  ' Най-висок приоритет
            numberPart = ExtractNumber(name)
            letterPart = "АВ"
            ' Проверка за "ДО." (допълнителен?)
        ElseIf name.Contains("ДО.") OrElse name.EndsWith("ДО") Then
            priority = "2"  ' Втори приоритет
            numberPart = ExtractNumber(name)
            letterPart = "ДО"
            ' Проверка за други букви + число (напр. "1А", "2Б", "А1")
        ElseIf HasNumberAndLetters(name) Then
            priority = "3"  ' Трети приоритет
            numberPart = ExtractNumber(name)
            letterPart = ExtractLetters(name)
            ' Проверка за само число (напр. "1", "2", "10")
        ElseIf IsNumeric(name) Then
            priority = "4"  ' Четвърти приоритет
            numberPart = name
            letterPart = ""
            ' Проверка за само букви (напр. "А", "Б", "LIGHT")
        Else
            priority = "5"  ' Най-нисък приоритет
            numberPart = ""
            letterPart = name
        End If
        ' ============================================================
        ' 2. СЪЗДАЙ КЛЮЧ ЗА СОРТИРАНЕ
        ' ============================================================
        ' Формат: Приоритет + Номер (с водещи нули) + Букви
        ' Пример: "10000000001АВ" за "1ав."
        If numberPart.Length > 0 Then
            ' Подравняване на числото с водещи нули (до 10 цифри)
            numberPart = numberPart.PadLeft(10, "0"c)
            Return priority & numberPart & letterPart
        Else
            ' Само букви - сортирай азбучно
            Return priority & "0000000000" & letterPart
        End If
    End Function
    ''' <summary>
    ''' Извлича числото от низ (напр. "1АВ" → "1", "А2Б" → "2")
    ''' </summary>
    Private Function ExtractNumber(text As String) As String
        Dim result As String = ""
        For Each c As Char In text
            If Char.IsDigit(c) Then
                result &= c
            End If
        Next
        Return result
    End Function
    ''' <summary>
    ''' Извлича буквите от низ (напр. "1АВ" → "АВ", "А2Б" → "АБ")
    ''' </summary>
    Private Function ExtractLetters(text As String) As String
        Dim result As String = ""
        For Each c As Char In text
            If Char.IsLetter(c) Then
                result &= c
            End If
        Next
        Return result
    End Function
    ''' <summary>
    ''' Проверява дали низът съдържа и букви и числа
    ''' </summary>
    Private Function HasNumberAndLetters(text As String) As Boolean
        Dim hasNumber As Boolean = False
        Dim hasLetter As Boolean = False
        For Each c As Char In text
            If Char.IsDigit(c) Then hasNumber = True
            If Char.IsLetter(c) Then hasLetter = True
        Next
        Return hasNumber AndAlso hasLetter
    End Function
    ''' <summary>
    ''' Проверява дали низът е само число
    ''' </summary>
    Private Function IsNumeric(text As String) As Boolean
        For Each c As Char In text
            If Not Char.IsDigit(c) Then Return False
        Next
        Return text.Length > 0
    End Function
    ''' <summary>
    ''' Сортира ListTokow по специалния приоритет
    ''' </summary>
    Private Sub SortCircuits()
        ListTokow = ListTokow.OrderBy(
            Function(t) t.Tablo
        ).ThenBy(
            Function(t) GetCircuitSortKey(t.ТоковКръг)
        ).ToList()
    End Sub
End Class