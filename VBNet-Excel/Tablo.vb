Imports System.Collections.Generic
Imports System.Diagnostics.Eventing
Imports System.Diagnostics.Eventing.Reader
Imports System.Drawing.Drawing2D
Imports System.Net
Imports System.Net.Security
Imports System.Reflection
Imports System.Security.Cryptography.X509Certificates
Imports System.Text.RegularExpressions
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.TrackBar
Imports System.Xml
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.DatabaseServices.Filters
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Internal.DatabaseServices
Imports Autodesk.AutoCAD.PlottingServices
Imports Autodesk.AutoCAD.Runtime
Imports Org.BouncyCastle.Asn1.Pkcs
Imports SWF = System.Windows.Forms

Public Class Tablo
    Dim PI As Double = 3.1415926535897931
    Dim cu As CommonUtil = New CommonUtil()
    Dim brTabla As Integer = 100
    Dim brTokKryg As Integer = 100
    Dim brKonsumator As Integer = 50
    Dim brKonsum As Integer = 10000
    Dim form_AS_tablo As New Form_Tablo()
    Dim appNameKonso As String = "EWG_KONSO"
    Dim appNameTablo As String = "EWG_TABLO"
    Dim za6t As Integer = 14
    Dim defkt As Integer = 8
    Dim innd_106 As Integer = 0
    Private isCellChangeTriggeredProgrammatically As Boolean = False
    '
    ' Електрически константи
    '
    Const Knti As Double = 1.125        ' коефициeнт за термичния изключвател

    Dim IcableDict As New Dictionary(Of String, Integer())
    Dim Kable_Size_L() As String
    Dim Kable_Size_N() As String
    Dim Breakers As New Dictionary(Of Integer, String)
    Dim Busbar_Cu As New Dictionary(Of Integer, String)
    Dim Busbar_Al As New Dictionary(Of Integer, String)
    Dim Cable_AlR_2 As New Dictionary(Of Integer, String)
    Dim Cable_AlR_4 As New Dictionary(Of Integer, String)
    Dim Disconnectors As New Dictionary(Of Integer, String)
    Dim RCD_Catalog As New List(Of strRCD)

    ' ===============================
    ' Глобални променливи за таблата
    ' ===============================
    Public widthColom As Double = 120      ' Ширина на всяка колона в таблицата
    Public heightRow As Double = 25        ' Височина на редовете
    Public widthText As Double = 140       ' Ширина на колоната за текст (напр. "Токов кръг")
    Public widthTextDim As Double = 40     ' Допълнителна ширина за текстова колона (напр. за единици)
    Public lengthProw As Double = 90       ' Дължина на вертикалните линии между текст и блокове
    Public lengthProwBlock As Double = 0   ' Дължина на линията под блока за прекъсвач (ще се изчислява по-късно)
    Public padingText As Double = 3        ' Отстояние на текста от линиите
    Public widthTablo As Double = 410      ' Ширина на цялото табло (за блокове и линии)
    Public heightText As Double = 12       ' Височина на текста, използван в блоковете
    Public Y_Шина As Double = 620          ' Вертикална позиция на шината (Y координата)

    Public Structure strTokow
        Dim ТоковКръг As String         ' Номер на токов кръг
        Dim brLamp As Integer           ' Брой лампи
        Dim brKontakt As Integer        ' Брой контакти
        Dim Мощност As Double           ' Мощност на токов кръг - в kW
        Dim Kabebel_Se4enie As String   ' Сечение на кабела
        Dim faza As String              ' Фаза
        Dim konsuator1 As String
        Dim konsuator2 As String
        'Изчислителни полета
        Dim BrojPol As String           ' Брой на полюсите
        ' Ток на токовия кръг
        ' За трифазен консуматор
        ' .Мощност * 1.2 / (0.38 * Math.Sqrt(3) * 0.9)
        '
        ' За монофазен консуматор
        ' .Мощност * 1.2 / (0.22 * 0.9)
        '
        Dim Tok As Double               ' Ток на токовия кръг
        ' Полета за защита
        Dim BlockName As String         ' Име на блок който се вмъква
        Dim Designation As String
        Dim ShortName As String         ' Кратко име - вид на апарата
        Dim Type As String              ' 
        Dim NumberPoles As String       ' Брой на модулите / 
        Dim RatedCurrent As String      ' Номинален ток
        Dim Curve As String             ' Крива
        Dim Current As String
        Dim Control As String
        Dim Sensitivity As String       ' Изключвателна възможност
        Dim Protection As String
        Dim RCD_Name As String          ' EZ9 RCCB или EZ9 RCBO
        Dim RCD_BlockName As String
        Dim Konsumator() As strKonsumator
    End Structure
    Public Structure strTablo
        Dim Name As String
        Dim countTokKryg As Integer
        Dim Tokowkryg() As strTokow
        Dim TabloType As String
    End Structure
    Public Structure strKonsumator
        Dim Name As String              ' Име на блока
        Dim ID_Block As ObjectId        ' Блок на елемента
        Dim ТоковКръг As String         ' Токов кръг към който е свързан
        Dim strМОЩНОСТ As String        ' Мощност от блока
        Dim doubМОЩНОСТ As Double       ' Изчислена мощност
        Dim ТАБЛО As String             ' Табло към което е включен токовия кръг
        Dim Pewdn As String             ' Предназначение 
        Dim PEWDN1 As String            ' Предназначение 
        Dim Dylvina_Led As Double       ' Дължина на LED лента
        Dim Visibility As String        ' 
    End Structure
    ' Декларация на структура ElectricalParameters
    Public Structure ElectricalParameters
        Dim TKryg As String         ' Име на токовия кръг
        Dim Power As Double         ' Мощност на токовия круг
        Dim Current As Double       ' Ток на кръга
        Dim CircuitType As String   ' Фаза
        Dim Phases As String        ' брой фази --- "1p"/"3p"
        Dim RCD As String           ' шина на която е ДЗТ
        Dim Columns As Integer      ' номер на колона
        Dim Bus As Boolean          ' True ако е на ралична шина
    End Structure
    ' Дефинираме структура за защитно устройство RCD
    Private Structure strRCD
        Public NominalCurrent As Double ' Номинален ток (A)
        Public Type As String           ' Тип: "AC", "A", "si"
        Public Poles As String          ' Брой полюси: "2p", "4p"
        Public Sensitivity As Double    ' Чувствителност (mA), ако е приложимо
        Public DeviceType As String     ' Вид устройство: "RCCB", "RCBO", "iID"
    End Structure
    <CommandMethod("Tablo")>
    Public Sub Tablo()
        Dim arrTablo(brTabla) As strTablo
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim edt As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database
        Dim SelectedSet = cu.GetObjects("INSERT", "Изберете блок")
        If SelectedSet Is Nothing Then
            MsgBox("НЕ Е маркиран нито един блок.")
            Exit Sub
        End If
        ' Запълва речниците и масивите с данни
        SetCatalog()
        Dim brTablo As Integer = 0
        Dim FEC_KRYG As Integer = 1
        Dim blkRecId As ObjectId = ObjectId.Null
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                Dim acBlkTbl As BlockTable
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim arrKonsum(brKonsum) As ObjectId
                Dim index As Integer = 0
                For Each sObj As SelectedObject In SelectedSet
                    blkRecId = sObj.ObjectId
                    arrKonsum(index) = sObj.ObjectId
                    index += 1
                Next
                For Each sObj As ObjectId In arrKonsum
                    'blkRecId = sObj
                    If sObj.IsNull Then Exit For
                    innd_106 = +1
                    Dim acBlkRef As BlockReference = DirectCast(acTrans.GetObject(sObj, OpenMode.ForWrite), BlockReference)
                    Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                    Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection

                    Dim ТоковКръг As String = ""
                    Dim strМОЩНОСТ As String = ""
                    Dim doubМОЩНОСТ As Double = 0.0
                    Dim ТАБЛО As String = ""
                    Dim Pewdn As String = ""
                    Dim PEWDN1 As String = ""
                    Dim countTokKryg As Integer = 0
                    Dim Dylvina_Led As Double
                    Dim Broj_Kontakti As Integer = 0
                    Dim Broj_Lampi As Integer = 0

                    Dim Kab_Tip As String = ""
                    Dim Kab_Se4 As String = ""
                    Dim Faza As String = ""
                    Dim Name_Blok As String = (CType(acBlkRef.DynamicBlockTableRecord.GetObject(OpenMode.ForRead), BlockTableRecord)).Name
                    For Each objID As ObjectId In attCol
                        Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                        Dim acAttRef As AttributeReference = dbObj
                        If Name_Blok = "Батерия" Then
                            ТоковКръг = FEC_KRYG.ToString

                            FEC_KRYG += 1
                        End If
                        If Name_Blok = "Табло и инвертор" Then
                            ТоковКръг = FEC_KRYG.ToString
                            FEC_KRYG += 1
                        End If

                        If acAttRef.Tag = "МОЩНОСТ" Then strМОЩНОСТ = acAttRef.TextString
                        If acAttRef.Tag = "LED" Then strМОЩНОСТ = acAttRef.TextString
                        If acAttRef.Tag = "КРЪГ" Then ТоковКръг = acAttRef.TextString
                        If acAttRef.Tag = "ТАБЛО" Then ТАБЛО = acAttRef.TextString
                        If acAttRef.Tag = "Pewdn" Then Pewdn = acAttRef.TextString
                        If acAttRef.Tag = "PEWDN1" Then PEWDN1 = acAttRef.TextString

                        If acAttRef.Tag = "ТОКОВ_КРЪГ" Then ТоковКръг = acAttRef.TextString
                        If acAttRef.Tag = "ГЛ.Р.Т." Then PEWDN1 = acAttRef.TextString
                        If acAttRef.Tag = "БРОЙ_ЛАМПИ" Then Broj_Lampi = acAttRef.TextString
                        If acAttRef.Tag = "БРОЙ_КОНТАКТИ" Then Broj_Kontakti = acAttRef.TextString
                        If acAttRef.Tag = "КАБЕЛ_ТИП" Then Kab_Tip = acAttRef.TextString
                        If acAttRef.Tag = "КАБЕЛ_СЕЧЕНИЕ" Then Kab_Se4 = acAttRef.TextString
                        If acAttRef.Tag = "ФАЗИ" Then Faza = acAttRef.TextString
                        If acAttRef.Tag = "ИМЕ" Then PEWDN1 = acAttRef.TextString
                    Next
                    Dim Visibility As String = ""
                    For Each prop As DynamicBlockReferenceProperty In props
                        If prop.PropertyName = "Visibility1" Then Visibility = prop.Value
                        If prop.PropertyName = "Visibility" Then Visibility = prop.Value
                        If prop.PropertyName = "Дължина" Then Dylvina_Led = prop.Value
                    Next
                    ' Проверка дали текстът съдържа "led/m" (т.е. LED лента)
                    If strМОЩНОСТ.ToLower().Contains("led/m") Then
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
                        doubМОЩНОСТ = (Dylvina_Led / 100) * мощностНаМетър
                        ' Записваме резултата като текст с 2 десетични места
                        strМОЩНОСТ = doubМОЩНОСТ.ToString("0.##")
                    End If

                    Select Case Visibility
                        Case "Само ключ",
                             "Лампион - рошав", "Лампион", "Настолна лампа - рошава",
                             "Настолна лампа", "Фотодатчик", "Датчик 360°", "Датчик насочен",
                             "Драйвер", "ПВ", "Линии", "Само текст", "Табло_Ново"
                            Continue For
                    End Select

                    Dim brМОЩНОСТ, moМОЩНОСТ As String
                    Dim poz As Integer = Math.Max(InStr(strМОЩНОСТ, "х"), InStr(strМОЩНОСТ, "x"))
                    If poz > 0 Then
                        brМОЩНОСТ = Mid(strМОЩНОСТ, 1, poz - 1)
                        moМОЩНОСТ = Mid(strМОЩНОСТ, poz + 1, Len(strМОЩНОСТ))
                        doubМОЩНОСТ = Val(brМОЩНОСТ) * Val(moМОЩНОСТ)
                    Else
                        doubМОЩНОСТ = Val(strМОЩНОСТ)
                    End If
                    If Not (Name_Blok = "Табло и инвертор") Then
                        doubМОЩНОСТ = doubМОЩНОСТ / 1000
                    End If

                    If ТАБЛО = "" Or ТАБЛО = "Табло" Then ТАБЛО = "Гл.Р.Т."
                    Dim iTablo As Integer = Array.FindIndex(arrTablo, Function(f) f.Name = ТАБЛО)
                    If iTablo = -1 Then
                        arrTablo(brTablo).Name = ТАБЛО
                        iTablo = brTablo
                        ReDim arrTablo(brTablo).Tokowkryg(brTokKryg)
                        arrTablo(brTablo).countTokKryg = 0
                        brTablo += 1
                    End If

                    Dim iKryg As Integer = Array.FindIndex(arrTablo(iTablo).Tokowkryg, Function(f) f.ТоковКръг = ТоковКръг)

                    If iKryg = -1 Then
                        With arrTablo(iTablo).Tokowkryg(arrTablo(iTablo).countTokKryg)
                            .ТоковКръг = ТоковКръг
                            .Мощност += doubМОЩНОСТ
                            .konsuator1 = Pewdn
                            .konsuator2 = PEWDN1
                        End With
                        iKryg = arrTablo(iTablo).countTokKryg
                        arrTablo(iTablo).countTokKryg += 1
                    Else
                        With arrTablo(iTablo).Tokowkryg(iKryg)
                            .Мощност += doubМОЩНОСТ
                            .konsuator1 = Pewdn
                            .konsuator2 = PEWDN1
                        End With
                    End If

                    'Name_Blok = (CType(acBlkRef.DynamicBlockTableRecord.GetObject(OpenMode.ForRead), BlockTableRecord)).Name

                    With arrTablo(iTablo).Tokowkryg(iKryg)
                        .BlockName = "s_c60_circ_break"
                        .Curve = "С"
                        .Sensitivity = "N"
                        Select Case Name_Blok
                            Case "Батерия"
                                .Tok = .Tok
                                .konsuator1 = "Батерия"
                                .NumberPoles = "3p"
                                .Tok = calc_Inom(.Мощност, .NumberPoles)
                            Case "Табло и инвертор"
                                .konsuator1 = Visibility
                                .NumberPoles = "3p"
                                .Tok = calc_Inom(.Мощност, .NumberPoles)
                            Case "Авария"
                                .brLamp += 1
                                .konsuator1 = "Аварийно"
                                .konsuator2 = "осветление"
                                .faza = "L"
                                .NumberPoles = "1p"
                                .Tok = calc_Inom(.Мощност, .NumberPoles)
                                .Kabebel_Se4enie = "3x1,5"
                                .RatedCurrent = "10"
                            Case "Линия МХЛ - 220V", "Полилей",
                                 "Авария_100", "LED_ULTRALUX", "LED_ULTRALUX_100",
                                 "Прожектор", "LED_lenta",
                                 "Металхаогенна лампа", "uli4no", "LED_DENIMA"
                                .brLamp += 1
                                .konsuator1 = "Общо"
                                .konsuator2 = "осветление"
                                .faza = "L"
                                .NumberPoles = "1p"
                                .Tok = calc_Inom(.Мощност, .NumberPoles)
                                .RatedCurrent = "10"
                                .Kabebel_Se4enie = "3x1,5"
                            Case "Плафони", "LED_луна"
                                .brLamp += 1
                                .NumberPoles = "1p"
                                .konsuator1 = "Общо"
                                .konsuator2 = "осветление"
                                .faza = "L"
                                .Tok = calc_Inom(.Мощност, .NumberPoles)
                                .Kabebel_Se4enie = "3x1,5"
                                .RatedCurrent = "10"
                            Case "Контакт"
                                .NumberPoles = "1p"
                                .brKontakt += 1
                                .konsuator1 = "Контакти"
                                .konsuator2 = ""
                                .Tok = calc_Inom(.Мощност, .NumberPoles)
                                .RatedCurrent = "20"
                                .faza = "L"
                                .Kabebel_Se4enie = "3x2,5"
                                Select Case Visibility
                                    Case "Двугнездов", "Двугнездов - противовлажен", "2xU"
                                        .brKontakt += 1
                                    Case "Тригнездов", "Тригнездов - противовлажен"
                                        .brKontakt += 2
                                    Case "Трифазен", "Трифазен - IP 54", "Трифазен - противовлажен", "ТР+2МФ"
                                        If Visibility = "ТР+2МФ" Then .brKontakt += 2
                                        .faza = "L1,L2,L3"
                                        .Kabebel_Se4enie = "5x2,5"
                                        .RatedCurrent = "25"
                                        .NumberPoles = "3p"
                                        .Tok = calc_Inom(.Мощност, .NumberPoles)
                                    Case "Твърда връзка", "Усилен"
                                        .konsuator1 = "Ел. печка"
                                        .faza = "L"
                                        .Kabebel_Se4enie = "3x4,0"
                                    Case Else
                                        .faza = "L"
                                        .Kabebel_Se4enie = "3x2,5"
                                End Select
                            Case "бойлерно табло"
                                Select Case Visibility
                                    Case "Ключ и контакт", "С два ключа и контакт"
                                        .brKontakt += 1
                                        .Tok = calc_Inom(.Мощност, .NumberPoles)
                                    Case "С два контакта и един ключ"
                                        .brKontakt += 2
                                        .Tok = calc_Inom(.Мощност, .NumberPoles)
                                End Select
                            Case "Бойлер"
                                .konsuator1 = .konsuator1
                                .konsuator2 = .konsuator2
                                Select Case Visibility
                                    Case "Хоризонтален", "Вертикален"
                                        .faza = "L"
                                        .konsuator1 = "Бойлер"
                                        .konsuator2 = ""
                                        .NumberPoles = "1p"
                                        .Tok = calc_Inom(.Мощност, .NumberPoles)
                                        .RatedCurrent = calc_breaker_EZ9(17)
                                        .Kabebel_Se4enie = calc_cable_Cu(.RatedCurrent, .NumberPoles)
                                        .Tok = calc_Inom(.Мощност, .NumberPoles)
                                    Case "Бойлер кухня"
                                        .faza = "L"
                                        .NumberPoles = "1p"
                                        .konsuator1 = "Бойлер"
                                        .konsuator2 = "10л"
                                        .RatedCurrent = calc_breaker_EZ9(17)
                                        .Kabebel_Se4enie = calc_cable_Cu(.RatedCurrent, .NumberPoles)
                                        .Tok = calc_Inom(.Мощност, .NumberPoles)
                                    Case "Проточен"
                                        .faza = "L"
                                        .konsuator1 = "Бойлер"
                                        .konsuator2 = "проточен"
                                        .NumberPoles = "1p"
                                        .Tok = calc_Inom(.Мощност, .NumberPoles)
                                        If .Мощност > 3 Then
                                            .RatedCurrent = calc_breaker_EZ9(.Tok)
                                        Else
                                            .RatedCurrent = calc_breaker_EZ9(17)
                                        End If
                                        .Kabebel_Se4enie = calc_cable_Cu(.RatedCurrent, .NumberPoles)
                                    Case "Изход 1p", "Сешоар", "Сешоар с контакт"
                                        .faza = "L"
                                        .NumberPoles = "1p"
                                        .Tok = calc_Inom(.Мощност, .NumberPoles)
                                        If .Мощност > 3 Then
                                            .RatedCurrent = calc_breaker_EZ9(.Tok)
                                        Else
                                            .RatedCurrent = calc_breaker_EZ9(17)
                                        End If
                                        .Kabebel_Se4enie = calc_cable_Cu(.RatedCurrent, .NumberPoles)
                                    Case "Изход 3p"
                                        .faza = "L1,L2,L3"
                                        .NumberPoles = "3p"
                                        .Tok = calc_Inom(.Мощност, .NumberPoles)
                                        .RatedCurrent = calc_breaker_EZ9(.Tok)
                                        .Kabebel_Se4enie = calc_cable_Cu(.RatedCurrent, .NumberPoles)
                                    Case "Хоризонтален - 380V", "Вертикален - 380V"
                                        .faza = "L1,L2,L3"
                                        .konsuator1 = "Бойлер"
                                        .konsuator2 = ""
                                        .NumberPoles = "3p"
                                        .Tok = calc_Inom(.Мощност, .NumberPoles)
                                        .RatedCurrent = calc_breaker_EZ9(.Tok)
                                        .Kabebel_Se4enie = calc_cable_Cu(.RatedCurrent, .NumberPoles)
                                    Case "Проточен - 380V"
                                        .faza = "L1,L2,L3"
                                        .konsuator1 = "Бойлер"
                                        .konsuator2 = "проточен"
                                        .NumberPoles = "3p"
                                        .Tok = calc_Inom(.Мощност, .NumberPoles)
                                        .RatedCurrent = calc_breaker_EZ9(.Tok)
                                        .Kabebel_Se4enie = calc_cable_Cu(.RatedCurrent, .NumberPoles)
                                    Case "Изход газ"
                                        .faza = "L"
                                        .NumberPoles = "1p"
                                        .konsuator1 = "Изход за"
                                        .konsuator2 = "газ"
                                        .Tok = calc_Inom(.Мощност, .NumberPoles)
                                        .Kabebel_Se4enie = "3x1,5"
                                        .RatedCurrent = "6"
                                    Case Else
                                        arrTablo(iTablo).Tokowkryg(iKryg).faza = "######"
                                End Select
                                .Kabebel_Se4enie = FixValue(.Kabebel_Se4enie)
                            Case "Вентилации"
                                Select Case Visibility
                                    Case "Вентилатор - канален 3P", "Вентилатор - прозоречен 3P"
                                        arrTablo(iTablo).Tokowkryg(iKryg).faza = "L1,L2,L3"
                                        arrTablo(iTablo).Tokowkryg(iKryg).Tok = arrTablo(iTablo).Tokowkryg(iKryg).Мощност * 1.2 / (0.38 * Math.Sqrt(3) * 0.9)
                                        arrTablo(iTablo).Tokowkryg(iKryg).NumberPoles = "3p"
                                        Select Case arrTablo(iTablo).Tokowkryg(iKryg).Tok
                                            Case < 6        ' АП - 6А
                                                arrTablo(iTablo).Tokowkryg(iKryg).Kabebel_Se4enie = "5x1,5"
                                                arrTablo(iTablo).Tokowkryg(iKryg).RatedCurrent = "6"
                                            Case < 10       ' АП - 10А
                                                arrTablo(iTablo).Tokowkryg(iKryg).Kabebel_Se4enie = "5x1,5"
                                                arrTablo(iTablo).Tokowkryg(iKryg).RatedCurrent = "10"
                                            Case < 16       ' АП - 60А
                                                arrTablo(iTablo).Tokowkryg(iKryg).Kabebel_Se4enie = "5x2,5"
                                                arrTablo(iTablo).Tokowkryg(iKryg).RatedCurrent = "16"
                                            Case < 20       ' АП - 20А
                                                arrTablo(iTablo).Tokowkryg(iKryg).Kabebel_Se4enie = "5x2,5"
                                                arrTablo(iTablo).Tokowkryg(iKryg).RatedCurrent = "20"
                                            Case < 25       ' АП - 25А
                                                arrTablo(iTablo).Tokowkryg(iKryg).Kabebel_Se4enie = "5x4,0"
                                                arrTablo(iTablo).Tokowkryg(iKryg).RatedCurrent = "25"
                                            Case < 32       ' АП - 32А
                                                arrTablo(iTablo).Tokowkryg(iKryg).Kabebel_Se4enie = "5x6,0"
                                                arrTablo(iTablo).Tokowkryg(iKryg).RatedCurrent = "32"
                                            Case < 40       ' АП - 40А
                                                arrTablo(iTablo).Tokowkryg(iKryg).Kabebel_Se4enie = "5x10,0"
                                                arrTablo(iTablo).Tokowkryg(iKryg).RatedCurrent = "40"
                                            Case < 50       ' АП - 50А
                                                arrTablo(iTablo).Tokowkryg(iKryg).Kabebel_Se4enie = "5x10,0"
                                                arrTablo(iTablo).Tokowkryg(iKryg).RatedCurrent = "50"
                                            Case < 63       ' АП - 63А
                                                arrTablo(iTablo).Tokowkryg(iKryg).Kabebel_Se4enie = "5x16,0"
                                                arrTablo(iTablo).Tokowkryg(iKryg).RatedCurrent = "63"
                                            Case < 80       ' АП - 80А
                                                arrTablo(iTablo).Tokowkryg(iKryg).Kabebel_Se4enie = "5x25,0"
                                                arrTablo(iTablo).Tokowkryg(iKryg).RatedCurrent = "80"
                                            Case < 100      ' АП - 100А
                                                arrTablo(iTablo).Tokowkryg(iKryg).Kabebel_Se4enie = "5x35,0"
                                                arrTablo(iTablo).Tokowkryg(iKryg).RatedCurrent = "100"
                                            Case < 125      ' АП - 125А
                                                arrTablo(iTablo).Tokowkryg(iKryg).Kabebel_Se4enie = "5x50,0"
                                                arrTablo(iTablo).Tokowkryg(iKryg).RatedCurrent = "125"
                                            Case Else
                                                arrTablo(iTablo).Tokowkryg(iKryg).Kabebel_Se4enie = "######"
                                        End Select
                                    Case Else
                                        arrTablo(iTablo).Tokowkryg(iKryg).faza = "L"
                                        arrTablo(iTablo).Tokowkryg(iKryg).Tok = arrTablo(iTablo).Tokowkryg(iKryg).Мощност * 1.2 / (0.22 * 0.9)
                                        arrTablo(iTablo).Tokowkryg(iKryg).NumberPoles = "1p"
                                        Select Case arrTablo(iTablo).Tokowkryg(iKryg).Tok
                                            Case < 6        ' АП - 6А
                                                arrTablo(iTablo).Tokowkryg(iKryg).Kabebel_Se4enie = "3x1,5"
                                                arrTablo(iTablo).Tokowkryg(iKryg).RatedCurrent = "6"
                                            Case < 10       ' АП - 10А
                                                arrTablo(iTablo).Tokowkryg(iKryg).Kabebel_Se4enie = "3x1,5"
                                                arrTablo(iTablo).Tokowkryg(iKryg).RatedCurrent = "10"
                                            Case < 16       ' АП - 16А
                                                arrTablo(iTablo).Tokowkryg(iKryg).Kabebel_Se4enie = "3x2,5"
                                                arrTablo(iTablo).Tokowkryg(iKryg).RatedCurrent = "16"
                                            Case < 20       ' АП - 20А
                                                arrTablo(iTablo).Tokowkryg(iKryg).Kabebel_Se4enie = "3x2,5"
                                            Case < 25       ' АП - 25А
                                                arrTablo(iTablo).Tokowkryg(iKryg).Kabebel_Se4enie = "3x4,0"
                                                arrTablo(iTablo).Tokowkryg(iKryg).RatedCurrent = "25"
                                            Case < 32       ' АП - 32А
                                                arrTablo(iTablo).Tokowkryg(iKryg).Kabebel_Se4enie = "3x6,0"
                                                arrTablo(iTablo).Tokowkryg(iKryg).RatedCurrent = "32"
                                            Case < 40       ' АП - 40А
                                                arrTablo(iTablo).Tokowkryg(iKryg).Kabebel_Se4enie = "3x10,0"
                                                arrTablo(iTablo).Tokowkryg(iKryg).RatedCurrent = "40"
                                            Case < 50       ' АП - 50А
                                                arrTablo(iTablo).Tokowkryg(iKryg).Kabebel_Se4enie = "3x10,0"
                                                arrTablo(iTablo).Tokowkryg(iKryg).RatedCurrent = "50"
                                            Case < 63       ' АП - 63А
                                                arrTablo(iTablo).Tokowkryg(iKryg).Kabebel_Se4enie = "3x16,0"
                                                arrTablo(iTablo).Tokowkryg(iKryg).RatedCurrent = "63"
                                            Case < 80       ' АП - 80А
                                                arrTablo(iTablo).Tokowkryg(iKryg).Kabebel_Se4enie = "3x25,0"
                                                arrTablo(iTablo).Tokowkryg(iKryg).RatedCurrent = "80"
                                            Case < 100      ' АП - 100А
                                                arrTablo(iTablo).Tokowkryg(iKryg).Kabebel_Se4enie = "3x35,0"
                                                arrTablo(iTablo).Tokowkryg(iKryg).RatedCurrent = "100"
                                            Case < 125      ' АП - 125А
                                                arrTablo(iTablo).Tokowkryg(iKryg).Kabebel_Se4enie = "3x50,0"
                                                arrTablo(iTablo).Tokowkryg(iKryg).RatedCurrent = "125"
                                            Case Else
                                                arrTablo(iTablo).Tokowkryg(iKryg).Kabebel_Se4enie = "######"
                                        End Select
                                End Select
                            Case "Табло_Главно"
                                .brLamp = Broj_Lampi
                                .brKontakt = Broj_Kontakti
                                .konsuator1 = "Табло"
                                .konsuator2 = PEWDN1
                                .faza = Faza
                                .NumberPoles = IIf(Faza = "L1,L2,L3", "3p", "1p")
                                .Tok = calc_Inom(.Мощност, .NumberPoles)
                                .RatedCurrent = calc_breaker_EZ9(.Tok)
                                .Kabebel_Se4enie = calc_cable_Cu(.RatedCurrent, .NumberPoles)
                        End Select
                        .Kabebel_Se4enie = FixValue(.Kabebel_Se4enie)
                    End With
                Next
                For i = 0 To arrTablo.Count - 1
                    If arrTablo(i).Name Is Nothing Then Exit For
                    For j = 0 To arrTablo(i).Tokowkryg.Count - 1
                        Dim TK As strTokow = arrTablo(i).Tokowkryg(j)
                        ' Да се направи поверка за консумато1 дали е "Табло" и ако да излиза от цикъка
                        If TK.ТоковКръг Is Nothing Then Exit For
                        If TK.konsuator1 = "Табло" Then Continue For
                        If TK.konsuator1 = "Ел. печка" Then Continue For

                        If TK.brLamp > 0 Then
                            If TK.konsuator1 <> "Аварийно" Then
                                TK.konsuator1 = "Общо"
                                TK.konsuator2 = "осветление"
                            End If
                            TK.faza = "L"
                            TK.Kabebel_Se4enie = "3x1,5"
                            TK.RatedCurrent = "10"
                        End If
                        If TK.brKontakt > 0 Then
                            If TK.faza = "L1,L2,L3" Then
                                TK.NumberPoles = "3p"
                                TK.konsuator1 = "Контакти"
                                TK.konsuator2 = ""
                                TK.RatedCurrent = "20"
                                TK.faza = "L1,L2,L3"
                                TK.Kabebel_Se4enie = "5x2,5"
                            Else
                                TK.NumberPoles = "1p"
                                TK.konsuator1 = "Контакти"
                                TK.konsuator2 = ""
                                TK.RatedCurrent = "20"
                                TK.faza = "L"
                                TK.Kabebel_Se4enie = "3x2,5"
                            End If
                        End If
                        arrTablo(i).Tokowkryg(j) = TK ' Записване на променения обект обратно
                    Next
                Next
                '
                ' Сортира токовите кръгове във всяко табло
                ' Използва се insertion sort върху масива Tokowkryg
                '
                For i = 0 To arrTablo.Count
                    ' Ако таблото няма име – прекратяваме обработката
                    If arrTablo(i).Name = Nothing Then Exit For
                    Dim pointer As Integer = 0
                    Dim posicion As Integer = 0
                    Dim curent As strTokow
                    Dim ind As Integer = 0
                    ' Определяме реалния брой на валидните токови кръгове
                    For Each TK As strTokow In arrTablo(i).Tokowkryg
                        ' Ако токовият кръг е празен – прекратяваме броенето
                        If TK.ТоковКръг = Nothing Then Exit For
                        ind += 1
                    Next
                    ' Insertion sort по име / номер на токов кръг
                    For pointer = 1 To ind - 1
                        ' Запазваме текущия елемент
                        curent = arrTablo(i).Tokowkryg(pointer)
                        posicion = pointer
                        ' Преместваме елементите надясно,
                        ' докато намерим правилната позиция
                        Do While posicion > 0 AndAlso
                            Compare(arrTablo(i).Tokowkryg(posicion - 1).ТоковКръг,
                                    curent.ТоковКръг)
                            arrTablo(i).Tokowkryg(posicion) =
                                arrTablo(i).Tokowkryg(posicion - 1)
                            posicion -= 1
                        Loop
                        ' Поставяме текущия токов кръг на намерената позиция
                        arrTablo(i).Tokowkryg(posicion) = curent
                    Next
                Next
                '
                ' Изтрива всички TabPages.
                ' В противен случай остават параметри в DataGridView
                '
                form_AS_tablo.TabControl1.TabPages.Clear()

                For i = 0 To arrTablo.Count
                    If arrTablo(i).Name = Nothing Then Exit For
                    If form_AS_tablo.TabControl1.TabPages.Count > i Then
                        form_AS_tablo.TabControl1.TabPages.Item(i).Name = arrTablo(i).Name
                    Else
                        form_AS_tablo.TabControl1.TabPages.Add(arrTablo(i).Name)
                    End If
                    form_AS_tablo.TabControl1.TabPages.Item(i).Text = arrTablo(i).Name
                    form_AS_tablo.TabControl1.TabPages.Item(i).Controls.Add(insSplitContainer(arrTablo(i).Name, arrTablo(i)))
                Next
                acTrans.Commit()
                Application.ShowModalDialog(form_AS_tablo)
            Catch ex As Exception
                innd_106 = innd_106
                MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
                acTrans.Abort()
            End Try
        End Using
    End Sub
    ''' <summary>
    ''' Проверява дали стойността след "x" (латиница или кирилица) е по-малка от 2,5.
    ''' Ако е по-малка, връща резултат с минимална стойност 2,5.
    ''' Винаги връща резултат с латинско "x" и запетая като десетичен разделител.
    ''' </summary>
    ''' <param name="input">Текстов низ във формат "aхb,c" или "ax b,c".</param>
    ''' <returns>Преобразуван низ във формат "a x b,c".</returns>
    Private Function FixValue(input As String) As String
        ' Проверка дали подаденият низ е празен или само с интервали
        If String.IsNullOrWhiteSpace(input) Then Return input
        ' Заместваме кирилското "х" (U+0445) с латинското "x" (U+0078)
        ' Това унифицира разделителя, за да работим еднакво с двата варианта
        Dim normalized As String = input.Replace("х"c, "x"c)
        ' Разделяме низа на две части по символа "x"
        ' Очакваме parts(0) = лявата част (напр. "5")
        ' и parts(1) = дясната част (напр. "1,5")
        Dim parts() As String = normalized.Split("x"c)
        ' Ако разделянето не дава точно две части, връщаме оригиналния текст
        If parts.Length <> 2 Then Return input
        ' Премахваме излишни интервали около двете части
        Dim leftPart As String = parts(0).Trim()
        Dim rightPart As String = parts(1).Trim()
        ' Опитваме се да конвертираме дясната част в число (Double)
        ' Първо заменяме запетаята с точка, за да е съвместимо с InvariantCulture
        Dim num As Double
        If Double.TryParse(rightPart.Replace(","c, "."c),
                       Globalization.NumberStyles.Any,
                       Globalization.CultureInfo.InvariantCulture, num) Then
            ' Ако стойността след "x" е по-малка от 2,5 → фиксираме я на 2,5
            If num < 2.5 Then num = 2.5
            ' Връщаме новия текст във формат:
            ' - лява част без промяна
            ' - латинско "x"
            ' - дясна част като число с една десетична позиция и запетая като разделител
            Return $"{leftPart}x{num.ToString("0.0", Globalization.CultureInfo.InvariantCulture).Replace("."c, ","c)}"
        End If
        ' Ако дясната част не може да се парсне като число, връщаме оригиналния текст
        Return input
    End Function
    ''' <summary>
    ''' Сравнява два токови кръга с цел сортиране.
    ''' Приоритетът е:
    ''' 1) Кръгове съдържащи "ав"
    ''' 2) Кръгове съдържащи "до"
    ''' 3) Всички останали (чисти числа или други означения).
    ''' При еднакъв тип се сравнява числовата стойност,
    ''' а при пълно съвпадение – лексикографски (без значение от регистъра).
    ''' </summary>
    ''' <param name="a">Първият токов кръг за сравнение.</param>
    ''' <param name="b">Вторият токов кръг за сравнение.</param>
    ''' <returns>
    ''' True, ако <paramref name="a"/> трябва да бъде подреден след <paramref name="b"/>,
    ''' False – в противен случай.
    ''' </returns>
    Function Compare(a As String, b As String) As Boolean
        ' Помощна функция, която определя приоритет на токовия кръг
        ' По-ниската стойност означава по-висок приоритет при сортиране
        Dim getPriority = Function(s As String) As Integer
                              ' Кръгове съдържащи "ав" (автоматични) – най-висок приоритет
                              If s.Contains("ав") Then Return 1
                              ' Кръгове съдържащи "до" – втори приоритет
                              If s.Contains("до") Then Return 2
                              ' Всички останали (чисти числа или други означения)
                              Return 3
                          End Function
        ' Определяме приоритета на двата сравнявани стринга
        Dim pA = getPriority(a.ToLower())
        Dim pB = getPriority(b.ToLower())
        ' 1. Ако двата токови кръга са от различен тип,
        ' сортираме ги по приоритет
        If pA <> pB Then
            ' Връща True, ако A трябва да застане СЛЕД B
            Return pA > pB
        End If
        ' 2. Ако са от един и същи тип,
        ' сравняваме числовата част на означението
        ' (извличаме само цифрите от стринга)
        Dim numA As Double = Val(Regex.Match(a, "\d+").Value)
        Dim numB As Double = Val(Regex.Match(b, "\d+").Value)
        If numA <> numB Then
            ' По-голямото число се подрежда след по-малкото
            Return numA > numB
        End If
        ' 3. Ако и числовите стойности съвпадат,
        ' правим финално текстово сравнение (без значение от регистъра)
        Return String.Compare(a, b, StringComparison.OrdinalIgnoreCase) > 0
    End Function
    <CommandMethod("InsertBlockTablo")>
    Public Sub InsertBlockTablo()
        ' Get the current database and start a transaction
        Dim acCurDb As Autodesk.AutoCAD.DatabaseServices.Database
        acCurDb = Application.DocumentManager.MdiActiveDocument.Database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim SelBlock = cu.GetObjects("INSERT", "Изберете БЛОК")
        Dim SelTablo = cu.GetObjects("INSERT", "Изберете ТАБЛО")

        If SelBlock Is Nothing Then
            MsgBox("Въпрос с повишена трудност." + vbCrLf + "Аз какво да установявам???")
            Exit Sub
        End If
        If SelTablo Is Nothing Then
            MsgBox("НЕ Е маркиран поне един блок за табло")
            Exit Sub
        End If
        If SelTablo.Count > 1 Then
            MsgBox("НЕ Е маркиран САМО един блок за табло")
            Exit Sub
        End If
        Try
            Dim blkRecIdTablo As ObjectId = SelTablo(0).ObjectId
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim strTablo As String = ""

                'blkRecIdTablo = sObj.ObjectId
                Dim acBlkRef As BlockReference = DirectCast(acTrans.GetObject(blkRecIdTablo,
                                                                                 OpenMode.ForWrite), BlockReference)
                Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                For Each objID As ObjectId In attCol
                    Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                    Dim acAttRef As AttributeReference = dbObj
                    If acAttRef.Tag = "ТАБЛО" Then strTablo = acAttRef.TextString
                Next

                Dim Xdata As ResultBuffer = GetXData(appNameKonso, blkRecIdTablo)

                Using rb As ResultBuffer = IIf(IsNothing(Xdata), New ResultBuffer, Xdata)
                    If IsNothing(Xdata) Then
                        rb.Add(New TypedValue(DxfCode.ExtendedDataRegAppName, appNameKonso))
                    End If
                    For Each acSSObj As SelectedObject In SelBlock
                        Dim blkRecIdBlock As ObjectId = acSSObj.ObjectId
                        Dim acBlkRef1 As BlockReference = DirectCast(acTrans.GetObject(blkRecIdBlock, OpenMode.ForWrite), BlockReference)
                        Dim attCol1 As AttributeCollection = acBlkRef.AttributeCollection
                        Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection
                        For Each objID As ObjectId In attCol
                            Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForWrite)
                            Dim acAttRef As AttributeReference = dbObj
                            If acAttRef.Tag = "ТАБЛО" Then
                                acAttRef.TextString = strTablo
                                rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, acSSObj.ObjectId.ToString))
                                Exit For
                            End If
                        Next
                    Next
                    If Not IsNothing(Xdata) Then
                        For Each tvXdata As TypedValue In Xdata
                            Dim yes As Boolean = True
                            If tvXdata.TypeCode = DxfCode.ExtendedDataRegAppName Then Continue For
                            For Each tvrb As TypedValue In rb
                                If tvXdata.Value = tvrb.Value Then
                                    yes = False
                                    Exit For
                                End If
                            Next
                            If yes Then rb.Add(New TypedValue(tvXdata.TypeCode, tvXdata.Value))
                        Next
                    End If
                    addXdataToSelectedObject(appNameKonso, rb, SelTablo(0))
                End Using
                acTrans.Commit()
            End Using
        Catch ex As Exception
            MsgBox("Възникна грешка:  " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
        End Try
    End Sub
    <CommandMethod("NewBlockTablo")>
    Public Sub NewBlockTablo()
        ' Get the current database and start a transaction
        Dim acCurDb As Autodesk.AutoCAD.DatabaseServices.Database
        acCurDb = Application.DocumentManager.MdiActiveDocument.Database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim SelBlock = cu.GetObjects("INSERT", "Изберете БЛОК")
        Dim SelTablo = cu.GetObjects("INSERT", "Изберете ТАБЛО")

        If SelBlock Is Nothing Then
            MsgBox("Въпрос с повишена трудност." + vbCrLf + "Аз какво да установявам???")
            Exit Sub
        End If
        If SelTablo Is Nothing Then
            MsgBox("НЕ Е маркиран поне един блок за табло")
            Exit Sub
        End If
        If SelTablo.Count > 1 Then
            MsgBox("НЕ Е маркиран САМО един блок за табло")
            Exit Sub
        End If
        Try
            ' Декларира променлива blkRecIdTablo, която съхранява ObjectId на избрания блок "Табло".
            Dim blkRecIdTablo As ObjectId = SelTablo(0).ObjectId
            ' Стартира транзакция за взаимодействие с AutoCAD обекта.
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                ' Декларира променлива strTablo за съхранение на стойността на атрибута с етикет "ТАБЛО".
                Dim strTablo As String = ""
                ' Цикъл през всеки избран обект в колекцията SelTablo.
                For Each sObj As SelectedObject In SelTablo
                    ' Взима ObjectId на текущия избран обект.
                    blkRecIdTablo = sObj.ObjectId
                    ' Взима референция към блока (BlockReference) на текущия обект, за да се работи с него.
                    Dim acBlkRef As BlockReference = DirectCast(acTrans.GetObject(blkRecIdTablo, OpenMode.ForWrite), BlockReference)
                    ' Взима колекцията с атрибути (AttributeCollection) на блока.
                    Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                    ' Цикъл през всеки обект в колекцията с атрибути.
                    For Each objID As ObjectId In attCol
                        ' Взима обект (DBObject) с текущия идентификатор (objID).
                        Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                        ' Преобразува обекта в AttributeReference, за да работи с атрибутите.
                        Dim acAttRef As AttributeReference = dbObj
                        ' Проверява дали тагът на атрибута е "ТАБЛО".
                        If acAttRef.Tag = "ТАБЛО" Then
                            ' Ако е, съхранява текста на атрибута в променливата strTablo.
                            strTablo = acAttRef.TextString
                        End If
                    Next
                Next
                ' Създава нов обект ResultBuffer, който ще съдържа допълнителни данни (Extended Data).
                Using rb As ResultBuffer = New ResultBuffer
                    ' Добавя регистрираното име на приложението (appNameKonso) към ResultBuffer.
                    rb.Add(New TypedValue(DxfCode.ExtendedDataRegAppName, appNameKonso))
                    ' Цикъл през всеки избран обект в колекцията SelBlock.
                    For Each acSSObj As SelectedObject In SelBlock
                        ' Взима ObjectId на текущия избран блок.
                        Dim blkRecIdBlock As ObjectId = acSSObj.ObjectId
                        ' Взима референция към блока (BlockReference) за работа с него.
                        Dim acBlkRef As BlockReference = DirectCast(acTrans.GetObject(blkRecIdBlock, OpenMode.ForWrite), BlockReference)
                        ' Взима колекцията с атрибути на блока.
                        Dim attCol As AttributeCollection = acBlkRef.AttributeCollection

                        ' Взима колекцията с динамични свойства на блока (ако е динамичен).
                        Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection
                        ' Цикъл през всеки атрибутен обект в колекцията с атрибути.
                        For Each objID As ObjectId In attCol
                            ' Взима обект (DBObject) за текущия идентификатор.
                            Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForWrite)
                            ' Преобразува обекта в AttributeReference за работа с атрибутите.
                            Dim acAttRef As AttributeReference = dbObj
                            ' Проверява дали тагът на атрибута е "ТАБЛО".
                            If acAttRef.Tag = "ТАБЛО" Then
                                ' Ако е, променя текста на атрибута, като го задава на strTablo.
                                acAttRef.TextString = strTablo
                                ' Добавя текущия обект (ObjectId) като разширени данни към ResultBuffer.
                                rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, acSSObj.ObjectId.ToString))
                                ' Излиза от цикъла, тъй като атрибутът "ТАБЛО" вече е намерен и актуализиран.
                                Exit For
                            End If
                        Next
                    Next
                    ' Добавя разширените данни към избрания обект в SelTablo(0) чрез функцията addXdataToSelectedObject.
                    addXdataToSelectedObject(appNameKonso, rb, SelTablo(0))
                End Using
                ' Потвърждава транзакцията, като записва всички промени.
                acTrans.Commit()
            End Using
        Catch ex As Exception
            MsgBox("Възникна грешка:  " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
        End Try
    End Sub
    <CommandMethod("getTablo")>
    Public Sub getTablo()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim edt As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database
        Dim SelBlock = cu.GetObjects("INSERT", "Изберете блок")
        Dim SelTablo = cu.GetObjects("INSERT", "Изберете ТАБЛО")

        If SelBlock Is Nothing Then
            MsgBox("Въпрос с повишена трудност." + vbCrLf + "Аз какво да установявам???")
            Exit Sub
        End If
        If SelTablo Is Nothing Then
            MsgBox("НЕ Е маркиран поне един блок за табло")
            Exit Sub
        End If
        If SelTablo.Count > 1 Then
            MsgBox("Маркирай САМО един блок за табло")
            Exit Sub
        End If

        Dim blkRecId As ObjectId = SelTablo(0).ObjectId
        Dim strTablo As String = ""
        Using trans As Transaction = acCurDb.TransactionManager.StartTransaction()

            Dim acBlkRef As BlockReference = DirectCast(trans.GetObject(blkRecId, OpenMode.ForRead), BlockReference)
            Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
            For Each objID As ObjectId In attCol
                Dim dbObj As DBObject = trans.GetObject(objID, OpenMode.ForRead)
                Dim acAttRef As AttributeReference = dbObj
                If acAttRef.Tag = "ТАБЛО" Then strTablo = acAttRef.TextString
            Next

            Dim blockRef As BlockReference = TryCast(trans.GetObject(blkRecId, OpenMode.ForRead), BlockReference)
            If Not blockRef.IsDynamicBlock Then Return
            Dim blockDef As BlockTableRecord = TryCast(trans.GetObject(blockRef.DynamicBlockTableRecord, OpenMode.ForRead), BlockTableRecord)
            If blockDef.ExtensionDictionary.IsNull Then Return
            Dim extDic As DBDictionary = TryCast(trans.GetObject(blockDef.ExtensionDictionary, OpenMode.ForRead), DBDictionary)
            Dim graph As EvalGraph = TryCast(trans.GetObject(extDic.GetAt("ACAD_ENHANCEDBLOCK"), OpenMode.ForRead), EvalGraph)
            Dim nodeIds As Integer() = graph.GetAllNodes()

            For Each sObj As SelectedObject In SelBlock
                blkRecId = sObj.ObjectId
                For Each nodeId As UInteger In nodeIds
                    Dim node As DBObject = graph.GetNode(nodeId, OpenMode.ForWrite, trans)
                    If Not (TypeOf node Is BlockPropertiesTable) Then Continue For
                    Dim table As BlockPropertiesTable = TryCast(node, BlockPropertiesTable)
                    Dim columns As Integer = table.Columns.Count
                    Dim currentRow As Integer = 0
                    For Each row As BlockPropertiesTableRow In table.Rows
                        edt.WriteMessage(vbLf & "[{0}]:" & vbTab, currentRow)
                        For currentColumn As Integer = 0 To columns - 1
                            Dim columnValue As TypedValue() = row(currentColumn).AsArray()
                            For Each tpVal As TypedValue In columnValue
                                edt.WriteMessage("{0}; ", tpVal.Value)
                            Next
                            edt.WriteMessage("|")
                        Next
                        currentRow += 1
                    Next
                Next
            Next
        End Using
    End Sub
    Private Function calc_breaker(Ikryg As Double' Върхов тok на токовия кръг
                                 ) As String            ' Избира автоматичен прекъсвач


        Ikryg = Knti * Ikryg

        ' Избира автоматичен прекъсвач на базата на зададения ток
        For Each breaker In Breakers
            If Ikryg <= breaker.Key Then
                Return breaker.Value & ", " & breaker.Key & "A"
            End If
        Next

        Return "НЕ Е подходящ прекъсвач"
    End Function
    Private Function calc_disconnector(Ikryg As Double) As String
        ' Речник за мощностни разединители на Schneider Electric
        ' Речник за всички товарови разединители


        Ikryg = Knti * Ikryg

        ' Избира мощностен разединител на базата на зададения ток
        For Each disconnector In Disconnectors
            If Ikryg <= disconnector.Key Then
                Return disconnector.Value & ", " & disconnector.Key & "A"
            End If
        Next
        Return "НЕ Е подходящ разединител"
    End Function
    ' Функцията избира подходящ автоматичен прекъсвач EZ9
    ' на базата на върховия ток на токовия кръг (Ikryg), умножен по корекционен коефициент Knti
    ' Връща стойност като текст (например "32") или "######", ако токът е извън стандартните прагове
    Private Function calc_breaker_EZ9(Ikryg As Double) As String
        ' Прилагаме корекционния коефициент върху подадения ток
        Ikryg = Knti * Ikryg
        ' Масив със стандартните токови стойности за серия EZ9
        Dim thresholds() As Double = {6, 10, 16, 20, 25, 32, 40, 50, 63, 80, 100, 125}
        ' Обхождаме праговете един по един
        For Each t In thresholds
            ' Ако токът е по-малък от дадения праг,
            ' това е първият прекъсвач, който може да поеме натоварването
            If Ikryg < t Then Return t.ToString()
        Next
        ' Ако токът е по-голям от всички прагове,
        ' връщаме "######", което показва, че няма подходящ прекъсвач
        Return "######"
    End Function
    Private Function calc_Inom(Pkryg As Double,                     ' мощност на токовия кръг
                           NumberPoles As String,                   ' брой фази на токовия кръг
                           Optional Motor As Boolean = False        ' Ако е двигател True - КПД и cos FI да са по 0,83
                           ) As Double                              ' Изчислява номинален ток за товар
        Dim CosFI As Double                                         ' Декларира променлива за cos φ (фактор на мощността)
        Dim KPD As Double                                           ' Декларира променлива за КПД (коефициент на полезно действие)
        Const U380 As Double = 0.38                                 ' Дефинира константа за напрежение при 380V, преобразувано в kV (киловолти)
        Const U220 As Double = 0.22                                 ' Дефинира константа за напрежение при 220V, преобразувано в kV (киловолти)
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
    Private Function calc_ISW(NumberPoles As String,                ' брой полюси на токовия кръг (въведен като низ)
                          Optional Pkryg As Double = 0,         ' Опционален параметър: мощността, предавана на токовия кръг (по подразбиране 0)
                          Optional Motor As Boolean = False,    ' Опционален параметър: дали е двигател (по подразбиране False); ако е True, КПД и cos ФИ са по 0,83
                          Optional Ikryg As Double = 0          ' Опционален параметър: токът, предаван на токовия кръг (по подразбиране 0)
                          ) As String                           ' Функцията връща като резултат низ, който представлява избрания мощностен разединител iSW
        Dim calc As String = ""                                     ' Променлива за съхранение на резултата от функцията
        Dim Inom As Double = 0                                      ' Номинален ток на токовия кръг

        ' Ако е зададена мощността (Pkryg > 0), изчислява номиналния ток с помощта на функцията calc_Inom
        If Pkryg > 0 Then
            Inom = calc_Inom(Pkryg, NumberPoles, Motor)             ' Изчислява номиналния ток според мощността, броя полюси и дали е двигател
        Else
            Inom = Ikryg                                            ' Ако не е зададена мощност, използва директно предадения ток Ikryg
        End If

        Ikryg = Inom * Knti                                         ' Изчислява коригирания ток, като умножава номиналния ток с коефициента Knti (не е дефиниран в кода)

        ' Избира мощностен разединител iSW в зависимост от стойността на коригирания ток Ikryg
        Select Case Ikryg
            Case < 20
                calc = "20"                                         ' Ако Ikryg е по-малко от 20, задава мощностен разединител 20
            Case < 32
                calc = "32"                                         ' Ако Ikryg е по-малко от 32, задава мощностен разединител 32
            Case < 40
                calc = "40"                                         ' Ако Ikryg е по-малко от 40, задава мощностен разединител 40
            Case < 63
                calc = "63"                                         ' Ако Ikryg е по-малко от 63, задава мощностен разединител 63
            Case < 100
                calc = "100"                                        ' Ако Ikryg е по-малко от 100, задава мощностен разединител 100
            Case < 125
                calc = "125"                                        ' Ако Ikryg е по-малко от 125, задава мощностен разединител 125
            Case Else
                calc = "######"                                     ' Ако Ikryg е по-голямо от всички изброени стойности, задава недефинирана стойност "######"
        End Select

        Return calc                                                 ' Връща резултата (избрания мощностен разединител) като низ
    End Function
    Private Function calc_INS(NumberPoles As String,                ' брой фази на токовия кръг
                              Optional Pkryg As Double = 0,         ' Ако се предава мощност на токовия кръг
                              Optional Motor As Boolean = False,    ' Ако е двигател True - КПД и cos FI да са по 0,83
                              Optional Ikryg As Double = 0          ' Ако се предава тока
                              ) As String                           ' Избира мощностен разединител iSW
        Dim calc As String = ""
        Dim Inom As Double = 0
        If Pkryg > 0 Then
            Inom = calc_Inom(Pkryg, NumberPoles, Motor)
        Else
            Inom = Ikryg
        End If

        Ikryg = Inom * Knti
        Select Case Ikryg
            Case < 40
                calc = "40"
            Case < 63
                calc = "63"
            Case < 80
                calc = "63"
            Case < 100
                calc = "100"
            Case < 125
                calc = "125"
            Case < 160
                calc = "160"
            Case < 200
                calc = "200"
            Case < 250
                calc = "250"
            Case < 320
                calc = "320"
            Case < 400
                calc = "400"
            Case < 500
                calc = "500"
            Case < 630
                calc = "630"
            Case < 800
                calc = "800"
            Case < 1000
                calc = "1000"
            Case < 1250
                calc = "1250"
            Case < 1600
                calc = "1600"
            Case < 2500
                calc = "2500"
            Case Else
                calc = "######"
        End Select
        Return calc
    End Function
    Private Function calc_IN(NumberPoles As String,                 ' брой фази на токовия кръг
                              Optional Pkryg As Double = 0,         ' Ако се предава мощност на токовия кръг
                              Optional Motor As Boolean = False,    ' Ако е двигател True - КПД и cos FI да са по 0,83
                              Optional Ikryg As Double = 0          ' Ако се предава тока
                              ) As String                           ' Избира мощностен разединител iSW
        Dim calc As String = ""
        Dim Inom As Double = 0
        If Pkryg > 0 Then
            Inom = calc_Inom(Pkryg, NumberPoles, Motor)
        Else
            Inom = Ikryg
        End If

        Ikryg = Inom * Knti
        Select Case Ikryg
            Case < 40
                calc = "40"
            Case < 63
                calc = "63"
            Case < 80
                calc = "63"
            Case < 100
                calc = "100"
            Case < 125
                calc = "125"
            Case < 160
                calc = "160"
            Case < 200
                calc = "200"
            Case < 250
                calc = "250"
            Case < 320
                calc = "320"
            Case < 400
                calc = "400"
            Case < 500
                calc = "500"
            Case < 630
                calc = "630"
            Case < 800
                calc = "800"
            Case < 1000
                calc = "1000"
            Case < 1250
                calc = "1250"
            Case < 1600
                calc = "1600"
            Case < 2500
                calc = "2500"
            Case Else
                calc = "######"
        End Select
        Return calc
    End Function
    Private Function calc_cable_Cu(Ibreaker As String,                  ' Ток на Автоматичния прекъсвач
                                   NumberPoles As String,               ' Брой на фазите
                                   Optional Polag As Integer = 0,       ' Начин на полагане 0 - във въздух; 1 - в земя
                                   Optional Broj_kab As Integer = 1,    ' Брой кабели положени паралелно на скара
                                   Optional Kabel As Integer = 0        ' 0 - Кабел; 1 - проводник
                                   ) As String                          ' Избира сечение на меден кабел


        Dim calc As String = ""
        Dim Inom As Double = Val(Ibreaker)
        Dim Idop As Double = 0
        Dim Kz As Double = 1
        Dim Q As Double = 70
        Dim Qokdef As Double = 30
        Dim Qok As Double
        Dim K2 As Double = 0
        Dim K1 As Double = 0


        If Polag = 0 Then
            Qok = 35
            K2 = Math.Sqrt((Q - Qok) / (Q - Qokdef))
        Else
            K2 = 1
        End If

        K1 = GetK1(Broj_kab, Polag)
        Idop = Inom / (K1 * K2)

        Select Case Kabel                       ' според това дали е кабел или проводник
            Case 0                              ' кабел - три/ пет жила в обща обвивка
                Select Case Polag               ' според начина на полагане във въздух или в земя
                    Case 0                      ' положени във въздух
                        Select Case Idop
                            Case < 19
                                calc = "1,5"
                            Case < 25
                                calc = "2,5"
                            Case < 34
                                calc = "4,0"
                            Case < 43
                                calc = "6,0"
                            Case < 59
                                calc = "10,0"
                            Case < 79
                                calc = "16,0"
                            Case < 105
                                calc = "25,0"
                            Case < 126
                                calc = "35,0"
                            Case < 157
                                calc = "50,0"
                            Case < 199
                                calc = "70,0"
                            Case < 246
                                calc = "95,0"
                            Case < 285
                                calc = "120,0"
                            Case < 326
                                calc = "150,0"
                            Case < 374
                                calc = "180,0"
                            Case < 445
                                calc = "240,0"
                            Case Else
                                calc = "######"
                        End Select
                    Case 1                      ' положени в земя
                        Select Case Idop
                            Case < 25
                                calc = "1,5"
                            Case < 34
                                calc = "2,5"
                            Case < 45
                                calc = "4,0"
                            Case < 55
                                calc = "6,0"
                            Case < 76
                                calc = "10,0"
                            Case < 96
                                calc = "16,0"
                            Case < 126
                                calc = "25,0"
                            Case < 151
                                calc = "35,0"
                            Case < 178
                                calc = "50,0"
                            Case < 225
                                calc = "70,0"
                            Case < 270
                                calc = "95,0"
                            Case < 306
                                calc = "120,0"
                            Case < 346
                                calc = "150,0"
                            Case < 390
                                calc = "180,0"
                            Case < 458
                                calc = "240,0"
                            Case Else
                                calc = "######"
                        End Select
                End Select
                Return IIf(NumberPoles = "3p", "5x", "3x") + calc
            Case 1                             ' проводник - едно жило
                Select Case Polag               ' според начина на полагане във въздух или в земя
                    Case 0                      ' положени във въздух
                        Select Case Idop
                            Case < 20
                                calc = "1,5"
                            Case < 27
                                calc = "2,5"
                            Case < 36
                                calc = "4,0"
                            Case < 45
                                calc = "6,0"
                            Case < 63
                                calc = "10,0"
                            Case < 82
                                calc = "16,0"
                            Case < 113
                                calc = "25,0"
                            Case < 138
                                calc = "35,0"
                            Case < 168
                                calc = "50,0"
                            Case < 210
                                calc = "70,0"
                            Case < 262
                                calc = "95,0"
                            Case < 307
                                calc = "120,0"
                            Case < 352
                                calc = "150,0"
                            Case < 405
                                calc = "180,0"
                            Case < 482
                                calc = "240,0"
                            Case < 555
                                calc = "300,0"
                            Case < 650
                                calc = "400,0"
                            Case < 500
                                calc = "750,0"
                            Case Else
                                calc = "######"
                        End Select
                    Case 1                      ' положени в земя
                        Select Case Idop
                            Case < 25
                                calc = "1,5"
                            Case < 34
                                calc = "2,5"
                            Case < 45
                                calc = "4,0"
                            Case < 55
                                calc = "6,0"
                            Case < 76
                                calc = "10,0"
                            Case < 96
                                calc = "16,0"
                            Case < 126
                                calc = "25,0"
                            Case < 151
                                calc = "35,0"
                            Case < 178
                                calc = "50,0"
                            Case < 225
                                calc = "70,0"
                            Case < 270
                                calc = "95,0"
                            Case < 306
                                calc = "120,0"
                            Case < 346
                                calc = "150,0"
                            Case < 390
                                calc = "180,0"
                            Case < 458
                                calc = "240,0"
                            Case Else
                                calc = "######"
                        End Select
                End Select
                Return IIf(NumberPoles = "3p", "4x1х", "2x1х") + calc
        End Select

    End Function
    ' Връща коефициент в зависимост от броя паралелно положени кабели
    Function GetK1(Broj_kab As Integer, Polag As Integer) As Double
        Dim K1_Air() As Double = {1, 0.75, 0.65, 0.55, 0.5, 0.45, 0.4, 0.35, 0.3}
        Dim K1_Ground() As Double = {1, 0.7, 0.6, 0.5, 0.45, 0.4, 0.35, 0.3, 0.25}
        Dim index As Integer

        Select Case Broj_kab
            Case < 5
                index = 0
            Case 5 To 6
                index = 1
            Case 7 To 9
                index = 2
            Case 10 To 13
                index = 3
            Case 14 To 18
                index = 4
            Case 19 To 23
                index = 5
            Case 24
                index = 6
            Case 25 To 40
                index = 7
            Case 41 To 61
                index = 8
        End Select

        If Polag = 0 Then ' Въздух
            Return K1_Air(index)
        Else ' Земя
            Return K1_Ground(index)
        End If
    End Function
    Private Function calc_cable_Al(Ibreaker As String,                  ' Ток на Автоматичния прекъсвач
                                   NumberPoles As String,               ' Брой на фазите
                                   Optional Polag As Integer = 0,       ' Начин на полагане 0 - във въздух; 1 - в земя
                                   Optional Broj_kab As Integer = 1,    ' Брой кабели положени паралелно на скара
                                   Optional Kabel As Integer = 0        ' 0 - Кабел; 1 - проводник
                                   ) As String                          ' Избира сечение на алуминиев кабел

        Dim calc As String = ""
        Dim Inom As Double = Val(Ibreaker)
        Dim Idop As Double = 0
        Dim Kz As Double = 1
        Dim Q As Double = 70
        Dim Qokdef As Double = 25
        Dim Qok As Double = IIf(Polag = 0, 35, 15)
        Dim K2 As Double = Math.Sqrt((Q - Qok) / (Q - Qokdef))
        Dim K1 As Double = 0

        Select Case Polag                       ' Според броя на паралелно положени кабели
            Case 0                              ' положени във въздух
                Select Case Broj_kab
                    Case < 5
                        K1 = 1
                    Case 5, 6
                        K1 = 0.75
                    Case 7, 8, 9
                        K1 = 0.65
                    Case 10, 11, 12, 13
                        K1 = 0.55
                    Case 14, 15, 16, 17, 18
                        K1 = 0.5
                    Case 19, 20, 21, 22, 23
                        K1 = 0.45
                    Case 24
                        K1 = 0.4
                    Case < 41
                        K1 = 0.35
                    Case < 62
                        K1 = 0.3
                End Select
            Case 1                              ' положени в земя
                Select Case Broj_kab
                    Case < 5
                        K1 = 1
                    Case 5, 6
                        K1 = 0.7
                    Case 7, 8, 9
                        K1 = 0.6
                    Case 10, 11, 12, 13
                        K1 = 0.5
                    Case 14, 15, 16, 17, 18
                        K1 = 0.45
                    Case 19, 20, 21, 22, 23
                        K1 = 0.4
                    Case 24
                        K1 = 0.35
                    Case < 41
                        K1 = 0.3
                    Case < 62
                        K1 = 0.25
                End Select
        End Select
        Idop = Inom / (K1 * K2)

        Select Case Kabel                       ' според това дали е кабел или проводник
            Case 0                              ' кабел - три/ пет жила в обща обвивка
                Select Case Polag               ' според начина на полагане във въздух или в земя
                    Case 0                      ' положени във въздух
                        Select Case Idop
                            Case < 19
                                calc = "1,5"
                            Case < 25
                                calc = "2,5"
                            Case < 34
                                calc = "4,0"
                            Case < 43
                                calc = "6,0"
                            Case < 59
                                calc = "10,0"
                            Case < 79
                                calc = "16,0"
                            Case < 105
                                calc = "25,0"
                            Case < 126
                                calc = "35,0"
                            Case < 157
                                calc = "50,0"
                            Case < 199
                                calc = "70,0"
                            Case < 246
                                calc = "95,0"
                            Case < 285
                                calc = "120,0"
                            Case < 326
                                calc = "150,0"
                            Case < 374
                                calc = "180,0"
                            Case < 445
                                calc = "240,0"
                            Case Else
                                calc = "######"
                        End Select
                    Case 1                      ' положени в земя
                        Select Case Idop
                            Case < 25
                                calc = "1,5"
                            Case < 34
                                calc = "2,5"
                            Case < 45
                                calc = "4,0"
                            Case < 55
                                calc = "6,0"
                            Case < 76
                                calc = "10,0"
                            Case < 96
                                calc = "16,0"
                            Case < 126
                                calc = "25,0"
                            Case < 151
                                calc = "35,0"
                            Case < 178
                                calc = "50,0"
                            Case < 225
                                calc = "70,0"
                            Case < 270
                                calc = "95,0"
                            Case < 306
                                calc = "120,0"
                            Case < 346
                                calc = "150,0"
                            Case < 390
                                calc = "180,0"
                            Case < 458
                                calc = "240,0"
                            Case Else
                                calc = "######"
                        End Select
                End Select
                Return IIf(NumberPoles = "3p", "5x", "3x") + calc
            Case 1                              ' проводник - едно жило
                Select Case Polag               ' според начина на полагане във въздух или в земя
                    Case 0                      ' положени във въздух
                        Select Case Idop
                            Case < 20
                                calc = "1,5"
                            Case < 27
                                calc = "2,5"
                            Case < 36
                                calc = "4,0"
                            Case < 45
                                calc = "6,0"
                            Case < 63
                                calc = "10,0"
                            Case < 82
                                calc = "16,0"
                            Case < 113
                                calc = "25,0"
                            Case < 138
                                calc = "35,0"
                            Case < 168
                                calc = "50,0"
                            Case < 210
                                calc = "70,0"
                            Case < 262
                                calc = "95,0"
                            Case < 307
                                calc = "120,0"
                            Case < 352
                                calc = "150,0"
                            Case < 405
                                calc = "180,0"
                            Case < 482
                                calc = "240,0"
                            Case < 555
                                calc = "300,0"
                            Case < 650
                                calc = "400,0"
                            Case < 500
                                calc = "750,0"
                            Case Else
                                calc = "######"
                        End Select
                    Case 1                      ' положени в земя
                        Select Case Idop
                            Case < 25
                                calc = "1,5"
                            Case < 34
                                calc = "2,5"
                            Case < 45
                                calc = "4,0"
                            Case < 55
                                calc = "6,0"
                            Case < 76
                                calc = "10,0"
                            Case < 96
                                calc = "16,0"
                            Case < 126
                                calc = "25,0"
                            Case < 151
                                calc = "35,0"
                            Case < 178
                                calc = "50,0"
                            Case < 225
                                calc = "70,0"
                            Case < 270
                                calc = "95,0"
                            Case < 306
                                calc = "120,0"
                            Case < 346
                                calc = "150,0"
                            Case < 390
                                calc = "180,0"
                            Case < 458
                                calc = "240,0"
                            Case Else
                                calc = "######"
                        End Select
                End Select
                Return IIf(NumberPoles = "3p", "4x1х", "2x1х") + calc
        End Select

    End Function
    Private Function calc_cable_AlR(Ibreaker As String,     ' Ток на Автоматичния прекъсвач
                                   NumberPoles As String    ' Брой на фазите
                                   ) As String              ' Избира сечение на алуминиев кабел тип Al/R

        Dim Polag As Integer = 0        ' Начин на полагане 0 - във въздух; 1 - в земя
        Dim Broj_kab As Integer = 1     ' Брой кабели положени паралелно на скара
        Dim Kabel As Integer = 0        ' 0 - Кабел; 1 - проводник
        Dim calc As String = ""
        Dim Inom As Double = Val(Ibreaker)
        Dim Idop As Double = 0
        Dim Kz As Double = 1
        Dim Q As Double = 70
        Dim Qokdef As Double = 25
        Dim Qok As Double = IIf(Polag = 0, 35, 15)
        Dim K2 As Double = Math.Sqrt((Q - Qok) / (Q - Qokdef))
        Dim K1 As Double = 1

        Idop = Inom / (K1 * K2)

        Select Case Idop
            Case < 83
                Return IIf(NumberPoles = "3p", "4x16,0", "2x16,0")
            Case < 111
                Return IIf(NumberPoles = "3p", "4x25,0", "2x25,0")
            Case < 138
                Return IIf(NumberPoles = "3p", "4x35,0", "----")
            Case < 164
                Return IIf(NumberPoles = "3p", "3x50+54", "----")
            Case < 213
                Return IIf(NumberPoles = "3p", "3x70+54", "----")
            Case < 258
                Return IIf(NumberPoles = "3p", "3x95+70", "----")
            Case < 320
                Return IIf(NumberPoles = "3p", "3x120+95", "----")
            Case < 344
                Return IIf(NumberPoles = "3p", "3x150+70", "----")
            Case Else
                Return "#####"
        End Select
    End Function
    Public Sub addXdataToSelectedObject(appName As String,          ' Име на таблицата
                                        Xdata As ResultBuffer,      ' Информация която се записва
                                        SelSet As SelectedObject    ' Блок в който се записва информацията                                        
                                        )
        ' Get the current database and start a transaction
        Dim acCurDb As Autodesk.AutoCAD.DatabaseServices.Database
        acCurDb = Application.DocumentManager.MdiActiveDocument.Database
        Try
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim acRegAppTbl As RegAppTable
                acRegAppTbl = acTrans.GetObject(acCurDb.RegAppTableId, OpenMode.ForRead)
                Dim acRegAppTblRec As RegAppTableRecord
                If acRegAppTbl.Has(appName) = False Then
                    acRegAppTblRec = New RegAppTableRecord
                    acRegAppTblRec.Name = appName
                    acTrans.GetObject(acCurDb.RegAppTableId, OpenMode.ForWrite)
                    acRegAppTbl.Add(acRegAppTblRec)
                    acTrans.AddNewlyCreatedDBObject(acRegAppTblRec, True)
                End If
                Dim acEnt As Entity = acTrans.GetObject(SelSet.ObjectId, OpenMode.ForWrite)
                acEnt.XData = Xdata
                acTrans.Commit()
            End Using
        Catch ex As Exception
            MsgBox("Възникна грешка:  " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
        End Try
    End Sub
    Private Function insSplitContainer(name As String, colArray As strTablo) As Windows.Forms.SplitContainer
        Dim SplitContainer As System.Windows.Forms.SplitContainer = New Windows.Forms.SplitContainer
        With SplitContainer
            .Name = name & "_SC_GB_DG"
            .Orientation = Windows.Forms.Orientation.Horizontal
            .Size = New System.Drawing.Size(475, 350)
            .Dock = Windows.Forms.DockStyle.Fill
            .SplitterDistance = 100
            .Panel1.Controls.Add(insGroupBox_BT(name))
            .Panel1.Controls.Add(insGroupBox_Tablo(name))
            .Panel2.Controls.Add(insGroupBox_DG(name, colArray))
        End With
        Return SplitContainer
    End Function
    Private Function insGroupBox_Tablo(name As String) As Windows.Forms.GroupBox
        Dim GroupBox As System.Windows.Forms.GroupBox = New Windows.Forms.GroupBox
        With GroupBox
            .Name = name & "_BT"
            .Size = New System.Drawing.Size(432, 250)
            .Location = New System.Drawing.Point(6, 6)
            .Font = New Drawing.Font("Arial", 12, Drawing.FontStyle.Bold)
            .Text = "Tабло '" & name & "'"
            '.Dock = Windows.Forms.DockStyle.Fill
            .BackColor = System.Drawing.SystemColors.Control
            .ForeColor = System.Drawing.SystemColors.WindowText
        End With

        With GroupBox.Controls
            .Add(insLabel(name & "/#/" & "1", "Вид на табло", 25, 6, 150))
            .Add(insLabel(name & "/#/" & "2", "Начин на монтаж", 50, 6, 150))
            .Add(insLabel(name & "/#/" & "3", "Размери", 75, 6, 150))
            .Add(insLabel(name & "/#/" & "4", "Захранващ кабел", 100, 6, 150))


            Dim ComboBoxItems(5) As String
            ComboBoxItems(0) = "Mini Pragma"
            ComboBoxItems(1) = "Pragma"
            ComboBoxItems(2) = "Kiedra"
            ComboBoxItems(3) = "Метално"
            ComboBoxItems(4) = "Метално"
            ComboBoxItems(5) = "Метално"

            Dim ComboBoxItems1(5) As String
            ComboBoxItems1(0) = "Изпъкнал монтаж"
            ComboBoxItems1(1) = "Вграден монтаж"
            ComboBoxItems1(2) = "На фундамент"
            ComboBoxItems1(3) = "Метално"
            ComboBoxItems1(4) = "Метално"
            ComboBoxItems1(5) = "Метално"

            Dim ComboBoxItems2(5) As String
            ComboBoxItems2(0) = "---"
            ComboBoxItems2(1) = "---"
            ComboBoxItems2(2) = "---"
            ComboBoxItems2(3) = "---"
            ComboBoxItems2(4) = "---"
            ComboBoxItems2(5) = "---"

            Dim ComboBoxItems3(5) As String
            ComboBoxItems3(0) = "СВТ"
            ComboBoxItems3(1) = "САВТ"
            ComboBoxItems3(2) = "Al/R"
            ComboBoxItems3(3) = "---"
            ComboBoxItems3(4) = "---"
            ComboBoxItems3(5) = "ПВВМ-Б1"

            Dim ComboBoxItems4(3) As String
            ComboBoxItems4(0) = "2"
            ComboBoxItems4(1) = "3"
            ComboBoxItems4(2) = "4"
            ComboBoxItems4(3) = "5"

            Dim ComboBoxItems5(11) As String
            ComboBoxItems5(0) = "4"
            ComboBoxItems5(1) = "6"
            ComboBoxItems5(2) = "10"
            ComboBoxItems5(3) = "16"
            ComboBoxItems5(4) = "25"
            ComboBoxItems5(5) = "35"
            ComboBoxItems5(6) = "50"
            ComboBoxItems5(7) = "10"
            ComboBoxItems5(8) = "10"
            ComboBoxItems5(9) = "10"
            ComboBoxItems5(10) = "10"
            ComboBoxItems5(11) = "185+35"

            .Add(insComboBox(name & "/#/" & "1", "Вид на ТАБЛОТО", 25, 160, ComboBoxItems, 150))
            .Add(insComboBox(name & "/#/" & "2", "Начин на монтаж", 50, 160, ComboBoxItems1, 150))
            .Add(insComboBox(name & "/#/" & "3", "Размери", 75, 160, ComboBoxItems2, 150))
            .Add(insComboBox(name & "/#/" & "4", "Tип", 100, 160, ComboBoxItems3, 75))
            .Add(insComboBox(name & "/#/" & "5", "бр", 100, 240, ComboBoxItems4, 40))
            .Add(insComboBox(name & "/#/" & "6", "сечение", 100, 283, ComboBoxItems5, 78))

        End With

        Return GroupBox
    End Function
    Private Function insGroupBox_DG(name As String, colArray As strTablo) As Windows.Forms.GroupBox
        Dim GroupBox As System.Windows.Forms.GroupBox = New Windows.Forms.GroupBox
        With GroupBox
            .Name = name & "_GB_DG"
            .Size = New System.Drawing.Size(432, 570)
            .Text = "Параметри на токовите кръгове на табло '" & name & "'"
            .Font = New Drawing.Font("Arial", 12, Drawing.FontStyle.Bold)
            .Dock = Windows.Forms.DockStyle.Fill
            .BackColor = System.Drawing.SystemColors.Control
            .ForeColor = System.Drawing.SystemColors.WindowText
            .Controls.Add(insDataGrid(name, colArray))
        End With
        Return GroupBox
    End Function
    Private Function insButtonn(nameButton As String, textButton As String, top As Integer, left As Integer) As System.Windows.Forms.Button
        Dim button As New System.Windows.Forms.Button
        With button
            .Name = nameButton
            .Size = New System.Drawing.Size(160, 25)
            .Text = textButton
            .Top = top
            .Font = New Drawing.Font("Arial", 12, Drawing.FontStyle.Regular)
            .Left = left
        End With
        AddHandler button.Click, AddressOf btnExport_Block_Click
        Return button
    End Function
    Private Function insLabel(nameLabel As String,
                              textLabel As String,
                              top As Integer,
                              left As Integer,
                              size As Integer
                              ) As System.Windows.Forms.Label
        Dim Label As New System.Windows.Forms.Label
        With Label
            .Name = nameLabel
            .Size = New System.Drawing.Size(size, 25)
            .Font = New Drawing.Font("Arial", 12, Drawing.FontStyle.Regular)
            .Text = textLabel
            .Top = top
            .Left = left
        End With
        Return Label
    End Function
    Private Function insTextBox(nameText As String, textText As String, top As Integer, left As Integer) As System.Windows.Forms.TextBox
        Dim TextBox As New System.Windows.Forms.TextBox
        With TextBox
            .Name = nameText
            .Size = New System.Drawing.Size(100, 25)
            .Text = textText
            .Top = top
            .Left = left
            .Font = New Drawing.Font("Arial", 10, Drawing.FontStyle.Regular)
        End With
        Return TextBox
    End Function
    Private Function insComboBox(nameComboBox As String,
                                 texComboBox As String,
                                 top As Integer,
                                 left As Integer,
                                 items() As String,
                                 SizeCombo As Integer
                                 ) As System.Windows.Forms.ComboBox
        Dim ComboBox As New System.Windows.Forms.ComboBox
        With ComboBox
            .Name = nameComboBox
            .Size = New System.Drawing.Size(SizeCombo, 25)
            .Text = texComboBox
            .Top = top
            .Left = left
            .Font = New Drawing.Font("Arial", 12, Drawing.FontStyle.Regular)
        End With
        For Each index In items
            ComboBox.Items.Add(index)
        Next
        AddHandler ComboBox.SelectedValueChanged, AddressOf comboSelectedValueChanged
        Return ComboBox
    End Function
    Private Sub comboSelectedValueChanged(sender As Object, e As EventArgs)
        Dim send As System.Windows.Forms.ComboBox = sender
        Dim name As String = ""
        Dim nam As String = ""
        Dim func As String = ""
        name = Mid(send.Name.ToString, 1, InStr(send.Name.ToString, "/#/") - 1)
        func = Mid(send.Name.ToString, InStr(send.Name.ToString, "/#/") + 3, Len(send.Name.ToString))

        '  MsgBox(send.Name.ToString & " NAME-> '" & name & "' FUNK-> " & func)

        '
        'Качва по стълбицата нагоре за да намери SplitContainer в който е DataGridView
        '
        Dim panel As System.Windows.Forms.SplitContainer = DirectCast(send.Parent.Parent.Parent, System.Windows.Forms.SplitContainer)
        Dim obList As List(Of Object) = New List(Of Object)
        Call FindChildren(panel, obList)
        Dim index As Integer
        For Each ctrl As Windows.Forms.Control In obList
            If TypeOf ctrl Is Windows.Forms.DataGridView Then
                Exit For
            End If
            index += 1
        Next
        Dim dagrid As System.Windows.Forms.DataGridView = DirectCast(obList(index), System.Windows.Forms.DataGridView)
    End Sub
    Private Function insCheckBox(nameCheckBox As String,
                                 texCheckBox As String,
                                 top As Integer,
                                 Left As Integer,
                                 Checked As Boolean
                                 ) As System.Windows.Forms.CheckBox
        Dim CheckBox As New System.Windows.Forms.CheckBox
        With CheckBox
            .Name = nameCheckBox
            .Size = New System.Drawing.Size(100, 25)
            .Text = texCheckBox
            .Top = top
            .Left = Left
            .Checked = Checked
        End With
        Return CheckBox
    End Function
    Private Function insDataGrid(name As String, colArray As strTablo) As Windows.Forms.DataGridView
        Dim dagrid As System.Windows.Forms.DataGridView = New Windows.Forms.DataGridView
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        With DataGridViewCellStyle1
            .Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
            .BackColor = System.Drawing.SystemColors.ControlDark
            .Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
            .ForeColor = System.Drawing.SystemColors.WindowText
            .SelectionBackColor = System.Drawing.SystemColors.Highlight
            .SelectionForeColor = System.Drawing.SystemColors.HighlightText
            .WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        End With
        With DataGridViewCellStyle2
            .BackColor = System.Drawing.Color.Silver
            .ForeColor = System.Drawing.Color.Black
            .Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
            .Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
            .SelectionBackColor = System.Drawing.SystemColors.Highlight
            .SelectionForeColor = System.Drawing.SystemColors.HighlightText
            .WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        End With
        With DataGridViewCellStyle3
            .Format = "N2"
            .NullValue = Nothing
        End With
        With dagrid
            .Name = name & "_DG"
            .BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
            .ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
            .RowHeadersDefaultCellStyle = DataGridViewCellStyle2
            .Size = New System.Drawing.Size(432, 450)
            .Dock = System.Windows.Forms.DockStyle.Fill
            .ColumnCount = 2 + colArray.countTokKryg + 1
            .ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
            With .Columns(0)
                .Width = 150
                .HeaderText = "Параметър"
                .Name = "Параметър"
                .Frozen = vbTrue
                .ReadOnly = vbTrue
                .SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
                .DefaultCellStyle = DataGridViewCellStyle3
            End With
            With .Columns(1)
                .DefaultCellStyle = DataGridViewCellStyle3
                .Width = 40
                .HeaderText = ""
                .Name = "Дим."
                .Frozen = vbTrue
                .ReadOnly = vbTrue
                .SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
            End With
            '
            ' Добавя редове
            ' Записва надписи в първа и втора колона 
            '
            With .Rows
                .Add({"Прекъсвач", " "})
                .Add({"Изчислен ток", "А"})
                .Add({"Тип на апарата", " "})
                .Add({"Номинален ток", "А"})
                .Add({"Изкл. възможност", " "})
                .Add({"Крива", " "})
                .Add({"Брой полюси", "бр."})
                .Add({"-----------", "---"})
                .Add({"ДТЗ", " "})
                .Add({"Вид на апарата", " "})
                .Add({"Клас на апарата", " "})
                .Add({"Номинален ток", "А"})
                .Add({"Изкл. възможност", "mA"})
                .Add({"Брой полюси", "бр."})
                .Add({"-----------", "---"})
                .Add({"Брой лампи", "бр."})
                .Add({"Брой контакти", "бр."})
                .Add({"Инст. мощност", "kW"})
                .Add({"Тип кабел", "---"})
                .Add({"Сечение", "---"})
                .Add({"Фаза", "---"})
                .Add({"Консуматор", "---"})
                .Add({"", "---"})
                .Add({"Управление", "---"})
                .Add({"Шина", "---"})
                .Add({"ДЗТ (RCD)", "---"})
            End With
            Dim i As Integer = 0
            Dim Мощност As Double = 0.0
            Dim brLamp As Integer = 0
            Dim brKontakt As Integer = 0
            Dim brCol As Integer = 0
            Dim ИзлазКонтакти As Integer = 0
            Dim ИзлазТок As Double = 0
            Dim ИзлазТрифази As Boolean = False
            Dim ТаблоТрифази As Boolean = False
            Dim Nula_RCD As Integer = 1
            Dim j As Integer = 0
            For i = 0 To colArray.countTokKryg
                With .Columns(2 + brCol)
                    .Width = 100
                    .HeaderText = colArray.Tokowkryg(i).ТоковКръг
                    .Name = colArray.Tokowkryg(i).ТоковКръг
                    .SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
                End With
                If colArray.Tokowkryg(i).faza <> "L1" And colArray.Tokowkryg(i).faza <> Nothing Then
                    ТаблоТрифази = True
                End If
                .Rows(1).Cells(2 + brCol).Value = colArray.Tokowkryg(i).Tok.ToString("0.00")
                .Rows(2).Cells(2 + brCol).Value = "EZ9 MCB"
                .Rows(3).Cells(2 + brCol).Value = colArray.Tokowkryg(i).RatedCurrent & "A"
                .Rows(4).Cells(2 + brCol).Value = colArray.Tokowkryg(i).Sensitivity
                .Rows(5).Cells(2 + brCol).Value = colArray.Tokowkryg(i).Curve
                .Rows(6).Cells(2 + brCol).Value = colArray.Tokowkryg(i).NumberPoles

                If colArray.Tokowkryg(i).konsuator1 = "Контакти" Then
                    .Rows(defkt).Cells(2 + brCol).Value = "N" + Nula_RCD.ToString
                    ИзлазКонтакти += 1
                    If colArray.Tokowkryg(i).faza <> "L1" And colArray.Tokowkryg(i).faza <> Nothing Then
                        ИзлазТок += colArray.Tokowkryg(i).Мощност * 1.2 / (0.38 * Math.Sqrt(3) * 0.9)
                    Else
                        ИзлазТок += colArray.Tokowkryg(i).Мощност * 1.2 / (0.22 * 0.9)
                    End If
                    If colArray.Tokowkryg(i).faza <> "L1" And colArray.Tokowkryg(i).faza <> Nothing Then
                        ИзлазТрифази = True
                    End If
                    If ИзлазКонтакти = 3 Then
                        .Rows(defkt + 1).Cells(2 + brCol).Value = "EZ9 RCCB"
                        .Rows(defkt + 2).Cells(2 + brCol).Value = "AC"
                        Select Case ИзлазТок
                            Case < 25
                                .Rows(defkt + 3).Cells(2 + brCol).Value = "25А"
                            Case < 40
                                .Rows(defkt + 3).Cells(2 + brCol).Value = "40А"
                            Case < 63
                                .Rows(defkt + 3).Cells(2 + brCol).Value = "63А"
                            Case Else
                                .Rows(defkt + 3).Cells(2 + brCol).Value = "#####"
                        End Select
                        .Rows(defkt + 4).Cells(2 + brCol).Value = "30mA"
                        .Rows(defkt + 5).Cells(2 + brCol).Value = IIf(ИзлазТрифази, "4p", "2p")
                        ИзлазТрифази = False
                        ИзлазКонтакти = 0
                        Nula_RCD += 1
                        ИзлазТок = 0
                    End If
                End If

                .Rows(za6t + 1).Cells(2 + brCol).Value = colArray.Tokowkryg(i).brLamp
                .Rows(za6t + 2).Cells(2 + brCol).Value = colArray.Tokowkryg(i).brKontakt
                .Rows(za6t + 3).Cells(2 + brCol).Value = colArray.Tokowkryg(i).Мощност.ToString("0.000")
                .Rows(za6t + 4).Cells(2 + brCol).Value = "СВТ"
                .Rows(za6t + 5).Cells(2 + brCol).Value = colArray.Tokowkryg(i).Kabebel_Se4enie
                .Rows(za6t + 6).Cells(2 + brCol).Value = colArray.Tokowkryg(i).faza
                .Rows(za6t + 7).Cells(2 + brCol).Value = colArray.Tokowkryg(i).konsuator1
                .Rows(za6t + 8).Cells(2 + brCol).Value = colArray.Tokowkryg(i).konsuator2

                Dim comboBoxCell_UP As New SWF.DataGridViewComboBoxCell()
                comboBoxCell_UP.ToolTipText = "Управление"
                comboBoxCell_UP.Items.AddRange(New String() {"Няма",
                                                            "Импулсно реле",
                                                            "Моторна защита",
                                                            "Контактор",
                                                            "Стълбищен автомат",
                                                            "Фото реле"})
                comboBoxCell_UP.Style.Alignment = SWF.DataGridViewContentAlignment.MiddleCenter
                .Rows(za6t + 9).Cells(2 + brCol) = comboBoxCell_UP

                Dim checkBoxCell As New SWF.DataGridViewCheckBoxCell()
                checkBoxCell.ToolTipText = "към друга щина"
                checkBoxCell.Value = False
                checkBoxCell.Style.Alignment = SWF.DataGridViewContentAlignment.MiddleCenter
                .Rows(za6t + 10).Cells(2 + brCol) = checkBoxCell

                Dim checkBoxCell_RCD As New SWF.DataGridViewCheckBoxCell()
                checkBoxCell_RCD.ToolTipText = "Добавя ДЗТ"
                checkBoxCell_RCD.Value = False
                checkBoxCell_RCD.Style.Alignment = SWF.DataGridViewContentAlignment.MiddleCenter
                .Rows(za6t + 11).Cells(2 + brCol) = checkBoxCell_RCD

                brLamp += colArray.Tokowkryg(i).brLamp
                brKontakt += colArray.Tokowkryg(i).brKontakt
                Мощност += colArray.Tokowkryg(i).Мощност
                brCol += 1
            Next

            With .Columns(1 + brCol)
                .DefaultCellStyle = DataGridViewCellStyle3
                .Width = 100
                .HeaderText = "ОБЩО"
                .Name = "ОБЩО"
                .SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
            End With

            .Rows(0).Cells("ОБЩО").Value = "за таблото"
            .Rows(1).Cells("ОБЩО").Value = calc_Inom(Мощност, IIf(ТаблоТрифази, "3p", "1p"))

            ' Създаване на комбинирана клетка за избор за "Тип на апарата"
            Dim comboBoxCell1 As New SWF.DataGridViewComboBoxCell()
            comboBoxCell1.Style.Alignment = SWF.DataGridViewContentAlignment.MiddleCenter
            comboBoxCell1.ToolTipText = "Тип на апарата"
            comboBoxCell1.Items.AddRange(New String() {"iSW", "INS", "IN", "EZ9 MCB", "C120", "NSX", "MTZ2"})


            Dim tokOptions As String() = {""}
            If ТаблоТрифази Then
                Select Case .Rows(1).Cells("ОБЩО").Value
                    Case < 75
                        tokOptions = {"20", "32", "40", "63", "100", "125"}
                        comboBoxCell1.Value = "iSW"
                    Case < 1100
                        tokOptions = {"40", "63", "80", "100", "125", "160", "200", "250", "320", "400", "500", "630", "1000", "1250", "1600"}
                        comboBoxCell1.Value = "INS"
                    Case > 1100
                        tokOptions = {"1000", "1600", "2500"}
                        comboBoxCell1.Value = "IN"
                End Select
            Else
                tokOptions = {"20", "32", "40", "63", "100", "125"}
                comboBoxCell1.Value = "iSW"
            End If

            ' Създаване на комбинирана клетка за избор за "Ток на апарата"
            Dim comboBoxCell2 As New SWF.DataGridViewComboBoxCell()
            comboBoxCell2.Style.Alignment = SWF.DataGridViewContentAlignment.MiddleCenter
            comboBoxCell2.ToolTipText = "Ток на апарата"
            comboBoxCell2.Items.AddRange(tokOptions)

            ' Създаване на комбинирана клетка за избор за "Тип на кабела"
            Dim comboBoxCell3 As New SWF.DataGridViewComboBoxCell()
            comboBoxCell3.Style.Alignment = SWF.DataGridViewContentAlignment.MiddleCenter
            comboBoxCell3.ToolTipText = "Тип на кабела"
            comboBoxCell3.Items.AddRange(New String() {"СВТ", "САВТ", "Al/R", "Al/R+СВТ"})
            comboBoxCell3.Value = "СВТ"

            ' Създаване на комбинирана клетка за избор за НАЧИН НА ВКЛЮЧВАНЕ
            Dim comboBoxCell4 As New SWF.DataGridViewComboBoxCell()
            comboBoxCell4.Style.Alignment = SWF.DataGridViewContentAlignment.MiddleCenter
            comboBoxCell4.ToolTipText = "Тип включване"
            comboBoxCell4.Items.AddRange(New String() {"Палец", "Задвижка"})
            comboBoxCell4.Value = "Палец"

            ' Задаване на клетките в съответните редове и колони
            .Rows(2).Cells("ОБЩО") = comboBoxCell1
            .Rows(3).Cells("ОБЩО") = comboBoxCell2

            comboBoxCell2.Value = calc_ISW(IIf(ТаблоТрифази, "3p", "1p"), Pkryg:=Мощност)
            .Rows(4).Cells("ОБЩО").Value = "-"
            .Rows(5).Cells("ОБЩО").Value = "-"
            .Rows(6).Cells("ОБЩО").Value = IIf(ТаблоТрифази, "3p", "1p")
            .Rows(7).Cells("ОБЩО") = comboBoxCell4

            .Rows(za6t + 1).Cells("ОБЩО").Value = brLamp
            .Rows(za6t + 2).Cells("ОБЩО").Value = brKontakt
            .Rows(za6t + 3).Cells("ОБЩО").Value = Мощност
            .Rows(za6t + 4).Cells("ОБЩО") = comboBoxCell3
            .Rows(za6t + 5).Cells("ОБЩО").Value = calc_cable_Cu(.Rows(1).Cells("ОБЩО").Value,
                                                                .Rows(6).Cells("ОБЩО").Value)
            .Rows(za6t + 6).Cells("ОБЩО").Value = IIf(ТаблоТрифази, "L1,L2,L3", "L")
            .Rows(za6t + 7).Cells("ОБЩО").Value = "Ке=" & FormatNumber(15 / Мощност, 2)
            .Rows(za6t + 8).Cells("ОБЩО").Value = "Рпр.=15кW"

            If ИзлазКонтакти = 1 Then
                For brCol = .ColumnCount To 2 Step -1
                    If colArray.Tokowkryg(brCol).konsuator1 = "Контакти" Then
                        If .Rows(defkt + 1).Cells(brCol - 1).Value = "EZ9 RCCB" Then
                        Else
                            .Rows(defkt + 1).Cells(2 + brCol - 1).Value = ""
                            .Rows(defkt + 2).Cells(2 + brCol - 1).Value = ""
                            .Rows(defkt + 3).Cells(2 + brCol - 1).Value = ""
                            .Rows(defkt + 4).Cells(2 + brCol - 1).Value = ""
                            .Rows(defkt + 5).Cells(2 + brCol - 1).Value = ""
                            .Rows(defkt + 1).Cells(2 + brCol).Value = "EZ9 RCCB"
                            .Rows(defkt + 2).Cells(2 + brCol).Value = "AC"
                            Select Case ИзлазТок
                                Case < 25
                                    .Rows(defkt + 3).Cells(2 + brCol).Value = "25А"
                                Case < 40
                                    .Rows(defkt + 3).Cells(2 + brCol).Value = "40А"
                                Case < 63
                                    .Rows(defkt + 3).Cells(2 + brCol).Value = "63А"
                            End Select
                            .Rows(defkt + 4).Cells(2 + brCol).Value = "30mA"
                            .Rows(defkt + 5).Cells(2 + brCol).Value = IIf(ИзлазТрифази, "4p", "2p")
                            Exit For
                        End If
                    End If
                Next
            End If
            If ИзлазКонтакти = 2 Then
                For brCol = .ColumnCount To 2 Step -1
                    If colArray.Tokowkryg(brCol).konsuator1 = "Контакти" Then
                        .Rows(defkt + 1).Cells(2 + brCol).Value = "EZ9 RCCB"
                        .Rows(defkt + 2).Cells(2 + brCol).Value = "AC"
                        Select Case ИзлазТок
                            Case < 25
                                .Rows(defkt + 3).Cells(2 + brCol).Value = "25А"
                            Case < 40
                                .Rows(defkt + 3).Cells(2 + brCol).Value = "40А"
                            Case < 63
                                .Rows(defkt + 3).Cells(2 + brCol).Value = "63А"
                        End Select
                        .Rows(defkt + 4).Cells(2 + brCol).Value = "30mA"
                        .Rows(defkt + 5).Cells(2 + brCol).Value = IIf(ИзлазТрифази, "4p", "2p")
                        Exit For
                    End If
                Next
            End If

        End With
        SetRCD(dagrid)
        '  AddHandler dagrid.Click, AddressOf btnExport_dagrid_Click
        AddHandler dagrid.CellValueChanged, AddressOf DataGridView1_CellValueChanged
        AddHandler dagrid.DataError,
            Sub(sender As Object, e As System.Windows.Forms.DataGridViewDataErrorEventArgs)
                e.ThrowException = False ' Пропускаме грешката, за да предотвратим изскачащото съобщение
            End Sub
        Return dagrid
    End Function
    Private Sub DataGridView1_CellValueChanged(ByVal sender As Object, ByVal e As SWF.DataGridViewCellEventArgs)
        Try
            Dim dagrid As SWF.DataGridView = CType(sender, SWF.DataGridView)
            Dim TabloName As String = Mid(dagrid.Name, 1, Len(dagrid.Name) - 3)
            'If isCellChangeTriggeredProgrammatically Then
            '    Return
            'End If
            isCellChangeTriggeredProgrammatically = True
            If e.ColumnIndex = dagrid.Columns("ОБЩО").Index AndAlso e.RowIndex = 19 Then
                Dim combo As String = ""
                ' Променяме цвета на фона на клетката
                MsgBox("Баш тази клетка не пишай или я изчисли правилно")
                Exit Try
            End If
            If e.ColumnIndex = dagrid.Columns("Общо").Index Then
                Dim rowIndex As Integer = dagrid.CurrentRow.Index
                Dim colIndex As Integer = dagrid.CurrentCell.ColumnIndex

                ' Клетка която извиква прекъсването
                Dim input As String = dagrid.Rows(rowIndex).Cells("ОБЩО").Value
                Dim Pinst As Double = Val(dagrid.Rows(za6t + 3).Cells("ОБЩО").Value)
                Dim Ibreaker As String = dagrid.Rows(3).Cells("ОБЩО").Value     ' Ток на Автоматичния прекъсвач
                Dim NumberPoles As String = dagrid.Rows(6).Cells("ОБЩО").Value  ' Брой на фазите

                If rowIndex = (za6t + 8) Then
                    Dim number As String = System.Text.RegularExpressions.Regex.Match(input, "\d+").Value
                    Dim result As Integer = Convert.ToInt32(number)

                    dagrid.Rows(rowIndex - 1).Cells("ОБЩО").Value = "Ke=" & (result / Pinst).ToString("0.00")

                    Return
                End If

                ' Избира кабел
                With True
                    'Ibreaker,      ' Ток на Автоматичния прекъсвач
                    'NumberPoles,   ' Брой на фазите
                    'layMethod:=0,  ' Начин на полагане 0 - във въздух; 1 - в земя
                    'Broj_Cable:=1, ' Брой кабели положени паралелно на скара
                    'Tipe_Cable:=0, ' 0 - Кабел; 1 - проводник
                    'matType:=0     ' 0 - Мед; 1 - Алуминии

                    Dim cable As String = ""
                    Dim index As Integer = 0
                    Dim calc_N As String
                    Dim Poles As String = ""
                    Dim Text As String = ""

                    Select Case dagrid.Rows(za6t + 4).Cells("ОБЩО").Value
                        Case "СВТ"
                            cable = calc_cable(Ibreaker, NumberPoles, Tipe_Cable:=0, matType:=0, RetType:=0)
                        Case "САВТ"
                            cable = calc_cable(Ibreaker, NumberPoles, Tipe_Cable:=0, matType:=1, RetType:=0)
                        Case "Al/R"
                            dagrid.Rows(za6t + 5).Cells("ОБЩО").Value = Get_cable_AlR(cable, NumberPoles)
                            Return
                        Case "Al/R+СВТ"
                            Text = Get_cable_AlR(cable, NumberPoles)
                            Text += "/"
                            Text += calc_cable(Ibreaker, NumberPoles, Tipe_Cable:=0, matType:=0, RetType:=0)
                            Return
                        Case Else
                            cable = calc_cable(Ibreaker, NumberPoles, Tipe_Cable:=0, matType:=1, RetType:=0)
                    End Select
                    index = Array.IndexOf(Kable_Size_L, cable)
                    calc_N = Kable_Size_N(index)
                    If NumberPoles = "1p" Then
                        If Val(cable) > 35 Then
                            Text = "НЯМА"
                        Else
                            If TabloName = "Гл.Р.Т." Or TabloName = "ГлРТ" Then
                                Text = "2x" & cable
                            Else
                                Text = "3x" & cable
                            End If

                        End If
                    Else
                        If calc_N = 0 Then
                            Poles = If(NumberPoles = "1p", "2x", "4x")
                            Text = "4x" & cable
                        Else
                            Poles = "3x"
                            Text = Poles & cable & "+" & calc_N
                        End If
                    End If
                    dagrid.Rows(za6t + 5).Cells("ОБЩО").Value = Text
                End With
            End If
            ' Проверка дали променената клетка е в колона "ОБЩО" и в ред 2
            If e.ColumnIndex = dagrid.Columns("ОБЩО").Index AndAlso e.RowIndex = 2 Then

                ' Получаване на новата стойност от клетката
                Dim selectedValue As String = dagrid.Rows(2).Cells("ОБЩО").Value.ToString()

                ' Динамично попълване на comboBoxCell2 спрямо избора
                Dim токOptions As New List(Of Integer)

                ' Проверка за съвпадение в речниците и добавяне на съответните стойности
                If Breakers.ContainsValue(selectedValue) Then
                    токOptions = Breakers.Where(Function(kv) kv.Value = selectedValue).Select(Function(kv) kv.Key).ToList()
                ElseIf Disconnectors.ContainsValue(selectedValue) Then
                    токOptions = Disconnectors.Where(Function(kv) kv.Value = selectedValue).Select(Function(kv) kv.Key).ToList()
                Else
                    ' Ако няма съвпадение, занули стойността в клетката и прекрати изпълнението
                    dagrid.Rows(2).Cells("ОБЩО").Value = Nothing
                    Return
                End If
                ' Изчистване и добавяне на новите стойности в comboBoxCell2
                Dim comboBoxCell2 = CType(dagrid.Rows(3).Cells("ОБЩО"), SWF.DataGridViewComboBoxCell)
                comboBoxCell2.Items.Clear()
                comboBoxCell2.Items.AddRange(токOptions.Select(Function(i) i.ToString()).ToArray())
            End If
            If e.RowIndex = 2 Then



            End If



        Finally
            isCellChangeTriggeredProgrammatically = False
        End Try
    End Sub
    Private Sub addColomDataGrid(dagrid As Windows.Forms.DataGridView)
        Dim dtCol As System.Windows.Forms.DataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        With dtCol
            .Width = 100
            .HeaderText = "т.к."
            .Frozen = vbFalse
            .ReadOnly = vbFalse
            .SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        End With
        dagrid.Columns.Add(dtCol)
    End Sub
    Private Sub table(arrTablo As Array)
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim edt As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database
        Dim pPtRes As PromptPointResult
        Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
        pPtOpts.Message = vbLf & "Изберете точка на вмъкване на Tablicata:  "
        pPtRes = acDoc.Editor.GetPoint(pPtOpts)

        If pPtRes.Status <> PromptStatus.OK Then Exit Sub

        Dim table As Table = New Table()

        table.TableStyle = acCurDb.Tablestyle
        table.SetSize(22, 5)
        table.SetRowHeight(1)

        table.Position = pPtRes.Value

        Dim j As Integer = 0
        For Each brTablo As strTablo In arrTablo
            If arrTablo(j).Name = Nothing Then Exit For
            table.Cells(j + 1, 0).TextString = arrTablo(j).Name
            j += 1
        Next

        table.GenerateLayout()
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

            Dim tr As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

            tr.AppendEntity(table)
            acTrans.AddNewlyCreatedDBObject(table, True)
        End Using

    End Sub
    Private Sub btnExport_Block_Click(sender As Object, e As EventArgs)
        Dim send As System.Windows.Forms.Button = sender
        Dim name As String = ""
        Dim nam As String = ""
        Dim func As String = ""
        name = Mid(send.Name.ToString, 1, InStr(send.Name.ToString, "/#/") - 1)
        func = Mid(send.Name.ToString, InStr(send.Name.ToString, "/#/") + 3, Len(send.Name.ToString))
        '
        'Качва по стълбицата нагоре за да намери SplitContainer в който е DataGridView
        '
        Dim panel As System.Windows.Forms.SplitContainer = DirectCast(send.Parent.Parent.Parent, System.Windows.Forms.SplitContainer)
        Dim obList As List(Of Object) = New List(Of Object)
        Call FindChildren(panel, obList)
        Dim index As Integer
        For Each ctrl As Windows.Forms.Control In obList
            If TypeOf ctrl Is Windows.Forms.DataGridView Then
                Exit For
            End If
            index += 1
        Next
        Dim dagrid As System.Windows.Forms.DataGridView = DirectCast(obList(index), System.Windows.Forms.DataGridView)
        Select Case func
            Case "1"
                CreateTablo(dagrid)
            Case "2"
                SetBreakers(dagrid)
            Case "3"
                SetRCD(dagrid)
            Case "4"
                SetBalance(dagrid)
            Case "5"
                Calculate_Faze(dagrid)
            Case "6"
                InsertTablo(dagrid)
            Case "7"
                InsertBus(dagrid)
        End Select
        form_AS_tablo.Visible = vbTrue
    End Sub
    Private Sub InsertBus(DataGridView As Windows.Forms.DataGridView)
        Dim boSecondBus As Boolean = False
        isCellChangeTriggeredProgrammatically = True
        Dim sumPkryg As Double = 0
        Dim ТаблоТрифази As Boolean = False
        For index As Integer = 2 To DataGridView.Columns.Count - 2
            If DataGridView.Rows(za6t + 10).Cells(index).Value Then
                sumPkryg += DataGridView.Rows(za6t + 3).Cells(index).Value                      ' Сумира мощностите на токовите кръгове
                If DataGridView.Rows(6).Cells(index).Value = "3p" And Not ТаблоТрифази Then
                    ТаблоТрифази = True
                End If
                If Not boSecondBus Then
                    boSecondBus = True
                End If
            Else
                If boSecondBus Then
                    boSecondBus = True
                    Dim col As New SWF.DataGridViewTextBoxColumn
                    Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
                    With DataGridViewCellStyle3
                        .Format = "N2"
                        .NullValue = Nothing
                    End With
                    With col
                        .Name = "Разединител"
                        .HeaderText = "Мощностен"
                        .DefaultCellStyle = DataGridViewCellStyle3
                        .Width = 100
                        .SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
                    End With
                    DataGridView.Columns.Insert(index, col)
                    DataGridView.Refresh()
                    SWF.Application.DoEvents()
                    For Each row As SWF.DataGridViewRow In DataGridView.Rows
                        row.Cells("Разединител").Value = String.Empty
                    Next
                    ' Създаване на комбинирана клетка за избор за "Тип на апарата"
                    Dim comboBoxCell1 As New SWF.DataGridViewComboBoxCell()

                    comboBoxCell1.Style.Alignment = SWF.DataGridViewContentAlignment.MiddleCenter
                    comboBoxCell1.ToolTipText = "Тип на апарата"
                    comboBoxCell1.Items.AddRange(New String() {"iSW", "INS", "IN"})

                    Dim tokOptions As String()
                    If ТаблоТрифази Then
                        Select Case DataGridView.Rows(1).Cells("Разединител").Value
                            Case < 75
                                tokOptions = {"20", "32", "40", "63", "100", "125"}
                                comboBoxCell1.Value = "iSW"
                            Case < 1100
                                tokOptions = {"40", "63", "80", "100", "125", "160", "200", "250", "320", "400", "500", "630", "1000", "1250", "1600"}
                                comboBoxCell1.Value = "INS"
                            Case > 1100
                                tokOptions = {"1000", "1600", "2500"}
                                comboBoxCell1.Value = "IN"
                        End Select
                    Else
                        tokOptions = {"20", "32", "40", "63", "100", "125"}
                        comboBoxCell1.Value = "iSW"
                    End If

                    ' Създаване на комбинирана клетка за избор за "Ток на апарата"
                    Dim comboBoxCell2 As New SWF.DataGridViewComboBoxCell()
                    comboBoxCell2.Style.Alignment = SWF.DataGridViewContentAlignment.MiddleCenter
                    comboBoxCell2.ToolTipText = "Ток на апарата"
                    comboBoxCell2.Items.AddRange(tokOptions)

                    ' Задаване на клетките в съответните редове и колони
                    DataGridView.Rows(2).Cells("Разединител") = comboBoxCell1
                    DataGridView.Rows(3).Cells("Разединител") = comboBoxCell2
                    DataGridView.Rows(za6t + 4).Cells("Разединител").Value = "Шина"

                    ' comboBoxCell2.Value = calc_ISW(IIf(ТаблоТрифази, "3p", "1p"), Pkryg:=Мощност)

                    DataGridView.Rows(0).Cells("Разединител").Value = "pазединител"
                    DataGridView.Rows(1).Cells("Разединител").Value = "-"                           ' Изчислен ток
                    DataGridView.Rows(4).Cells("Разединител").Value = "-"                           ' изключвателна възможност
                    DataGridView.Rows(5).Cells("Разединител").Value = "-"                           ' крива
                    DataGridView.Rows(6).Cells("Разединител").Value = IIf(ТаблоТрифази, "3p", "1p") ' брой полюси
                    boSecondBus = False
                    Exit For
                End If
            End If
        Next
        Calculate_Bus(DataGridView)
        isCellChangeTriggeredProgrammatically = False
    End Sub
    Private Sub Calc_Bus(DataGridView As Windows.Forms.DataGridView,
                         calcColumn As String                          ' Колона в която се записва редултата Очаква се или "Разединител" или "ОБЩО"
                         )

        For index As Integer = 2 To DataGridView.Columns.Count - 2



        Next
    End Sub
    Private Sub SetBreakers(DataGridView As Windows.Forms.DataGridView)

    End Sub
    Private Sub InsertTablo(DataGridView As Windows.Forms.DataGridView)
        Dim brColums As Integer = DataGridView.Columns.Count - 1
        Dim Name_Tablo As String = Mid(DataGridView.Name, 1, Len(DataGridView.Name) - 3)
        Dim GRT_Name As String = "Гл.Р.Т."
        Dim Токов_кръг As String = ""
        Dim Брой_лампи As String = IIf(DataGridView.Rows(za6t + 1).Cells(brColums).Value = 0,
                          "----",
                          DataGridView.Rows(za6t + 1).Cells(brColums).Value.ToString)
        Dim Брой_контакти As String = IIf(DataGridView.Rows(za6t + 2).Cells(brColums).Value = 0,
                                  "----",
                                  DataGridView.Rows(za6t + 2).Cells(brColums).Value.ToString)
        Dim МОЩНОСТ As String = (DataGridView.Rows(za6t + 3).Cells(brColums).Value * 1000).ToString
        Dim Кабел_тип As String = DataGridView.Rows(za6t + 4).Cells(brColums).Value.ToString
        Dim Кабел_сечение As String = DataGridView.Rows(za6t + 5).Cells(brColums).Value.ToString
        Dim Фази As String = DataGridView.Rows(za6t + 6).Cells(brColums).Value.ToString
        Try
            '
            ' Поставя блок на табло с даннни
            '
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database
            Dim ptBasePointRes As PromptPointResult
            Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
            '' Prompt for the start point
            form_AS_tablo.Visible = vbFalse

            pPtOpts.Message = vbLf & "Изберете точка на вмъкване на блока за таблото: "
            ptBasePointRes = acDoc.Editor.GetPoint(pPtOpts)

            If ptBasePointRes.Status = PromptStatus.Cancel Then Exit Sub
            Dim ptBasePoint As Point3d = ptBasePointRes.Value

            Dim blkRecId_ As ObjectId = ObjectId.Null
            blkRecId_ = cu.InsertBlock("Табло_Главно",
                           New Point3d(ptBasePoint.X, ptBasePoint.Y, 0),
                           "EL_ТАБЛА",
                           New Scale3d(2, 2, 2)
                           )

            Using trans As Transaction = acDoc.TransactionManager.StartTransaction()

                Dim acBlkTbl As BlockTable = trans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim acBlkRef As BlockReference =
                                DirectCast(trans.GetObject(blkRecId_, OpenMode.ForWrite), BlockReference)

                Dim Index As Integer = DataGridView.Columns.Count - 1
                Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                For Each objID As ObjectId In attCol
                    Dim dbObj As DBObject = trans.GetObject(objID, OpenMode.ForWrite)
                    Dim acAttRef As AttributeReference = dbObj
                    If acAttRef.Tag = "ТАБЛО" Then acAttRef.TextString = GRT_Name
                    If acAttRef.Tag = "ГЛ.Р.Т." Then acAttRef.TextString = Name_Tablo
                    If acAttRef.Tag = "ТОКОВ_КРЪГ" Then acAttRef.TextString = Name_Tablo
                    If acAttRef.Tag = "БРОЙ_ЛАМПИ" Then acAttRef.TextString = Брой_лампи
                    If acAttRef.Tag = "БРОЙ_КОНТАКТИ" Then acAttRef.TextString = Брой_контакти
                    If acAttRef.Tag = "МОЩНОСТ" Then acAttRef.TextString = МОЩНОСТ
                    If acAttRef.Tag = "КАБЕЛ_ТИП" Then acAttRef.TextString = Кабел_тип
                    If acAttRef.Tag = "КАБЕЛ_СЕЧЕНИЕ" Then acAttRef.TextString = Кабел_сечение
                    If acAttRef.Tag = "ФАЗИ" Then acAttRef.TextString = Фази
                Next
                trans.Commit()
            End Using
        Catch ex As Exception
            MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
        End Try
    End Sub
    Private Sub SetRCD(DataGridView As Windows.Forms.DataGridView)
        ' Локална променлива за сумиране на токове
        Dim sum As Double = 0
        ' Обхожда всички колони от втората до предпоследната
        For index As Integer = 2 To DataGridView.Columns.Count - 2
            ' Проверява дали редът (za6t + 7) съдържа "Контакти"
            ' Ако не е "Контакти", пропуска текущата колона
            If Not DataGridView.Rows(za6t + 7).Cells(index).Value = "Контакти" Then Continue For
            ' Всички следващи действия са върху DataGridView
            With DataGridView
                ' Добавя към сумата резултата от функцията calc_Inom,
                ' която изчислява стойност на база на две клетки от различни редове
                sum += calc_Inom(.Rows(za6t + 3).Cells(index).Value, .Rows(6).Cells(index).Value)
                ' Занулява ред 1 за текущата колона (вероятно ред за резултат)
                .Rows(1).Cells(index).Value = 0
                ' Изчиства предишни стойности в зоната за дефектнотокова защита
                .Rows(defkt + 1).Cells(index).Value = Nothing
                .Rows(defkt + 2).Cells(index).Value = Nothing
                .Rows(defkt + 3).Cells(index).Value = Nothing
                .Rows(defkt + 4).Cells(index).Value = Nothing
                .Rows(defkt + 5).Cells(index).Value = Nothing
                ' --- Проверка 1: ако стойността в ред defkt е различна от лявата и дясната колона ---
                If .Rows(defkt).Cells(index - 1).Value <> .Rows(defkt).Cells(index).Value And
                    .Rows(defkt).Cells(index + 1).Value <> .Rows(defkt).Cells(index).Value Then
                    ' Изчиства съдържанието в редове 2 до 6
                    .Rows(2).Cells(index).Value = ""
                    .Rows(3).Cells(index).Value = ""
                    .Rows(4).Cells(index).Value = ""
                    .Rows(5).Cells(index).Value = ""
                    .Rows(6).Cells(index).Value = ""

                    ' Поставя дефектнотокова защита (RCBO) със зададени параметри
                    .Rows(defkt).Cells(index).Value = "N"                 ' Неутрален проводник
                    .Rows(defkt + 1).Cells(index).Value = "EZ9 RCBO"      ' Тип защита
                    .Rows(defkt + 2).Cells(index).Value = "AC"            ' Тип чувствителност
                    .Rows(defkt + 3).Cells(index).Value = "20А"           ' Номинален ток
                    .Rows(defkt + 4).Cells(index).Value = "30mA"          ' Чувствителност
                    ' Определя броя на полюсите в зависимост от това дали е трифазно
                    .Rows(defkt + 5).Cells(index).Value = IIf(.Rows(za6t + 6).Cells(index).Value = "L1,L2,L3", "4p", "2p")

                    ' Записва сумарната стойност във формат 0.00
                    .Rows(1).Cells(index).Value = sum.ToString("0.00")
                    ' Нулира акумулираната сума
                    sum = 0
                End If
                ' --- Проверка 2: ако текущата стойност е еднаква с лявата и дясната ---
                ' Тук не се извършва действие, вероятно е оставено за бъдеща логика
                If .Rows(defkt).Cells(index - 1).Value = .Rows(defkt).Cells(index).Value And
                    .Rows(defkt).Cells(index + 1).Value = .Rows(defkt).Cells(index).Value Then
                End If
                ' --- Проверка 3: ако лявата е различна, а дясната е еднаква ---
                ' Също празен блок, вероятно резервиран за бъдещи правила
                If .Rows(defkt).Cells(index - 1).Value <> .Rows(defkt).Cells(index).Value And
                .Rows(defkt).Cells(index + 1).Value = .Rows(defkt).Cells(index).Value Then

                End If
                ' --- Проверка 4: ако лявата е еднаква, а дясната е различна ---
                ' В този случай се задава RCCB (само дефектнотокова защита)
                If .Rows(defkt).Cells(index - 1).Value = .Rows(defkt).Cells(index).Value And
                    .Rows(defkt).Cells(index + 1).Value <> .Rows(defkt).Cells(index).Value Then

                    .Rows(defkt + 1).Cells(index).Value = "EZ9 RCCB"      ' Тип устройство
                    .Rows(defkt + 2).Cells(index).Value = "AC"            ' Тип чувствителност
                    .Rows(defkt + 3).Cells(index).Value = "25А"           ' Номинален ток
                    .Rows(defkt + 4).Cells(index).Value = "30mA"          ' Чувствителност
                    ' Определя полюсите според това дали е трифазно
                    .Rows(defkt + 5).Cells(index).Value =
                    IIf(.Rows(za6t + 6).Cells(index).Value = "L1,L2,L3", "4p", "2p")
                    ' Записва сумарната стойност във формат 0.00
                    .Rows(1).Cells(index).Value = sum.ToString("0.00")
                    ' Нулира сумата
                    sum = 0
                End If
            End With
        Next
        ' --- Втори цикъл: разпределяне/наследяване на фазата ---
        Dim faza As String = ""
        ' Обхожда колоните в обратен ред, за да разпространи стойността на фазата
        For index As Integer = DataGridView.Columns.Count - 2 To 2 Step -1
            ' Работи само с колони "Контакти"
            If Not DataGridView.Rows(za6t + 7).Cells(index).Value = "Контакти" Then Continue For
            ' Ако в ред 1 има изчислена стойност (по-голяма от 0)
            ' Записва текущата фаза в променливата faza
            If DataGridView.Rows(1).Cells(index).Value > 0 Then
                faza = DataGridView.Rows(za6t + 6).Cells(index).Value
            Else
                ' Ако няма стойност, колона наследява фазата от предходната
                DataGridView.Rows(za6t + 6).Cells(index).Value = faza
            End If
        Next
        ' -------------------------------------------------------------------
        ' Търси и поставя ДЗТ (дефектнотокова защита) на бойлерите
        ' -------------------------------------------------------------------
        With DataGridView
            ' Обхожда отново всички колони
            For index As Integer = 2 To .Columns.Count - 2
                ' Проверява дали колоната е от тип "Бойлер"
                ' Ако колоната не е "Бойлер", прескачаме
                If Not .Rows(za6t + 7).Cells(index).Value = "Бойлер" Then Continue For
                ' Изчиства ненужни редове в колоната
                .Rows(2).Cells(index).Value = ""
                .Rows(3).Cells(index).Value = ""
                .Rows(4).Cells(index).Value = ""
                .Rows(5).Cells(index).Value = ""
                .Rows(6).Cells(index).Value = ""
                ' Поставя стандартна RCBO защита за бойлера
                .Rows(defkt).Cells(index).Value = "N"
                .Rows(defkt + 1).Cells(index).Value = "EZ9 RCBO"
                .Rows(defkt + 2).Cells(index).Value = "AC"
                .Rows(defkt + 3).Cells(index).Value = "20А"
                .Rows(defkt + 4).Cells(index).Value = "30mA"
                ' Определя броя на полюсите (4p за трифазно, 2p за еднофазно)
                .Rows(defkt + 5).Cells(index).Value = IIf(.Rows(za6t + 6).Cells(index).Value = "L1,L2,L3", "4p", "2p")
            Next
        End With
        ' -------------------------------------------------------------------
        ' Търси и поставя ДЗТ (дефектнотокова защита) на маркирани консуатори
        ' -------------------------------------------------------------------
        With DataGridView
            ' Обхождаме всички колони, започвайки от 2 до предпоследната
            For index As Integer = 2 To .Columns.Count - 2
                ' Опитваме се да вземем клетката с чекбокс от реда za6t + 11 и текущата колона
                Dim checkBoxCell As SWF.DataGridViewCheckBoxCell = TryCast(.Rows(za6t + 11).Cells(index), SWF.DataGridViewCheckBoxCell)
                ' Ако клетката не е чекбокс или не е отметната, пропускаме итерацията
                If Not checkBoxCell.Value = True Then Continue For
                Dim fazi As String = IIf(.Rows(za6t + 6).Cells(index).Value = "L1,L2,L3", "4p", "2p")
                Dim Itk As Double = Val(.Rows(1).Cells(index).Value)
                Dim RCD As strRCD = FindRCDPriority(Itk, fazi, "RCBO")
                ' Изчиства ненужни редове в колоната
                .Rows(2).Cells(index).Value = ""
                .Rows(3).Cells(index).Value = ""
                .Rows(4).Cells(index).Value = ""
                .Rows(5).Cells(index).Value = ""
                .Rows(6).Cells(index).Value = ""
                ' Поставя избраната RCBO защита
                .Rows(defkt).Cells(index).Value = "N"
                .Rows(defkt + 1).Cells(index).Value = If(RCD.DeviceType = "iID", RCD.DeviceType, "EZ9 " + RCD.DeviceType)
                .Rows(defkt + 2).Cells(index).Value = RCD.Type
                .Rows(defkt + 3).Cells(index).Value = RCD.NominalCurrent
                .Rows(defkt + 4).Cells(index).Value = RCD.Sensitivity
                .Rows(defkt + 5).Cells(index).Value = RCD.Poles
            Next
        End With
    End Sub
    ''' <summary>
    ''' Намира подходящо защитно устройство (RCCB, RCBO или iID) според предпочитан тип,
    ''' върхов ток на кръга и брой полюси.
    ''' </summary>
    Private Function FindRCDPriority(Ikryg As Double, Poles As String, PreferredDevice As String) As strRCD

        ' --- 1. Проверка за RCCB ---
        If PreferredDevice.ToUpper() = "RCCB" AndAlso Ikryg > 63 Then
            Return New strRCD With {
            .NominalCurrent = 0,
            .Type = "AC",
            .Poles = Poles,
            .Sensitivity = 30,
            .DeviceType = PreferredDevice
        }
        End If
        ' --- 2. Проверка за валиден брой полюси ---
        If Poles.ToLower() <> "2p" AndAlso Poles.ToLower() <> "4p" Then
            Return New strRCD With {
            .NominalCurrent = 0,
            .Type = "AC",
            .Poles = Poles,
            .Sensitivity = 30,
            .DeviceType = PreferredDevice
        }
        End If
        ' --- 3. Определяне на реда на търсене ---
        Dim searchTypes As New List(Of String)
        If PreferredDevice.ToUpper() = "RCBO" Then
            searchTypes.Add("RCBO")  ' първо RCBO
            searchTypes.Add("iID")   ' ако няма подходящо RCBO → търсим iID
        Else
            searchTypes.Add(PreferredDevice)
        End If
        ' --- 4. Търсене по типовете ---
        For Each t In searchTypes
            Dim filteredCatalog = RCD_Catalog.Where(Function(x) x.Type.ToUpper() = "AC" _
                                                AndAlso x.Sensitivity = 30 _
                                                AndAlso x.Poles.ToLower() = Poles.ToLower() _
                                                AndAlso x.DeviceType.ToUpper() = t.ToUpper()) _
                                          .OrderBy(Function(x) x.NominalCurrent)
            ' Първото устройство с NominalCurrent >= Ikryg
            Dim rcd = filteredCatalog.FirstOrDefault(Function(x) x.NominalCurrent >= Ikryg)
            If rcd.NominalCurrent > 0 Then
                Return rcd
            End If
        Next
        ' --- 5. Ако няма нищо подходящо ---
        Return New strRCD With {
        .NominalCurrent = 0,
        .Type = "AC",
        .Poles = Poles,
        .Sensitivity = 30,
        .DeviceType = PreferredDevice}
    End Function
    ''' <summary>
    ''' Създава и конфигурира GroupBox за управление на дадено табло.
    ''' GroupBox-ът съдържа бутони за различни действия, които могат да се приложат към това табло:
    ''' вмъкване на табло, избор на защита, поправка на ДЗТ, балансиране на фази, изчисляване на фази,
    ''' вмъкване на блок табло и разделяне на шина.
    ''' </summary>
    ''' <param name="name">
    ''' Името на таблото. Използва се за:
    ''' - Заглавието на GroupBox-а (в текста му)  
    ''' - Генериране на уникални имена за бутоните в GroupBox-а
    ''' </param>
    ''' <returns>
    ''' Връща напълно инициализиран System.Windows.Forms.GroupBox с размери, шрифт, цветове и всички бутони, готов за добавяне в форма или контейнер.
    ''' </returns>
    Private Function insGroupBox_BT(name As String) As Windows.Forms.GroupBox
        ' Създаваме нов GroupBox
        Dim GroupBox As System.Windows.Forms.GroupBox = New Windows.Forms.GroupBox
        ' Конфигуриране на основните свойства на GroupBox-а
        With GroupBox
            .Name = name & "_BT"                                    ' Уникално име за контрола, комбинирайки името на таблото с "_BT"
            .Size = New System.Drawing.Size(432, 250)               ' Размер на GroupBox
            .Location = New System.Drawing.Point(500, 6)            ' Позиция в родителския контейнер
            .Font = New Drawing.Font("Arial", 12, Drawing.FontStyle.Bold) ' Шрифт за заглавието
            .Text = "Действие за табло '" & name & "'"              ' Заглавие, показва името на таблото
            .BackColor = System.Drawing.SystemColors.Control        ' Фонов цвят
            .ForeColor = System.Drawing.SystemColors.WindowText     ' Цвят на текста
        End With
        ' Добавяне на бутони към GroupBox-а
        With GroupBox.Controls
            .Add(insButtonn(name & "/#/" & "1", "Вмъкни табло", 25, 6))       ' Бутона за създаване/вмъкване на табло
            .Add(insButtonn(name & "/#/" & "2", "Избери защита", 50, 6))         ' Бутона за избор на защита (прекъсвач)
            .Add(insButtonn(name & "/#/" & "3", "Поправи ДЗТ", 75, 6))           ' Бутона за поправка на ДЗТ (RCD)
            .Add(insButtonn(name & "/#/" & "4", "Балансирай фази", 100, 6))      ' Бутона за балансиране на фазите на таблото
            .Add(insButtonn(name & "/#/" & "5", "Изчисли фази", 100, 170))       ' Бутона за изчисление на фазите
            .Add(insButtonn(name & "/#/" & "6", "Вмъкни блок табло", 125, 6))    ' Бутона за вмъкване на блок табло
            .Add(insButtonn(name & "/#/" & "7", "Раздели шина", 150, 6))         ' Бутона за разделяне на шина
        End With
        ' Връщаме готовия GroupBox
        Return GroupBox
    End Function
    ''' <summary>
    ''' Балансира фазите на електрическо табло, като разпределя потребителските токови кръгове (консуматори)
    ''' по фазите L1, L2 и L3 с цел оптимално натоварване.
    ''' Процедурата извършва следните основни стъпки:
    ''' 1. Проверява дали таблото съдържа поне един трифазен консуматор.
    '''    Ако няма, пита потребителя дали да балансира еднофазно табло.
    ''' 2. Извиква SetRCD за разпределение на RCD (Residual Current Device) преди балансирането.
    ''' 3. Създава масив от структури ElectricalParameters, в който записва данните за токовите кръгове,
    '''    включително мощност, ток, фази, RCD и свързана шина.
    ''' 4. Сортира масива по токовете на консуматорите (най-големите първи).
    ''' 5. Балансира фазите чрез цикъл, като добавя консуматорите към най-малко натоварената фаза.
    ''' 6. Сумира токовете по фази за всички консуматори и записва резултатите в DataGridView.
    ''' 7. Пренасочва фазите така, че първият токов кръг да започва с L1.
    ''' 8. Повтаря разпределението на RCD след приключване на фазовото балансиране.
    ''' 9. Извършва допълнителни изчисления за консуматори, разпределени на различни шини.
    ''' </summary>
    ''' <param name="DataGridView">
    ''' DataGridView, съдържащ всички данни за токовите кръгове, фази, токове, RCD и шини.
    ''' Редовете трябва да съдържат следните типове данни:
    ''' - Rед za6t + 3: мощност
    ''' - Rед 1: ток
    ''' - Ред za6t + 6: фаза или CircuitType
    ''' - Rед defkt: RCD
    ''' - Ред 6: Phases
    ''' - Ред za6t + 10: Bus (шина)
    ''' Методът модифицира стойностите в тези редове и добавя резултати за общите токове по фази.
    ''' </param>
    Private Sub SetBalance(DataGridView As Windows.Forms.DataGridView)
        Dim brColums As Integer = DataGridView.Columns.Count
        Dim Faza_Tablo As Boolean = False
        Dim Faza As String = ""

        For index As Integer = 2 To brColums - 1
            Faza = DataGridView.Rows(za6t + 6).Cells(index).Value.ToString
            If Faza <> "L" And Not Faza_Tablo And Faza <> Nothing Then Faza_Tablo = True
        Next

        If Not Faza_Tablo Then
            If MsgBox("Да балансирам ли еднофазно табло?",
                      vbOKCancel,
                      "НЕ Е нито един трифазен консуматор") = MsgBoxResult.Cancel Then
                Exit Sub
            End If
        End If
        '
        ' Разпределя RCD за да балансира правилно фазите
        '
        SetRCD(DataGridView)
        Dim columnCount As Integer = DataGridView.Columns.Count
        Dim ЕlArray(columnCount - 1) As ElectricalParameters
        '
        ' Инициализация на всеки елемент от масива
        '
        For i As Integer = 2 To columnCount - 2
            If DataGridView.Rows(0).Cells(i).Value = "pазединител" Then Continue For
            ЕlArray(i).TKryg = DataGridView.Columns(i).Name
            ЕlArray(i).Power = DataGridView.Rows(za6t + 3).Cells(i).Value
            ЕlArray(i).CircuitType = IIf(DataGridView.Rows(za6t + 6).Cells(i).Value = "L1,L2,L3", "L1,L2,L3", "")
            ЕlArray(i).Phases = DataGridView.Rows(6).Cells(i).Value
            ЕlArray(i).RCD = DataGridView.Rows(defkt).Cells(i).Value
            ЕlArray(i).Current = DataGridView.Rows(1).Cells(i).Value
            ЕlArray(i).Columns = i
            ЕlArray(i).Bus = DataGridView.Rows(za6t + 10).Cells(i).Value
        Next
        '
        ' Сортиране на масива по полето "ток"
        '
        Array.Sort(ЕlArray, Function(x, y) y.Current.CompareTo(x.Current))
        '
        ' Балансира фазите по ток
        '
        Dim arr(2) As Double    ' Сума по токове по фази
        Dim faze As Integer = 0 ' Инициализираме индекса на най-малкото число
        For i As Integer = 0 To UBound(ЕlArray)
            '
            ' намира най-ненатоварената фаза
            '
            For ind As Integer = 0 To 2
                If arr(ind) < arr(faze) Then
                    faze = ind
                End If
            Next
            '
            ' записва номера на фазата в масива
            '
            ЕlArray(i).CircuitType = IIf(ЕlArray(i).CircuitType = "L1,L2,L3",
                                         ЕlArray(i).CircuitType,
                                         "L" & (faze + 1).ToString)
            '
            ' Сумираме натоварването на фазите
            '
            arr(0) = 0
            arr(1) = 0
            arr(2) = 0
            For j = 0 To UBound(ЕlArray)
                Select Case ЕlArray(j).CircuitType
                    Case "L1,L2,L3"
                        arr(0) += ЕlArray(j).Current
                        arr(1) += ЕlArray(j).Current
                        arr(2) += ЕlArray(j).Current
                    Case "L1"
                        arr(0) += ЕlArray(j).Current
                    Case "L2"
                        arr(1) += ЕlArray(j).Current
                    Case "L3"
                        arr(2) += ЕlArray(j).Current
                End Select
            Next
        Next
        '
        ' Сумира токовете по фази за консуматори общо за таблото
        '
        arr(0) = 0
        arr(1) = 0
        arr(2) = 0
        For index As Integer = 0 To UBound(ЕlArray)
            If ЕlArray(index).Bus Then Continue For
            Select Case ЕlArray(index).CircuitType
                Case "L1,L2,L3"
                    arr(0) += ЕlArray(index).Current
                    arr(1) += ЕlArray(index).Current
                    arr(2) += ЕlArray(index).Current
                Case "L1"
                    arr(0) += ЕlArray(index).Current
                Case "L2"
                    arr(1) += ЕlArray(index).Current
                Case "L3"
                    arr(2) += ЕlArray(index).Current
            End Select
        Next
        ' намира най-натоварената фаза
        For ind As Integer = 0 To 2
            If arr(ind) > arr(faze) Then
                faze = ind
            End If
        Next
        With DataGridView
            .Rows(1).Cells("ОБЩО").Value = arr(faze)
            .Rows(2).Cells("ОБЩО").Value = "iSW"
            .Rows(3).Cells(("ОБЩО")).Value = calc_ISW("3p", Ikryg:=arr(faze))
            .Rows(6).Cells("ОБЩО").Value = "3p"
            .Rows(za6t + 5).Cells(("ОБЩО")).Value = calc_cable_Cu(.Rows(3).Cells(("ОБЩО")).Value, "3p")
            .Rows(za6t + 6).Cells(("ОБЩО")).Value = "L1,L2,L3"
            .Rows(defkt).Cells("ОБЩО").Value = "Ток фази"
            .Rows(defkt + 1).Cells("ОБЩО").Value = "L1->" & arr(0).ToString("0.00") & "A"
            .Rows(defkt + 2).Cells("ОБЩО").Value = "L2->" & arr(1).ToString("0.00") & "A"
            .Rows(defkt + 3).Cells("ОБЩО").Value = "L3->" & arr(2).ToString("0.00") & "A"
        End With
        '
        '  Записва данните в DataGridView
        '
        For i As Integer = 0 To UBound(ЕlArray)
            If DataGridView.Rows(0).Cells(i).Value = "pазединител" Then Continue For
            If ЕlArray(i).Columns = 0 Then Continue For
            DataGridView.Rows(za6t + 6).Cells(ЕlArray(i).Columns).Value = ЕlArray(i).CircuitType
        Next
        '
        ' Разпределя RCD след разпределение на фазите
        '
        SetRCD(DataGridView)
        '
        ' Подрежда фазите -> на първи токов кръг да е L1
        '
        If InStr(DataGridView.Rows(za6t + 6).Cells(2).Value, "L1") = 0 Then
            Dim Pom_Fase As String = DataGridView.Rows(za6t + 6).Cells(2).Value
            For i As Integer = 2 To columnCount - 2
                If DataGridView.Rows(0).Cells(i).Value = "pазединител" Then Continue For
                If DataGridView.Rows(za6t + 6).Cells(i).Value = "L1,L2,L3" Then Continue For
                If DataGridView.Rows(za6t + 6).Cells(i).Value = Pom_Fase Then
                    DataGridView.Rows(za6t + 6).Cells(i).Value = "L1"
                    Continue For
                End If
                If DataGridView.Rows(za6t + 6).Cells(i).Value = "L1" Then DataGridView.Rows(za6t + 6).Cells(i).Value = Pom_Fase
            Next
        End If
        Calculate_Faze(DataGridView)
        '
        ' Сумира токовете по фази за консуматори които са на различна шина
        '
        If DataGridView.Columns.Contains("Разединител") Then Calculate_Bus(DataGridView)
        Calculate_Faze(DataGridView)
    End Sub
    ''' <summary>
    ''' Изчислява натоварването на шина (Bus) в DataGridView, като сумира токовете по фази и 
    ''' попълва стойности за разединител, ток, тип и кабел.
    ''' </summary>
    ''' <param name="DataGridView">DataGridView, съдържащ параметрите на електрическите кръгове.</param>
    Private Sub Calculate_Bus(DataGridView As Windows.Forms.DataGridView)
        ' Брой колони в DataGridView
        Dim columnCount As Integer = DataGridView.Columns.Count
        ' Масив от структури ElectricalParameters, един елемент за всяка колона
        Dim ЕlArray(columnCount - 1) As ElectricalParameters
        ' Инициализация на всеки елемент от масива ЕlArray:
        ' Попълване на свойства като ток на кръга, мощност, вид на кръга, фази, RCD, текущ ток и шина.
        ' Пропуска колони, които съдържат "pазединител" в първия ред.
        For i As Integer = 2 To columnCount - 2
            If DataGridView.Rows(0).Cells(i).Value = "pазединител" Then Continue For

            ЕlArray(i).TKryg = DataGridView.Columns(i).Name                       ' Име на кръга
            ЕlArray(i).Power = DataGridView.Rows(za6t + 3).Cells(i).Value         ' Мощност на товара
            ЕlArray(i).CircuitType = DataGridView.Rows(za6t + 6).Cells(i).Value   ' Вид на кръга (L1, L2, L3, L1,L2,L3)
            ЕlArray(i).Phases = DataGridView.Rows(6).Cells(i).Value               ' Брой фази
            ЕlArray(i).RCD = DataGridView.Rows(defkt).Cells(i).Value              ' RCD/защита
            ЕlArray(i).Current = DataGridView.Rows(1).Cells(i).Value              ' Ток на кръга
            ЕlArray(i).Columns = i                                               ' Индекс на колоната
            ЕlArray(i).Bus = DataGridView.Rows(za6t + 10).Cells(i).Value          ' Свързаност към шината (True/False)
        Next

        ' Масив за сумиране на токовете по фази
        Dim arr(2) As Double
        Dim faze As Integer = 0 ' Индекс на фазата с най-голям ток
        arr(0) = 0
        arr(1) = 0
        arr(2) = 0

        ' Обхождане на всички елементи в ЕlArray
        ' Сумиране на токовете по фази според типа на кръга:
        ' - "L1,L2,L3" -> добавя текущ ток към всички фази
        ' - "L1" -> добавя към L1, и т.н.
        ' Пропуска кръгове, които не са свързани към шината (Bus=False)
        For index As Integer = 0 To UBound(ЕlArray)
            If Not ЕlArray(index).Bus Then Continue For

            Select Case ЕlArray(index).CircuitType
                Case "L1,L2,L3"
                    arr(0) += ЕlArray(index).Current
                    arr(1) += ЕlArray(index).Current
                    arr(2) += ЕlArray(index).Current
                Case "L1"
                    arr(0) += ЕlArray(index).Current
                Case "L2"
                    arr(1) += ЕlArray(index).Current
                Case "L3"
                    arr(2) += ЕlArray(index).Current
            End Select
        Next
        ' Намира най-натоварената фаза (с най-голям ток)
        For ind As Integer = 0 To 2
            If arr(ind) > arr(faze) Then
                faze = ind
            End If
        Next
        ' Попълване на стойности в колоната "Разединител":
        ' - Ток на най-натоварената фаза
        ' - Тип на защита "iSW"
        ' - Изчислен ток за iSW
        ' - Попълване на стойности за общо, кабел, ток по фази
        With DataGridView
            .Rows(1).Cells("Разединител").Value = arr(faze)
            .Rows(2).Cells("Разединител").Value = "iSW"
            .Rows(3).Cells("Разединител").Value = calc_ISW(.Rows(6).Cells("ОБЩО").Value, Ikryg:= .Rows(1).Cells("Разединител").Value)
            .Rows(6).Cells("Разединител").Value = .Rows(6).Cells("ОБЩО").Value
            .Rows(za6t + 5).Cells("Разединител").Value = calc_cable_Cu(.Rows(3).Cells("Разединител").Value, "3p")
            .Rows(za6t + 6).Cells("Разединител").Value = .Rows(za6t + 6).Cells("ОБЩО").Value
            .Rows(defkt).Cells("Разединител").Value = "Ток фази"
            .Rows(defkt + 1).Cells("Разединител").Value = "L1->" & arr(0).ToString("0.00") & "A"
            .Rows(defkt + 2).Cells("Разединител").Value = "L2->" & arr(1).ToString("0.00") & "A"
            .Rows(defkt + 3).Cells("Разединител").Value = "L3->" & arr(2).ToString("0.00") & "A"
        End With
    End Sub
    ''' <summary>
    ''' Изчислява натоварването по фази (L1, L2, L3) за всички кръгове в DataGridView
    ''' и попълва сумарния ток в ред "ОБЩО", както и токовете по отделните фази.
    ''' </summary>
    ''' <param name="DataGridView">DataGridView, съдържащ параметрите на електрическите кръгове.</param>
    Private Sub Calculate_Faze(DataGridView As Windows.Forms.DataGridView)
        ' Брой колони в DataGridView
        Dim brColums As Integer = DataGridView.Columns.Count
        ' Масив за сумиране на токовете по фази: arr(0)=L1, arr(1)=L2, arr(2)=L3
        Dim arr(2) As Double
        ' Индекс на фазата с най-голям ток, начална стойност 1 (L2)
        Dim faze As Integer = 1
        ' Обхождане на колоните, започвайки от 2 до предпоследната
        ' Пропуска колони, които съдържат "pазединител" в първия ред
        For index As Integer = 2 To brColums - 2
            If DataGridView.Rows(0).Cells(index).Value = "pазединител" Then Continue For
            ' В зависимост от вида на кръга, добавя токовете към съответните фази
            Select Case DataGridView.Rows(za6t + 6).Cells(index).Value
                Case "L1,L2,L3"
                    arr(0) += DataGridView.Rows(1).Cells(index).Value
                    arr(1) += DataGridView.Rows(1).Cells(index).Value
                    arr(2) += DataGridView.Rows(1).Cells(index).Value
                Case "L1"
                    arr(0) += DataGridView.Rows(1).Cells(index).Value
                Case "L2"
                    arr(1) += DataGridView.Rows(1).Cells(index).Value
                Case "L3"
                    arr(2) += DataGridView.Rows(1).Cells(index).Value
            End Select
        Next
        ' <summary>
        ' Намира фазата с най-голям ток
        ' </summary>
        For ind As Integer = 0 To 2
            If arr(ind) > arr(faze) Then
                faze = ind
            End If
        Next
        ' Попълване на сумарния ток на най-натоварената фаза
        ' в колоната "ОБЩО" и токовете по отделните фази
        DataGridView.Rows(1).Cells("ОБЩО").Value = arr(faze).ToString("0.00") & "A"
        DataGridView.Rows(defkt + 1).Cells("ОБЩО").Value = "L1->" & arr(0).ToString("0.00") & "A"
        DataGridView.Rows(defkt + 2).Cells("ОБЩО").Value = "L2->" & arr(1).ToString("0.00") & "A"
        DataGridView.Rows(defkt + 3).Cells("ОБЩО").Value = "L3->" & arr(2).ToString("0.00") & "A"
    End Sub
    Private Function Calculate_GV2(Мощност As String, ' Мощност на двигателя
                                     Връща As Integer   ' Какво да върне функцита                                     
                                     ) As String
        ' Връща 1 - Тип на защитата
        '       2 - Pдвиг(400V)
        '       3 - Настройки
        Dim P_GV As Double = Val(Мощност)
        Dim GV_Тип As String = ""
        Dim GV_Pдвиг As String = ""
        Dim GV_Наст As String = ""
        Dim In_GV As Double = 1.2 * calc_Inom(P_GV, "3p", True)
        Select Case In_GV
            Case 0.1 To 0.16
                GV_Тип = "GV2-ME"
                GV_Pдвиг = "<0,06kW"
                GV_Наст = "0.1-0.16A"
            Case 0.16 To 0.25
                GV_Тип = "GV2-ME"
                GV_Pдвиг = "0,06kW"
                GV_Наст = "0.16-0.25A"
            Case 0.25 To 0.4
                GV_Тип = "GV2-ME"
                GV_Pдвиг = "0,09kW"
                GV_Наст = "0.25-0.40A"
            Case 0.4 To 0.63
                GV_Тип = "GV2-ME"
                GV_Pдвиг = "0,12kW"
                GV_Наст = "0.4-0.63A"
            Case 0.63 To 1
                GV_Тип = "GV2-ME"
                GV_Pдвиг = "0,25kW"
                GV_Наст = "0.63-1.0A"
            Case 1 To 1.6
                GV_Тип = "GV2-ME"
                GV_Pдвиг = "0,37kW"
                GV_Наст = "1.0-1.6A"
            Case 1.6 To 2.5
                GV_Тип = "GV2-ME"
                GV_Pдвиг = "0,75kW"
                GV_Наст = "1.6-2.5A"
            Case 2.5 To 4.0
                GV_Тип = "GV2-ME"
                GV_Наст = "2.5-4.0A"
                GV_Pдвиг = "1,1kW"
            Case 4.0 To 6.3
                GV_Тип = "GV2-ME"
                GV_Наст = "4.0-6.3A"
                GV_Pдвиг = "2,2kW"
            Case 6.0 To 10
                GV_Тип = "GV2-ME"
                GV_Наст = "6.0-10A"
                GV_Pдвиг = "3,0kW"
                GV_Pдвиг = "4,0kW"
            Case 9.0 To 14
                GV_Тип = "GV2-ME"
                GV_Наст = "9.0-14A"
                GV_Pдвиг = "5,5kW"
            Case 13.0 To 18
                GV_Тип = "GV2-ME"
                GV_Наст = "13-18A"
                GV_Pдвиг = "7,5kW"
            Case 17.0 To 23
                GV_Тип = "GV2-ME"
                GV_Наст = "17-23A"
                GV_Pдвиг = "9,0kW"
            Case 17.0 To 25
                GV_Тип = "GV3-P"
                GV_Наст = "17-25A"
                GV_Pдвиг = "11,0kW"
            Case 23.0 To 32
                GV_Тип = "GV3-P"
                GV_Наст = "23-32A"
                GV_Pдвиг = "15,0kW"
            Case 30.0 To 40
                GV_Тип = "GV3-P"
                GV_Наст = "30-40A"
                GV_Pдвиг = "18,5kW"
            Case 20 To 50
                GV_Тип = "GV4P"
                GV_Pдвиг = "11-22kW"
                GV_Наст = "20-50A"
            Case 40 To 80
                GV_Тип = "GV4P"
                GV_Pдвиг = "22-37kW"
                GV_Наст = "40-80A"
            Case 65 To 115
                GV_Тип = "GV4P"
                GV_Pдвиг = "37-55kW"
                GV_Наст = "65-115A"
        End Select
        Select Case Връща
            Case 1
                Return GV_Тип
            Case 2
                Return GV_Pдвиг
            Case 3
                Return GV_Наст
        End Select
    End Function
    ''' <summary>
    ''' Чертане на рамката на електрическото табло (външни контури + вертикални линии)
    ''' </summary>
    Private Sub DrawPanelFrame(ptBasePoint As Point3d,
                               brColums As Integer,
                               twoBus As Boolean
                               )
        ' =====================================================
        ' 1️⃣ Определяне на основни координати и масив за линии
        ' =====================================================
        ' prX: крайната X координата за хоризонталните линии
        Dim prX As Double = ptBasePoint.X + widthText + widthTextDim + (brColums - IIf(twoBus, 3, 2)) * widthColom
        Dim prY As Double ' ще се използва по-късно за текста
        ' Массив за начални и крайни точки на линиите
        Dim arrPoint(17, 1) As Point3d
        ' --- Хоризонтални линии на редовете на таблицата ---
        arrPoint(0, 0) = New Point3d(ptBasePoint.X, ptBasePoint.Y, 0)
        arrPoint(0, 1) = New Point3d(prX, ptBasePoint.Y, 0)
        arrPoint(1, 0) = New Point3d(ptBasePoint.X, ptBasePoint.Y + 3 * heightRow, 0)
        arrPoint(1, 1) = New Point3d(prX, ptBasePoint.Y + 3 * heightRow, 0)
        arrPoint(2, 0) = New Point3d(ptBasePoint.X, ptBasePoint.Y + 4 * heightRow, 0)
        arrPoint(2, 1) = New Point3d(prX, ptBasePoint.Y + 4 * heightRow, 0)
        arrPoint(3, 0) = New Point3d(ptBasePoint.X, ptBasePoint.Y + 5 * heightRow, 0)
        arrPoint(3, 1) = New Point3d(prX, ptBasePoint.Y + 5 * heightRow, 0)
        arrPoint(4, 0) = New Point3d(ptBasePoint.X, ptBasePoint.Y + 6 * heightRow, 0)
        arrPoint(4, 1) = New Point3d(prX, ptBasePoint.Y + 6 * heightRow, 0)
        arrPoint(5, 0) = New Point3d(ptBasePoint.X, ptBasePoint.Y + 7 * heightRow, 0)
        arrPoint(5, 1) = New Point3d(prX, ptBasePoint.Y + 7 * heightRow, 0)
        arrPoint(6, 0) = New Point3d(ptBasePoint.X, ptBasePoint.Y + 8 * heightRow, 0)
        arrPoint(6, 1) = New Point3d(prX, ptBasePoint.Y + 8 * heightRow, 0)
        arrPoint(7, 0) = New Point3d(ptBasePoint.X, ptBasePoint.Y + 9 * heightRow, 0)
        arrPoint(7, 1) = New Point3d(prX, ptBasePoint.Y + 9 * heightRow, 0)
        arrPoint(8, 0) = New Point3d(ptBasePoint.X, ptBasePoint.Y + 10 * heightRow, 0)
        arrPoint(8, 1) = New Point3d(prX, ptBasePoint.Y + 10 * heightRow, 0)
        ' --- Вертикални линии на таблицата ---
        arrPoint(9, 0) = New Point3d(ptBasePoint.X, ptBasePoint.Y, 0)
        arrPoint(9, 1) = New Point3d(ptBasePoint.X, ptBasePoint.Y + 10 * heightRow, 0)
        arrPoint(10, 0) = New Point3d(ptBasePoint.X + widthText, ptBasePoint.Y, 0)
        arrPoint(10, 1) = New Point3d(ptBasePoint.X + widthText, ptBasePoint.Y + 10 * heightRow, 0)
        arrPoint(11, 0) = New Point3d(ptBasePoint.X + widthText + widthTextDim, ptBasePoint.Y, 0)
        arrPoint(11, 1) = New Point3d(ptBasePoint.X + widthText + widthTextDim, ptBasePoint.Y + 10 * heightRow, 0)
        ' --- Допълнителни линии за рамка на блока и шина ---
        arrPoint(12, 0) = New Point3d(ptBasePoint.X + widthText, ptBasePoint.Y + 10 * heightRow + lengthProw, 0)
        arrPoint(12, 1) = New Point3d(ptBasePoint.X + widthText, ptBasePoint.Y + 10 * heightRow + lengthProw + widthTablo, 0)
        arrPoint(13, 0) = New Point3d(ptBasePoint.X + widthText, ptBasePoint.Y + 10 * heightRow + lengthProw, 0)
        arrPoint(13, 1) = New Point3d(prX, ptBasePoint.Y + 10 * heightRow + lengthProw, 0)
        arrPoint(14, 0) = New Point3d(prX, ptBasePoint.Y + 10 * heightRow + lengthProw, 0)
        arrPoint(14, 1) = New Point3d(prX, ptBasePoint.Y + 10 * heightRow + lengthProw + widthTablo, 0)
        arrPoint(15, 0) = New Point3d(ptBasePoint.X + widthText, ptBasePoint.Y + 10 * heightRow + lengthProw + widthTablo, 0)
        arrPoint(15, 1) = New Point3d(prX, ptBasePoint.Y + 10 * heightRow + lengthProw + widthTablo, 0)
        ' --- Линии за маркировка (червен кръст) ---
        arrPoint(16, 0) = New Point3d(ptBasePoint.X + widthText + 18, ptBasePoint.Y + 10 * heightRow + lengthProw + widthTablo, 0)
        arrPoint(16, 1) = New Point3d(ptBasePoint.X + widthText + 18, ptBasePoint.Y + 10 * heightRow + lengthProw + widthTablo - 36, 0)
        arrPoint(17, 0) = New Point3d(ptBasePoint.X + widthText, ptBasePoint.Y + 10 * heightRow + lengthProw + widthTablo - 18, 0)
        arrPoint(17, 1) = New Point3d(ptBasePoint.X + widthText + 36, ptBasePoint.Y + 10 * heightRow + lengthProw + widthTablo - 18, 0)
        ' =====================================================
        ' 2️⃣ Чертеж на линиите
        ' =====================================================
        For index = 0 To UBound(arrPoint)
            Select Case index
                Case < 12
                    ' Чертае основната таблица (хоризонтални и вертикални линии)
                    cu.DrowLine(arrPoint(index, 0), arrPoint(index, 1), "EL_ТАБЛА", LineWeight.ByLayer, "ByLayer")
                Case 12 To 15
                    ' Чертае рамка за блока с информация за шина
                    cu.DrowLine(arrPoint(index, 0), arrPoint(index, 1), "EL_ТАБЛА", LineWeight.ByLayer, "CENTER")
                Case 16, 17
                    ' Чертае червен кръст за маркиране на таблото
                    cu.DrowLine(arrPoint(index, 0), arrPoint(index, 1), "Defpoints", LineWeight.ByLayer, "ByLayer", 1)
            End Select
        Next
        ' --- Вертикални линии за колоните на таблицата ---
        For index = 1 To brColums - IIf(twoBus, 3, 2)
            Dim X As Double = ptBasePoint.X + widthText + widthTextDim + index * widthColom
            cu.DrowLine(New Point3d(X, ptBasePoint.Y, 0),
                    New Point3d(X, ptBasePoint.Y + 10 * heightRow, 0),
                    "EL_ТАБЛА",
                    LineWeight.ByLayer,
                    "ByLayer")
        Next
        ' =====================================================
        ' 3️⃣ Вмъкване на текстове (първа колона)
        ' =====================================================
        prX = ptBasePoint.X + padingText
        prY = ptBasePoint.Y + (heightRow - heightText) / 2
        cu.InsertText("Токов кръг", New Point3d(prX, prY + 9 * heightRow, 0),
                  "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
        cu.InsertText("Брой лампи", New Point3d(prX, prY + 8 * heightRow, 0),
                  "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
        cu.InsertText("Брой контакти", New Point3d(prX, prY + 7 * heightRow, 0),
                  "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
        cu.InsertText("Инстал. мощност", New Point3d(prX, prY + 6 * heightRow, 0),
                  "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
        cu.InsertText("Тип кабел", New Point3d(prX, prY + 5 * heightRow, 0),
                  "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
        cu.InsertText("Сечение кабел", New Point3d(prX, prY + 4 * heightRow, 0),
                  "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
        cu.InsertText("Фаза", New Point3d(prX, prY + 3 * heightRow, 0),
                  "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
        cu.InsertText("Консуматор", New Point3d(prX, prY + 2 * heightRow, 0),
                  "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
        ' =====================================================
        ' 4️⃣ Вмъкване на текстове (втора колона)
        ' =====================================================
        prX = prX + widthText
        prY = ptBasePoint.Y + (heightRow - heightText) / 2
        cu.InsertText("№", New Point3d(prX, prY + 9 * heightRow, 0),
                  "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
        cu.InsertText("бр.", New Point3d(prX, prY + 8 * heightRow, 0),
                  "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
        cu.InsertText("бр.", New Point3d(prX, prY + 7 * heightRow, 0),
                  "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
        cu.InsertText("kW", New Point3d(prX, prY + 6 * heightRow, 0),
                  "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
        cu.InsertText("---", New Point3d(prX, prY + 5 * heightRow, 0),
                  "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
        cu.InsertText("mm²", New Point3d(prX, prY + 4 * heightRow, 0),
                  "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
        cu.InsertText("---", New Point3d(prX, prY + 3 * heightRow, 0),
                  "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
        cu.InsertText("---", New Point3d(prX, prY + 2 * heightRow, 0),
                  "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
    End Sub
    Private Sub CreateTablo(DataGridView As Windows.Forms.DataGridView)
        '' Get the current database and start the Transaction Manager
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ptBasePointRes As PromptPointResult
        Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
        form_AS_tablo.Visible = vbFalse

        pPtOpts.Message = vbLf & "Изберете долен ляв ъгъл на таблото: "
        ptBasePointRes = acDoc.Editor.GetPoint(pPtOpts)

        Dim brColums As Integer = DataGridView.Columns.Count
        If ptBasePointRes.Status = PromptStatus.Cancel Then Exit Sub
        Dim ptBasePoint As Point3d = ptBasePointRes.Value

        Dim TabloName As String = Mid(DataGridView.Name, 1, Len(DataGridView.Name) - 3)

        Dim blkRecId_D As ObjectId = ObjectId.Null
        Dim blkRecId_L As ObjectId = ObjectId.Null
        Dim index_D As Integer = 0

        Dim Faza_Tablo As Boolean = False
        Dim brTokKrygoweNa6ina As Integer = 0

        Dim index As Integer
        Dim twoBus As Boolean = False

        If DataGridView.Columns.Contains("Разединител") Then twoBus = True

        Try

            ' Чертане на рамката на електрическото табло (външни контури + вертикални линии)
            DrawPanelFrame(ptBasePoint, brColums, twoBus)

            Dim X As Double = 0
            Dim LeftPoint As Point3d = New Point3d(0, 0, 0)
            Dim RightPoint As Point3d = New Point3d(0, 0, 0)
            Dim Coloni_Broj As Integer = 2
            Dim Coloni_Mo6t As Integer = 0

            For index = 2 To brColums - 1
                If DataGridView.Columns(index).HeaderText.ToString = "Мощностен" Then
                    Coloni_Mo6t = index
                    Continue For
                End If
                Dim ТоковКръг As String = DataGridView.Columns(index).HeaderText.ToString
                Dim brLap As String = IIf(DataGridView.Rows(za6t + 1).Cells(index).Value = 0,
                                          "----",
                                          DataGridView.Rows(za6t + 1).Cells(index).Value.ToString)
                Dim brKontakt As String = IIf(DataGridView.Rows(za6t + 2).Cells(index).Value = 0,
                                          "----",
                                          DataGridView.Rows(za6t + 2).Cells(index).Value.ToString)
                Dim Мощност As String = CDbl(DataGridView.Rows(za6t + 3).Cells(index).Value).ToString("0.000")
                Dim typeKabel As String = DataGridView.Rows(za6t + 4).Cells(index).Value.ToString
                Dim sechKabel As String = DataGridView.Rows(za6t + 5).Cells(index).Value.ToString
                Dim Faza As String = DataGridView.Rows(za6t + 6).Cells(index).Value.ToString
                Dim Broj_N As String
                If DataGridView.Rows(defkt).Cells(index).Value = Nothing Then
                    Broj_N = ""
                Else
                    Broj_N = DataGridView.Rows(defkt).Cells(index).Value.ToString
                End If

                Dim konsuator1 As String = DataGridView.Rows(za6t + 7).Cells(index).Value
                Dim konsuator2 As String = DataGridView.Rows(za6t + 8).Cells(index).Value
                '
                ' Попълва текстове в таблицата
                '
                X = ptBasePoint.X + widthText + widthTextDim + (Coloni_Broj - 2) * widthColom + widthColom / 2
                cu.InsertText(ТоковКръг,
                              New Point3d(X + padingText,
                                          ptBasePoint.Y + 9 * heightRow + heightRow / 2, 0),
                              "EL__DIM", heightText, TextHorizontalMode.TextMid, TextVerticalMode.TextBase)
                cu.InsertText(brLap,
                              New Point3d(X + padingText,
                                          ptBasePoint.Y + 8 * heightRow + heightRow / 2, 0),
                              "EL__DIM", heightText, TextHorizontalMode.TextMid, TextVerticalMode.TextBase)
                cu.InsertText(brKontakt,
                              New Point3d(X + padingText,
                                          ptBasePoint.Y + 7 * heightRow + heightRow / 2, 0),
                              "EL__DIM", heightText, TextHorizontalMode.TextMid, TextVerticalMode.TextBase)
                cu.InsertText(Мощност,
                              New Point3d(X + padingText,
                                          ptBasePoint.Y + 6 * heightRow + heightRow / 2, 0),
                              "EL__DIM", heightText, TextHorizontalMode.TextMid, TextVerticalMode.TextBase)
                cu.InsertText(typeKabel,
                              New Point3d(X + padingText,
                                          ptBasePoint.Y + 5 * heightRow + heightRow / 2, 0),
                              "EL__DIM", heightText, TextHorizontalMode.TextMid, TextVerticalMode.TextBase)
                cu.InsertText(sechKabel,
                              New Point3d(X + padingText,
                                          ptBasePoint.Y + 4 * heightRow + heightRow / 2, 0),
                              "EL__DIM", heightText, TextHorizontalMode.TextMid, TextVerticalMode.TextBase)
                cu.InsertText(Faza,
                              New Point3d(X + padingText,
                                          ptBasePoint.Y + 3 * heightRow + heightRow / 2, 0),
                              "EL__DIM", heightText, TextHorizontalMode.TextMid, TextVerticalMode.TextBase)
                cu.InsertText(konsuator1,
                              New Point3d(X - widthColom / 2 + padingText,
                                          ptBasePoint.Y + 2 * heightRow + (heightRow - heightText) / 2, 0),
                              "EL__DIM", 12, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
                cu.InsertText(konsuator2,
                              New Point3d(X - widthColom / 2 + padingText,
                                          ptBasePoint.Y + 1 * heightRow + (heightRow - heightText) / 2, 0),
                              "EL__DIM", 12, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)

                '
                ' Поставя знак за прекъсвач
                '
                If Faza <> "L1" And Not Faza_Tablo And Faza <> Nothing Then
                    Faza_Tablo = True
                End If
                Dim doc As Document = Application.DocumentManager.MdiActiveDocument
                Dim blkRecId As ObjectId = ObjectId.Null

                If index = brColums - 1 Then Exit For ' Ако е последната колона излиза от цикъла
                '
                ' Изчислява позицията на блока за вмъкване
                '
                X = ptBasePoint.X + widthText + widthTextDim - widthColom / 2 + (Coloni_Broj - 1) * widthColom

                Select Case konsuator1
                    Case "Бойлер", "Проточен"
                        blkRecId = cu.InsertBlock("s_dpnn_vigi_circ_break",
                                           New Point3d(X, ptBasePoint.Y + Y_Шина, 0),
                                           "EL_ТАБЛА",
                                           New Scale3d(5, 5, 5)
                                           )
                        lengthProwBlock = 132.5
                        LeftPoint = New Point3d(0, 0, 0)
                        RightPoint = New Point3d(0, 0, 0)
                    Case "Контакти"
                        If DataGridView.Rows(defkt).Cells(index - 1).Value <> DataGridView.Rows(defkt).Cells(index).Value And
                           DataGridView.Rows(defkt).Cells(index + 1).Value <> DataGridView.Rows(defkt).Cells(index).Value Then

                            LeftPoint = New Point3d(0, 0, 0)
                            RightPoint = New Point3d(0, 0, 0)

                            lengthProwBlock = 132.5
                            blkRecId = cu.InsertBlock("s_dpnn_vigi_circ_break",
                                           New Point3d(X, ptBasePoint.Y + Y_Шина, 0),
                                           "EL_ТАБЛА",
                                           New Scale3d(5, 5, 5)
                                           )
                        End If
                        If DataGridView.Rows(defkt).Cells(index - 1).Value = DataGridView.Rows(defkt).Cells(index).Value And
                           DataGridView.Rows(defkt).Cells(index + 1).Value = DataGridView.Rows(defkt).Cells(index).Value Then

                            lengthProwBlock = 27.5
                            blkRecId = cu.InsertBlock("s_c60_circ_break",
                                           New Point3d(X, ptBasePoint.Y + Y_Шина - 117.5, 0),
                                           "EL_ТАБЛА",
                                           New Scale3d(5, 5, 5)
                                           )
                        End If
                        If DataGridView.Rows(defkt).Cells(index - 1).Value <> DataGridView.Rows(defkt).Cells(index).Value And
                           DataGridView.Rows(defkt).Cells(index + 1).Value = DataGridView.Rows(defkt).Cells(index).Value Then

                            lengthProwBlock = 27.5
                            blkRecId = cu.InsertBlock("s_c60_circ_break",
                                           New Point3d(X, ptBasePoint.Y + Y_Шина - 117.5, 0),
                                           "EL_ТАБЛА",
                                           New Scale3d(5, 5, 5)
                                           )

                            LeftPoint = New Point3d(X - widthColom / 4, ptBasePoint.Y + Y_Шина - 117.5, 0)
                            RightPoint = New Point3d(0, 0, 0)

                        End If
                        If DataGridView.Rows(defkt).Cells(index - 1).Value = DataGridView.Rows(defkt).Cells(index).Value And
                           DataGridView.Rows(defkt).Cells(index + 1).Value <> DataGridView.Rows(defkt).Cells(index).Value Then

                            blkRecId = cu.InsertBlock("s_c60_circ_break",
                                           New Point3d(X, ptBasePoint.Y + Y_Шина - 117.5, 0),
                                           "EL_ТАБЛА",
                                           New Scale3d(5, 5, 5)
                                           )

                            RightPoint = New Point3d(X + widthColom / 4, ptBasePoint.Y + Y_Шина - 117.5, 0)

                            blkRecId_D = cu.InsertBlock("s_id_res_circ_break",
                                           New Point3d((LeftPoint.X + RightPoint.X) / 2, ptBasePoint.Y + Y_Шина, 0),
                                           "EL_ТАБЛА",
                                           New Scale3d(5, 5, 5)
                                           )

                            lengthProwBlock = 27.5

                        End If
                    Case Else
                        ' Проверява дали трявба да е Моторна защита
                        If DataGridView.Rows(za6t + 9).Cells(index).Value = "Моторна защита" Then
                            blkRecId = cu.InsertBlock("s_GV2",
                                                      New Point3d(X, ptBasePoint.Y + Y_Шина, 0),
                                                      "EL_ТАБЛА",
                                                      New Scale3d(5, 5, 5)
                                                      )
                        Else
                            blkRecId = cu.InsertBlock("s_c60_circ_break",
                                                      New Point3d(X, ptBasePoint.Y + Y_Шина, 0),
                                                      "EL_ТАБЛА",
                                                      New Scale3d(5, 5, 5)
                                                      )
                        End If
                        lengthProwBlock = 145
                        LeftPoint = New Point3d(0, 0, 0)
                        RightPoint = New Point3d(0, 0, 0)
                End Select
                '
                ' Попълва параметрите на блока с автоматичния прекъсвач или ДЗТ
                '
                If Not IsNothing(blkRecId.Database) Then
                    Using trans As Transaction = doc.TransactionManager.StartTransaction()
                        Dim acBlkTbl As BlockTable = trans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                        Dim acBlkRef As BlockReference =
                        DirectCast(trans.GetObject(blkRecId, OpenMode.ForWrite), BlockReference)

                        Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                        For Each objID As ObjectId In attCol
                            Dim dbObj As DBObject = trans.GetObject(objID, OpenMode.ForWrite)
                            Dim acAttRef As AttributeReference = dbObj

                            If DataGridView.Rows(defkt + 1).Cells(index).Value = "EZ9 RCBO" Then
                                If acAttRef.Tag = "1" Then acAttRef.TextString = DataGridView.Rows(defkt + 2).Cells(index).Value ' АЦ
                                If acAttRef.Tag = "2" Then acAttRef.TextString = DataGridView.Rows(defkt + 5).Cells(index).Value ' 2п
                                If acAttRef.Tag = "3" Then acAttRef.TextString = "C"
                                If acAttRef.Tag = "4" Then acAttRef.TextString = "20A"
                                If acAttRef.Tag = "5" Then acAttRef.TextString = DataGridView.Rows(defkt + 4).Cells(index).Value ' 30мА
                                If acAttRef.Tag = "SHORTNAME" Then acAttRef.TextString = DataGridView.Rows(defkt + 1).Cells(index).Value
                            Else
                                If acAttRef.Tag = "1" Then acAttRef.TextString = ""
                                If acAttRef.Tag = "2" Then acAttRef.TextString = DataGridView.Rows(5).Cells(index).Value
                                If acAttRef.Tag = "3" Then acAttRef.TextString = DataGridView.Rows(6).Cells(index).Value
                                If acAttRef.Tag = "4" Then acAttRef.TextString = DataGridView.Rows(3).Cells(index).Value
                                If acAttRef.Tag = "SHORTNAME" Then acAttRef.TextString = DataGridView.Rows(2).Cells(index).Value
                            End If

                            If DataGridView.Rows(za6t + 9).Cells(index).Value = "Моторна защита" Then

                                If acAttRef.Tag = "1" Then acAttRef.TextString = Calculate_GV2(
                                                                                 DataGridView.Rows(za6t + 3).Cells(index).Value, 3)
                                If acAttRef.Tag = "2" Then acAttRef.TextString = "3P"
                                If acAttRef.Tag = "3" Then acAttRef.TextString = Calculate_GV2(
                                                                                 DataGridView.Rows(za6t + 3).Cells(index).Value, 2)
                                If acAttRef.Tag = "4" Then acAttRef.TextString = ""
                                If acAttRef.Tag = "5" Then acAttRef.TextString = ""
                                If acAttRef.Tag = "SHORTNAME" Then acAttRef.TextString = Calculate_GV2(
                                                                                 DataGridView.Rows(za6t + 3).Cells(index).Value, 1)
                            End If

                            If acAttRef.Tag = "DESIGNATION" Then acAttRef.TextString = ""
                            If acAttRef.Tag = "REFNB" Then acAttRef.TextString = TabloName

                        Next
                        trans.Commit()
                        blkRecId = ObjectId.Null
                    End Using
                End If
                '
                ' Попълва атрибутии на ДЗТ 
                '
                If Not IsNothing(blkRecId_D.Database) Then
                    Using trans As Transaction = doc.TransactionManager.StartTransaction()
                        Dim acBlkTbl As BlockTable = trans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                        Dim acBlkRef As BlockReference =
                                DirectCast(trans.GetObject(blkRecId_D, OpenMode.ForWrite), BlockReference)

                        Dim attCol As AttributeCollection = acBlkRef.AttributeCollection

                        For Each objID As ObjectId In attCol
                            Dim dbObj As DBObject = trans.GetObject(objID, OpenMode.ForWrite)
                            Dim acAttRef As AttributeReference = dbObj
                            If acAttRef.Tag = "DESIGNATION" Then acAttRef.TextString = ""
                            If acAttRef.Tag = "1" Then acAttRef.TextString = DataGridView.Rows(defkt + 2).Cells(index).Value ' АЦ
                            If acAttRef.Tag = "2" Then acAttRef.TextString = DataGridView.Rows(defkt + 5).Cells(index).Value ' 2п
                            If acAttRef.Tag = "3" Then acAttRef.TextString = DataGridView.Rows(defkt + 3).Cells(index).Value
                            If acAttRef.Tag = "4" Then acAttRef.TextString = "Мигновена"
                            If acAttRef.Tag = "5" Then acAttRef.TextString = DataGridView.Rows(defkt + 4).Cells(index).Value ' 30мА
                            If acAttRef.Tag = "REFNB" Then acAttRef.TextString = TabloName
                            If acAttRef.Tag = "SHORTNAME" Then acAttRef.TextString = DataGridView.Rows(defkt + 1).Cells(index).Value
                        Next
                        trans.Commit()
                        blkRecId_D = ObjectId.Null
                    End Using
                End If
                '
                ' Поставя знак за управление
                '
                Dim blkRecId_U As ObjectId = ObjectId.Null
                Dim Rows9Value As Boolean = False
                Dim kvadrat As Boolean = False
                Dim Rows9Name As String = ""
                Dim str_1 As String = ""
                Dim str_2 As String = ""
                Dim str_3 As String = ""
                Dim str_4 As String = ""
                Dim str_5 As String = ""
                Dim str_SHORTNAME As String = ""
                Select Case DataGridView.Rows(za6t + 9).Cells(index).Value
                    Case "Импулсно реле"
                        Rows9Name = "s_tl"
                        Rows9Value = True
                        kvadrat = True
                        str_1 = ""
                        str_2 = "1p"
                        str_3 = "16A"
                        str_4 = "220VAC"
                        str_5 = ""
                        str_SHORTNAME = "iTL"
                    Case "Контактор"
                        Rows9Name = "s_ct_cont_no"
                        Rows9Value = True
                        kvadrat = True
                        str_1 = "1НО"
                        str_2 = ""
                        str_3 = "16A"
                        str_4 = "220VAC"
                        str_5 = ""
                        str_SHORTNAME = "iCT"
                    Case "Стълбищен автомат"
                        Rows9Name = "s_min"
                        Rows9Value = True
                        kvadrat = True
                        Rows9Name = "s_ct_cont_no"
                        Rows9Value = True
                        kvadrat = True
                        str_1 = ""
                        str_2 = ""
                        str_3 = "0.5-20min"
                        str_4 = "16A"
                        str_5 = ""
                        str_SHORTNAME = "MINp"
                    Case "Фото реле"
                        Rows9Name = "s_switch_light_sens"
                        Rows9Value = True
                        kvadrat = False
                        str_1 = ""
                        str_2 = "2-100 Lx"
                        str_3 = "10 А"
                        str_4 = ""
                        str_5 = ""
                        str_SHORTNAME = "IC100"
                    Case "Моторна защита"
                        Rows9Name = "s_tesys_cont_no"
                        Rows9Value = True
                        kvadrat = True
                        str_1 = "3NO"
                        str_2 = "1NO+1NC"
                        str_3 = "9A"
                        str_4 = "230VAC"
                        str_5 = "Винтова"
                        str_SHORTNAME = "LC1D"
                    Case "Няма", Nothing
                End Select
                If Rows9Value Then
                    blkRecId_U = cu.InsertBlock(Rows9Name,
                                           New Point3d(X, ptBasePoint.Y + Y_Шина - 135, 0),
                                           "EL_ТАБЛА",
                                           New Scale3d(5, 5, 5)
                                           )
                    If kvadrat Then
                        cu.InsertBlock("Ключ_квадрат",
                                           New Point3d(X - 32, ptBasePoint.Y + Y_Шина - 330, 0),
                                           "EL_ТАБЛА",
                                           New Scale3d(1, 1, 1))
                        cu.DrowLine(New Point3d(X - 32,
                                                ptBasePoint.Y + Y_Шина - 197,
                                                0),
                                    New Point3d(X - 32,
                                                ptBasePoint.Y + Y_Шина - 305,
                                                0),
                                    "EL_ТАБЛА",
                                    LineWeight.ByLayer,
                                    "ByLayer")
                    End If

                    Using trans As Transaction = doc.TransactionManager.StartTransaction()
                        Dim acBlkTbl As BlockTable = trans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                        Dim acBlkRef As BlockReference =
                        DirectCast(trans.GetObject(blkRecId_U, OpenMode.ForWrite), BlockReference)

                        Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                        For Each objID As ObjectId In attCol
                            Dim dbObj As DBObject = trans.GetObject(objID, OpenMode.ForWrite)
                            Dim acAttRef As AttributeReference = dbObj
                            If acAttRef.Tag = "1" Then acAttRef.TextString = str_1
                            If acAttRef.Tag = "2" Then acAttRef.TextString = str_2
                            If acAttRef.Tag = "3" Then acAttRef.TextString = str_3
                            If acAttRef.Tag = "4" Then acAttRef.TextString = str_4
                            If acAttRef.Tag = "5" Then acAttRef.TextString = str_5
                            If acAttRef.Tag = "SHORTNAME" Then acAttRef.TextString = str_SHORTNAME
                            If acAttRef.Tag = "DESIGNATION" Then acAttRef.TextString = ""
                            If acAttRef.Tag = "REFNB" Then acAttRef.TextString = TabloName
                        Next
                        trans.Commit()
                        blkRecId_U = ObjectId.Null
                    End Using
                End If
                '
                ' Поставя Хоризонтална линия и надпис над нея
                ' над шината за дефектно токовите
                '
                If RightPoint.X <> 0 Then
                    cu.DrowLine(LeftPoint,
                                RightPoint,
                                "EL_ТАБЛА",
                                LineWeight.LineWeight070,
                                "ByLayer")

                    Dim text As String
                    text = DataGridView.Rows(za6t + 6).Cells(index).Value.ToString
                    text += ","
                    text += DataGridView.Rows(defkt).Cells(index).Value
                    text += ",PE"
                    cu.InsertText(text, New Point3d(LeftPoint.X, LeftPoint.Y + 2 * padingText, 0),
                                  "EL__DIM", heightText,
                                  TextHorizontalMode.TextLeft,
                                  TextVerticalMode.TextBase)
                End If
                '
                ' Поставя вертикалната линия под блока за прекъсвача
                '
                X = ptBasePoint.X + widthText + widthTextDim + (Coloni_Broj - 1) * widthColom
                cu.DrowLine(New Point3d(X - widthColom / 2,
                                        ptBasePoint.Y + 10 * heightRow,
                                        0),
                            New Point3d(X - widthColom / 2,
                                        ptBasePoint.Y + 10 * heightRow + lengthProw + lengthProwBlock +
                                        IIf(Rows9Value, -95, 0),
                                        0),
                            "EL_ТАБЛА",
                            LineWeight.ByLayer,
                            "ByLayer")
                '
                ' Проверва дали е сложена отметка за различна шина
                '
                If (DataGridView.Rows(za6t + 10).Cells(index).Value) Then
                    brTokKrygoweNa6ina += 1
                End If
                Coloni_Broj += 1
            Next

            '-------------------------------------------------------------------------------------------------------------------------------------------------
            X = ptBasePoint.X + widthText + widthTextDim
            '
            ' Поставя името на таблото
            '
            cu.InsertText(TabloName,
                          New Point3d(X + (brColums - IIf(twoBus, 4, 3)) * widthColom,
                                      ptBasePoint.Y + Y_Шина + 95,
                                      0),
                          "EL__DIM", heightText + 5, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            '
            ' Чертае линита за шина
            '
            cu.InsertText(IIf(Faza_Tablo, "L1,L2,L3,N,PE", "L,N,PE"),
                          New Point3d(X, ptBasePoint.Y + Y_Шина + 2 * padingText, 0),
                          "EL__DIM",
                          heightText,
                          TextHorizontalMode.TextLeft,
                          TextVerticalMode.TextBase)
            '
            '   Шина 1 е общата, но когато са дви шини X_6ina1 - е втората
            '   Тъпо е но не ми се мисли как да сменяам двете шини
            '
            '
            ' Изчислява положрнието на двата разединителя
            '
            Dim X_6ina1 As Double = (X + brTokKrygoweNa6ina * widthColom + 30 + X + (brColums - 3) * widthColom) / 2
            Dim X_6ina2 As Double = (X + X + brTokKrygoweNa6ina * widthColom - 30) / 2
            '
            '   Проверява броя на шините
            '
            If Not twoBus Then
                cu.DrowLine(New Point3d(X, ptBasePoint.Y + Y_Шина, 0),
                            New Point3d(X + (brColums - 3) * widthColom, ptBasePoint.Y + Y_Шина, 0),
                            "EL_ТАБЛА",
                            LineWeight.LineWeight070,
                            "ByLayer")
            Else
                cu.DrowLine(New Point3d(X, ptBasePoint.Y + Y_Шина, 0),
                            New Point3d(X + brTokKrygoweNa6ina * widthColom - 30, ptBasePoint.Y + Y_Шина, 0),
                            "EL_ТАБЛА",
                            LineWeight.LineWeight070,
                            "ByLayer")

                cu.InsertText(IIf(Faza_Tablo, "L1,L2,L3,N,PE", "L,N,PE"),
                          New Point3d(X + brTokKrygoweNa6ina * widthColom + 30, ptBasePoint.Y + Y_Шина + 2 * padingText, 0),
                          "EL__DIM",
                          heightText,
                          TextHorizontalMode.TextLeft,
                          TextVerticalMode.TextBase)

                cu.DrowLine(New Point3d(X + brTokKrygoweNa6ina * widthColom + 30, ptBasePoint.Y + Y_Шина, 0),
                            New Point3d(X + (brColums - 4) * widthColom, ptBasePoint.Y + Y_Шина, 0),
                            "EL_ТАБЛА",
                            LineWeight.LineWeight070,
                            "ByLayer")
                '
                ' Поставя линия между двата товарови прекъсвача
                '
                cu.DrowLine(New Point3d(X_6ina1, ptBasePoint.Y + Y_Шина + 95, 0),
                            New Point3d(X_6ina2, ptBasePoint.Y + Y_Шина + 95, 0),
                            "EL_ТАБЛА",
                            LineWeight.ByLayer,
                            "ByLayer")
                '
                ' Поставя товаров прекъсвач на шината, ако са две шини на първата шина
                '
                Dim doc As Document = Application.DocumentManager.MdiActiveDocument
                Dim blkRecId As ObjectId = ObjectId.Null
                blkRecId = cu.InsertBlock("s_i_ng_switch_disconn",
                               New Point3d(X_6ina2,
                                           ptBasePoint.Y + Y_Шина + 95,
                                           0),
                               "EL_ТАБЛА",
                               New Scale3d(5, 5, 5)
                               )
                Using trans As Transaction = doc.TransactionManager.StartTransaction()
                    Dim acBlkTbl As BlockTable = trans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                    Dim acBlkRef As BlockReference =
                                    DirectCast(trans.GetObject(blkRecId, OpenMode.ForWrite), BlockReference)

                    Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                    For Each objID As ObjectId In attCol
                        Dim dbObj As DBObject = trans.GetObject(objID, OpenMode.ForWrite)
                        Dim acAttRef As AttributeReference = dbObj
                        If acAttRef.Tag = "DESIGNATION" Then acAttRef.TextString = ""
                        If acAttRef.Tag = "1" Then acAttRef.TextString = ""
                        If acAttRef.Tag = "2" Then acAttRef.TextString = DataGridView.Rows(6).Cells("Разединител").Value
                        If acAttRef.Tag = "3" Then acAttRef.TextString = DataGridView.Rows(3).Cells("Разединител").Value + "А"
                        If acAttRef.Tag = "4" Then acAttRef.TextString = ""
                        If acAttRef.Tag = "5" Then acAttRef.TextString = ""
                        If acAttRef.Tag = "REFNB" Then acAttRef.TextString = TabloName
                        If acAttRef.Tag = "SHORTNAME" Then acAttRef.TextString = DataGridView.Rows(2).Cells("Разединител").Value
                    Next
                    trans.Commit()
                End Using
            End If
            '
            ' Поставя товаров прекъсвач на шината, ако са две шини на втората шина
            '
            Dim doc_ As Document = Application.DocumentManager.MdiActiveDocument
            Dim blkRecId_ As ObjectId = ObjectId.Null
            If DataGridView.Rows(7).Cells("ОБЩО").Value = "Палец" Then
                blkRecId_ = cu.InsertBlock("s_i_ng_switch_disconn",
                                           New Point3d(X_6ina1, ptBasePoint.Y + Y_Шина + 95, 0),
                                           "EL_ТАБЛА",
                                           New Scale3d(5, 5, 5)
                                           )
            Else
                blkRecId_ = cu.InsertBlock("s_ns100_motor_fixed",
                           New Point3d(X_6ina1, ptBasePoint.Y + Y_Шина + 95 + 40, 0),
                           "EL_ТАБЛА",
                           New Scale3d(5, 5, 5)
                           )
                CreateTablo_PV(blkRecId_)
            End If

            Using trans As Transaction = doc_.TransactionManager.StartTransaction()
                Dim acBlkTbl As BlockTable = trans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim acBlkRef As BlockReference =
                                DirectCast(trans.GetObject(blkRecId_, OpenMode.ForWrite), BlockReference)

                Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                For Each objID As ObjectId In attCol
                    Dim dbObj As DBObject = trans.GetObject(objID, OpenMode.ForWrite)
                    Dim acAttRef As AttributeReference = dbObj
                    If acAttRef.Tag = "DESIGNATION" Then acAttRef.TextString = ""
                    If acAttRef.Tag = "1" Then acAttRef.TextString = ""
                    If acAttRef.Tag = "2" Then acAttRef.TextString = DataGridView.Rows(6).Cells("ОБЩО").Value
                    If acAttRef.Tag = "3" Then acAttRef.TextString = DataGridView.Rows(3).Cells("ОБЩО").Value + "А"
                    If acAttRef.Tag = "4" Then acAttRef.TextString = ""
                    If acAttRef.Tag = "5" Then acAttRef.TextString = ""
                    If acAttRef.Tag = "REFNB" Then acAttRef.TextString = TabloName
                    If acAttRef.Tag = "SHORTNAME" Then acAttRef.TextString = DataGridView.Rows(2).Cells("ОБЩО").Value
                Next
                trans.Commit()
            End Using
            '
            ' Поставя линия над товаров прекъсвач 
            '
            cu.DrowLine(New Point3d(X_6ina1, ptBasePoint.Y + Y_Шина + 95, 0),
                        New Point3d(X_6ina1, ptBasePoint.Y + Y_Шина + 220, 0),
                        "EL_ТАБЛА",
                        LineWeight.ByLayer,
                        "ByLayer")
            '
            ' Поставя надпис над товаров прекъсвач 
            '
            blkRecId_ = cu.InsertBlock("Кабел",
                           New Point3d(X_6ina1,
                                       ptBasePoint.Y + Y_Шина + 95 + 90,
                                       0),
                           "EL__DIM",
                           New Scale3d(1, 1, 1)
                           )
            Using trans As Transaction = doc_.TransactionManager.StartTransaction()
                Dim acBlkTbl As BlockTable = trans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim acBlkRef As BlockReference =
                                DirectCast(trans.GetObject(blkRecId_, OpenMode.ForWrite), BlockReference)
                Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection
                For Each prop As DynamicBlockReferenceProperty In props
                    'This Is where you change states based on input
                    If prop.PropertyName = "Visibility" Then prop.Value = "Точка"
                Next

                Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                For Each objID As ObjectId In attCol
                    Dim dbObj As DBObject = trans.GetObject(objID, OpenMode.ForWrite)
                    Dim acAttRef As AttributeReference = dbObj
                    If acAttRef.Tag = "NA4IN_0" Then acAttRef.TextString =
                            DataGridView.Rows(za6t + 4).Cells(brColums - 1).Value + " " +
                            DataGridView.Rows(za6t + 5).Cells(brColums - 1).Value + "mm²"
                    If acAttRef.Tag = "NA4IN_1" Then acAttRef.TextString =
                            IIf(TabloName = "Гл.Р.Т." Or TabloName = "ГлРТ", "от електромерно табло", "от табло Гл.Р.Т.")
                    If acAttRef.Tag = "NA4IN_2" Then acAttRef.TextString = ""
                    If acAttRef.Tag = "NA4IN_3" Then acAttRef.TextString = ""
                    If acAttRef.Tag = "NA4IN_4" Then acAttRef.TextString = ""
                    If acAttRef.Tag = "NA4IN_5" Then acAttRef.TextString = ""
                    If acAttRef.Tag = "NA4IN_6" Then acAttRef.TextString = ""
                    If acAttRef.Tag = "NA4IN_7" Then acAttRef.TextString = ""
                    If acAttRef.Tag = "NA4IN_8" Then acAttRef.TextString = ""
                    If acAttRef.Tag = "NA4IN_9" Then acAttRef.TextString = ""
                    If acAttRef.Tag = "NA4IN_10" Then acAttRef.TextString = ""
                Next
                trans.Commit()
            End Using
            cu.EditDynamicBlockReferenceKabel(blkRecId_)
            '
            ' Поставя товаров прекъсвач на първата шина когато са две шини, ако е една шина не прави нищо
            '
            If twoBus Then
                blkRecId_ = cu.InsertBlock("s_i_ng_switch_disconn",
                                           New Point3d(X_6ina2, ptBasePoint.Y + Y_Шина + 95, 0),
                                           "EL_ТАБЛА",
                                           New Scale3d(5, 5, 5))

                cu.DrowLine(New Point3d(X_6ina1, ptBasePoint.Y + Y_Шина + 95, 0),
                            New Point3d(X_6ina2, ptBasePoint.Y + Y_Шина + 95, 0),
                            "EL_ТАБЛА",
                            LineWeight.ByLayer,
                            "ByLayer")

                Using trans As Transaction = doc_.TransactionManager.StartTransaction()
                    Dim acBlkTbl As BlockTable = trans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                    Dim acBlkRef As BlockReference =
                                    DirectCast(trans.GetObject(blkRecId_, OpenMode.ForWrite), BlockReference)

                    Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                    For Each objID As ObjectId In attCol
                        Dim dbObj As DBObject = trans.GetObject(objID, OpenMode.ForWrite)
                        Dim acAttRef As AttributeReference = dbObj
                        If acAttRef.Tag = "DESIGNATION" Then acAttRef.TextString = ""
                        If acAttRef.Tag = "1" Then acAttRef.TextString = ""
                        If acAttRef.Tag = "2" Then acAttRef.TextString = DataGridView.Rows(6).Cells(Coloni_Mo6t).Value ' 2п
                        If acAttRef.Tag = "3" Then acAttRef.TextString = DataGridView.Rows(3).Cells(Coloni_Mo6t).Value + "А"
                        If acAttRef.Tag = "4" Then acAttRef.TextString = ""
                        If acAttRef.Tag = "5" Then acAttRef.TextString = ""
                        If acAttRef.Tag = "REFNB" Then acAttRef.TextString = TabloName
                        If acAttRef.Tag = "SHORTNAME" Then acAttRef.TextString = DataGridView.Rows(2).Cells(Coloni_Mo6t).Value
                    Next
                    trans.Commit()
                End Using

            End If
            '
            ' Поставя знак за заземление
            '
            If TabloName = "Гл.Р.Т." Or TabloName = "ГлРТ" Then

                cu.DrowLine(New Point3d(X, ptBasePoint.Y + Y_Шина, 0),
                        New Point3d(X - widthColom, ptBasePoint.Y + Y_Шина, 0),
                        "EL_ТАБЛА",
                        LineWeight.ByLayer,
                        "ByLayer")

                cu.InsertText("R<30Ω",
                          New Point3d(X - widthColom, ptBasePoint.Y + Y_Шина + 2 * padingText, 0),
                          "EL__DIM",
                          heightText,
                          TextHorizontalMode.TextLeft,
                          TextVerticalMode.TextBase)

                Dim blkRecId = cu.InsertBlock("Заземление",
                               New Point3d(X - widthColom, ptBasePoint.Y + Y_Шина, 0),
                               "EL_ТАБЛА",
                               New Scale3d(0.21, 0.21, 0.21)
                               )

                Dim doc As Document = Application.DocumentManager.MdiActiveDocument
                Using trans As Transaction = doc.TransactionManager.StartTransaction()
                    Dim acBlkTbl As BlockTable = trans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                    Dim acBlkRef As BlockReference =
                        DirectCast(trans.GetObject(blkRecId, OpenMode.ForWrite), BlockReference)

                    Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection

                    For Each prop As DynamicBlockReferenceProperty In props
                        'This Is where you change states based on input
                        If prop.PropertyName = "Visibility" Then prop.Value = "Заземител-БЕЗ контролна клема"
                        If prop.PropertyName = "Position1 X" Then prop.Value = -10.0
                        If prop.PropertyName = "Position1 Y" Then prop.Value = -80.0
                        If prop.PropertyName = "Angle1" Then prop.Value = 0.0
                    Next
                    Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                    For Each objID As ObjectId In attCol
                        Dim dbObj As DBObject = trans.GetObject(objID, OpenMode.ForWrite)
                        Dim acAttRef As AttributeReference = dbObj
                        If acAttRef.Tag = "ТАБЛО" Then acAttRef.TextString = "2к"
                    Next

                    trans.Commit()
                End Using
                cu.InsertText("PE",
                          New Point3d(X - widthColom + 3 * padingText, ptBasePoint.Y + Y_Шина - heightText - padingText, 0),
                          "EL__DIM",
                          heightText,
                          TextHorizontalMode.TextLeft,
                          TextVerticalMode.TextBase)
            End If

            Dim Zabelevka As String = "1. Таблото да се изпълни в съответствие с изискванията на БДС EN 61439-1."
            Zabelevka += vbCrLf & "2. Aпаратурата и тоководящите части да бъдат монтирани зад защитни капаци. "
            Zabelevka += vbCrLf & "3. Достъпа до палците и ръкохватките на комутационните апарати се осигурява посредством отвори в защитните капаци."
            Zabelevka += vbCrLf & "4. Апаратурата е избрана по каталога на SCHNEIDER ELECTRIC."
            Zabelevka += vbCrLf & "5. Изборът на автоматичните прекъсвачи е съобразен с токовете на к.с., спазени са изискванията за селективност."
            Zabelevka += vbCrLf & "6. При замяна типа на апаратурата да се преизчисли схемата."
            Zabelevka += vbCrLf & "7. При замяна номиналният ток на апаратурата да се преизчисли сечението на кабелите."

            cu.InsertMText("ЗАБЕЛЕЖКИ:",
                                     New Point3d(ptBasePoint.X,
                                                 ptBasePoint.Y - 20, 0),
                                     "EL__DIM", 10)
            cu.InsertMText(Zabelevka,
                                     New Point3d(ptBasePoint.X + 30,
                                                 ptBasePoint.Y - 20 - heightRow, 0),
                                     "EL__DIM", 10)
            Dim pol As Integer = 0
            For index = 2 To brColums - 1
                pol += Val(DataGridView.Rows(6).Cells(index).Value)
                pol += Val(DataGridView.Rows(defkt + 5).Cells(index).Value)
                pol += IIf(DataGridView.Rows(za6t + 9).Cells(index).Value = "" Or
                           DataGridView.Rows(za6t + 9).Cells(index).Value = "Няма", 0, 1)
            Next

            DataGridView.Rows(defkt + 5).Cells("ОБЩО").Value = "Полюси->" & pol.ToString(0)

            cu.InsertMText("Полюси -> " & pol.ToString(0),
                         New Point3d(ptBasePoint.X + 160,
                                     ptBasePoint.Y + 900, 0),
                         "Defpoints", 20, 1)
        Catch ex As Exception
            MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
        End Try
    End Sub
    Private Sub CreateTablo_PV(blkRecId_ As ObjectId)
        Dim doc_ As Document = Application.DocumentManager.MdiActiveDocument
        Using trans As Transaction = doc_.TransactionManager.StartTransaction()
            Dim blkRef As BlockReference = CType(trans.GetObject(blkRecId_, OpenMode.ForRead), BlockReference)
            ' Вземи точката на вмъкване
            Dim insertionPoint As Point3d = blkRef.Position

            ' Запиши координатите в променливи
            Dim Block_X As Double = insertionPoint.X
            Dim Block_Y As Double = insertionPoint.Y
            Dim Block_Z As Double = insertionPoint.Z

            Dim blkRecId = cu.InsertBlock("s_c60_circ_break",
                                       New Point3d(insertionPoint.X + 120, insertionPoint.Y + 60, 0),
                                       "EL_ТАБЛА",
                                       New Scale3d(5, 5, 5)
                                       )
            Dim acBlkRef As BlockReference =
                                DirectCast(trans.GetObject(blkRecId, OpenMode.ForWrite), BlockReference)
            Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
            For Each objID As ObjectId In attCol
                Dim dbObj As DBObject = trans.GetObject(objID, OpenMode.ForWrite)
                Dim acAttRef As AttributeReference = dbObj
                If acAttRef.Tag = "DESIGNATION" Then acAttRef.TextString = ""
                If acAttRef.Tag = "1" Then acAttRef.TextString = ""
                If acAttRef.Tag = "2" Then acAttRef.TextString = "C"
                If acAttRef.Tag = "3" Then acAttRef.TextString = "3p"
                If acAttRef.Tag = "4" Then acAttRef.TextString = "20A"
                If acAttRef.Tag = "5" Then acAttRef.TextString = ""
                If acAttRef.Tag = "REFNB" Then acAttRef.TextString = "TabloName"
                If acAttRef.Tag = "SHORTNAME" Then acAttRef.TextString = "EZ9 MCB"
            Next


            Dim arrPoint(10, 1) As Point3d
            arrPoint(0, 0) = New Point3d(insertionPoint.X - 36.5, insertionPoint.Y - 43, 0)
            arrPoint(0, 1) = New Point3d(insertionPoint.X - 36.5, insertionPoint.Y + 80, 0)
            arrPoint(1, 0) = New Point3d(insertionPoint.X - 36.5, insertionPoint.Y + 80, 0)
            arrPoint(1, 1) = New Point3d(insertionPoint.X + 240, insertionPoint.Y + 80, 0)
            arrPoint(2, 0) = New Point3d(insertionPoint.X + 0, insertionPoint.Y + 60, 0)
            arrPoint(2, 1) = New Point3d(insertionPoint.X + 120, insertionPoint.Y + 60, 0)
            arrPoint(3, 0) = New Point3d(insertionPoint.X + 120, insertionPoint.Y - 75, 0)
            arrPoint(3, 1) = New Point3d(insertionPoint.X + 200, insertionPoint.Y - 75, 0)
            arrPoint(4, 0) = New Point3d(insertionPoint.X + 120, insertionPoint.Y - 95, 0)
            arrPoint(4, 1) = New Point3d(insertionPoint.X + 220, insertionPoint.Y - 95, 0)
            arrPoint(5, 0) = New Point3d(insertionPoint.X + 120, insertionPoint.Y - 115, 0)
            arrPoint(5, 1) = New Point3d(insertionPoint.X + 240, insertionPoint.Y - 115, 0)
            arrPoint(6, 0) = New Point3d(insertionPoint.X + 120, insertionPoint.Y - 75, 0)
            arrPoint(6, 1) = New Point3d(insertionPoint.X + 120, insertionPoint.Y - 115, 0)

            arrPoint(7, 0) = New Point3d(insertionPoint.X + 340, insertionPoint.Y - 100, 0)
            arrPoint(7, 1) = New Point3d(insertionPoint.X + 420, insertionPoint.Y - 100, 0)
            arrPoint(8, 0) = New Point3d(insertionPoint.X + 340, insertionPoint.Y - 85, 0)
            arrPoint(8, 1) = New Point3d(insertionPoint.X + 400, insertionPoint.Y - 85, 0)
            arrPoint(9, 0) = New Point3d(insertionPoint.X + 340, insertionPoint.Y - 70, 0)
            arrPoint(9, 1) = New Point3d(insertionPoint.X + 380, insertionPoint.Y - 70, 0)
            arrPoint(10, 0) = New Point3d(insertionPoint.X + 340, insertionPoint.Y - 55, 0)
            arrPoint(10, 1) = New Point3d(insertionPoint.X + 360, insertionPoint.Y - 55, 0)







            For index = 0 To UBound(arrPoint)

                cu.DrowLine(arrPoint(index, 0),
                            arrPoint(index, 1),
                            layerLine:="EL_ТАБЛА",
                            WeightLine:=LineWeight.ByLayer,
                            LineTipe:="ByLayer",
                            LineColor:=90)

            Next

            'Public Function DrowLine(pt1 As Point3d,                    '   Начална точка на линията
            '         pt2 As Point3d,                        '   Крайна точка на линията
            '         layerLine As String,                   '   Слой в който се чертае линията
            '         WeightLine As LineWeight,              '   Дебелина на линия
            '         LineTipe As String,                    '   Тип на линията 
            '         Optional LineColor As Integer = 256    '   Цвят на линията (по подразбиране е 256, което означава "по слой")
            '         ) As ObjectId

            trans.Commit()
        End Using
    End Sub
    Public Sub FindChildren(ByVal parentCtrl As Windows.Forms.Control, ByRef children As List(Of Object))
        '
        ' Рекурсивна функция за извличане на контролите в дълбочина
        '
        If parentCtrl.HasChildren Then
            For Each ctrl As Windows.Forms.Control In parentCtrl.Controls
                If TypeOf ctrl Is Windows.Forms.ToolStrip Then
                    Dim toll As Windows.Forms.ToolStrip
                    toll = ctrl
                    For Each item In toll.Items
                        children.Add(item)
                    Next item
                End If
                children.Add(ctrl)
                Call FindChildren(ctrl, children)
            Next ctrl
        End If
    End Sub
    <CommandMethod("readBlockTable")>
    Public Shared Sub CmdReadBlockTable()
        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
        Dim peo As PromptEntityOptions = New PromptEntityOptions("Select a dynamic block reference: ")
        peo.SetRejectMessage("Only block reference")
        peo.AddAllowedClass(GetType(BlockReference), False)
        Dim per As PromptEntityResult = ed.GetEntity(peo)
        If per.Status <> PromptStatus.OK Then Return
        Dim blockRefId As ObjectId = per.ObjectId
        Dim db As Database = Application.DocumentManager.MdiActiveDocument.Database

        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Dim blockRef As BlockReference = TryCast(trans.GetObject(blockRefId, OpenMode.ForRead), BlockReference)
            If Not blockRef.IsDynamicBlock Then Return
            Dim blockDef As BlockTableRecord = TryCast(trans.GetObject(blockRef.DynamicBlockTableRecord, OpenMode.ForRead), BlockTableRecord)
            If blockDef.ExtensionDictionary.IsNull Then Return
            Dim extDic As DBDictionary = TryCast(trans.GetObject(blockDef.ExtensionDictionary, OpenMode.ForRead), DBDictionary)
            Dim graph As EvalGraph = TryCast(trans.GetObject(extDic.GetAt("ACAD_ENHANCEDBLOCK"), OpenMode.ForRead), EvalGraph)
            Dim nodeIds As Integer() = graph.GetAllNodes()

            For Each nodeId As UInteger In nodeIds
                Dim node As DBObject = graph.GetNode(nodeId, OpenMode.ForRead, trans)
                If Not (TypeOf node Is BlockPropertiesTable) Then Continue For
                Dim table As BlockPropertiesTable = TryCast(node, BlockPropertiesTable)
                Dim columns As Integer = table.Columns.Count
                Dim currentRow As Integer = 0
                For Each row As BlockPropertiesTableRow In table.Rows
                    ed.WriteMessage(vbLf & "[{0}]:" & vbTab, currentRow)
                    For currentColumn As Integer = 0 To columns - 1
                        Dim columnValue As TypedValue() = row(currentColumn).AsArray()
                        For Each tpVal As TypedValue In columnValue
                            ed.WriteMessage("{0}; ", tpVal.Value)
                        Next
                        ed.WriteMessage("|")
                    Next
                    currentRow += 1
                Next
            Next
        End Using
    End Sub
    Function GetXData(appName As String,        ' Име на таблицата
                      SelSet As ObjectId        ' Блок от който се чете информацията                                        
                      ) As ResultBuffer
        Dim acCurDb As Autodesk.AutoCAD.DatabaseServices.Database
        acCurDb = Application.DocumentManager.MdiActiveDocument.Database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim Xdata As ResultBuffer = New ResultBuffer
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acEnt As Entity = acTrans.GetObject(SelSet, OpenMode.ForRead)
            Xdata = acEnt.GetXDataForApplication(appName)
            acTrans.Abort()
        End Using
        Return Xdata
    End Function
    <CommandMethod("AttachXDataToSelectionSetObjects")>
    Public Sub AttachXDataToSelectionSetObjects()
        ' Get the current database and start a transaction
        Dim acCurDb As Autodesk.AutoCAD.DatabaseServices.Database
        acCurDb = Application.DocumentManager.MdiActiveDocument.Database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim appName As String = "MY_APP"
        Dim xdataStr As String = "This is some xdata"

        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            ' Request objects to be selected in the drawing area
            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
            ' If the prompt status is OK, objects were selected
            If acSSPrompt.Status = PromptStatus.OK Then
                ' Open the Registered Applications table for read
                Dim acRegAppTbl As RegAppTable
                acRegAppTbl = acTrans.GetObject(acCurDb.RegAppTableId, OpenMode.ForRead)
                ' Check to see if the Registered Applications table record for the custom app exists
                Dim acRegAppTblRec As RegAppTableRecord
                If acRegAppTbl.Has(appName) = False Then
                    acRegAppTblRec = New RegAppTableRecord
                    acRegAppTblRec.Name = appName
                    acTrans.GetObject(acCurDb.RegAppTableId, OpenMode.ForWrite)
                    acRegAppTbl.Add(acRegAppTblRec)
                    acTrans.AddNewlyCreatedDBObject(acRegAppTblRec, True)
                End If
                ' Define the Xdata to add to each selected object
                Using rb As New ResultBuffer
                    rb.Add(New TypedValue(DxfCode.ExtendedDataRegAppName, appName))
                    rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, xdataStr))
                    Dim acSSet As SelectionSet = acSSPrompt.Value
                    ' Step through the objects in the selection set
                    For Each acSSObj As SelectedObject In acSSet
                        ' Open the selected object for write
                        Dim acEnt As Entity = acTrans.GetObject(acSSObj.ObjectId,
                                                            OpenMode.ForWrite)
                        rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, acEnt.ObjectId.ToString))
                        ' Append the extended data to each object
                        acEnt.XData = rb
                    Next

                End Using
            End If
            ' Save the new object to the database
            acTrans.Commit()
            ' Dispose of the transaction
        End Using

    End Sub
    <CommandMethod("ViewXData")>
    Public Sub ViewXData()
        ' Get the current database and start a transaction
        Dim acCurDb As Autodesk.AutoCAD.DatabaseServices.Database
        acCurDb = Application.DocumentManager.MdiActiveDocument.Database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim appName As String = "EWG_TABLO"
        Dim msgstr As String = ""
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            ' Request objects to be selected in the drawing area
            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
            ' If the prompt status is OK, objects were selected
            If acSSPrompt.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt.Value
                ' Step through the objects in the selection set
                For Each acSSObj As SelectedObject In acSSet
                    ' Open the selected object for read
                    Dim acEnt As Entity = acTrans.GetObject(acSSObj.ObjectId,
                                                            OpenMode.ForRead)
                    ' Get the extended data attached to each object for MY_APP
                    Dim rb As ResultBuffer = acEnt.GetXDataForApplication(appName)
                    ' Make sure the Xdata is not empty
                    If Not IsNothing(rb) Then
                        ' Get the values in the xdata
                        For Each typeVal As TypedValue In rb
                            msgstr = msgstr & vbCrLf & typeVal.TypeCode.ToString() & ":" & typeVal.Value
                        Next
                    Else
                        msgstr = "NONE"
                    End If
                    ' Display the values returned
                    MsgBox(appName & " xdata on " & VarType(acEnt).ToString() & ":" & vbCrLf & msgstr)

                    msgstr = ""
                Next
            End If
            ' Ends the transaction and ensures any changes made are ignored
            acTrans.Abort()
            ' Dispose of the transaction
        End Using
    End Sub
    <CommandMethod("DeleteXData")>
    Public Shared Sub DeleteXData_Method()
        Dim db As Database = HostApplicationServices.WorkingDatabase
        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor

        Try
            Dim prEntRes As PromptEntityResult = ed.GetEntity(vbLf & "Select an Entity to delete XDATA from")

            If prEntRes.Status = PromptStatus.OK Then

                Using tr As Transaction = db.TransactionManager.StartTransaction()
                    Dim ent As Entity = CType(tr.GetObject(prEntRes.ObjectId, OpenMode.ForRead), Entity)

                    If ent.XData Is Nothing Then
                        ed.WriteMessage(vbLf & "No XData in the entity.")
                        Return
                    End If

                    Dim appName As String = ed.GetString(vbLf &
                                                         "AppName for the XData to delete (Enter for ALL)").StringResult

                    'If String.IsNullOrEmpty(appName) And
                    'MsgBox("Woult you really like to delete all XData associated with the entity?",
                    '                    "Delete All XData", MsgBoxStyle.YesNo) = MsgBoxResult.No Then Return

                    If String.IsNullOrEmpty(appName) Then

                        For Each tv As TypedValue In ent.XData.AsArray().Where(Function(e) e.TypeCode = 1001)
                            ent.UpgradeOpen()
                            ent.XData = New ResultBuffer(New TypedValue(1001, tv.Value))
                            ent.DowngradeOpen()
                        Next

                        ed.WriteMessage(vbLf & "All XData have been deleted.")
                    Else
                        Dim rb As ResultBuffer = ent.GetXDataForApplication(appName)

                        If rb IsNot Nothing Then
                            ent.UpgradeOpen()
                            ent.XData = New ResultBuffer(New TypedValue(1001, appName))
                            ent.DowngradeOpen()
                            ed.WriteMessage(vbLf & "XData with AppName {0} have been deleted.", appName)
                        Else
                            ed.WriteMessage(vbLf & "It doesn't have XData with AppName {0}.", appName)
                        End If
                    End If

                    tr.Commit()
                End Using
            End If

        Catch ex As System.Exception
            ed.WriteMessage(ex.ToString())
        End Try
    End Sub
    ' Избира сечение на кабел
    Private Function calc_cable(Ibreaker As String,                 ' Ток на Автоматичния прекъсвач
                                NumberPoles As String,              ' Брой на фазите
                                Optional layMethod As Integer = 0,  ' Начин на полагане 0 - във въздух; 1 - в земя
                                Optional Broj_Cable As Integer = 1, ' Брой кабели положени паралелно на скара
                                Optional Tipe_Cable As Integer = 0, ' 0 - Кабел; 1 - проводник
                                Optional matType As Integer = 0,    ' 0 - Мед; 1 - Алуминии
                                Optional RetType As Integer = 1     ' 0 - Връща само сечението на фазовото жило; 1 - Връща целия текст
                                ) As String                         ' Избира сечение на кабел

        Dim calc As String = "######"
        Dim Inom As Double = Val(Ibreaker)
        Dim Idop As Double = 0
        Dim Kz As Double = 1
        Dim Q As Double = 65
        Dim Qokdef As Double = 25
        Dim Qok As Double = IIf(layMethod = 0, 30, 15)
        Dim K2 As Double = Math.Sqrt((Q - Qok) / (Q - Qokdef))
        Dim K1 As Double = 1

        K1 = GetK1(Broj_Cable, layMethod)
        Idop = Inom / (K1 * K2)

        ' Използване на речника
        Dim Icable() As Integer = IcableDict(layMethod.ToString() & "_" & matType.ToString() & "_" & Tipe_Cable.ToString())

        For i As Integer = 0 To Icable.Length - 1
            If Icable(i) > Idop Then
                calc = Kable_Size_L(i)
                Exit For
            End If
        Next

        ' Начално предположение за брой кабели
        Dim numCables As Integer = 1

        ' Ако не е намерено подходящо сечение
        If calc = "######" Then
            ' Намиране на поредния номер на "185,0" в масива Kable_Size
            Dim index185 As Integer = Array.IndexOf(Kable_Size_L, "185")
            ' Ако намерим "185", изчисляваме броя кабели
            If index185 >= 0 Then
                ' Увеличаване на броя на кабелите, докато не удовлетворим условието
                Do While Idop > (Icable(index185) * numCables)
                    numCables += 1
                Loop
                ' Изчисляване на calc за съответното сечение и брой кабели
                calc = Kable_Size_L(index185)
            End If
        End If

        If RetType = 0 Then Return calc

        Dim calc_N As String = ""
        Dim Poles As String = If(NumberPoles = "1p", "3x", "5x")
        If Val(calc) > 16 Then
            Poles = "4х"
            Dim index As Integer = Array.IndexOf(Kable_Size_L, calc)
            calc_N = Kable_Size_N(index)
        End If

        Dim Text As String = ""
        Text = If(numCables > 1, numCables & "x", "")
        Text += If(matType = 0, "СВТ", "САВТ")

        If Poles = "4х" Then
            Text += "3х" & calc & "+" & calc_N
        Else
            Text += Poles & calc
        End If

        Text += "mm²"
        Return Text
    End Function
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
        Breakers = New Dictionary(Of Integer, String) From {
            {10, "EZ9 MCB"},
            {16, "EZ9 MCB"},
            {20, "EZ9 MCB"},
            {25, "EZ9 MCB"},
            {32, "EZ9 MCB"},
            {40, "EZ9 MCB"},
            {50, "EZ9 MCB"},
            {63, "EZ9 MCB"},
            {80, "C120"},
            {100, "C120"},
            {125, "NSX"},
            {160, "NSX"},
            {200, "NSX"},
            {250, "NSX"},
            {400, "NSX"},
            {630, "NSX"},
            {800, "NSX"},
            {1000, "NSX"},
            {1250, "NSX"},
            {1600, "NSX"},
            {2000, "MTZ2 20"},
            {2500, "MTZ2 25"},
            {3200, "MTZ2 32"},
            {4000, "MTZ2 40"}
            }
        ' Речник за всички товарови прекъсвачи
        Disconnectors = New Dictionary(Of Integer, String) From {
                        {20, "iSW"},
                        {32, "iSW"},
                        {40, "iSW"},
                        {63, "iSW"},
                        {100, "INS"},
                        {160, "INS"},
                        {250, "INS"},
                        {400, "INS"},
                        {630, "INS"},
                        {800, "INS"},
                        {1000, "INS"},
                        {1250, "INS"},
                        {1600, "IN"},
                        {2500, "IN"}
                }
        'Допустими токови натоварвания на МЕДНИ ШИНИ
        Busbar_Cu = New Dictionary(Of Integer, String) From {
            {210, "15x3"},
            {275, "20x3"},
            {340, "25x3"},
            {475, "30x4"},
            {625, "40x4"},
            {700, "40x5"},
            {860, "50x5"},
            {955, "50x6"},
            {1125, "60x6"},
            {1480, "80x6"},
            {1810, "100x6"},
            {1320, "60x8"},
            {1690, "80x8"},
            {2080, "100x8"},
            {2400, "120x8"},
            {1475, "60x10"},
            {1900, "80x10"},
            {2310, "100x10"},
            {2650, "120x10"}
            }
        'Допустими токови натоварвания на МЕДНИ ШИНИ
        Busbar_Al = New Dictionary(Of Integer, String) From {
            {165, "15x3"},
            {215, "20x3"},
            {265, "25x3"},
            {365, "30x4"},
            {480, "40x4"},
            {540, "40x5"},
            {665, "50x5"},
            {740, "50x6"},
            {870, "60x6"},
            {1150, "80x6"},
            {1425, "100x6"},
            {1025, "60x8"},
            {1320, "80x8"},
            {1625, "100x8"},
            {1900, "120x8"},
            {1155, "60x10"},
            {1480, "80x10"},
            {1820, "100x10"},
            {2070, "120x10"}
            }
        'Допустими токови натоварвания на na Кабел Al/R 2 ЖИЛА
        Cable_AlR_2 = New Dictionary(Of Integer, String) From {
            {83, "2x16"},
            {111, "2x25"}
            }
        'Допустими токови натоварвания на na Кабел Al/R 4 ЖИЛА
        Cable_AlR_4 = New Dictionary(Of Integer, String) From {
            {83, "4x16"},
            {111, "4x25"},
            {138, "3х35+54"},
            {164, "3х50+54"},
            {213, "3х70+54"},
            {258, "3х95+70"},
            {344, "3х150+70"}
            }
        RCD_Catalog = New List(Of strRCD) From {
    New strRCD With {.NominalCurrent = 25, .Type = "AC", .Poles = "2p", .Sensitivity = 30, .DeviceType = "RCCB"},
    New strRCD With {.NominalCurrent = 25, .Type = "AC", .Poles = "4p", .Sensitivity = 30, .DeviceType = "RCCB"},
    New strRCD With {.NominalCurrent = 40, .Type = "AC", .Poles = "2p", .Sensitivity = 30, .DeviceType = "RCCB"},
    New strRCD With {.NominalCurrent = 40, .Type = "AC", .Poles = "4p", .Sensitivity = 30, .DeviceType = "RCCB"},
    New strRCD With {.NominalCurrent = 63, .Type = "AC", .Poles = "2p", .Sensitivity = 30, .DeviceType = "RCCB"},
    New strRCD With {.NominalCurrent = 63, .Type = "AC", .Poles = "4p", .Sensitivity = 30, .DeviceType = "RCCB"},
    New strRCD With {.NominalCurrent = 6, .Type = "AC", .Poles = "2p", .Sensitivity = 30, .DeviceType = "RCBO"},
    New strRCD With {.NominalCurrent = 10, .Type = "AC", .Poles = "2p", .Sensitivity = 30, .DeviceType = "RCBO"},
    New strRCD With {.NominalCurrent = 16, .Type = "AC", .Poles = "2p", .Sensitivity = 30, .DeviceType = "RCBO"},
    New strRCD With {.NominalCurrent = 20, .Type = "AC", .Poles = "2p", .Sensitivity = 30, .DeviceType = "RCBO"},
    New strRCD With {.NominalCurrent = 25, .Type = "AC", .Poles = "2p", .Sensitivity = 30, .DeviceType = "RCBO"},
    New strRCD With {.NominalCurrent = 32, .Type = "AC", .Poles = "2p", .Sensitivity = 30, .DeviceType = "RCBO"},
    New strRCD With {.NominalCurrent = 40, .Type = "AC", .Poles = "2p", .Sensitivity = 30, .DeviceType = "RCBO"},
    New strRCD With {.NominalCurrent = 16, .Type = "AC", .Poles = "2p", .Sensitivity = 30, .DeviceType = "iID"},
    New strRCD With {.NominalCurrent = 25, .Type = "AC", .Poles = "2p", .Sensitivity = 10, .DeviceType = "iID"},
    New strRCD With {.NominalCurrent = 25, .Type = "AC", .Poles = "2p", .Sensitivity = 30, .DeviceType = "iID"},
    New strRCD With {.NominalCurrent = 25, .Type = "AC", .Poles = "2p", .Sensitivity = 100, .DeviceType = "iID"},
    New strRCD With {.NominalCurrent = 25, .Type = "AC", .Poles = "2p", .Sensitivity = 300, .DeviceType = "iID"},
    New strRCD With {.NominalCurrent = 25, .Type = "AC", .Poles = "2p", .Sensitivity = 500, .DeviceType = "iID"},
    New strRCD With {.NominalCurrent = 25, .Type = "AC", .Poles = "4p", .Sensitivity = 30, .DeviceType = "iID"},
    New strRCD With {.NominalCurrent = 25, .Type = "AC", .Poles = "4p", .Sensitivity = 300, .DeviceType = "iID"},
    New strRCD With {.NominalCurrent = 25, .Type = "AC", .Poles = "4p", .Sensitivity = 500, .DeviceType = "iID"},
    New strRCD With {.NominalCurrent = 40, .Type = "AC", .Poles = "2p", .Sensitivity = 30, .DeviceType = "iID"},
    New strRCD With {.NominalCurrent = 40, .Type = "AC", .Poles = "2p", .Sensitivity = 100, .DeviceType = "iID"},
    New strRCD With {.NominalCurrent = 40, .Type = "AC", .Poles = "2p", .Sensitivity = 300, .DeviceType = "iID"},
    New strRCD With {.NominalCurrent = 40, .Type = "AC", .Poles = "2p", .Sensitivity = 500, .DeviceType = "iID"},
    New strRCD With {.NominalCurrent = 40, .Type = "AC", .Poles = "4p", .Sensitivity = 30, .DeviceType = "iID"},
    New strRCD With {.NominalCurrent = 40, .Type = "AC", .Poles = "4p", .Sensitivity = 300, .DeviceType = "iID"},
    New strRCD With {.NominalCurrent = 40, .Type = "AC", .Poles = "4p", .Sensitivity = 500, .DeviceType = "iID"},
    New strRCD With {.NominalCurrent = 63, .Type = "AC", .Poles = "2p", .Sensitivity = 30, .DeviceType = "iID"},
    New strRCD With {.NominalCurrent = 63, .Type = "AC", .Poles = "2p", .Sensitivity = 100, .DeviceType = "iID"},
    New strRCD With {.NominalCurrent = 63, .Type = "AC", .Poles = "2p", .Sensitivity = 300, .DeviceType = "iID"},
    New strRCD With {.NominalCurrent = 63, .Type = "AC", .Poles = "2p", .Sensitivity = 500, .DeviceType = "iID"},
    New strRCD With {.NominalCurrent = 63, .Type = "AC", .Poles = "4p", .Sensitivity = 30, .DeviceType = "iID"},
    New strRCD With {.NominalCurrent = 63, .Type = "AC", .Poles = "4p", .Sensitivity = 300, .DeviceType = "iID"},
    New strRCD With {.NominalCurrent = 63, .Type = "AC", .Poles = "4p", .Sensitivity = 500, .DeviceType = "iID"},
    New strRCD With {.NominalCurrent = 80, .Type = "AC", .Poles = "2p", .Sensitivity = 30, .DeviceType = "iID"},
    New strRCD With {.NominalCurrent = 80, .Type = "AC", .Poles = "2p", .Sensitivity = 100, .DeviceType = "iID"},
    New strRCD With {.NominalCurrent = 80, .Type = "AC", .Poles = "4p", .Sensitivity = 30, .DeviceType = "iID"},
    New strRCD With {.NominalCurrent = 80, .Type = "AC", .Poles = "4p", .Sensitivity = 100, .DeviceType = "iID"},
    New strRCD With {.NominalCurrent = 80, .Type = "AC", .Poles = "4p", .Sensitivity = 300, .DeviceType = "iID"},
    New strRCD With {.NominalCurrent = 100, .Type = "AC", .Poles = "2p", .Sensitivity = 30, .DeviceType = "iID"},
    New strRCD With {.NominalCurrent = 100, .Type = "AC", .Poles = "2p", .Sensitivity = 100, .DeviceType = "iID"},
    New strRCD With {.NominalCurrent = 100, .Type = "AC", .Poles = "2p", .Sensitivity = 300, .DeviceType = "iID"},
    New strRCD With {.NominalCurrent = 100, .Type = "AC", .Poles = "4p", .Sensitivity = 30, .DeviceType = "iID"},
    New strRCD With {.NominalCurrent = 100, .Type = "AC", .Poles = "4p", .Sensitivity = 100, .DeviceType = "iID"},
    New strRCD With {.NominalCurrent = 100, .Type = "AC", .Poles = "4p", .Sensitivity = 300, .DeviceType = "iID"}
}
    End Sub
    Private Function Get_cable_AlR(cable As String,        ' Сечение на фазовото жило
                                   NumberPoles As String   ' Брой на фазите
                                   ) As String

        'Dim current As Integer
        Dim text As String = ""
        ' Подходящо сечение за минимални стойности
        cable = If(Val(cable) < 15, "16", cable)

        '' Преобразуване на сечението на кабела в цяло число

        'If Not Integer.TryParse(cable, current) Then
        '    Return "Невалидно сечение на кабела"
        'End If

        Select Case NumberPoles
            Case "1p"  ' За 2-фазен кабел
                For Each kvp In Cable_AlR_2
                    ' Проверяваме дали подаденото сечение се съдържа в значението (value) на речника
                    If kvp.Value.Contains(cable) Then
                        text = kvp.Value
                        Exit For
                    End If
                Next
            Case "3p"  ' За 4-фазен кабел
                For Each kvp In Cable_AlR_4
                    ' Проверяваме дали подаденото сечение се съдържа в значението (value) на речника
                    If kvp.Value.Contains(cable) Then
                        text = kvp.Value
                        Exit For
                    End If
                Next
            Case Else
                text = "Невалиден брой фази"
        End Select

        ' Ако не е намерено подходящо сечение
        If String.IsNullOrEmpty(text) Then
            text = "НЕ Е подходящ кабел за това сечение"
        End If

        Return text
    End Function
End Class

'Using rec As New Xrecord()
'rec.Data = New ResultBuffer(
'                New TypedValue(CInt(DxfCode.Text), "This is a test"),
'                New TypedValue(CInt(DxfCode.Int8), 0),
'                New TypedValue(CInt(DxfCode.Int16), 1),
'                New TypedValue(CInt(DxfCode.Int32), 2),
'                New TypedValue(CInt(DxfCode.HardPointerId), db.BlockTableId),
'                New TypedValue(CInt(DxfCode.BinaryChunk), New Byte() {0, 1, 2, 3, 4}),
'                New TypedValue(CInt(DxfCode.ArbitraryHandle), db.BlockTableId.Handle),
'                New TypedValue(CInt(DxfCode.UcsOrg), New Point3d(0, 0, 0)))
'End Using

'искам да добавя още малко логика в избора на сечение:
'избраното сечението се записва е променлива calc
'когато избраното сечението е по-малко от 185 (има една грешка която съм допуснал сеченито 180 трябва да е 185) 
'трбва да остане 185
'но когато е по-голямо не е удобно да се работи с такива големи сечения
'тогава се избират два кабела и се поставят в паралел
'например ако избраното сечение е трябва да е 
