Imports System.ComponentModel
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Text.RegularExpressions
Imports ACSMCOMPONENTS24Lib
Imports Autodesk.AutoCAD
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Windows.Data
Imports Autodesk.AutoCAD.Interop.Common

' Дефинирате AcApp като съкращение за AutoCAD Application
Imports AcApp = Autodesk.AutoCAD.ApplicationServices.Application
Public Class SheetSet_new
    ' Постоянни данни от твоята информация
    Private Const Set_Desc As String = "Създадено от Бат Генчо"
    Private Const DefaultProjectName As String = "Project"

    ' Пътят до нашия DST файл (за тестовете приемаме, че е до DWG файла)
    Private dstPath As String = ""
    Structure srtSheetSet
        Dim nameSheet As String             ' Името на групата листи (Subset)
        Dim nameSubSheet As String          ' Името на подлист (Sub-subset)
        Dim nameLayoutForSheet As String    ' Името, което ще се вижда в Sheet Set (Title)
        Dim nameLayout As String            ' Името на Layout в Autocad
        Dim Number As Double                ' Номерът на етажа/котата за сортиране
        Dim nameFile As String              ' Пътят до DWG файла
        Dim objectType As String            ' Типът на обекта
        Dim objectID As Object              ' Променено от IAcSmObjectId на Object
    End Structure
    Dim numbers As New Dictionary(Of Integer, String) From {
        {1, "Първи етаж"},
        {2, "Втори етаж"},
        {3, "Трети етаж"},
        {4, "Четвърти етаж"},
        {5, "Пети етаж"},
        {6, "Шести етаж"},
        {7, "Седми етаж"},
        {8, "Осми етаж"},
        {9, "Девети етаж"},
        {10, "Десети етаж"},
        {11, "Единадесети етаж"},
        {12, "Дванадесети етаж"},
        {13, "Тринадесети етаж"},
        {14, "Четиринадесети етаж"},
        {15, "Петнадесети етаж"},
        {16, "Шестнадесети етаж"},
        {17, "Седемнадесети етаж"},
        {18, "Осемнадесети етаж"},
        {19, "Деветнадесети етаж"},
        {20, "Двадесети етаж"}
}
    Dim Sheets As New Dictionary(Of String, Integer) From {
        {"Съдържание", 1},
        {"Ел.захранване НН", 2},
        {"Заснемане и демонтаж", 3},
        {"Осветителна инсталация", 4},
        {"Евакуационно осветление", 5},
        {"Фасадно осветление", 6},
        {"Силова инсталация", 7},
        {"Ел.инсталация контакти", 8},
        {"План покрив", 9},
        {"Инсталации интернет и кабелна телевизия", 10},
        {"Слаботокова инсталация", 11},
        {"Кабелни скари и кабелни канали", 12},
        {"Защитна заземителна инсталация", 13},
        {"Мълниезащитна инсталация", 14},
        {"Еднолинейна схема на табло", 15},
        {"Котировки ел. инсталации", 16}
    }
    Dim Installations As New Dictionary(Of String, String) From {
        {"ВЪН", "Ел.захранване НН"},
        {"ОСВ", "Осветителна инсталация"},
        {"ЕВА", "Евакуационно осветление"},
        {"ФАС", "Фасадно осветление"},
        {"КОН", "Ел.инсталация контакти"},
        {"СИЛ", "Силова инсталация"},
        {"СЛА", "Инсталации интернет и кабелна телевизия"}, ' Това да се използва за къщите
        {"МЪЛ", "Мълниезащитна инсталация"},
        {"ЗАЗ", "Защитна заземителна инсталация"},
        {"ТАБ", "Еднолинейна схема на табло"},
        {"КОТ", "Котировки ел. инсталации"},
        {"ПОК", "План покрив"},
        {"ЗАС", "Заснемане и демонтаж"},
        {"СКА", "Кабелни скари и кабелни канали"},
        {"ИНТ", "Слаботокова инсталация"},
        {"КАБ", "Слаботокова инсталация"},
        {"ТЕЛ", "Слаботокова инсталация"},
        {"ИНК", "Слаботокова инсталация"},
        {"ИКТ", "Слаботокова инсталация"},
        {"БОЛ", "Слаботокова инсталация"},
        {"ДОС", "Слаботокова инсталация"},
        {"ОПО", "Слаботокова инсталация"},
        {"ВИД", "Слаботокова инсталация"},
        {"СОТ", "Слаботокова инсталация"},
        {"ПИЦ", "Слаботокова инсталация"},
        {"ДОМ", "Слаботокова инсталация"},
        {"НАС", "Настройки"}
    }
    Dim Slabotokowa As New Dictionary(Of String, String) From {
        {"ИНТ", "Инсталация интернет"},
        {"КАБ", "Инсталация кабелна телевизия"},
        {"ТЕЛ", "Телефонна инсталация"},
        {"ИНК", "Инсталации интернет и кабелна телевизия"},
        {"ИКТ", "Инсталации интернет кабелна телевизия и телефон"},
        {"ВИД", "Инсталация видеонаблюдение"},
        {"БОЛ", "Болнична повиквателна инсталация"},
        {"ДОС", "Инсталация контрол на достъпа"},
        {"ОПО", "Оповестителна инсталация"},
        {"СОТ", "Сигнално-охранителна инсталация"},
        {"ПИЦ", "Пожароизвестителна инсталация"},
        {"ДОМ", "Домофона инсталация"}
    }
    ' Всички видове инсталации с трибуквен код ---
    Private ReadOnly LisAll As New Dictionary(Of String, String) From {
    {"ВЪН", "OUT"},               ' Outdoor / външно
    {"Ел.захранване НН", "PWR"},  ' Power supply / Ел.захранване
    {"Осветителна инсталация", "LGT"},  ' Lighting / Осветление
    {"Евакуационно осветление", "EVC"},  ' Evacuation lighting / Осветление за евакуация
    {"Фасадно осветление", "FAC"},  ' Facade lighting / Фасадно осветление
    {"Ел.инсталация контакти", "CON"},  ' Contacts / Ел.инсталация контакти
    {"Силова инсталация", "POW"},  ' Power / Силова инсталация
    {"Мълниезащитна инсталация", "LTP"},  ' Lightning protection / Мълниезащита
    {"Защитна заземителна инсталация", "GND"},  ' Grounding / Заземяване
    {"Еднолинейна схема на табло", "DBS"},  ' Distribution board schematic / Табло
    {"Котировки ел. инсталации", "ELV"},  ' Elevations / Котировки
    {"План покрив", "RFL"},  ' Roof plan / План покрив
    {"Заснемане и демонтаж", "DEM"},  ' Demolition / Заснемане и демонтаж
    {"Настройки", "CFG"},  ' Configuration / Настройки
    {"Инсталация интернет", "NET"},  ' Internet installation
    {"Инсталация кабелна телевизия", "CAT"},  ' Cable TV
    {"Телефонна инсталация", "TEL"},  ' Telephone installation
    {"Инсталации интернет и кабелна телевизия", "INT"},  ' Combined Internet + TV
    {"Инсталации интернет кабелна телевизия и телефон", "ICT"},  ' Internet + TV + Telephone
    {"Инсталация видеонаблюдение", "VID"},  ' Video surveillance
    {"Болнична повиквателна инсталация", "NCS"},  ' Nurse Call System 
    {"Инсталация контрол на достъпа", "ACS"},  ' Access Control System
    {"Оповестителна инсталация", "PAS"},  ' Public Address System
    {"Сигнално-охранителна инсталация", "SEC"},  ' Security / Охрана
    {"Пожароизвестителна инсталация", "FIR"},  ' Fire alarm / Пожар
    {"Домофона инсталация", "DIC"}  ' Door Intercom Communication / Система с табло на входа, звънци по апартаменти/офиси, отваряне на врата
    }
    ' ГЛАВНИ ПАПКИ: Име (Key) и Пореден номер/Индекс (Value)
    Dim MainSubsets As New Dictionary(Of String, Integer) From {
    {"Ел. захранване НН", 0},
    {"Осветителна инсталация", 1},
    {"Силова инсталация", 2},
    {"Слаботокова инсталация", 3},
    {"Заземителна инсталация", 4},
    {"Мълниезащитна инсталация", 5},
    {"Кабелни скари и кабелни канали", 6},
    {"Еднолинейна схема на", 7},
    {"Котировки ел. инсталации", 8}
}
    ' ПОД-ПАПКИ: Само за Слаботоковите инсталации
    ' Програмата ще знае: ако инсталацията е в този списък, сложи я вътре в "Слаботокови инсталации"
    Dim LowVoltageSubsets As New List(Of String) From {
    "Пожароизвестяване",
    "Интернет телевизия телефони",
    "Сигнално-охранителна",
    "Домофонна",
    "Оповестителна",
    "Видеонаблюдение"
}
    <CommandMethod("UpdateSSM")>
    Public Sub RunUpdate()
        Dim acDoc As Document = AcApp.DocumentManager.MdiActiveDocument ' Активен документ
        Dim acDb As Database = acDoc.Database
        ' 1. Взимаме пълния път на текущия DWG файл
        Dim dwgPath As String = acDb.Filename
        ' Извикваме процедурата за името
        Dim buildingName As String = GetBuildingName(acDoc)
        ' Ако потребителят е отказал име, спираме до тук
        'If buildingName = "CANCELLED" Then Return
        Dim name_file As String = acDoc.Name                                    ' Име на DWG файла
        Dim File_Path As String = Path.GetDirectoryName(name_file)              ' Път до папката
        Dim Path_Name As String = Path.GetFileName(File_Path)                   ' Име на папката (име на проекта)
        Dim File_DST As String = Path.Combine(File_Path, Path_Name & ".dst")    ' Пълен път до DST файла
        Dim Set_Desc As String = "Създадено от Бат Генчо"                        ' Описание на Sheet Set-а

        Dim listSheetSet As New List(Of srtSheetSet)                             ' Нови Layout-и за добавяне
        Dim sheetSetManager As IAcSmSheetSetMgr = New AcSmSheetSetMgr            ' Sheet Set Manager
        Dim sheetSetDatabase As AcSmDatabase
        ' Проверка дали DST файлът съществува
        If System.IO.File.Exists(File_DST) Then
            sheetSetDatabase = sheetSetManager.OpenDatabase(File_DST, False)    ' Отваряме съществуващ DST
        Else
            sheetSetDatabase = sheetSetManager.CreateDatabase(File_DST, "", True) ' Създаваме нов DST
        End If
        Try
            Dim sheetSet As AcSmSheetSet = sheetSetDatabase.GetSheetSet()
            If LockDatabase(sheetSetDatabase, True) = False Then                 ' Заключване за запис
                MsgBox("Sheet set не може да бъде отворен за четене.")
                Exit Sub
            End If
            ' Връща списък с всички Sheet-и от дадена Sheet Set база данни (DST).
            Dim sheetsInFile As List(Of srtSheetSet) = GetSheetsFromDatabase(sheetSetDatabase)  ' Съществуващи Sheet-и
            sheetSet.SetName(Path_Name)                                                         ' Име на Sheet Set-а
            sheetSet.SetDesc(Set_Desc)                                                          ' Описание на Sheet Set-а
            ' Почистване на всички Sheet-и от текущия DWG и премахване на празни папки (Subsets) в Sheet Set базата данни.
            CleanOldSheetsFromCurrentDWG(sheetSetDatabase, name_file)

            sheetsInFile = GetSheetsFromDatabase(sheetSetDatabase)  ' Съществуващи Sheet-и
            'Обхожда Layout-ите в чертежа и анализира имената им спрямо зададените речници.
            listSheetSet = CollectLayoutsData(acDoc, name_file, sheetsInFile)
            ' --- 5. СОРТИРАНЕ ---
            Dim sortedList As New List(Of srtSheetSet)
            sortedList = BuildSortedSheetList(listSheetSet)
            ' --- 6. ЗАПИС В DST ---
            saveDST(acDoc, sheetSetDatabase, sortedList, name_file)
            ' --- 7. ГЕНЕРИРАНЕ НА НОМЕРАЦИЯТА ---
            GenerateSheetNumbers(acDoc, sheetSetDatabase)

            'LogSheetSetContent(sheetSetDatabase)

            'ProcessSheetSetContent(sheetSetDatabase, 1)

        Catch ex As Exception
            MsgBox("Грешка: " & ex.Message)
        Finally
            If sheetSetDatabase IsNot Nothing Then LockDatabase(sheetSetDatabase, False) ' Отключване на DST
        End Try
        MsgBox("Sheet Set Name: " & sheetSetDatabase.GetSheetSet().GetName() & vbCrLf &
           "Sheet Set Description: " & sheetSetDatabase.GetSheetSet().GetDesc())
    End Sub
    ''' <summary>
    ''' Изгражда финален, напълно подреден списък от Sheet-и за AutoCAD Sheet Set Manager.
    ''' Подреждането се извършва на ТРИ логически нива:
    '''
    ''' 1) Първо ниво – nameSheet
    '''    • Основни секции (инсталации)
    '''    • Подредени според дефиниран ред (Dictionary Sheets)
    '''
    ''' 2) Второ ниво – nameSubSheet
    '''    • Подсекции в рамките на всяка основна секция
    '''    • Запазва реалната структура от данните
    '''
    ''' 3) Трето ниво – nameLayoutForSheet
    '''    • Реално сортиране на листовете (Кота / Етаж / Ниво)
    '''    • Съобразено със специфичното поведение на AutoCAD SSM
    ''' </summary>
    ''' <param name="listSheetSet">Списък с всички листове (srtSheetSet), събрани от DWG</param>
    ''' <returns>Краен списък, готов за импорт в Sheet Set</returns>
    Private Function BuildSortedSheetList(listSheetSet As List(Of srtSheetSet)) As List(Of srtSheetSet)
        ' =========================================================================
        ' 1. ПОДГОТОВКА И ПРОВЕРКИ
        ' =========================================================================
        ' Списък, който ще държи резултата след второто ниво на подреждане
        Dim secondLevelList As New List(Of srtSheetSet)
        ' Ако входният списък е празен или Nothing – няма какво да сортираме
        If listSheetSet Is Nothing OrElse listSheetSet.Count = 0 Then
            Return secondLevelList
        End If
        ' =========================================================================
        ' 2. ПРЕДВАРИТЕЛНО ПОЧИСТВАНЕ НА ДАННИТЕ
        ' =========================================================================
        ' • Премахваме ненужни секции
        ' • Коригираме празни или некоректни стойности
        ' Цикълът е отзад-напред, за да можем безопасно да махаме елементи
        For i As Integer = listSheetSet.Count - 1 To 0 Step -1
            Dim s As srtSheetSet = listSheetSet(i)
            ' 2.1 Премахваме секцията "Настройки" – тя не трябва да влиза в Sheet Set-а
            If String.Equals(s.nameSheet, "Настройки", StringComparison.OrdinalIgnoreCase) Then
                listSheetSet.RemoveAt(i)
                Continue For
            End If
            ' 2.2 Ако липсва име на основна секция,
            ' използваме името на Layout-а като резервен вариант
            If String.IsNullOrWhiteSpace(s.nameSheet) Then
                s.nameSheet = s.nameLayout
                listSheetSet(i) = s
            End If
        Next
        ' =========================================================================
        ' 3. ПЪРВО НИВО – ПОДРЕЖДАНЕ ПО ОСНОВНИ СЕКЦИИ (nameSheet)
        ' =========================================================================
        ' Тук ще съберем всички Sheet-и, подредени по дефинирания ред в Dictionary Sheets
        Dim firstLevelList As New List(Of srtSheetSet)
        ' 3.1 Минаваме по реда, зададен в Dictionary Sheets
        ' Това гарантира еднакъв и контролиран ред всеки път
        For Each pair In Sheets
            For Each item In listSheetSet
                If item.nameSheet = pair.Key Then
                    firstLevelList.Add(item)
                End If
            Next
        Next
        ' 3.2 Добавяме и всички останали секции,
        ' които НЕ присъстват в Dictionary Sheets
        For Each item In listSheetSet
            If Not Sheets.ContainsKey(item.nameSheet) Then
                If Not firstLevelList.Contains(item) Then
                    firstLevelList.Add(item)
                End If
            End If
        Next
        ' =========================================================================
        ' 4. ВТОРО НИВО – ПОДРЕЖДАНЕ ПО ПОДСЕКЦИИ (nameSubSheet)
        ' =========================================================================
        ' 4.1 Вземаме всички реално използвани секции от данните
        Dim realSectionsInData =
        firstLevelList.Select(Function(x) x.nameSheet).Distinct().ToList()
        ' 4.2 Подреждаме секциите според тежестта им в Dictionary Sheets
        ' Ако секцията не съществува там – тя отива най-отзад
        Dim orderedSections =
        realSectionsInData.OrderBy(Function(name)
                                       If Sheets.ContainsKey(name) Then
                                           Return Sheets(name)
                                       Else
                                           Return 999
                                       End If
                                   End Function).ToList()
        ' 4.3 Изграждаме secondLevelList,
        ' като групираме по секция → подсекция
        For Each sectionName In orderedSections
            Dim currentSectionSheets =
            firstLevelList.Where(Function(x) x.nameSheet = sectionName).ToList()
            ' Вземаме уникалните подсекции за текущата секция
            Dim uniqueSubs =
            currentSectionSheets.Select(Function(x) If(x.nameSubSheet, "")).Distinct().ToList()
            For Each subName In uniqueSubs
                For Each s In currentSectionSheets
                    If If(s.nameSubSheet, "") = subName Then
                        secondLevelList.Add(s)
                    End If
                Next
            Next
        Next
        ' =========================================================================
        ' 5. ТРЕТО НИВО – ЛОГИЧЕСКО СОРТИРАНЕ НА ЛИСТОВЕТЕ
        ' =========================================================================
        ' Това е финалният списък, който ще бъде върнат
        Dim returnList As New List(Of srtSheetSet)
        ' 5.1 Вземаме реда на секциите такъв, какъвто вече е изграден
        Dim sectionOrder =
        secondLevelList.Select(Function(x) x.nameSheet).Distinct().ToList()
        For Each secName In sectionOrder
            ' Всички листове за текущата секция
            Dim sectionItems =
            secondLevelList.Where(Function(x) x.nameSheet = secName).ToList()
            ' Подсекции в рамките на тази секция
            Dim uniqueSubs =
            sectionItems.Select(Function(x) If(x.nameSubSheet, "")).Distinct().ToList()
            For Each subName In uniqueSubs
                ' Всички листове за текущата подсекция
                Dim subGroupItems = sectionItems.Where(Function(x) If(x.nameSubSheet, "").Trim() = subName.Trim()).ToList()
                ' Проверяваме дали имаме коти
                Dim hasElevations =
                subGroupItems.Any(Function(x) x.nameLayoutForSheet.Contains("Кота"))
                Dim sortedSubGroup As List(Of srtSheetSet)
                If hasElevations Then
                    ' -------------------------------------------------------------
                    ' СОРТИРАНЕ ПО КОТА (числово)
                    ' -------------------------------------------------------------
                    sortedSubGroup =
                    subGroupItems.OrderBy(Function(x)
                                              Dim val As Double = 0
                                              Dim layoutName = If(x.nameLayoutForSheet, "")
                                              Dim numPart =
                                                   layoutName.Replace("Кота", "").Replace(",", ".").Trim()

                                              Double.TryParse(numPart,
                                                               Globalization.NumberStyles.Any,
                                                               Globalization.CultureInfo.InvariantCulture,
                                                               val)
                                              Return val
                                          End Function).ToList()
                Else
                    ' -------------------------------------------------------------
                    ' СОРТИРАНЕ ПО ЕТАЖ / НИВО
                    ' -------------------------------------------------------------
                    ' --- СОРТИРАНЕ ПО ЕТАЖ / НИВО ---
                    sortedSubGroup = subGroupItems.OrderBy(Function(x As srtSheetSet)
                                                               Dim nameLower = x.nameLayoutForSheet.ToLower().Trim()
                                                               Dim match = numbers.FirstOrDefault(Function(p) p.Value.ToLower() = nameLower)
                                                               If match.Value IsNot Nothing Then Return CDbl(match.Key)
                                                               Dim m = System.Text.RegularExpressions.Regex.Match(nameLower, "-?\d+")
                                                               If m.Success Then
                                                                   Dim levelNum As Double = 0
                                                                   If Double.TryParse(m.Value, Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, levelNum) Then Return levelNum
                                                               End If
                                                               Return 9999.0
                                                           End Function).ThenBy(Function(x) x.nameLayoutForSheet).ToList()
                End If
                ' =========================================================================
                ' СПЕЦИФИКА НА AUTOCAD SHEET SET MANAGER
                ' =========================================================================
                ' AutoCAD добавя всеки нов Sheet НАЙ-ОТГОРЕ в списъка.
                ' За да получим правилен визуален ред,
                ' обръщаме подредения списък преди импорта.
                sortedSubGroup.Reverse()
                ' Добавяме към финалния резултат
                returnList.AddRange(sortedSubGroup)
            Next
        Next
        ' === Преномериране на сортираните листове от 1 до N (Structure-safe) ===
        Dim idx As Integer = 1
        For i As Integer = 0 To returnList.Count - 1
            Dim s As srtSheetSet = returnList(i)
            s.Number = idx
            returnList(i) = s
            idx += 1
        Next
        ' Връщаме напълно подредения списък
        Return returnList
    End Function
    ''' <summary>
    ''' Връща всички компоненти (листове, subsets и др.) от дадена SheetSet база данни.
    ''' </summary>
    ''' <param name="db">SheetSet база данни (AcSmDatabase)</param>
    ''' <returns>Списък с всички IAcSmComponent обекти в базата</returns>
    Public Function FindAllComponents(db As AcSmDatabase) As List(Of IAcSmComponent)
        Dim comps As New List(Of IAcSmComponent)
        Try
            ' Получаваме итератор за всички обекти в базата
            Dim iter As IAcSmEnumPersist = db.GetEnumerator()
            Dim obj As IAcSmPersist = iter.Next()
            ' Обхождаме всички обекти
            While obj IsNot Nothing
                ' Проверка дали обектът е компонент
                If TypeOf obj Is IAcSmComponent Then
                    comps.Add(DirectCast(obj, IAcSmComponent))
                End If
                obj = iter.Next()
            End While
        Catch ex As Exception
            ' Показваме грешка (може да се заменя с логиране)
            MsgBox($"Грешка при обхождане на компонентите: {ex.Message}")
        End Try
        Return comps
    End Function

    ''  UpdateSSM

    ''' <summary>
    ''' Рекурсивна процедура за запис на йерархията на Sheet Set в лог файл.
    ''' Предполага се, че базата данни (db) е отворена и заключена от външния код.
    ''' </summary>
    Public Sub LogSheetSetContent(db As AcSmDatabase)
        If db Is Nothing Then Exit Sub
        ' 1. Определяне на пътя за лог файла (същата папка, където е DST)
        Dim dstPath As String = db.GetFileName()
        Dim logPath As String = Path.ChangeExtension(dstPath, ".log")
        Try
            ' 2. Записване на лога с UTF8 кодировка за правилна кирилица
            Using writer As New StreamWriter(logPath, False, System.Text.Encoding.UTF8)
                writer.WriteLine("===== ДЪМП НА КОМПОНЕНТИ =====")
                writer.WriteLine("Файл: " & Path.GetFileName(dstPath))
                writer.WriteLine("Дата: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
                writer.WriteLine(New String("-"c, 40))
                ' Вземаме основния проект (корена)
                Dim ss As IAcSmSheetSet = db.GetSheetSet()
                ' Стартираме рекурсивното обхождане
                WriteComponentToLog(ss, writer, 0)
                writer.WriteLine(New String("-"c, 40))
                writer.WriteLine("===== КРАЙ НА ЛОГА =====")
            End Using
            ' По желание: MsgBox("Логът е готов: " & logPath)
        Catch ex As Exception
            MsgBox("Грешка при запис в лог файла: " & ex.Message)
        End Try
    End Sub
    ''' <summary>
    ''' Рекурсивна процедура за обхождане на компонентите на SheetSet и запис в лог файл.
    ''' </summary>
    ''' <param name="comp">Компонент (SheetSet, Subset, Sheet или друг)</param>
    ''' <param name="writer">StreamWriter за записване на лог</param>
    ''' <param name="indent">Ниво на отстъп за визуализация на дървото</param>
    Private Sub WriteComponentToLog(comp As IAcSmComponent, writer As StreamWriter, indent As Integer)
        If comp Is Nothing Then Return

        Const IndentSize As Integer = 4
        Dim prefix As String = New String(" "c, indent * IndentSize)
        Dim name As String = ""

        ' --- Опитваме се да вземем името ---
        Try
            name = comp.GetName()
        Catch
            name = "Unknown Name"
        End Try

        ' --- Определяме типа за лог ---
        Dim typeLabel As String
        If TypeOf comp Is IAcSmSheetSet Then
            typeLabel = "[SheetSet]"
        ElseIf TypeOf comp Is IAcSmSubset Then
            typeLabel = "[Subset]"
        ElseIf TypeOf comp Is IAcSmSheet Then
            typeLabel = "[Sheet]"
        Else
            typeLabel = "[Component]"
        End If

        ' --- Записваме реда в лог ---
        writer.WriteLine($"{prefix}{typeLabel} {name}")

        ' --- Обхождаме децата ---
        Try
            If TypeOf comp Is IAcSmSubset Then
                ' Subset: листовете
                Dim subset As IAcSmSubset = CType(comp, IAcSmSubset)
                Dim iter As IAcSmEnumComponent = subset.GetSheetEnumerator()
                If iter IsNot Nothing Then
                    Dim child As IAcSmComponent = iter.Next()
                    While child IsNot Nothing
                        WriteComponentToLog(child, writer, indent + 1)
                        child = iter.Next()
                    End While
                End If

            ElseIf TypeOf comp Is IAcSmSheetSet Then
                ' SheetSet: директни деца (Subset-и и Sheet-и)
                Dim sheetSet As IAcSmSheetSet = CType(comp, IAcSmSheetSet)
                Dim children As Array = Nothing
                sheetSet.GetDirectlyOwnedObjects(children)
                If children IsNot Nothing Then
                    For Each childObj In children
                        Dim child As IAcSmComponent = TryCast(childObj, IAcSmComponent)
                        If child IsNot Nothing Then
                            WriteComponentToLog(child, writer, indent + 1)
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            writer.WriteLine(prefix & "    ! Грешка при обхождане на децата: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Команда за AutoCAD: Генерира номерата на листовете в DST файла
    ''' Извиква основния метод GenerateSheetNumbers.
    ''' </summary>
    <CommandMethod("GenerateSheetNumbers")>
    Public Sub GenerateSheetNumbersCommand()
        Dim sheetSetDatabase As AcSmDatabase
        Try
            ' --- 1. Взимаме активния документ и базата данни ---
            Dim acDoc As Document = AcApp.DocumentManager.MdiActiveDocument
            Dim acDb As Database = acDoc.Database
            ' --- 3. Определяме пътя до DST файла ---
            Dim dwgFolder As String = Path.GetDirectoryName(acDb.Filename)  ' Папката на DWG файла
            Dim projectName As String = Path.GetFileName(dwgFolder)        ' Име на проекта (папката)
            Dim dstPath As String = Path.Combine(dwgFolder, projectName & ".dst") ' Пълен път до DST файла
            ' --- 4. Инициализираме Sheet Set Manager ---
            Dim sheetSetManager As IAcSmSheetSetMgr = New AcSmSheetSetMgr()
            ' --- 5. Проверяваме дали DST файлът съществува ---
            If System.IO.File.Exists(dstPath) Then            ' Отваряме съществуващ DST файл
                sheetSetDatabase = sheetSetManager.OpenDatabase(dstPath, False)
            Else
                ' Ако DST файлът не съществува, извеждаме съобщение и спираме
                MsgBox("DST файлът не съществува: " & dstPath, MsgBoxStyle.Exclamation, "Внимание")
                Return
            End If
            If LockDatabase(sheetSetDatabase, True) = False Then                 ' Заключване за запис
                MsgBox("Sheet set не може да бъде отворен за четене.")
                Exit Sub
            End If
            ' --- 6. Извикваме основната логика за генериране на номера на листовете ---
            GenerateSheetNumbers(acDoc, sheetSetDatabase)
        Catch ex As Exception
            MsgBox("Грешка: " & ex.Message)
        Finally
            If sheetSetDatabase IsNot Nothing Then LockDatabase(sheetSetDatabase, False) ' Отключване на DST
        End Try
    End Sub
    Public Sub GenerateSheetNumbers(acDoc As Document, dstDatabase As AcSmDatabase)
        ' --- 1. Получаваме списъка с листове от DST ---
        Dim dstSheets As List(Of srtSheetSet) = GetSheetsFromDatabase(dstDatabase)
        ' --- 2. Взимаме BuildingName от активния DWG ---
        Dim buildingName As String = GetOrCreateBuildingName(acDoc)
        ' --- 3. Генериране на номерация според режима ---
        Try
            Dim sheetSet As IAcSmSheetSet = dstDatabase.GetSheetSet()
            Dim werwer = FindAllComponents(dstDatabase)

            ' If True Then SetSheetCount()
            'Dim werwer = FindAllComponents(sheetSet)
            If buildingName = "BuildingName" Then
                ' -----------------------------
                ' СТАНДАРТЕН РЕЖИМ - > една сграда
                '------------------------------
                Dim pko As New PromptKeywordOptions(vbLf & "Изберете начин на номериране на листовете:")
                pko.Keywords.Add("Последователно 01, 02, ... , N")          ' за нас → Global
                pko.Keywords.Add("По кодове формат XXX-01-00")              ' за нас → ByInstallation
                pko.Keywords.Default = "Последователно 01, 02, ... , N"

                Dim pkr As PromptResult = acDoc.Editor.GetKeywords(pko)

                Dim numberingMode As String
                If pkr.Status = PromptStatus.OK Then
                    If pkr.StringResult = "Последователно 01, 02, ... , N" Then
                        numberingMode = "Global"
                    Else
                        numberingMode = "ByInstallation"
                    End If
                Else
                    MsgBox("Номерирането е прекъснато.")
                    Exit Sub
                End If
            Else
                ' -----------------------------
                ' РАЗШИРЕН РЕЖИМ - > много сгради
                '------------------------------
            End If
        Catch ex As Exception
            ' 7. Ако възникне грешка, показваме съобщение
            MsgBox("Грешка при номериране на Sheet Set файла: " & ex.Message)
        End Try
    End Sub
    ''' <summary>
    ''' Създава Sheet Set файл (DST) и добавя листовете според подадения сортиран списък.
    ''' </summary>
    ''' <param name="sheetSetDatabase">Отворената Sheet Set база данни (DST)</param>
    ''' <param name="sortedList">Списък от листове (srtSheetSet), сортирани по групи</param>
    ''' <param name="name_file">DWG файл, който се добавя към листовете</param>
    Public Sub saveDST(acDoc As Document,
                       sheetSetDatabase As AcSmDatabase,
                       sortedList As List(Of srtSheetSet),
                       name_file As String)
        Try
            Dim rootSubset As AcSmSubset = Nothing
            Dim useBuildingRoot As Boolean = False
            ' --- Четене на BuildingName от DWG без промпт към потребителя ---
            Dim buildingName As String = GetOrCreateBuildingName(acDoc)
            ' --- Проверка на режима ---
            If buildingName = "BuildingName" Then
                ' -----------------------------
                ' СТАНДАРТЕН РЕЖИМ - > една сграда
                '------------------------------
                acDoc.Editor.WriteMessage(vbLf & "СТАНДАРТЕН РЕЖИМ - > една сграда")
            Else
                ' -----------------------------
                ' РАЗШИРЕН РЕЖИМ - > много сгради
                '------------------------------
                ' Корен на DST става buildingName
                acDoc.Editor.WriteMessage(vbLf & "РАЗШИРЕН РЕЖИМ - > много сгради:" & buildingName)
                rootSubset = CreateSubset(sheetSetDatabase, buildingName, "", "", "", "", True)
                useBuildingRoot = True
            End If
            ' --- Променливи за текущата и главната папка (Subset) ---
            Dim mainSubset As AcSmSubset = Nothing
            Dim currentSubset As AcSmSubset = Nothing
            Dim prevNameSheet As String = ""
            Dim prevNameSubSheet As String = ""
            ' --- Обхождане на списъка с листове ---
            For i As Integer = 0 To sortedList.Count - 1
                Dim current = sortedList(i)
                ' --- Създаване на mainSubset (nameSheet) ако се е сменило ---
                If current.nameSheet <> prevNameSheet Then
                    If useBuildingRoot Then
                        ' Разширен режим: mainSubset е под buildingName
                        mainSubset = rootSubset.CreateSubset(current.nameSheet, "")
                    Else
                        ' Стандартен режим: mainSubset е директно root
                        mainSubset = CreateSubset(sheetSetDatabase, current.nameSheet, "", "", "", "", True)
                    End If
                    currentSubset = mainSubset
                    prevNameSheet = current.nameSheet
                    prevNameSubSheet = "" ' нулираме предходната под-папка
                End If
                ' --- Създаване на под-папка (nameSubSheet) ако има такава ---
                If Not String.IsNullOrEmpty(current.nameSubSheet) Then
                    If current.nameSubSheet <> prevNameSubSheet Then
                        currentSubset = mainSubset.CreateSubset(current.nameSubSheet, "")
                        prevNameSubSheet = current.nameSubSheet
                    End If
                Else
                    ' Ако няма под-папка, листът отива директно под mainSubset
                    currentSubset = mainSubset
                End If
                ' --- Импортиране на листа ---
                ImportASheet(currentSubset, current.nameLayoutForSheet, "", (i + 1).ToString("D2"), current.nameFile, current.nameLayout)
            Next
            ' --- Съобщение за успешно записан DST ---
            MsgBox("Sheet Set файлът е запазен успешно!")
        Catch ex As Exception
            ' 7. Ако възникне грешка, показваме съобщение
            MsgBox("Грешка при запазване на Sheet Set файла: " & ex.Message)
        End Try
    End Sub
    ''' <summary>
    ''' Обхожда Layout-ите в чертежа и анализира имената им спрямо зададените речници.
    ''' Събира данни за всички Layout-и в даден DWG документ.
    ''' Използва информацията за имената и групиране според инсталации и подкатегории.
    ''' </summary>
    ''' <param name="acDoc">DWG документ, от който се извличат Layout-и</param>
    ''' <param name="name_file">Пълен път до DWG файла</param>
    ''' <param name="sheetsInFile">Списък със съществуващи Layout-и (DST), за да не се добавят дубли</param>
    ''' <returns>Списък от структури srtSheetSet с обработените Layout-и</returns>
    Public Function CollectLayoutsData(acDoc As Document, name_file As String, sheetsInFile As List(Of srtSheetSet)) As List(Of srtSheetSet)
        ' 1. Създаваме празен списък за резултатите
        Dim listSheetSet As New List(Of srtSheetSet)
        ' 2. Стартираме транзакция за безопасен достъп до DWG обектите
        Using acTrans As Transaction = acDoc.TransactionManager.StartTransaction()
            ' 3. Вземаме Layout Dictionary от базата данни
            Dim laye As DBDictionary = acTrans.GetObject(acDoc.Database.LayoutDictionaryId, OpenMode.ForRead)
            ' 4. Обхождаме всички Layout-и
            For Each Item As DBDictionaryEntry In laye
                ' 5. Проверки за валидност на името (игнорираме MODEL или твърде кратки/празни имена)
                If Item.Key.ToUpper() = "MODEL" OrElse String.IsNullOrEmpty(Item.Key) OrElse Item.Key.Length < 3 Then
                    Continue For
                End If
                ' 6. Създаваме нов обект за текущия Layout
                Dim currentItem As New srtSheetSet
                currentItem.nameLayout = Item.Key          ' Име на Layout-а
                currentItem.nameFile = name_file           ' Пълен път до DWG файла
                ' 7. Вземаме първите 3 символа за групиране (напр. ОСВ, СИЛ, ПИЦ)
                Dim Instal As String = Item.Key.ToUpper().Substring(0, 3)
                ' 8. Използваме твои речници за определяне на групата и подсекцията
                currentItem.nameSheet = If(Installations.ContainsKey(Instal), Installations(Instal), "")
                currentItem.nameSubSheet = If(Slabotokowa.ContainsKey(Instal), Slabotokowa(Instal), "")
                ' 9. Логика за имената на Layout-а според ключови думи (КОТА / ТАБЛО / ЕТАЖ / СУТ) 
                currentItem.nameLayoutForSheet = GetLayoutName(Item.Key)
                ' 10. Добавяме в списъка само ако Layout-а е нов (не съществува в DST)
                If IsNewLayout(name_file, currentItem.nameLayout, sheetsInFile) Then
                    listSheetSet.Add(currentItem)
                End If
            Next
            ' 11. Потвърждаваме транзакцията
            acTrans.Commit()
        End Using
        ' 12. Връщаме списъка с обработени Layout-и
        Return listSheetSet
    End Function
    ''' <summary>
    ''' Определя името на layout-а за Sheet на база ключов текст.
    ''' Анализира ключови думи като КОТА, ТАБЛО, ЕТАЖ, СУТЕРЕН.
    ''' Ако ключът не се разпознае, връща ясно обозначение за проблемен layout.
    ''' </summary>
    ''' <param name="key">Оригиналният ключ (напр. име от речник)</param>
    ''' Речник с номера на етажи:
    ''' Key = номер (Integer),
    ''' Value = текст (напр. "Първи етаж")
    ''' <returns>Форматирано име за nameLayoutForSheet</returns>
    Private Function GetLayoutName(key As String) As String
        ' Преобразуваме ключа в главни букви
        ' за да избегнем проблеми с малки/големи букви
        Dim upperKey As String = key.ToUpper()
        ' Помощна променлива за извлечени стойности
        Dim result As String = ""
        ' Основна логика за разпознаване по ключови думи
        Select Case True
        ' ===============================
        ' КОТА
        ' ===============================
            Case upperKey.Contains("КОТА")
                ' Позиция на думата "КОТА"
                Dim kotaIdx As Integer = upperKey.IndexOf("КОТА")
                ' Позиция на "+" и "-"
                Dim pIdx As Integer = key.IndexOf("+")
                Dim mIdx As Integer = key.IndexOf("-")
                ' Определяме откъде започва котата
                Select Case True
                    Case pIdx > 0 AndAlso pIdx > kotaIdx
                        result = key.Substring(pIdx).Trim()
                    Case mIdx > 0 AndAlso mIdx > kotaIdx
                        result = key.Substring(mIdx).Trim()
                    Case Else
                        ' Ако няма + или -, взимаме текста след последния интервал
                        Dim sIdx As Integer = key.Trim().LastIndexOf(" ")
                        result = If(sIdx > -1, "+" & key.Substring(sIdx).Trim(), "+0.00")
                End Select
                ' Връщаме форматираното име за layout
                Return "Кота " & result
        ' ===============================
        ' ТАБЛО
        ' ===============================
            Case upperKey.Contains("ТАБЛО")
                ' Намираме последния интервал
                Dim lastSpace As Integer = key.Trim().LastIndexOf(" ")
                ' Вземаме текста след него
                result = If(lastSpace > -1, key.Substring(lastSpace).Trim(), " ")
                ' Връщаме форматираното име
                Return "Табло ''" & Trim(result) & "''"
        ' ===============================
        ' ЕТАЖ
        ' ===============================
            Case upperKey.Contains("ЕТАЖ")
                ' Опит за извличане на цифра (напр. "2")
                Dim m As Match = Regex.Match(upperKey, "\d+")
                ' Тук ще запишем намерения етаж
                Dim foundFloor As String = ""
                If m.Success Then
                    ' Ако има цифра – търсим я в речника
                    Dim fNum As Integer = CInt(m.Value)
                    If numbers.ContainsKey(fNum) Then
                        foundFloor = numbers(fNum)
                    End If
                Else
                    ' Ако няма цифри – търсим текстово съвпадение
                    For Each kvp In numbers
                        If upperKey.Contains(kvp.Value.ToUpper()) Then
                            foundFloor = kvp.Value
                            Exit For
                        End If
                    Next
                End If
                ' Проверка дали е намерен валиден етаж
                If Not String.IsNullOrEmpty(foundFloor) Then
                    Return foundFloor
                Else
                    ' Маркираме проблемен етаж
                    Return "###### ЕТАЖ"
                End If
        ' ===============================
        ' СУТЕРЕН
        ' ===============================
            Case upperKey.Contains("СУТ")
                Return "Сутерен"

                ' ===============================
                ' ПО ПОДРАЗБИРАНЕ
                ' ===============================
            Case Else
                ' Ако няма разпозната ключова дума,
                ' връщаме оригиналния ключ
                Return key
        End Select
    End Function
    ''' <summary>
    ''' Проверява дали Layout-ът е нов за Sheet Set-а.
    ''' Връща TRUE, ако НЯМА съвпадение (трябва да се добави).
    ''' Връща FALSE, ако ИМА съвпадение (вече съществува).
    ''' </summary>
    Private Function IsNewLayout(ByVal currentFile As String, ByVal currentLayout As String, ByVal existingSheets As List(Of srtSheetSet)) As Boolean
        ' Обхождаме всички записи, които вече сме прочели от DST файла
        For Each existing In existingSheets
            ' Сравняваме Пътя на файла И Името на Layout-а
            ' Използваме StringComparison.OrdinalIgnoreCase, за да не правим разлика между главни и малки букви
            If String.Equals(existing.nameFile, currentFile, StringComparison.OrdinalIgnoreCase) AndAlso
           String.Equals(existing.nameLayout, currentLayout, StringComparison.OrdinalIgnoreCase) Then
                ' Намерихме съвпадение! Значи НЕ е нов.
                Return False
            End If
        Next
        ' Ако сме преминали през целия списък и не сме намерили съвпадение
        Return True
    End Function
    ''' <summary>
    ''' Връща текущото име на сграда, записано в DWG файла (Custom Property "BuildingName").
    ''' Ако няма записано име, създава свойството с начална стойност "BuildingName" и го връща.
    ''' </summary>
    ''' <param name="doc">Текущ документ (Document) на AutoCAD</param>
    ''' <returns>Име на сградата като String</returns>
    Private Function GetOrCreateBuildingName(doc As Document) As String
        Dim buildingName As String = "BuildingName" ' Начална стойност
        Try
            Dim db As Database = doc.Database
            Dim infoBuilder As New DatabaseSummaryInfoBuilder(db.SummaryInfo)
            Dim customProps As System.Collections.IDictionary = infoBuilder.CustomPropertyTable
            Dim propKey As String = "BuildingName"
            ' Проверка дали ключът съществува
            If customProps.Contains(propKey) Then
                ' Вземаме вече съществуващото име
                buildingName = customProps(propKey).ToString().Trim()
            Else
                ' Ако няма записано, създаваме свойството с начална стойност
                customProps.Add(propKey, buildingName)
                ' Записваме обратно в базата данни на чертежа
                db.SummaryInfo = infoBuilder.ToDatabaseSummaryInfo()
                doc.Editor.WriteMessage(vbLf & "Създадено свойство BuildingName със стойност: " & buildingName)
            End If
        Catch ex As Exception
            ' В случай на грешка, връщаме стойността по подразбиране
            doc.Editor.WriteMessage(vbLf & "Грешка при четене/създаване на BuildingName: " & ex.Message)
        End Try
        Return buildingName
    End Function
    ''' <summary>
    ''' Връща името на сградата от DWG файла.
    ''' Използва GetOrCreateBuildingName за автоматично създаване, ако няма property.
    ''' След това пита потребителя дали иска промяна, като показва пояснителен текст за режима.
    ''' Ако потребителят откаже или избере "Не", връща "CANCELLED".
    ''' </summary>
    ''' <param name="doc">Текущ документ (Document) на AutoCAD</param>
    ''' <returns>Име на сградата като String или "CANCELLED" при отказ</returns>
    Private Function GetBuildingName(doc As Document) As String
        ' Взимаме Editor за съобщения и промпти
        Dim ed As Editor = doc.Editor
        Dim propKey As String = "BuildingName" ' Името на Custom Property, което следим
        ' --- 1. Извикваме функцията, която проверява и създава BuildingName, ако липсва ---
        ' GetOrCreateBuildingName гарантира, че property винаги съществува
        Dim bName As String = GetOrCreateBuildingName(doc)
        ' --- 2. Сглобяване на пояснителен текст за потребителя ---
        ' Тук съобщаваме текущото име и режима на работа (Стандартен/Разширен)
        Dim msg As String = vbLf & "Текущ обект: [" & bName & "]"
        If bName.Equals(propKey, StringComparison.OrdinalIgnoreCase) Then
            ' Стойността е стандартната -> единична сграда
            msg &= " -> СТАНДАРТЕН РЕЖИМ (ЕДНА СГРАДА)."
        Else
            ' Иначе -> разширен режим с много сгради
            msg &= " -> РАЗШИРЕН РЕЖИМ (МНОГО СГРАДИ)."
        End If
        msg &= " Желаете ли промяна? " ' Въпрос към потребителя
        ' --- 3. Пита потребителя дали иска промяна ---
        Dim pko As New PromptKeywordOptions(msg)
        pko.Keywords.Add("Да")
        pko.Keywords.Add("Не")
        pko.Keywords.Default = "Не"
        Dim pkr As PromptResult = ed.GetKeywords(pko)
        ' --- 4. Обратна логика: ако НЕ избере "Да" или статус <> OK, връщаме "CANCELLED" ---
        ' Това означава, че процесът може да спре или да се игнорира промяната
        If pkr.Status <> PromptStatus.OK OrElse pkr.StringResult <> "Да" Then
            Return "CANCELLED"
        End If
        ' --- 5. Потребителят избра "Да" → въвежда ново име ---
        Dim pso As New PromptStringOptions(vbLf & "Въведете ново име на сградата (или 'BuildingName' за общ режим, Enter за пропуск): ")
        pso.AllowSpaces = True
        Dim pr As PromptResult = ed.GetString(pso)
        ' Ако потребителят натисне ESC, връщаме "CANCELLED"
        If pr.Status <> PromptStatus.OK Then Return "CANCELLED"
        ' Ако потребителят натисне само Enter (празен низ), запазваме старото име
        If Not String.IsNullOrWhiteSpace(pr.StringResult) Then
            bName = pr.StringResult.Trim()
        End If
        ' --- 6. Обновяване на Custom Property в DWG ---
        Try
            Dim infoBuilder As New DatabaseSummaryInfoBuilder(doc.Database.SummaryInfo)
            Dim customProps As System.Collections.IDictionary = infoBuilder.CustomPropertyTable
            ' Присвояваме стойността (ако е празно, остава старата)
            customProps(propKey) = bName
            ' Записваме обратно в DWG
            doc.Database.SummaryInfo = infoBuilder.ToDatabaseSummaryInfo()
            ed.WriteMessage(vbLf & "Параметърът е обновен на: " & bName)
        Catch ex As Exception
            ed.WriteMessage(vbLf & "Грешка при обновяване на BuildingName: " & ex.Message)
        End Try
        ' --- 7. Връщаме текущото име на сградата ---
        Return bName
    End Function
    ''' <summary>
    ''' Заключва или отключва дадена база данни на Sheet Set (AcSmDatabase).
    ''' </summary>
    ''' <param name="database">Обектът AcSmDatabase, който искаме да заключим/отключим</param>
    ''' <param name="lockFlag">True = заключване, False = отключване</param>
    ''' <returns>
    ''' Връща True, ако операцията е успешна (базата е заключена/отключена), False иначе
    ''' </returns>
    Public Function LockDatabase(
                                 database As AcSmDatabase,
                                 lockFlag As Boolean) As Boolean
        ' Променлива за състоянието на заключване
        Dim dbLock As Boolean = False
        ' Ако искаме да заключим и базата в момента е отключена
        If lockFlag = True And
            database.GetLockStatus() = AcSmLockStatus.AcSmLockStatus_UnLocked Then
            ' Заключваме базата
            database.LockDb(database)
            dbLock = True
            ' Ако искаме да отключим и базата е локално заключена
        ElseIf lockFlag = False And
            database.GetLockStatus() = AcSmLockStatus.AcSmLockStatus_Locked_Local Then
            ' Отключваме базата
            database.UnlockDb(database)
            dbLock = True
            ' Във всички останали случаи операцията не е приложима
        Else
            dbLock = False
        End If
        ' Връщаме резултата от операцията
        LockDatabase = dbLock
    End Function
    ''' <summary>
    ''' Изчислява общия брой листове във всички отворени Sheet Set бази данни
    ''' и записва резултата като персонализирано свойство "Общ брой листове".
    ''' </summary>
    Public Sub SetSheetCount()
        ' Инициализира брояч на листовете с 0
        Dim nSheetCount As Integer = 0
        ' Създава обект от тип IAcSmSheetSetMgr, който представлява Sheet Set Manager
        Dim sheetSetManager As IAcSmSheetSetMgr
        sheetSetManager = New AcSmSheetSetMgr
        ' Използва Sheet Set Manager за получаване на перечислител за заредените бази данни
        Dim enumDatabase As IAcSmEnumDatabase
        enumDatabase = sheetSetManager.GetDatabaseEnumerator()
        ' Получава първата отворена база данни
        Dim item As IAcSmPersist
        item = enumDatabase.Next()
        Dim sheetSetDatabase As AcSmDatabase
        ' Обхожда всички отворени бази данни
        Do While Not item Is Nothing
            sheetSetDatabase = item
            ' Опитва се да заключи базата данни за редактиране
            If LockDatabase(sheetSetDatabase, True) = True Then
                On Error Resume Next
                Dim enumerator As IAcSmEnumPersist
                Dim itemSheetSet As IAcSmPersist
                ' Получава перечислител за обектите в множеството листове
                enumerator = sheetSetDatabase.GetEnumerator()
                itemSheetSet = enumerator.Next()
                ' Обхожда всички обекти в множеството листове
                Do While Not itemSheetSet Is Nothing
                    ' Увеличава брояча, ако обектът е лист
                    If itemSheetSet.GetTypeName() = "AcSmSheet" Then
                        nSheetCount = nSheetCount + 1
                    End If
                    ' Получава следващия обект
                    itemSheetSet = enumerator.Next()
                Loop
                ' Създава персонализирано свойство на множеството листове
                SetCustomProperty(sheetSetDatabase.GetSheetSet(),
                              "Общ брой листове",
                              CStr(nSheetCount),
                              PropertyFlags.CUSTOM_SHEETSET_PROP)
                ' Отключва базата данни
                LockDatabase(sheetSetDatabase, False)
                ' Нулира брояча на листовете за следващата база данни
                nSheetCount = 0
            Else
                ' Показва съобщение за грешка, ако не успее да заключи базата данни
                MsgBox("Unable to access " & sheetSetDatabase.GetSheetSet().GetName())
            End If
            ' Проверява за следваща отворена база данни
            item = enumDatabase.Next
        Loop
    End Sub
    ''' <summary>
    ''' Създава или актуализира персонализирано свойство (Custom Property) за лист или Sheet Set.
    ''' </summary>
    ''' <param name="owner">
    ''' Обектът, за който се задава свойството. Може да е лист (AcSmSheet) или набор от листове (AcSmSheetSet).
    ''' </param>
    ''' <param name="propertyName">Името на персонализираното свойство</param>
    ''' <param name="propertyValue">Стойността на персонализираното свойство</param>
    ''' <param name="sheetSetFlag">Флаг, указващ типа или предназначението на атрибута</param>
    Private Sub SetCustomProperty(owner As IAcSmPersist,
                              propertyName As String,
                              propertyValue As Object,
                              sheetSetFlag As PropertyFlags)
        ' 1. Създава референция към чантата с персонализирани атрибути (Custom Property Bag)
        Dim customPropertyBag As AcSmCustomPropertyBag
        ' 2. Проверява типа на обекта и получава съответната чанта
        If owner.GetTypeName() = "AcSmSheet" Then
            ' Ако обектът е лист (sheet)
            Dim sheet As AcSmSheet = owner
            customPropertyBag = sheet.GetCustomPropertyBag()
        Else
            ' Ако обектът е набор от листове (sheet set)
            Dim sheetSet As AcSmSheetSet = owner
            customPropertyBag = sheetSet.GetCustomPropertyBag()
        End If
        ' 3. Създава референция към персонализирана стойност на атрибут (Custom Property Value)
        Dim customPropertyValue As AcSmCustomPropertyValue = New AcSmCustomPropertyValue()
        customPropertyValue.InitNew(owner)
        ' 4. Задава флаг за атрибута
        customPropertyValue.SetFlags(sheetSetFlag)
        ' 5. Задава стойността на атрибута
        customPropertyValue.SetValue(propertyValue)
        ' 6. Създава или обновява персонализираното свойство в чантата
        customPropertyBag.SetProperty(propertyName, customPropertyValue)
    End Sub
    ''' <summary>
    ''' Преобразува цяло число в текстово представяне,
    ''' използвайки предварително дефиниран речник (numbers).
    ''' </summary>
    ''' <param name="number">Цяло число, което трябва да се преобразува в текст</param>
    ''' <returns>
    ''' Текстовото представяне на числото, ако съществува в речника,
    ''' или "Невалидно число", ако липсва.
    ''' </returns>
    Function NumberToText(ByVal number As Integer) As String
        ' Проверяваме дали речникът съдържа подаденото число като ключ
        If numbers.ContainsKey(number) Then
            ' Връщаме текстовата стойност, съответстваща на числото
            Return numbers(number)
        Else
            ' Ако числото не съществува в речника
            Return "Невалидно число"
        End If
    End Function
    ''' <summary>
    ''' Връща списък с всички Sheet-и от дадена Sheet Set база данни (DST).
    ''' </summary>
    ''' <param name="sheetSetDatabase">Sheet Set база данни, от която се извличат листовете</param>
    ''' <returns>Списък от структури srtSheetSet с информация за всеки Sheet</returns>
    Public Function GetSheetsFromDatabase(sheetSetDatabase As AcSmDatabase) As List(Of srtSheetSet)
        ' 1. Създаваме празен списък за съхраняване на вече съществуващите листове
        Dim existingSheets As New List(Of srtSheetSet)
        ' 2. Получаваме enumerator за всички persist обекти в DST файла
        Dim iter As IAcSmEnumPersist = sheetSetDatabase.GetEnumerator()
        ' 3. Вземаме първия обект от enumerator-а
        Dim item As IAcSmPersist = iter.Next()
        ' 4. Обхождаме всички обекти в Sheet Set базата
        While item IsNot Nothing
            ' 5. Проверяваме дали текущият обект е Sheet
            If TypeOf item Is IAcSmSheet Then
                ' 5a. Каст към IAcSmSheet
                Dim smSheet As IAcSmSheet = DirectCast(item, IAcSmSheet)
                ' 5b. Създаваме нов обект за съхранение на данните
                Dim data As New srtSheetSet
                ' 5c. Записваме основни свойства на Sheet-а
                data.Number = smSheet.GetNumber()                        ' Номер на листа
                data.nameLayoutForSheet = smSheet.GetTitle()             ' Заглавие на листа (Sheet Title)
                data.objectType = smSheet.GetTypeName()                  ' Тип на обекта
                data.objectID = smSheet.GetObjectId()                    ' ObjectId
                ' 5d. Вземаме референцията към Layout-а в чертежа
                Dim layoutRef As IAcSmAcDbLayoutReference = smSheet.GetLayout()
                If layoutRef IsNot Nothing Then
                    data.nameLayout = layoutRef.GetName()                ' Име на Layout-а
                    data.nameFile = layoutRef.GetFileName()              ' Име на DWG файла
                End If
                ' 5e. Добавяме готовия обект към списъка с вече съществуващи Sheet-и
                existingSheets.Add(data)
            End If
            ' 6. Преминаваме към следващия обект в enumerator-а
            item = iter.Next()
        End While
        ' 7. Връщаме списъка с всички съществуващи Sheet-и
        Return existingSheets
    End Function
    ''' <summary>
    ''' Почистване на всички Sheet-и от текущия DWG и премахване на празни папки (Subsets) в Sheet Set базата данни.
    ''' </summary>
    ''' <param name="db">Sheet Set база данни (DST), която ще се почиства</param>
    ''' <param name="targetPath">Път до DWG файла, за който се изтриват листовете</param>
    Private Sub CleanOldSheetsFromCurrentDWG(ByRef db As AcSmDatabase, ByVal targetPath As String)
        Try
            ' 1. Списък с Sheet-ове, които ще бъдат изтрити
            Dim sheetsToDelete As New List(Of IAcSmSheet)
            ' 2. Списък с папки (Subsets), които потенциално могат да станат празни
            Dim subsetsToCheck As New List(Of IAcSmSubset)
            ' 3. Обхождаме всички persist обекти в Sheet Set базата
            Dim iter As IAcSmEnumPersist = db.GetEnumerator()
            Dim item As IAcSmPersist = iter.Next()
            While item IsNot Nothing
                ' A) Ако текущият обект е Sheet
                If TypeOf item Is IAcSmSheet Then
                    Dim smSheet As IAcSmSheet = DirectCast(item, IAcSmSheet)
                    Try
                        ' Вземаме Layout референцията към DWG файла
                        Dim layoutRef As IAcSmAcDbLayoutReference = smSheet.GetLayout()
                        If layoutRef IsNot Nothing Then
                            ' Проверяваме дали Sheet-ът сочи към текущия DWG
                            If String.Equals(layoutRef.GetFileName(),
                                         targetPath,
                                         StringComparison.OrdinalIgnoreCase) Then
                                ' Маркираме Sheet-а за изтриване
                                sheetsToDelete.Add(smSheet)
                            End If
                        End If
                    Catch
                        ' Ако Sheet-ът е повреден или има проблем с Layout референцията,
                        ' го маркираме директно за изтриване
                        sheetsToDelete.Add(smSheet)
                    End Try
                    ' B) Ако текущият обект е папка (Subset)
                ElseIf TypeOf item Is IAcSmSubset Then
                    ' Запазваме папката за по-късна проверка за празнота
                    subsetsToCheck.Add(DirectCast(item, IAcSmSubset))
                End If
                ' Преминаваме към следващия persist обект
                item = iter.Next()
            End While
            ' 4. Изтриваме всички Sheet-ове, които бяха маркирани
            For Each sheet In sheetsToDelete
                Try
                    Dim owner As IAcSmSubset = TryCast(sheet.GetOwner(), IAcSmSubset)
                    If owner IsNot Nothing Then
                        owner.RemoveSheet(sheet)
                    End If
                Catch
                    ' Игнорираме грешки при изтриване на отделни Sheet-ове
                End Try
            Next
            ' 5. Изтриване на празните папки (Subsets)
            ' Използваме Do цикъл, защото след всяко изтриване
            ' структурата на Sheet Set-а се променя
            Dim somethingWasDeleted As Boolean
            Do
                somethingWasDeleted = False
                ' 5a. Винаги вземаме АКТУАЛЕН списък на всички Subsets
                Dim currentSubsets As New List(Of IAcSmSubset)
                Dim iterSubset As IAcSmEnumPersist = db.GetEnumerator()
                Dim itemSubset As IAcSmPersist = iterSubset.Next()
                While itemSubset IsNot Nothing
                    If TypeOf itemSubset Is IAcSmSubset Then
                        currentSubsets.Add(DirectCast(itemSubset, IAcSmSubset))
                    End If
                    itemSubset = iterSubset.Next()
                End While
                ' 5b. Проверяваме папките отзад-напред (от най-вътрешните към външните)
                For i As Integer = currentSubsets.Count - 1 To 0 Step -1
                    Dim subFolder = currentSubsets(i)
                    ' Проверяваме дали папката е празна (няма Sheet-ове или подпапки)
                    Dim subIter As IAcSmEnumComponent = subFolder.GetSheetEnumerator()
                    If subIter Is Nothing OrElse subIter.Next() Is Nothing Then
                        Try
                            Dim owner As Object = subFolder.GetOwner()
                            ' Премахваме папката според нейния собственик
                            If TypeOf owner Is IAcSmSubset Then
                                DirectCast(owner, IAcSmSubset).RemoveSubset(subFolder)
                                somethingWasDeleted = True
                            ElseIf TypeOf owner Is IAcSmSheetSet Then
                                DirectCast(owner, IAcSmSheetSet).RemoveSubset(subFolder)
                                somethingWasDeleted = True
                            End If
                            If somethingWasDeleted Then
                                Debug.Print("--- Изтрита празна папка (Subset).")
                                Exit For ' Излизаме, за да обновим списъка в Do цикъла
                            End If
                        Catch
                            ' Игнорираме грешки при триене на конкретна папка
                        End Try
                    End If
                Next
            Loop While somethingWasDeleted ' Повтаряме, докато има какво да се трие
        Catch ex As Exception
            ' Логваме общата грешка в Debug конзолата
            Debug.Print("Грешка при почистване: " & ex.Message)
        End Try
    End Sub
    ''' <summary>
    ''' Създава нов Subset (папка) в Sheet Set-а с предоставено име и описание.
    ''' Може да зададе местоположение за нови листове и шаблони.
    ''' </summary>
    ''' <param name="sheetSetDatabase">Sheet Set база данни (DST)</param>
    ''' <param name="name">Име на Subset-а</param>
    ''' <param name="description">Описание на Subset-а</param>
    ''' <param name="newSheetLocation">Път за нови DWG файлове (по избор)</param>
    ''' <param name="newSheetDWTLocation">Път до DWT шаблон (по избор)</param>
    ''' <param name="newSheetDWTLayout">Име на Layout в DWT шаблона (по избор)</param>
    ''' <param name="promptForDWT">Дали да се пита за шаблон при създаване на нов лист</param>
    ''' <returns>Връща създадения AcSmSubset</returns>
    Private Function CreateSubset(sheetSetDatabase As AcSmDatabase,
                                  name As String,
                                  description As String,
                                  Optional newSheetLocation As String = "",
                                  Optional newSheetDWTLocation As String = "",
                                  Optional newSheetDWTLayout As String = "",
                                  Optional promptForDWT As Boolean = False) As AcSmSubset

        ' 1. Създаваме Subset с име и описание
        Dim subset As AcSmSubset = sheetSetDatabase.GetSheetSet().CreateSubset(name, description)
        ' 2. Вземаме папката, в която се намира Sheet Set-ът
        Dim sheetSetFolder As String
        sheetSetFolder = Mid(sheetSetDatabase.GetFileName(), 1, InStrRev(sheetSetDatabase.GetFileName(), "\"))
        ' 3. Създаваме File Reference обект за нови листове
        Dim fileReference As IAcSmFileReference
        fileReference = subset.GetNewSheetLocation()
        ' 4. Ако е зададен път за нови листове, го използваме, иначе използваме папката на Sheet Set-а
        If newSheetLocation <> "" Then
            fileReference.SetFileName(newSheetLocation)
        Else
            fileReference.SetFileName(sheetSetFolder)
        End If
        ' 5. Задаваме местоположението за нови листове в Subset-а
        subset.SetNewSheetLocation(fileReference)
        ' 6. Вземаме Layout Reference обекта за дефиниране на шаблон
        Dim layoutReference As AcSmAcDbLayoutReference
        layoutReference = subset.GetDefDwtLayout
        ' 7. Ако е зададен шаблон, задаваме неговото име и път
        If newSheetDWTLocation <> "" Then
            layoutReference.SetFileName(newSheetDWTLocation)
            layoutReference.SetName(newSheetDWTLayout)
            ' 8. Присвояваме Layout Reference на Subset-а
            subset.SetDefDwtLayout(layoutReference)
        End If
        ' 9. Задаваме дали да се пита за шаблон при създаване на нов лист
        subset.SetPromptForDwt(promptForDWT)
        ' 10. Връщаме създадения Subset
        CreateSubset = subset
    End Function
    ''' <summary>
    ''' Импортира нов лист (Sheet) в Subset или Sheet Set.
    ''' Настройва Layout, DWG файл и свойства на листа.
    ''' </summary>
    ''' <param name="component">Обектът, в който ще се импортира (Subset или Sheet Set)</param>
    ''' <param name="title">Заглавие на листа</param>
    ''' <param name="description">Описание на листа</param>
    ''' <param name="number">Номер на листа</param>
    ''' <param name="fileName">DWG файл за импортиране</param>
    ''' <param name="layout">Layout в DWG файла</param>
    ''' <returns>Връща импортирания AcSmSheet</returns>
    Private Sub ImportASheet(component As IAcSmComponent,
                                  title As String,
                                  description As String,
                                  number As String,
                                  fileName As String,
                                  layout As String)
        Try
            ' Ако заглавието е празно, използваме името на Layout-а
            If IsNothing(title) Then title = layout
            Dim sheet As AcSmSheet
            ' 1. Създаваме Layout Reference обект
            Dim layoutReference As New AcSmAcDbLayoutReference
            layoutReference.InitNew(component)
            ' 2. Настройваме DWG файл и Layout за листа
            layoutReference.SetFileName(fileName)
            layoutReference.SetName(layout)
            ' 3. Импортираме листа в Subset или Sheet Set
            If component.GetTypeName = "AcSmSubset" Then
                Debug.Print("Опит за импорт: File=" & fileName & " Layout=" & layout)
                ' Проверка дали DWG файлът физически съществува
                If Not System.IO.File.Exists(fileName) Then
                    MsgBox("Грешка: Файлът не е намерен на този път: " & fileName)
                End If
                Dim subset As AcSmSubset = component
                sheet = subset.ImportSheet(layoutReference)
                subset.InsertComponent(sheet, Nothing)
            Else
                Dim sheetSetDatabase As AcSmDatabase = component
                sheet = sheetSetDatabase.GetSheetSet().ImportSheet(layoutReference)
                sheetSetDatabase.GetSheetSet().InsertComponent(sheet, Nothing)
            End If
            ' 4. Настройваме свойства на листа
            sheet.SetDesc(description)
            sheet.SetTitle(title)
            sheet.SetNumber(number)
            ' 5. Връщаме импортирания лист
        Catch ex As Exception
            ' Показваме съобщение при възникнала грешка
            MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
        End Try
    End Sub
    ' --- Основен метод за стартиране на нумерацията на Sheet Set ---
    Public Sub ProcessSheetSetContent(db As AcSmDatabase, ByRef currentNumber As Integer)
        ' Ако базата данни е нищо, излизаме
        If db Is Nothing Then Exit Sub
        ' Вземаме главния SheetSet от базата
        Dim ss As IAcSmSheetSet = db.GetSheetSet()
        ' Стартираме рекурсивното обхождане от корена (SheetSet)
        IterateAndNumber(ss, currentNumber)
    End Sub
    ' --- Рекурсивна процедура за обход и нумерация ---
    Private Sub IterateAndNumber(comp As IAcSmComponent, ByRef number As Integer)
        ' Проверка за нищо
        If comp Is Nothing Then Return
        ' --- Ако компонентът е Sheet (лист) ---
        If TypeOf comp Is IAcSmSheet Then
            Dim sheet As IAcSmSheet = CType(comp, IAcSmSheet)
            ' Задаваме номер с формат D2 (01, 02, 03 …)
            sheet.SetNumber(number.ToString("D2"))
            number += 1
        End If
        ' --- Рекурсивно обработваме Subset-и и SheetSet-и ---
        Try
            ' Ако компонентът е Subset
            If TypeOf comp Is IAcSmSubset Then
                Dim subset As IAcSmSubset = CType(comp, IAcSmSubset)
                ' Вземаме Enumerator за листовете в Subset
                Dim iter As IAcSmEnumComponent = subset.GetSheetEnumerator()
                Dim child As IAcSmComponent = iter.Next()
                ' Обхождаме всички листове/подкомпоненти в Subset
                While child IsNot Nothing
                    IterateAndNumber(child, number)
                    child = iter.Next()
                End While
                ' Ако компонентът е SheetSet (корен на йерархията)
            ElseIf TypeOf comp Is IAcSmSheetSet Then
                Dim sheetSet As IAcSmSheetSet = CType(comp, IAcSmSheetSet)
                Dim children As Array = Nothing
                ' Вземаме директно собствените обекти на SheetSet
                sheetSet.GetDirectlyOwnedObjects(children)
                ' Ако има такива, обхождаме ги един по един
                If children IsNot Nothing Then
                    For Each childObj In children
                        Dim child As IAcSmComponent = TryCast(childObj, IAcSmComponent)
                        If child IsNot Nothing Then IterateAndNumber(child, number)
                    Next
                End If
            End If
        Catch
            ' Ако има грешка при достъп до компонентите, игнорираме за момента
            ' Може да се добави лог или съобщение
        End Try
    End Sub
End Class