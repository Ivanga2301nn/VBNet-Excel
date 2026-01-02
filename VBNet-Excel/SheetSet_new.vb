Imports System.ComponentModel
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports ACSMCOMPONENTS24Lib
Imports Autodesk.AutoCAD
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports Microsoft.Office.Interop.Excel

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
        {"Ел.захранване НН", 1},
        {"Заснемане и демонтаж", 2},
        {"Осветителна инсталация", 3},
        {"Евакуационно осветление", 4},
        {"Фасадно осветление", 5},
        {"Силова инсталация", 6},
        {"Ел.инсталация контакти", 7},
        {"План покрив", 8},
        {"Инсталации интернет и кабелна телевизия", 9},
        {"Слаботокова инсталация", 10},
        {"Кабелни скари и кабелни канали", 11},
        {"Защитна заземителна инсталация", 12},
        {"Мълниезащитна инсталация", 13},
        {"Еднолинейна схема на табло", 14},
        {"Котировки ел. инсталации", 15}
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
        {"ВИД", "Инсталации видеонаблюдение"},
        {"БОЛ", "Болнична повиквателна инсталация"},
        {"ДОС", "Инсталация контрол на достъпа"},
        {"ОПО", "Оповестителна инсталация"},
        {"СОТ", "Сигнално-охранителна инсталация"},
        {"ПИЦ", "Пожароизвестителна инсталация"},
        {"ДОМ", "Домофона инсталация"}
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
        Dim buildingName As String = GetBuildingName(acDb)
        ' Ако потребителят е отказал име, спираме до тук
        If buildingName = "CANCELLED" Then Return
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
            'Обхожда Layout-ите в чертежа и анализира имената им спрямо зададените речници.
            listSheetSet = CollectLayoutsData(acDoc, name_file, sheetsInFile)

            ' --- 5. СОРТИРАНЕ ---
            Dim sortedList As New List(Of srtSheetSet)
            sortedList = BuildSortedSheetList(listSheetSet)

            Dim a As Integer = 1
            a = a + 1
            ' --- 6. ЗАПИС В DST ---
            saveDST(sheetSetDatabase, File_DST, sortedList, name_file)
        Catch ex As Exception
            MsgBox("Грешка: " & ex.Message)
        Finally
            If sheetSetDatabase IsNot Nothing Then LockDatabase(sheetSetDatabase, False) ' Отключване на DST
        End Try
        SetSheetCount()                                                          ' Обновяване на броя листове
        MsgBox("Sheet Set Name: " & sheetSetDatabase.GetSheetSet().GetName() & vbCrLf &
           "Sheet Set Description: " & sheetSetDatabase.GetSheetSet().GetDesc())
    End Sub



    ''' <summary>
    ''' Изгражда списък от листове, подреден по трите нива:
    ''' 1) nameSheet (първо ниво, ред от Sheets)
    ''' 2) nameSubSheet (второ ниво, последователно)
    ''' 3) nameLayoutForSheet (трето ниво, логически ред)
    ''' </summary>
    ''' <param name="listSheetSet">Списък с листове (srtSheetSet)</param>
    ''' <returns>Подреден списък от листове</returns>
    Private Function BuildSortedSheetList(listSheetSet As List(Of srtSheetSet)) _
                                      As List(Of srtSheetSet)
        Dim secondLevelList As New List(Of srtSheetSet)
        If listSheetSet Is Nothing OrElse listSheetSet.Count = 0 Then
            Return secondLevelList
        End If
        ' === Първо ниво ===
        Dim firstLevelList As New List(Of srtSheetSet)
        For Each pair In Sheets
            For Each item In listSheetSet
                If item.nameSheet = pair.Key Then
                    firstLevelList.Add(item)
                End If
            Next
        Next
        ' === Второ ниво ===
        ' Създаваме нов списък, който ще съдържа подредените елементи по nameSubSheet

        Dim subSheetNames As New List(Of String)
        ' Събираме уникални nameSubSheet от първо ниво
        For Each s In firstLevelList
            If Not subSheetNames.Contains(s.nameSubSheet) Then
                subSheetNames.Add(s.nameSubSheet)
            End If
        Next
        ' Подреждаме първо ниво по второ ниво
        For Each subName In subSheetNames
            For Each s In firstLevelList
                If s.nameSubSheet = subName Then
                    secondLevelList.Add(s)
                End If
            Next
        Next
        Return secondLevelList
    End Function





    ''' <summary>
    ''' Сортира списък от обекти srtSheetSet спрямо подредбата на ключовете в Dictionary.
    ''' </summary>
    ''' <param name="sourceList">Оригиналният списък с данни.</param>
    ''' <param name="orderDict">Dictionary, чиито ключове определят новата подредба.</param>
    Public Function SortSheetSet(sourceList As List(Of srtSheetSet), orderDict As Dictionary(Of String, Object)) As List(Of srtSheetSet)

        ' 1. Индексираме оригиналния списък в Dictionary за O(1) достъп.
        ' Използваме GroupBy, в случай че имаш повече от един елемент с едно и също име.
        Dim lookup = sourceList.GroupBy(Function(x) x.nameSheet).ToDictionary(Function(g) g.Key, Function(g) g.ToList())

        Dim result As New List(Of srtSheetSet)

        ' 2. Обхождаме само желания ред (от Sheets)
        For Each key In orderDict.Keys
            If lookup.ContainsKey(key) Then
                ' Добавяме всички намерени елементи с това име
                result.AddRange(lookup(key))
            End If
        Next

        Return result
    End Function
    ''' <summary>
    ''' Създава Sheet Set файл (DST) и добавя листовете според подадения сортиран списък.
    ''' </summary>
    ''' <param name="sheetSetDatabase">Отворената Sheet Set база данни (DST)</param>
    ''' <param name="dstPath">Път, където ще се запази DST файлът</param>
    ''' <param name="sortedList">Списък от листове (srtSheetSet), сортирани по групи</param>
    ''' <param name="name_file">DWG файл, който се добавя към листовете</param>
    Public Sub saveDST(sheetSetDatabase As AcSmDatabase,
                   dstPath As String,
                   sortedList As List(Of srtSheetSet),
                   name_file As String)
        Try
            ' --- 6. ЗАПИС В DST ---
            ' 1. Променливи за текущата и главната папка (Subset)
            Dim mainSubset As AcSmSubset = Nothing
            Dim currentSubset As AcSmSubset = Nothing
            ' 2. Обхождаме сортирания списък с листове
            For i As Integer = 0 To sortedList.Count - 1
                Dim current = sortedList(i)
                ' 2а. Номерираме листа според реда му в списъка
                'current.Number = (i + 1).ToString()
                ' 3. Създаваме нова главна папка (Subset), ако е първи елемент
                ' или ако името на основния лист (Sheet) се е променило спрямо предходния
                If i = 0 OrElse current.nameSheet <> sortedList(i - 1).nameSheet Then
                    mainSubset = CreateSubset(sheetSetDatabase, current.nameSheet, "", "", "", "", True)
                    currentSubset = mainSubset
                End If
                ' 4. Ако има под-папка (SubSheet), създаваме я
                If Not String.IsNullOrEmpty(current.nameSubSheet) Then
                    ' Създаваме под-папка, ако името й се различава от предходното
                    If i = 0 OrElse current.nameSubSheet <> sortedList(i - 1).nameSubSheet Then
                        currentSubset = mainSubset.CreateSubset(current.nameSubSheet, "")
                    End If
                Else
                    ' Ако няма под-папка, текущата папка е главната
                    currentSubset = mainSubset
                End If

                ' 5. Импортираме листа в текущата папка (Subset)
                ImportASheet(currentSubset, current.nameLayoutForSheet, "", current.Number, name_file, current.nameLayout)
            Next
            ' 6. Показваме съобщение за успешно записан DST файл
            MsgBox("Sheet Set файлът е запазен успешно: " & dstPath)
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
    ''' Връща името на сградата от свойствата на чертежа (Database).
    ''' Ако потребителят иска, позволява въвеждане на ново име.
    ''' Използва "BuildingName" като ключово свойство.
    ''' </summary>
    ''' <param name="acDb">Обект Database на текущия AutoCAD документ</param>
    ''' <returns>
    ''' Връща текущото или ново име на сградата.
    ''' Ако потребителят анулира въвеждането (ESC), връща "CANCELLED".
    ''' </returns>
    Private Function GetBuildingName(ByVal acDb As Database) As String
        ' Ключ за потребителското свойство
        Dim propKey As String = "BuildingName"
        ' Стандартна стойност, ако не е намерено име
        Dim bName As String = propKey
        ' Editor за взаимодействие с потребителя
        Dim ed As Editor = AcApp.DocumentManager.MdiActiveDocument.Editor
        ' 1. Достъп до свойствата на чертежа (DWG PROPS)
        Dim infoBuilder As New DatabaseSummaryInfoBuilder(acDb.SummaryInfo)
        Dim customProps As System.Collections.IDictionary = infoBuilder.CustomPropertyTable
        ' Проверка дали свойството вече съществува
        If customProps.Contains(propKey) Then
            bName = customProps(propKey).ToString().Trim()
        End If
        ' 2. Сглобяване на пояснителен текст за потребителя
        Dim msg As String = vbLf & "Текущ обект: [" & bName & "]"
        If bName.Equals(propKey, StringComparison.OrdinalIgnoreCase) Then
            ' Ако стойността е стандартната -> едно име на сграда
            msg &= " -> СТАНДАРТЕН РЕЖИМ (ЕДНА СГРАДА)."
        Else
            ' Иначе -> разширен режим с много сгради
            msg &= " -> РАЗШИРЕН РЕЖИМ (МНОГО СГРАДИ)."
        End If
        msg &= " Желаете ли промяна? "
        ' 3. Питане на потребителя с ключови думи Yes/No
        Dim pko As New PromptKeywordOptions(msg)
        pko.Keywords.Add("Да")
        pko.Keywords.Add("Не")
        pko.Keywords.Default = "Не"
        Dim pkr As PromptResult = ed.GetKeywords(pko)
        ' 4. Ако потребителят избере "Yes", питаме за ново име
        If pkr.Status = PromptStatus.OK AndAlso pkr.StringResult = "Да" Then
            Dim pso As New PromptStringOptions(vbLf & "Въведете име на сграда (или '" & propKey & "' за общ режим): ")
            pso.AllowSpaces = True
            Dim pr As PromptResult = ed.GetString(pso)
            ' Ако потребителят натисне ESC или откаже -> прекратяваме
            If pr.Status <> PromptStatus.OK Then Return "CANCELLED"
            ' Вземаме въведеното име и премахваме излишните интервали
            bName = pr.StringResult.Trim()
            ' 5. Записваме новото или промененото име в свойствата на чертежа
            If customProps.Contains(propKey) Then
                customProps(propKey) = bName
            Else
                customProps.Add(propKey, bName)
            End If
            ' 6. Записваме промените в базата данни на чертежа
            acDb.SummaryInfo = infoBuilder.ToDatabaseSummaryInfo()
            ' Информация към потребителя
            ed.WriteMessage(vbLf & "Параметърът е обновен на: " & bName)
        End If
        ' 7. Връщаме текущото или ново име
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
        ' 1. Инициализация на брояч за листовете
        Dim nSheetCount As Integer = 0
        ' 2. Създаваме обект Sheet Set Manager за достъп до базите данни
        Dim sheetSetManager As IAcSmSheetSetMgr
        sheetSetManager = New AcSmSheetSetMgr
        ' 3. Получаваме перечислител за всички заредени Sheet Set бази данни
        Dim enumDatabase As IAcSmEnumDatabase
        enumDatabase = sheetSetManager.GetDatabaseEnumerator()
        ' 4. Получаваме първата отворена база данни
        Dim item As IAcSmPersist
        item = enumDatabase.Next()
        Dim sheetSetDatabase As AcSmDatabase
        ' 5. Обхождаме всички отворени бази данни
        Do While Not item Is Nothing
            sheetSetDatabase = item
            ' 6. Опитваме се да заключим базата данни за редактиране
            If LockDatabase(sheetSetDatabase, True) = True Then
                On Error Resume Next ' Игнорира грешки при обхождане
                ' 7. Получаваме перечислител за обектите в Sheet Set
                Dim enumerator As IAcSmEnumPersist
                Dim itemSheetSet As IAcSmPersist
                enumerator = sheetSetDatabase.GetEnumerator()
                itemSheetSet = enumerator.Next()
                ' 8. Обхождаме всички обекти в множеството листове
                Do While Not itemSheetSet Is Nothing
                    ' Ако обектът е лист (AcSmSheet), увеличаваме брояча
                    If itemSheetSet.GetTypeName() = "AcSmSheet" Then
                        nSheetCount = nSheetCount + 1
                    End If
                    ' Взимаме следващия обект
                    itemSheetSet = enumerator.Next()
                Loop
                ' 9. Записваме броя листове като персонализирано свойство на Sheet Set
                SetCustomProperty(sheetSetDatabase.GetSheetSet(),
                              "Общ брой листове",
                              CStr(nSheetCount),
                              PropertyFlags.CUSTOM_SHEETSET_PROP)
                ' 10. Отключваме базата данни след обработката
                LockDatabase(sheetSetDatabase, False)
                ' 11. Нулираме брояча за следващата база данни
                nSheetCount = 0
            Else
                ' 12. Ако не успеем да заключим базата, показваме съобщение
                MsgBox("Unable to access " & sheetSetDatabase.GetSheetSet().GetName())
            End If
            ' 13. Продължаваме към следващата отворена база данни
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
    Private Function ImportASheet(component As IAcSmComponent,
                                  title As String,
                                  description As String,
                                  number As String,
                                  fileName As String,
                                  layout As String) As AcSmSheet
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
            ImportASheet = sheet
        Catch ex As Exception
            ' Показваме съобщение при възникнала грешка
            MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
        End Try
    End Function
End Class