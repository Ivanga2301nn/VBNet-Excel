Imports System.Collections.Generic
Imports System.Diagnostics.Eventing.Reader
Imports System.Drawing.Drawing2D
Imports System.IO
Imports System.Net
Imports System.Security.Cryptography
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.TrackBar
Imports System.Windows.Input
Imports ACSMCOMPONENTS24Lib
Imports Autodesk.AutoCAD
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.DatabaseServices.Filters
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.GraphicsInterface
Imports Autodesk.AutoCAD.Internal
Imports Autodesk.AutoCAD.Runtime
Imports AXDBLib
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application
Imports excel = Microsoft.Office.Interop.Excel

' https://help.autodesk.com/view/OARX/2022/ENU/?guid=GUID-56F608AE-CEB3-471E-8A64-8C909B989F24
' https://adndevblog.typepad.com/autocad/2013/09/using-sheetset-manager-api-in-vbnet.html
Public Class SheetSet
    Dim cu As CommonUtil = New CommonUtil()
    ' Create a new EXCEL sheet set
    Private NoMyExcelProcesses() As Process
    Dim excel_Workbook As excel.Workbook
    Dim wsObekri As excel.Worksheet
    Structure srtSheetSet
        Dim nameSheet As String                 ' Името на групата листи в комплекта от листа
        Dim nameSubSheet As String              ' Името на подлист (ако е приложимо)
        Dim nameLayoutForSheet As String        ' Името на Layout w
        Dim nameLayout As String                ' Името на Layout в Autocad
        Dim Number As Double                    ' Номерът на етажа/котата 
        Dim nameFile As String                  ' Името на Файла в който е листа
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
    ''' <summary>
    ''' Създава или обновява Sheet Set (.dst) за активния проект,
    ''' като записва всички отворени чертежи, открива нови Layout-и
    ''' и ги добавя в подходящи групи и подгрупи.
    ''' </summary>
    <CommandMethod("ADSK_CreateSheetSet")>
    Public Sub ADSK_CreateSheetSet()
        'Dim docs As DocumentCollection = Application.DocumentManager   ' Колекция от всички отворени документи
        '' Записваме всички отворени документи
        'For Each doc As Document In docs
        '    ' Проверка дали документът не е само за четене
        '    If doc.IsReadOnly Then Continue For

        '    Try
        '        Using docLock As DocumentLock = doc.LockDocument()
        '            ' Използваме FullFileName или Filename за по-голяма сигурност
        '            Dim fileName As String = doc.Database.Filename

        '            ' Ако чертежът никога не е записван (Drawing1.dwg), Filename може да е празен
        '            If String.IsNullOrEmpty(fileName) Then fileName = doc.Name

        '            doc.Database.SaveAs(fileName, DwgVersion.Current)
        '            doc.Editor.WriteMessage(vbLf & "Бат Генчо записа: " & doc.Name)
        '        End Using
        '    Catch ex As System.Exception
        '        MsgBox("Грешка при файл: " & doc.Name & vbCrLf & "Път: " & doc.Database.Filename & vbCrLf & "Грешка: " & ex.Message)
        '        ' Тук вече ще хванете специфичното съобщение
        '        doc.Editor.WriteMessage(vbLf & "Грешка при запис на " & doc.Name & ": " & ex.Message)
        '    End Try
        'Next
        ' Даваме време на Windows да опресни файловата система
        System.Windows.Forms.Application.DoEvents()
        ' --- 1. ПЪТИЩА И ИМЕНА ---
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument   ' Активен документ
        Dim name_file As String = acDoc.Name                                    ' Име на DWG файла
        Dim File_Path As String = Path.GetDirectoryName(name_file)              ' Път до папката
        Dim Path_Name As String = Path.GetFileName(File_Path)                   ' Име на папката (име на проекта)

        Dim File_DST As String = Path.Combine(File_Path, Path_Name & ".dst")    ' Пълен път до DST файла
        Dim Set_Desc As String = "Създадено от Бат Генчо"                        ' Описание на Sheet Set-а

        ' --- 2. ДЕКЛАРАЦИИ И РЕЧНИЦИ ---
        Dim arSubset(10, 10) As String                                          ' Дефиниция на подгрупи
        arSubset(0, 0) = "Ел. захранване НН"
        arSubset(1, 0) = "Осветителна инсталация"
        arSubset(2, 0) = "Силова инсталация"
        arSubset(3, 0) = "Слаботокови инсталации"
        arSubset(3, 1) = "Пожароизвестяване"
        arSubset(3, 2) = "Интернет телевизия телефони"
        arSubset(3, 3) = "Сигнално-охранителна"
        arSubset(3, 4) = "Домофонна"
        arSubset(3, 5) = "Оповестителна"
        arSubset(4, 0) = "Заземителна инсталация"
        arSubset(5, 0) = "Мълниезащитна инсталация"
        arSubset(6, 0) = "Еднолинейна схема на"
        arSubset(7, 0) = "Которовки ел. инсталации"
        arSubset(8, 0) = "Кабелни скари и кабелни канали"

        Dim listSheetSet As New List(Of srtSheetSet)                             ' Нови Layout-и за добавяне
        Dim sheetSetManager As IAcSmSheetSetMgr = New AcSmSheetSetMgr            ' Sheet Set Manager
        Dim sheetSetDatabase As AcSmDatabase

        ' Проверка дали DST файлът съществува
        If System.IO.File.Exists(File_DST) Then
            sheetSetDatabase = sheetSetManager.OpenDatabase(File_DST, False)    ' Отваряме съществуващ DST
        Else
            sheetSetDatabase = sheetSetManager.CreateDatabase(File_DST, "", True) ' Създаваме нов DST
        End If
        Dim sheetSet As AcSmSheetSet = sheetSetDatabase.GetSheetSet()            ' Основният Sheet Set

        ' --- 3. РАБОТА С БАЗАТА ДАННИ ---
        Try
            If LockDatabase(sheetSetDatabase, True) = False Then                 ' Заключване за запис
                MsgBox("Sheet set не може да бъде отворен за четене.")
                Exit Sub
            End If
            Dim sheetsInFile As List(Of srtSheetSet) = GetSheetsFromDatabase(sheetSetDatabase) ' Съществуващи Sheet-и


            sheetSet.SetName(Path_Name)                                          ' Име на Sheet Set-а
            sheetSet.SetDesc(Set_Desc)                                           ' Описание на Sheet Set-а
            ' Обхождаме Layout-ите в чертежа
            Using acTrans As Transaction = acDoc.TransactionManager.StartTransaction()
                Dim laye As DBDictionary = acTrans.GetObject(acDoc.Database.LayoutDictionaryId, OpenMode.ForRead)
                For Each Item As DBDictionaryEntry In laye
                    If Item.Key.ToUpper() = "MODEL" OrElse String.IsNullOrEmpty(Item.Key) OrElse Item.Key.Length < 3 Then Continue For
                    Dim currentItem As New srtSheetSet
                    currentItem.nameLayout = Item.Key                            ' Име на Layout-а
                    Dim Instal As String = Item.Key.ToUpper().Substring(0, 3)     ' Код на инсталацията

                    currentItem.nameSheet = If(Installations.ContainsKey(Instal), Installations(Instal), "") ' Основна група
                    currentItem.nameSubSheet = If(Slabotokowa.ContainsKey(Instal), Slabotokowa(Instal), "")   ' Подгрупа

                    ' ЛОГИКА ЗА ИМЕНА (КОТА / ТАБЛО / ЕТАЖ)
                    Dim result As String = ""
                    Select Case True
                        Case Item.Key.ToUpper().Contains("КОТА")
                            Dim kotaIdx = Item.Key.ToUpper().IndexOf("КОТА")
                            Dim pIdx = Item.Key.IndexOf("+")
                            Dim mIdx = Item.Key.IndexOf("-")
                            Select Case True
                                Case pIdx > 0 And pIdx > kotaIdx
                                    result = Item.Key.Substring(pIdx).Trim()
                                Case mIdx > 0 And mIdx > kotaIdx
                                    result = Item.Key.Substring(mIdx).Trim()
                                Case Else
                                    Dim sIdx = Item.Key.Trim().LastIndexOf(" ")
                                    result = If(sIdx > -1, "+" & Item.Key.Substring(sIdx).Trim(), "+0.00")
                            End Select
                            currentItem.nameLayoutForSheet = "Кота " & result
                            currentItem.Number = Val(result.Replace(",", "."))
                        Case Item.Key.ToUpper().Contains("ТАБЛО")
                            Dim lastSpace = Item.Key.Trim().LastIndexOf(" ")
                            result = If(lastSpace > -1, Item.Key.Substring(lastSpace).Trim(), " ")
                            currentItem.nameLayoutForSheet = "Табло ''" & Trim(result) & "''"
                        Case Item.Key.ToUpper().Contains("ЕТАЖ")
                            Dim m As Match = Regex.Match(Item.Key.ToUpper(), "\d+")
                            If m.Success AndAlso Integer.TryParse(m.Value, Nothing) Then
                                currentItem.nameLayoutForSheet = NumberToText(CInt(m.Value))
                            Else
                                currentItem.nameLayoutForSheet = "###### ЕТАЖ"
                            End If
                        Case Item.Key.ToUpper().Contains("Сут")
                            currentItem.nameLayoutForSheet = "Сутерен"
                        Case Else
                            currentItem.nameLayoutForSheet = If(Not String.IsNullOrEmpty(currentItem.nameSheet), currentItem.nameSheet, Item.Key)
                    End Select
                    currentItem.nameFile = name_file                              ' DWG файл
                    If IsNewLayout(name_file, currentItem.nameLayout, sheetsInFile) Then
                        listSheetSet.Add(currentItem)                             ' Добавяме само нови Layout-и
                    End If
                Next
            End Using
            ' --- 5. СОРТИРАНЕ ---
            Dim sortedList As New List(Of srtSheetSet)
            For Each pair In Sheets
                For Each item In listSheetSet
                    If item.nameSheet = pair.Key Then sortedList.Add(item)
                Next
            Next
            ' --- 6. ЗАПИС В DST ---
            Dim mainSubset As AcSmSubset = Nothing
            Dim currentSubset As AcSmSubset = Nothing

            For i As Integer = 0 To sortedList.Count - 1
                Dim current = sortedList(i)
                current.Number = (i + 1).ToString()                               ' Номериране на листовете

                If i = 0 OrElse current.nameSheet <> sortedList(i - 1).nameSheet Then
                    mainSubset = CreateSubset(sheetSetDatabase, current.nameSheet, "", "", "", "", True)
                    currentSubset = mainSubset
                End If

                If Not String.IsNullOrEmpty(current.nameSubSheet) Then
                    If i = 0 OrElse current.nameSubSheet <> sortedList(i - 1).nameSubSheet Then
                        currentSubset = mainSubset.CreateSubset(current.nameSubSheet, "")
                    End If
                Else
                    currentSubset = mainSubset
                End If

                ImportASheet(currentSubset, current.nameLayoutForSheet, "", current.Number, name_file, current.nameLayout)
            Next

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
    ''' Извлича всички Sheet-и от подаден Sheet Set Database (.dst)
    ''' и ги връща като списък от структури srtSheetSet.
    ''' </summary>
    Public Function GetSheetsFromDatabase(sheetSetDatabase As AcSmDatabase) As List(Of srtSheetSet)
        Dim existingSheets As New List(Of srtSheetSet)        ' Списък с вече съществуващите листове
        Dim iter As IAcSmEnumPersist = sheetSetDatabase.GetEnumerator()  ' Enumerator за всички persist обекти в DST
        Dim item As IAcSmPersist = iter.Next()                 ' Вземаме първия елемент от enumerator-а
        ' Обхождаме всички обекти в Sheet Set базата
        While item IsNot Nothing
            ' Проверяваме дали текущият обект е Sheet
            If TypeOf item Is IAcSmSheet Then
                Dim smSheet As IAcSmSheet = DirectCast(item, IAcSmSheet) ' Каст към Sheet
                Dim data As New srtSheetSet                              ' Нов обект за съхранение на данните
                data.Number = smSheet.GetNumber()                        ' Номер на листа
                data.nameLayoutForSheet = smSheet.GetTitle()             ' Заглавие на листа (Sheet Title)
                ' Вземаме референцията към Layout-а в чертежа
                Dim layoutRef As IAcSmAcDbLayoutReference = smSheet.GetLayout()
                If layoutRef IsNot Nothing Then
                    data.nameLayout = layoutRef.GetName()                ' Име на Layout-а
                    data.nameFile = layoutRef.GetFileName()              ' Име на DWG файла
                End If
                existingSheets.Add(data)                                 ' Добавяме листа в списъка
            End If
            item = iter.Next()                                           ' Преминаваме към следващия обект
        End While
        Return existingSheets                                            ' Връщаме списъка с листове
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
    Function NumberToText(ByVal number As Integer) As String
        If numbers.ContainsKey(number) Then
            Return numbers(number)
        Else
            Return "Невалидно число"
        End If
    End Function
    Function TextТоNumber(ByVal floorName As String) As Integer
        ' Convert the floor name to uppercase for case-insensitive comparison
        floorName = floorName.ToUpper()

        ' Iterate through the dictionary and find the corresponding number
        For Each entry In numbers
            If entry.Value = floorName Then
                Return entry.Key
            End If
        Next
        ' If no match is found, return -1
        Return -1
    End Function
    ' Used to add a sheet to a sheet set or subset
    ' Note: This function is dependent on a Default Template and Storage location
    ' being set for the sheet set or subset.
    Private Function AddSheet(ByVal component As IAcSmComponent,
                              ByVal name As String,
                              ByVal description As String,
                              ByVal title As String,
                              ByVal number As String) As AcSmSheet

        Dim sheet As AcSmSheet
        ' Check to see if the component is a sheet set or subset, 
        ' and create the new sheet based on the component's type
        If component.GetTypeName = "AcSmSubset" Then
            Dim subset As AcSmSubset = component
            sheet = subset.AddNewSheet(name, description)
            ' Add the sheet as the first one in the subset
            subset.InsertComponent(sheet, Nothing)
        Else
            sheet = component.GetDatabase().GetSheetSet().AddNewSheet(name, description)
            ' Add the sheet as the first one in the sheet set
            component.GetDatabase().GetSheetSet().InsertComponent(sheet, Nothing)
        End If
        ' Set the number and title of the sheet
        sheet.SetNumber(number)
        sheet.SetTitle(title)
        AddSheet = sheet
    End Function
    ' Used to lock/unlock a sheet set database
    Public Function LockDatabase(
                                 database As AcSmDatabase,
                                 lockFlag As Boolean) As Boolean
        Dim dbLock As Boolean = False
        ' If lockFalg equals True then attempt to lock the database, otherwise
        ' attempt to unlock it.
        If lockFlag = True And
            database.GetLockStatus() = AcSmLockStatus.AcSmLockStatus_UnLocked Then
            database.LockDb(database)
            dbLock = True
        ElseIf lockFlag = False And
            database.GetLockStatus = AcSmLockStatus.AcSmLockStatus_Locked_Local Then
            database.UnlockDb(database)
            dbLock = True
        Else
            dbLock = False
        End If
        LockDatabase = dbLock
    End Function
    ' Used to add a subset to a sheet set
    Private Function CreateSubset(sheetSetDatabase As AcSmDatabase,
                                  name As String,
                                  description As String,
                                  Optional newSheetLocation As String = "",
                                  Optional newSheetDWTLocation As String = "",
                                  Optional newSheetDWTLayout As String = "",
                                  Optional promptForDWT As Boolean = False) As AcSmSubset
        ' Create a subset with the provided name and description
        Dim subset As AcSmSubset = sheetSetDatabase.GetSheetSet().CreateSubset(name, description)
        ' Get the folder the sheet set is stored in
        Dim sheetSetFolder As String
        sheetSetFolder = Mid(sheetSetDatabase.GetFileName(), 1, InStrRev(sheetSetDatabase.GetFileName(), "\"))
        ' Create a reference to a File Reference object
        Dim fileReference As IAcSmFileReference
        fileReference = subset.GetNewSheetLocation()
        ' Check to see if a path was provided, if not default
        ' to the location of the sheet set
        If newSheetLocation <> "" Then
            fileReference.SetFileName(newSheetLocation)
        Else
            fileReference.SetFileName(sheetSetFolder)
        End If
        ' Set the location for new sheets added to the subset
        subset.SetNewSheetLocation(fileReference)
        ' Create a reference to a Layout Reference object
        Dim layoutReference As AcSmAcDbLayoutReference
        layoutReference = subset.GetDefDwtLayout
        ' Check to see that a default DWT location and name was provided
        If newSheetDWTLocation <> "" Then
            ' Set the template location and name of the layout
            ' for the Layout Reference object
            layoutReference.SetFileName(newSheetDWTLocation)
            layoutReference.SetName(newSheetDWTLayout)
            ' Set the Layout Reference for the subset
            subset.SetDefDwtLayout(layoutReference)
        End If
        ' Set the Prompt for Template option of the subset
        subset.SetPromptForDwt(promptForDWT)
        CreateSubset = subset
    End Function
    ' Set the default properties of a sheet set
    Private Sub SetSheetSetDefaults(sheetSetDatabase As AcSmDatabase,
                                    name As String,
                                    description As String,
                                    Optional newSheetLocation As String = "",
                                    Optional newSheetDWTLocation As String = "",
                                    Optional newSheetDWTLayout As String = "",
                                    Optional promptForDWT As Boolean = False)
        ' Set the Name and Description for the sheet set
        sheetSetDatabase.GetSheetSet().SetName(name)
        sheetSetDatabase.GetSheetSet().SetDesc(description)
        ' Check to see if a Storage Location was provided
        If newSheetLocation <> "" Then
            ' Get the folder the sheet set is stored in
            Dim sheetSetFolder As String
            sheetSetFolder = Mid(sheetSetDatabase.GetFileName(), 1, InStrRev(sheetSetDatabase.GetFileName(), "\"))
            ' Create a reference to a File Reference object
            Dim fileReference As IAcSmFileReference
            fileReference = sheetSetDatabase.GetSheetSet().GetNewSheetLocation()
            ' Set the default storage location based on the location of the sheet set
            fileReference.SetFileName(sheetSetFolder)
            ' Set the new Sheet location for the sheet set
            sheetSetDatabase.GetSheetSet().SetNewSheetLocation(fileReference)
        End If
        ' Check to see if a Template was provided
        If newSheetDWTLocation <> "" Then
            ' Set the Default Template for the sheet set
            Dim layoutReference As AcSmAcDbLayoutReference
            layoutReference = sheetSetDatabase.GetSheetSet().GetDefDwtLayout()
            ' Set the template location and name of the layout
            ' for the Layout Reference object
            layoutReference.SetFileName(newSheetDWTLocation)
            layoutReference.SetName(newSheetDWTLayout)
            ' Set the Layout Reference for the sheet set
            sheetSetDatabase.GetSheetSet().SetDefDwtLayout(layoutReference)
        End If
        ' Set the Prompt for Template option of the subset
        sheetSetDatabase.GetSheetSet().SetPromptForDwt(promptForDWT)
    End Sub
    ' Import a sheet into a sheet set or subset
    Private Function ImportASheet(component As IAcSmComponent,
                                  title As String,
                                  description As String,
                                  number As String,
                                  fileName As String,
                                  layout As String) As AcSmSheet
        Try
            If IsNothing(title) Then title = layout
            Dim sheet As AcSmSheet
            ' Create a reference to a Layout Reference object
            Dim layoutReference As New AcSmAcDbLayoutReference
            layoutReference.InitNew(component)
            ' Set the layout and drawing file to use for the sheet
            layoutReference.SetFileName(fileName)
            layoutReference.SetName(layout)
            ' Import the sheet into the sheet set
            ' Check to see if the Component is a Subset or Sheet Set
            If component.GetTypeName = "AcSmSubset" Then
                ' Сложи това на ред 729
                Debug.Print("Опит за импорт: File=" & fileName & " Layout=" & layout)

                ' Провери дали файлът физически съществува в този момент
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
            ' Set the properties of the sheet
            sheet.SetDesc(description)
            sheet.SetTitle(title)
            sheet.SetNumber(number)
            ImportASheet = sheet
        Catch ex As Exception
            ' Показване на съобщение за грешка, ако такава възникне
            MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
        End Try
    End Function
    ' Настройва/създава персонализиран атрибут на лист или набор от листове
    Private Sub SetCustomProperty(owner As IAcSmPersist,
                                  propertyName As String,
                                  propertyValue As Object,
                                  sheetSetFlag As PropertyFlags)
        ' Създава референция към чантата с персонализирани атрибути (Custom Property Bag)
        Dim customPropertyBag As AcSmCustomPropertyBag

        If owner.GetTypeName() = "AcSmSheet" Then
            ' Ако обектът е лист (sheet), получава чантата с персонализирани атрибути за този лист
            Dim sheet As AcSmSheet = owner
            customPropertyBag = sheet.GetCustomPropertyBag()
        Else
            ' Ако обектът е набор от листове (sheet set), получава чантата с персонализирани атрибути за този набор
            Dim sheetSet As AcSmSheetSet = owner
            customPropertyBag = sheetSet.GetCustomPropertyBag()
        End If
        ' Създава референция към персонализирана стойност на атрибут (Custom Property Value)
        Dim customPropertyValue As AcSmCustomPropertyValue = New AcSmCustomPropertyValue()
        customPropertyValue.InitNew(owner)

        ' Задава флаг за атрибута
        customPropertyValue.SetFlags(sheetSetFlag)
        ' Задава стойност за атрибута
        customPropertyValue.SetValue(propertyValue)
        ' Създава атрибута
        customPropertyBag.SetProperty(propertyName, customPropertyValue)
    End Sub
    Private Sub Form_Closed()
        If IsNothing(excel_Workbook) Then
            Exit Sub
        End If
        Try
            excel_Workbook.Save()
            excel_Workbook.Close()
            excel_Workbook = Nothing
            'excel_Workbook.Quit()
        Catch ex As Exception
            MsgBox("Файла вече е затворен")
        End Try
    End Sub
    <CommandMethod("Excel_Name_Progect")>
    Public Sub Excel_Name_Progect()
        Dim name_file As String = Application.DocumentManager.MdiActiveDocument.Name
        Dim File_Path As String = Path.GetDirectoryName(name_file)
        Dim Zapis(18) As String
        Zapis(0) = cu.GetObjects_TEXT("Изберете Наименование на ОБЕКТА")
        Zapis(1) = cu.GetObjects_TEXT("Изберете Местоположение на ОБЕКТА")
        Zapis(2) = cu.GetObjects_TEXT("Изберете ВЪЗЛОЖИТЕЛ на проекта")
        Zapis(3) = cu.GetObjects_TEXT("Изберете СОСТВЕНИК на обекта")
        Zapis(4) = cu.GetObjects_TEXT("Изберете ФАЗА на проекта")
        Zapis(5) = cu.GetObjects_TEXT("Изберете ДАТА на проекта")
        Zapis(6) = cu.GetObjects_TEXT("Изберете съгласувал част АРХИТЕКТУРА")
        Zapis(7) = cu.GetObjects_TEXT("Изберете съгласувал част КОНСТРУКЦИИ")
        Zapis(8) = cu.GetObjects_TEXT("Изберете съгласувал част ТЕХНОЛОГИЯ")
        Zapis(9) = cu.GetObjects_TEXT("Изберете съгласувал част ВиК")
        Zapis(10) = cu.GetObjects_TEXT("Изберете съгласувал част ОВ")
        Zapis(11) = cu.GetObjects_TEXT("Изберете съгласувал част Геодезия")
        Zapis(12) = cu.GetObjects_TEXT("Изберете съгласувал част ВП")
        Zapis(13) = cu.GetObjects_TEXT("Изберете съгласувал част ЕЕФ")
        Zapis(14) = cu.GetObjects_TEXT("Изберете съгласувал част ПБ")
        Zapis(15) = cu.GetObjects_TEXT("Изберете съгласувалчаст ПБЗ")
        Zapis(16) = cu.GetObjects_TEXT("Изберете съгласувал част ПУСО")
        Zapis(17) = "инж. М.Тонкова-Генчева"
        Zapis(18) = File_Path
        '
        '
        '
        Dim nameExcel As String = "\\MONIKA\Monika\_НАСТРОЙКИ\Обекти.xlsx"
        '
        ' Проверява дали EXCEL е отворен
        '
        Dim stream As FileStream = Nothing
        Try
            stream = File.Open(nameExcel, FileMode.Open, FileAccess.ReadWrite, FileShare.None)
            stream.Close()
        Catch ex As Exception
            MsgBox("Отворен е файл с име : " + Chr(13) + Chr(13) +
                   nameExcel + Chr(13) + Chr(13) +
                   "Моля затворете го преди да продължите!")
            Exit Sub
        End Try
        '
        'Get all currently running process Ids for Excel applications
        '
        NoMyExcelProcesses = Process.GetProcessesByName("Excel")

        Dim objExcel As excel.Application = New excel.Application()
        excel_Workbook = objExcel.Workbooks.Open(nameExcel)

        objExcel.Visible = vbTrue
        wsObekri = excel_Workbook.Worksheets("Обекти")
        Dim Red As Integer
        For i As Integer = 2 To 10000
            If Len(wsObekri.Range("A" & i.ToString).Value) = 0 Then
                Red = i
                Exit For
            End If
        Next
        For i = 0 To UBound(Zapis)
            wsObekri.Cells(Red, i + 1).Value = Zapis(i)
        Next

        ' Close the EXCEL
        Form_Closed()
    End Sub
    ' Counts up the sheets for all the open sheet sets
    <CommandMethod("ADSK_SetSheetCount")>
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
    ' Обход на всички отворени набори от листове
    <CommandMethod("ADSK_StepThroughTheOpenSheetSets")>
    Public Sub StepThroughTheOpenSheetSets()
        ' Взема референция към Sheet Set Manager обекта
        Dim sheetSetManager As IAcSmSheetSetMgr
        sheetSetManager = New AcSmSheetSetMgr

        ' Взема заредените бази данни
        Dim enumDatabase As IAcSmEnumDatabase
        enumDatabase = sheetSetManager.GetDatabaseEnumerator()

        ' Взема първата отворена база данни
        Dim item As IAcSmPersist
        item = enumDatabase.Next()
        Dim customMessage As String = ""
        ' Ако има отворена база — продължи
        If Not item Is Nothing Then
            Dim count As Integer = 0
            ' Обход на енумератора на базите данни
            Do While Not item Is Nothing
                ' Добавя името на файла на отворения sheet set към изходния низ
                customMessage = customMessage + vbLf +
                            item.GetDatabase().GetFileName()
                ' Взема следващата отворена база и увеличава брояча
                item = enumDatabase.Next()
                count = count + 1
            Loop
            customMessage = "Sheet sets open: " + count.ToString() +
                        customMessage
        Else
            customMessage = "No sheet sets are currently open."
        End If
        ' Показва съобщението
        MsgBox(customMessage)
    End Sub
    ' Синхронизира свойствата на лист с тези на набора от листове
    Private Sub SyncProperties(ByVal sheetSetDatabase As IAcSmDatabase)
        ' Вземи обектите в набора от листове
        Dim enumerator As IAcSmEnumPersist = sheetSetDatabase.GetEnumerator()
        ' Вземи първия обект от енумератора
        Dim item As IAcSmPersist
        item = enumerator.Next()
        ' Премини през всички обекти в набора от листове
        Do While Not item Is Nothing
            Dim sheet As IAcSmSheet = Nothing
            ' Провери дали обектът е лист
            If item.GetTypeName() = "AcSmSheet" Then
                sheet = item
                ' Създай референция към енумератора на свойства за 
                ' чантата с персонализирани свойства
                Dim enumeratorProperty As IAcSmEnumProperty
                enumeratorProperty = item.GetDatabase().GetSheetSet().GetCustomPropertyBag().GetPropertyEnumerator()
                ' Вземи стойностите от набора от листове, за да ги прехвърлиш към листовете
                Dim name As String = ""
                Dim customPropertyValue As AcSmCustomPropertyValue = Nothing
                ' Вземи първото свойство
                enumeratorProperty.Next(name, customPropertyValue)
                ' Премини през всяко от свойствата
                Do While Not customPropertyValue Is Nothing
                    ' Провери дали свойството е за лист
                    If customPropertyValue.GetFlags() =
                    PropertyFlags.CUSTOM_SHEET_PROP Then
                        SetCustomProperty(sheet, name, customPropertyValue.GetValue(), customPropertyValue.GetFlags())
                    End If
                    ' Вземи следващото свойство
                    enumeratorProperty.Next(name, customPropertyValue)
                Loop
            End If
            ' Вземи следващия лист
            item = enumerator.Next()
        Loop
    End Sub
End Class
