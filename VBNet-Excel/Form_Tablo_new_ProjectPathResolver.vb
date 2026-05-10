
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Newtonsoft.Json
Imports System.IO
Imports System.Text.RegularExpressions
''' <summary>
''' Клас за управление на проектни данни за електрически табла.
''' Отговаря за:
''' - Зареждане на данни от JSON файлове
''' - Запис на данни в JSON файлове
''' - Генериране на пътища до проектните файлове
''' - Извличане на име на проект/сграда от DWG файл
''' - Възстановяване и поправка на данни при зареждане
''' 
''' Класът служи като централизирана система за
''' сериализация и управление на проектната информация,
''' свързана с ListTokow и структурата на таблата.
''' </summary>
Public Class Form_Tablo_new_ProjectPathResolver
    ''' <summary>
    ''' Извлича име на сграда от името на DWG файла.
    ''' Логиката:
    ''' 1. Взима името на файла без разширението
    ''' 2. Търси шаблон за сграда с номер
    ''' 3. Ако няма шаблон → търси произволно число
    ''' 4. Ако няма число → генерира безопасно име
    ''' 5. Връща крайното име на проекта
    ''' </summary>
    Public Function GetBuildingNameFromDwg(dwgFullPath As String) As String
        ' Взимаме името на файла без разширението
        Dim fileName As String =
        Path.GetFileNameWithoutExtension(dwgFullPath)
        ' Шаблон за търсене:
        ' сграда_1
        ' sgr-2
        ' blk 3
        ' bldg4
        Dim pattern As String =
        "(?i)(?:сграда|sgr|block|blk|bldg)[\s_-]*(\d+)"
        ' Търсим съвпадение по шаблона
        Dim match As Match = Regex.Match(fileName, pattern)
        ' Ако намерим номер на сграда
        If match.Success AndAlso match.Groups.Count > 1 Then Return $"Сграда_{match.Groups(1).Value}"
        ' Ако няма шаблон → търсим произволно число
        Dim digitsPattern As String = "\d+"
        Dim digitMatch As Match =
        Regex.Match(fileName, digitsPattern)
        ' Ако намерим число
        If digitMatch.Success Then Return $"Сграда_{digitMatch.Value}"
        ' Ако няма числа → създаваме безопасно име
        ' Премахваме неподходящите символи
        Dim safeName As String =
        Regex.Replace(fileName, "[^a-zA-Z0-9_-]", "_")
        ' Ако резултатът е празен → използваме резервно име
        If String.IsNullOrEmpty(safeName) Then safeName = "Project"
        ' Връщаме безопасното име
        Return safeName
    End Function
    ''' <summary>
    ''' Генерира пълния път до JSON файла за текущия проект.
    ''' Логиката:
    ''' 1. Извлича папката на DWG файла
    ''' 2. Проверява дали пътят е валиден
    ''' 3. Извлича името на сградата от DWG файла
    ''' 4. Използва резервни стойности при липсващи данни
    ''' 5. Генерира крайния път до JSON файла
    ''' </summary>
    Public Function GetJsonTargetPath(dwgFullPath As String) As String
        ' Извличаме папката на DWG файла
        Dim directory As String = Path.GetDirectoryName(dwgFullPath)
        ' Ако пътят е невалиден → използваме текущата директория
        If String.IsNullOrEmpty(directory) Then directory = Environment.CurrentDirectory
        ' Извличаме името на сградата от DWG файла
        Dim buildingName As String = GetBuildingNameFromDwg(dwgFullPath)
        ' Ако няма валидно име → използваме резервно име
        If String.IsNullOrEmpty(buildingName) Then buildingName = "Project"
        ' Генерираме и връщаме пълния път до JSON файла
        Return Path.Combine(directory, $"{buildingName}_Tokowi.json")
    End Function
    ''' <summary>
    ''' Записва проекта в JSON файл.
    ''' Логиката:
    ''' 1. Определя пътя до JSON файла
    ''' 2. Проверява дали папката съществува
    ''' 3. Сериализира данните в JSON формат
    ''' 4. Записва JSON файла на диска
    ''' 5. Връща True при успешен запис
    ''' 6. Връща False при грешка
    ''' </summary>
    Public Function SaveProject(data As List(Of Form_Tablo_new.strTokow),
                            dwgFullPath As String) As Boolean

        Try
            ' Определяме пътя до JSON файла
            Dim targetPath As String = GetJsonTargetPath(dwgFullPath)
            ' Ако пътят е невалиден → прекратяваме
            If String.IsNullOrEmpty(targetPath) Then Return False
            ' Взимаме папката на файла
            Dim dir As String = IO.Path.GetDirectoryName(targetPath)
            ' Ако папката не съществува → прекратяваме
            If Not IO.Directory.Exists(dir) Then Return False
            ' Сериализираме данните в JSON формат
            Dim json As String =
            JsonConvert.SerializeObject(data, Formatting.Indented)
            ' Записваме JSON файла на диска
            IO.File.WriteAllText(targetPath, json)
            ' Успешен запис
            Return True
        Catch
            ' При грешка връщаме False
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Зарежда проект от JSON файл и възстановява данните в ListTokow.
    ''' Логиката:
    ''' 1. Определя пътя до JSON файла
    ''' 2. Проверява дали файлът съществува
    ''' 3. Прочита съдържанието на файла
    ''' 4. Десериализира JSON данните
    ''' 5. Подготвя и поправя заредените записи
    ''' 6. При грешка запазва текущите данни без промяна
    ''' </summary>
    Public Sub LoadProject(ByRef targetList As List(Of Form_Tablo_new.strTokow),
                       dwgFullPath As String,
                       acDb As Database)
        Try
            ' Определяме пътя до JSON файла за текущия DWG
            Dim targetPath As String = GetJsonTargetPath(dwgFullPath)
            ' Ако няма валиден път или файлът не съществува
            If String.IsNullOrEmpty(targetPath) OrElse
           Not IO.File.Exists(targetPath) Then
                ' Стартираме обработка с празни данни
                ProcessAndRepairList(targetList, Nothing, acDb)
                Return
            End If
            ' Прочитаме съдържанието на JSON файла
            Dim json As String = IO.File.ReadAllText(targetPath)
            ' Ако файлът е празен
            If String.IsNullOrEmpty(json) Then
                ' Стартираме обработка с празни данни
                ProcessAndRepairList(targetList, Nothing, acDb)
                Return
            End If
            ' Десериализираме JSON данните към списък от strTokow
            Dim loadedData As List(Of Form_Tablo_new.strTokow) =
            JsonConvert.DeserializeObject(Of List(Of Form_Tablo_new.strTokow))(json)
            ' Обработваме и поправяме заредените данни
            ProcessAndRepairList(targetList, loadedData, acDb)
        Catch ex As Exception
            ' Ако възникне грешка:
            ' - НЕ променяме targetList
            ' - Старите данни остават запазени
            ' - Програмата продължава работа без загуба на информация
        End Try
    End Sub
    ''' <summary>
    ''' Обработва заредените данни: изчиства старите, добавя новите и оправя ID-тата.
    ''' ТОВА Е МЯСТОТО ЗА БЪДЕЩИ ПРОВЕРКИ (версии, статуси, merge логика).
    ''' </summary>
    Private Sub ProcessAndRepairList(ByRef targetList As List(Of Form_Tablo_new.strTokow),
                                     sourceList As List(Of Form_Tablo_new.strTokow),
                                     acDb As Database)
        Try
            ' 1. Изчистване на текущия списък (или бъдещ Merge логика тук)
            'targetList.Clear()
            ' 2. Добавяне на новите данни
            If sourceList IsNot Nothing Then
                targetList.AddRange(sourceList)
            End If
            ' 3. ОПРАВЯНЕ НА ID-ТАТА (Handle -> ObjectId)
            If acDb IsNot Nothing Then
                For Each t In targetList
                    If t.Konsumator Is Nothing Then Continue For
                    For Each k In t.Konsumator
                        If Not String.IsNullOrEmpty(k.Handle_Block) Then
                            Try
                                Dim h As New Handle(Convert.ToInt64(k.Handle_Block, 16))
                                k.ID_Block = acDb.GetObjectId(False, h, 0)
                            Catch
                                k.ID_Block = ObjectId.Null
                            End Try
                        End If
                    Next
                Next
            End If
        Catch ex As Exception
            ' АКО ИМА ГРЕШКА: НЕ пипаме targetList. Той остава със старите си данни.
            ' Програмата продължава нормално, без загуба на работа.
        End Try
    End Sub
End Class
