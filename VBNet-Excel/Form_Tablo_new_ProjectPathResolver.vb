Imports System.Collections.Generic
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Linq
Imports System.Security.Cryptography
Imports System.Security.Cryptography.X509Certificates
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Autodesk.AutoCAD
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.DatabaseServices.Filters
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.GraphicsSystem
Imports Autodesk.AutoCAD.Internal.DatabaseServices
Imports Autodesk.AutoCAD.PlottingServices
Imports Autodesk.AutoCAD.Runtime
Imports AXDBLib
Imports iTextSharp.text.pdf
Imports Microsoft.Office.Interop.Word
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Org.BouncyCastle.Asn1.Cmp
Imports Org.BouncyCastle.Math.EC.ECCurve
Imports Button = System.Windows.Forms.Button
Imports Font = System.Drawing.Font
Imports System.IO
Imports System.Text.RegularExpressions





Public Class Form_Tablo_new_ProjectPathResolver






    ''' <summary>
    ''' Извлича името на сградата от пълния път на DWG файла.
    ''' Връща чист текст, готов за използване в имена на файлове.
    ''' </summary>
    Public Function GetBuildingNameFromDwg(dwgFullPath As String) As String
        ' 1. Вземи само името на файла (без пътя и разширението)
        Dim fileName As String = Path.GetFileNameWithoutExtension(dwgFullPath)
        ' 2. Търси модел "Сграда" + число (напр. Сграда_1, Sgr2, Block 3)
        ' Regex опции: IgnoreCase (да не прави разлика между главни/малки букви)
        Dim pattern As String = "(?i)(?:сграда|sgr|block|blk|bldg)[\s_-]*(\d+)"
        Dim match As Match = Regex.Match(fileName, pattern)
        If match.Success AndAlso match.Groups.Count > 1 Then
            ' Ако намерим (напр. "Project_Sgr1"), връщаме "Сграда_1"
            Return $"Сграда_{match.Groups(1).Value}"
        End If
        ' 3. Ако няма ключова дума, търси просто число в името (напр. "DWG_02")
        Dim digitsPattern As String = "\d+"
        Dim digitMatch As Match = Regex.Match(fileName, digitsPattern)
        If digitMatch.Success Then
            Return $"Сграда_{digitMatch.Value}"
        End If
        ' 4. Fallback: Ако не намерим нищо, използваме името на файла (почистено)
        Dim safeName As String = Regex.Replace(fileName, "[^a-zA-Z0-9_-]", "_")
        If String.IsNullOrEmpty(safeName) Then safeName = "Project"
        Return safeName
    End Function
    ''' <summary>
    ''' Генерира пълен път до JSON файла за запазване/зареждане.
    ''' Формат: "{ПапкаНаDWG}\{ИмеНаСграда}_Tokowi.json"
    ''' </summary>
    Public Function GetJsonTargetPath(dwgFullPath As String) As String
        ' 1. Извличаме папката, в която се намира DWG файлът
        Dim directory As String = Path.GetDirectoryName(dwgFullPath)
        If String.IsNullOrEmpty(directory) Then
            directory = Environment.CurrentDirectory ' Fallback, ако пътят е невалиден
        End If
        ' 2. Извличаме името на сградата чрез първия метод
        Dim buildingName As String = GetBuildingNameFromDwg(dwgFullPath)
        If String.IsNullOrEmpty(buildingName) Then
            buildingName = "Project" ' Fallback, ако не успеем да разчетем име
        End If
        ' 3. Комбинираме в краен път: C:\...\Папка\Сграда_1_Tokowi.json
        Return Path.Combine(directory, $"{buildingName}_Tokowi.json")
    End Function
    ''' <summary>
    ''' Записва списъка с данни в JSON файл по подаден DWG път.
    ''' Връща True при успешен запис, False при грешка.
    ''' </summary>
    Public Function SaveProject(data As List(Of Form_Tablo_new.strTokow), dwgFullPath As String) As Boolean
        Try
            ' 1. Генерираме целевия път
            Dim targetPath As String = GetJsonTargetPath(dwgFullPath)
            If String.IsNullOrEmpty(targetPath) Then Return False
            ' 2. Проверяваме дали папката съществува (поправен синтаксис)
            Dim dir As String = IO.Path.GetDirectoryName(targetPath)
            If Not IO.Directory.Exists(dir) Then Return False
            ' 3. Сериализираме и записваме
            Dim json As String = JsonConvert.SerializeObject(data, Formatting.Indented)
            IO.File.WriteAllText(targetPath, json)
            Return True
        Catch
            ' При всяка грешка (достъп, сериализация, диск) връщаме False
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Чете JSON файла и извиква процедурата за обработка и поправка на данните.
    ''' </summary>
    Public Sub LoadProject(ByRef targetList As List(Of Form_Tablo_new.strTokow), dwgFullPath As String, acDb As Database)
        Try
            Dim targetPath As String = GetJsonTargetPath(dwgFullPath)
            ' Ако няма файл, просто изчистваме списъка и излизаме
            If String.IsNullOrEmpty(targetPath) OrElse Not IO.File.Exists(targetPath) Then
                ProcessAndRepairList(targetList, Nothing, acDb)
                Return
            End If
            ' 1. Четем файла
            Dim json As String = IO.File.ReadAllText(targetPath)
            If String.IsNullOrEmpty(json) Then
                ProcessAndRepairList(targetList, Nothing, acDb)
                Return
            End If
            ' 2. Десериализираме във временен списък
            Dim loadedData As List(Of Form_Tablo_new.strTokow) = JsonConvert.DeserializeObject(Of List(Of Form_Tablo_new.strTokow))(json)
            ' 3. Извикваме новата процедура за обработка и поправка на ID-та
            ProcessAndRepairList(targetList, loadedData, acDb)
        Catch ex As Exception
            ' При фатална грешка изчистваме списъка
            targetList.Clear()
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
            targetList.Clear()
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
