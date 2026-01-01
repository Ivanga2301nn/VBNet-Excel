Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports System.Collections.Generic
Imports System.Linq

Public Class AutoCADAPI
    <CommandMethod("Scale_Layouts")>
    Public Sub ProcessLayouts()
        ' Вземаме текущия документ и база данни
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Try
            ' Използваме Transaction, за да работим с базата данни
            Using tr As Transaction = db.TransactionManager.StartTransaction()
                ' Вземаме LayoutDictionary, който съдържа информация за всички Layouts
                Dim layoutDict As DBDictionary = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead)
                For Each layoutEntry As DBDictionaryEntry In layoutDict
                    ' Получаваме Layout обекта от идентификатора
                    Dim layout As Layout = tr.GetObject(layoutEntry.Value, OpenMode.ForRead)
                    Dim layoutName As String = layout.LayoutName

                    ' Пропускаме Layout-ите, които започват с "Настрой" или "Model"
                    If layoutName.StartsWith("Настрой") Or layoutName.StartsWith("Model") Then Continue For

                    ' Вземаме BlockTableRecord, който съдържа всички блокове в Layout
                    Dim btr As BlockTableRecord = tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead)
                    Dim annotationScales As New HashSet(Of String)()
                    Dim viewports As New List(Of Viewport)()

                    ' Обхождаме всички обекти в BlockTableRecord
                    For Each objId As ObjectId In btr
                        Dim ent As Entity = tr.GetObject(objId, OpenMode.ForRead)
                        ' Проверяваме дали елементът е Viewport
                        If TypeOf ent Is Viewport Then
                            Dim vp As Viewport = CType(ent, Viewport)
                            Dim scaleName As String = vp.AnnotationScale.Name
                            ' Отваряме Viewport за писане
                            vp.UpgradeOpen()
                            ' Изключваме мащаб 1:1 от списъка на уникални мащаби
                            If scaleName <> "1:1" Then
                                annotationScales.Add(scaleName)
                                ' Прехвърляме Viewport в слой Defpoints
                                vp.Layer = "Defpoints"
                            Else
                                ' Прехвърляме Viewport в слой EL_РАМКА
                                vp.Layer = "EL_РАМКА"
                            End If
                            ' Затваряме Viewport за писане
                            vp.DowngradeOpen()
                            viewports.Add(vp)
                        End If
                    Next

                    ' Ако има повече от един Viewport, питаме потребителя да избере AnnotationScale
                    Dim selectedScale As String = ""
                    If annotationScales.Count > 1 Then
                        ' Създаваме PromptKeywordOptions с предварително зададени ключови стойности
                        Dim pDouOpts As PromptKeywordOptions = New PromptKeywordOptions("")
                        pDouOpts.Message = $"За Layout: {layoutName} Изберете МАЩАБ от следните: "
                        pDouOpts.AllowNone = False

                        ' Добавяме ключови стойности (мащаби) към опцията
                        For Each scale In annotationScales
                            pDouOpts.Keywords.Add(scale)
                        Next

                        ' Задаваме ключова стойност по подразбиране, ако има такава
                        If annotationScales.Count > 0 Then
                            pDouOpts.Keywords.Default = annotationScales.First()
                        End If

                        ' Получаваме избора на потребителя
                        Dim result = ed.GetKeywords(pDouOpts)

                        If result.Status = PromptStatus.OK AndAlso annotationScales.Contains(result.StringResult) Then
                            selectedScale = result.StringResult
                        Else
                            ed.WriteMessage(vbLf & "Невалиден избор или не беше направен избор.")
                            Continue For
                        End If
                    ElseIf annotationScales.Count = 1 Then
                        ' Ако има само една стойност, използваме я автоматично
                        selectedScale = annotationScales.First()
                    End If

                    ' Обхождаме отново блоковете в Layout-а, за да актуализираме атрибутите
                    For Each objId As ObjectId In btr
                        Dim ent As Entity = tr.GetObject(objId, OpenMode.ForRead)
                        ' Проверяваме дали елементът е BlockReference
                        If TypeOf ent Is BlockReference Then
                            Dim blockRef As BlockReference = CType(ent, BlockReference)
                            ' Извличаме атрибутите на блока
                            For Each attId As ObjectId In blockRef.AttributeCollection
                                Dim attRef As AttributeReference = tr.GetObject(attId, OpenMode.ForWrite)
                                ' Ако атрибутът е "МАЩАБ"
                                If attRef.Tag.ToUpper() = "МАЩАБ" Then
                                    ' Променяме слоя на блока на "EL_РАМКА"
                                    blockRef.UpgradeOpen() ' Отваряме блока за писане
                                    blockRef.Layer = "EL_РАМКА"
                                    blockRef.DowngradeOpen() ' Затваряме блока за писане
                                    If layoutName.StartsWith("Табло") Then
                                        ' Ако Layout започва с "Табло", задаваме "---"
                                        attRef.TextString = "----"
                                    Else
                                        ' Ако Layout не започва с "Табло", извличаме числото след двоеточието от selectedScale
                                        Dim scaleValueStr As String = selectedScale.Split(":"c).Last()
                                        ' Използваме CultureInfo за правилно парсване на дробните числа със запетая
                                        Dim cultureInfo As System.Globalization.CultureInfo = System.Globalization.CultureInfo.InvariantCulture
                                        ' Преобразуваме извлеченото число в тип Double и го умножаваме по 10
                                        Dim scaleValue As Double = Double.Parse(scaleValueStr.Replace(",", "."), cultureInfo)
                                        ' Записваме резултата в атрибута в формат "1:{число}"
                                        attRef.TextString = $"1:{(scaleValue * 10).ToString(cultureInfo)}"
                                    End If
                                End If
                            Next
                        End If
                    Next
                Next
                ' Финализираме транзакцията
                tr.Commit()
            End Using
        Catch ex As Exception
            ' Показваме съобщение за грешка, ако такава възникне
            MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
        End Try
    End Sub
End Class


