Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.TrackBar
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.DatabaseServices.Filters
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Microsoft.Office.Interop
Imports excel = Microsoft.Office.Interop.Excel
Public Class Insert_Signature
    Dim cu As CommonUtil = New CommonUtil()
    Dim Zapis(26) As String
    Dim File_Path As String
    ' --- Създаване на речника ---
    Dim Data As New Dictionary(Of String, String)

    Public Shared Sub Set_Signature(blockNames As String())
        '
        ' Този код клонира блокове от един DWG файл в друг, като използва имената на блоковете, предоставени в масива blockNames.
        ' Ако име на блок не съществува в изходния DWG файл, то той просто се пропуска.
        ' В случай на грешка, кодът показва съобщение за грешка с подробности за проблема.
        '
        Try
            ' Получаване на активния документ
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = doc.Editor
            ' Отваряне на базата данни
            Using OpenDb As Database = New Database(False, True)
                ' Четене на DWG файла
                OpenDb.ReadDwgFile("\\MONIKA\Monika\_НАСТРОЙКИ\PODPISI.dwg", System.IO.FileShare.ReadWrite, True, "")
                ' Създаване на колекция от ObjectId
                Dim ids As ObjectIdCollection = New ObjectIdCollection()
                Dim DocblockNames As New List(Of String)
                'Обхождане на всички блоковете в базата данни
                Using tr As Transaction = OpenDb.TransactionManager.StartTransaction()
                    Dim bt As BlockTable = tr.GetObject(OpenDb.BlockTableId, OpenMode.ForRead)
                    For Each btrId As ObjectId In bt
                        Dim btr As BlockTableRecord = tr.GetObject(btrId, OpenMode.ForRead)
                        If Not btr.IsAnonymous And Not btr.Name.StartsWith("*") Then
                            'Добавяне на името на блока в масива
                            DocblockNames.Add(btr.Name)
                        End If
                    Next
                    tr.Commit()
                End Using
                Using tran As Transaction = OpenDb.TransactionManager.StartTransaction()
                    ' Получаване на таблицата с блокове
                    Dim bt As BlockTable
                    bt = CType(tran.GetObject(OpenDb.BlockTableId, OpenMode.ForRead), BlockTable)
                    ' Обхождане на всички имена на блокове
                    For Each blockName In blockNames
                        Dim engineerIndex As Integer = blockName.ToLower().IndexOf("инж.")
                        Dim architectIndex As Integer = blockName.ToLower().IndexOf("арх.")
                        If Not (engineerIndex >= 0 Or architectIndex >= 0) Then Continue For
                        If Not bt.Has(blockName) Then
                            blockName = SelectBlock(DocblockNames, blockName)
                        End If
                        If blockName <> "" Then
                            ids.Add(bt(blockName))
                        End If
                    Next
                    ' Приключване на транзакцията
                    tran.Commit()
                End Using
                ' Ако има намерени блокове
                If ids.Count <> 0 Then
                    ' Получаване на текущата база данни
                    Dim destdb As Database = doc.Database
                    ' Създаване на нова карта за клониране
                    Dim iMap As IdMapping = New IdMapping()
                    ' Клониране на обектите в текущата база данни
                    destdb.WblockCloneObjects(ids, destdb.BlockTableId, iMap, DuplicateRecordCloning.Ignore, False)
                End If
            End Using
        Catch ex As Exception
            ' Показване на съобщение за грешка, ако такава възникне
            MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
        End Try
    End Sub
    Public Shared Function SelectBlock(blockNamesIds As List(Of String), blockNames As String) As String
        Dim result As String = ""
        ' Намиране на позицията на "&"
        Dim indexOfAmpersand As Integer = blockNames.IndexOf("&")

        ' Извличане на текст
        Dim selectedText As String
        If indexOfAmpersand >= 0 Then
            selectedText = blockNames.Substring(indexOfAmpersand + 1)
        Else
            Return result
        End If

        ' Претърсване на blockNamesIds
        Dim selectedBlocks As New List(Of String)
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim pDouOpts As PromptDoubleOptions = New PromptDoubleOptions("")
        For Each blockId In blockNamesIds
            Dim sss As String = blockId
            If blockId.ToLower().Contains(selectedText.ToLower()) Then
                pDouOpts.Keywords.Add(blockId)
            End If
        Next

        With pDouOpts
            .Message = vbCrLf & "Изберете: "
            .AllowZero = False
            .AllowNegative = False
        End With
        Dim pKeyRes As PromptDoubleResult = acDoc.Editor.GetDouble(pDouOpts)
        If pKeyRes.Status = PromptStatus.Keyword Then
            result = pKeyRes.StringResult
        Else
            result = pKeyRes.Value
        End If
        Return result
    End Function

    ' Чете данните за обекта и ги записва в блок Insert_Signature
    <CommandMethod("Insert_Signature")>
    Public Sub Insert_Signature()
        Dim name_file As String = Application.DocumentManager.MdiActiveDocument.Name
        File_Path = Path.GetDirectoryName(name_file)
        Dim File_name As String = Path.GetFileName(name_file)
        Dim Path_Name As String = Mid(File_Path, InStrRev(File_Path, "\") + 1, Len(File_Path))
        ' --- Основни полета ---
        Data("ОБЕКТ") = cu.GetObjects_TEXT("Изберете Наименование на ОБЕКТА", vbFalse)
        If Data("ОБЕКТ") = "  #####  " Then
            MsgBox("Няма избран текст за Наименование на ОБЕКТА.")
            Exit Sub
        End If
        Data("МЕСТОПОЛОЖЕНИЕ") = cu.GetObjects_TEXT("Изберете Местоположение на ОБЕКТА", vbFalse)
        Data("ВЪЗЛОЖИТЕЛ") = cu.GetObjects_TEXT("Изберете ВЪЗЛОЖИТЕЛ на проекта", vbFalse)
        Data("СОБСТВЕНИК") = cu.GetObjects_TEXT("Изберете СОСТВЕНИК на обекта", vbFalse)
        Data("ФАЗА") = cu.GetObjects_TEXT("Изберете ФАЗА на проекта", vbFalse)
        Data("ДАТА") = cu.GetObjects_TEXT("Изберете ДАТА на проекта", vbFalse)
        Data("АРХИТЕКТ") = cu.GetObjects_TEXT("Изберете съгласувал част АРХИТЕКТУРА")
        Data("КОНСТРУКТОР") = cu.GetObjects_TEXT("Изберете съгласувал част КОНСТРУКЦИИ")
        Data("ТЕХНОЛОГИЯ") = cu.GetObjects_TEXT("Изберете съгласувал част ТЕХНОЛОГИЯ")
        Data("ВИК") = cu.GetObjects_TEXT("Изберете съгласувал част ВиК")
        Data("ОВ") = cu.GetObjects_TEXT("Изберете съгласувал част ОВ")
        Data("ГЕОДЕЗИЯ") = cu.GetObjects_TEXT("Изберете съгласувал част Геодезия")
        Data("ВП") = cu.GetObjects_TEXT("Изберете съгласувал част ВП")
        Data("ЕЕФ") = cu.GetObjects_TEXT("Изберете съгласувал част ЕЕФ")
        Data("ПБ") = cu.GetObjects_TEXT("Изберете съгласувал част ПБ")
        Data("ПБЗ") = cu.GetObjects_TEXT("Изберете съгласувалчаст ПБЗ")
        Data("ПУСО") = cu.GetObjects_TEXT("Изберете съгласувал част ПУСО")
        ' Проектант с дефолт
        Dim projektant As String = cu.GetObjects_TEXT("Изберете ПРОЕКТАНТ")
        Data("ПРОЕКТАНТ") = If(projektant = "  #####  ", "инж. М.Тонкова-Генчева", projektant)
        ' Път към файл
        Data("FILE_PATH") = File_Path
        ' Допълнителни полета
        Data("Ном.заявление") = cu.GetObjects_TEXT("Изберете Номер на заявлене")
        Data("Дата_заявление") = cu.GetObjects_TEXT("Изберете Дата на изготвяне")
        Data("SAP") = cu.GetObjects_TEXT("Изберете съгласувал SAP номер")
        ' Дружество с дефолт
        Dim drujestvo As String = cu.GetObjects_TEXT("Изберете дружество")
        Data("Дружество") = If(drujestvo = "  #####  ",
                       Chr(34) & "Електроразпределителни мрежи " & Chr(34) & "Запад" & Chr(34) & " ЕАД",
                       drujestvo)
        Data("ТОЧКА_3") = cu.GetObjects_TEXT("Изберете текст съдържащ описанието по т.3 от становището")
        Data("ИМЕ") = cu.GetObjects_TEXT("Изберете Име")
        Data("АДРЕС") = cu.GetObjects_TEXT("Изберете Адрес")
        Data("ПАРТИДА") = cu.GetObjects_TEXT("Изберете Партида")
        Try
            ' Получаване на активния документ
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            ' Получаване на базата данни на активния документ
            Dim acCurDb As Database = acDoc.Database
            ' Започване на транзакция
            Using actrans As Transaction = acDoc.TransactionManager.StartTransaction()
                ' Отваряне на LayoutDictionary
                Dim layoutDict As DBDictionary = CType(actrans.GetObject(acCurDb.LayoutDictionaryId, OpenMode.ForRead), DBDictionary)
                ' Инициализиране на брояч за layout-ите
                Dim layoutCount As Integer = 0
                ' Инициализиране на списък, който ще съдържа имената на layout-ите
                Dim layoutNamesList As New List(Of String)()
                ' Обхождане на layout-ите в речника
                For Each entry As DBDictionaryEntry In layoutDict
                    ' Получаване на Layout обекта
                    Dim layout As Layout = CType(actrans.GetObject(entry.Value, OpenMode.ForRead), Layout)
                    If layout IsNot Nothing Then
                        ' Вземане на името на Layout-а
                        Dim layoutName As String = layout.LayoutName
                        ' Проверка дали името на Layout не е "Model" и не започва с "Настройки"
                        If layoutName <> "Model" AndAlso Not layoutName.StartsWith("Настройки") Then
                            ' Увеличаване на брояча и добавяне на името в списъка
                            layoutCount += 1
                            layoutNamesList.Add(layoutName)
                        End If
                    End If
                Next
                ' Конвертиране на списъка в масив
                Dim layoutNamesArray As String() = layoutNamesList.ToArray()
                ' Получаване на таблицата на блоковете
                Dim acBlkTbl As BlockTable = actrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                ' Получаване на ID на записа на блока "Insert_Signature" в таблицата на блоковете
                Dim blkRecId As ObjectId = acBlkTbl("Insert_Signature")
                ' Получаване на записа на блока
                Dim acBlkTblRec As BlockTableRecord = actrans.GetObject(blkRecId, OpenMode.ForRead)
                ' Обхождане на всички блокови референции за блока "Insert_Signature"
                For Each blkRefId As ObjectId In acBlkTblRec.GetBlockReferenceIds(True, True)
                    ' Получаване на блоковата референция
                    Dim acBlkRef As BlockReference = actrans.GetObject(blkRefId, OpenMode.ForWrite)
                    ' Получаване на колекцията от атрибути на блоковата референция
                    Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                    ' Обхождане на всички атрибути
                    'For Each objID As ObjectId In attCol
                    '    ' Получаване на атрибута
                    '    Dim dbObj As DBObject = actrans.GetObject(objID, OpenMode.ForWrite)
                    '    Dim acAttRef As AttributeReference = dbObj
                    '    ' Проверка тагът на атрибута и промяна на текста на атрибута
                    '    If acAttRef.Tag = "ОБЕКТ" Then acAttRef.TextString = Zapis(0)
                    '    If acAttRef.Tag = "МЕСТОПОЛОЖЕНИЕ" Then acAttRef.TextString = Zapis(1)
                    '    If acAttRef.Tag = "ВЪЗЛОЖИТЕЛ" Then acAttRef.TextString = Zapis(2)
                    '    If acAttRef.Tag = "СОБСТВЕНИК" Then acAttRef.TextString = Zapis(3)
                    '    If acAttRef.Tag = "ФАЗА" Then acAttRef.TextString = Zapis(4)
                    '    If acAttRef.Tag = "ДАТА" Then acAttRef.TextString = Zapis(5)
                    '    If acAttRef.Tag = "АРХИТЕКТ" Then acAttRef.TextString = Zapis(6)
                    '    If acAttRef.Tag = "КОНСТРУКТОР" Then acAttRef.TextString = Zapis(7)
                    '    If acAttRef.Tag = "ТЕХНОЛОГИЯ" Then acAttRef.TextString = Zapis(8)
                    '    If acAttRef.Tag = "ВИК" Then acAttRef.TextString = Zapis(9)
                    '    If acAttRef.Tag = "ОВ" Then acAttRef.TextString = Zapis(10)
                    '    If acAttRef.Tag = "ГЕОДЕЗИЯ" Then acAttRef.TextString = Zapis(11)
                    '    If acAttRef.Tag = "ВП" Then acAttRef.TextString = Zapis(12)
                    '    If acAttRef.Tag = "ЕЕФ" Then acAttRef.TextString = Zapis(13)
                    '    If acAttRef.Tag = "ПБ" Then acAttRef.TextString = Zapis(14)
                    '    If acAttRef.Tag = "ПБЗ" Then acAttRef.TextString = Zapis(15)
                    '    If acAttRef.Tag = "ПУСО" Then acAttRef.TextString = Zapis(16)
                    '    If acAttRef.Tag = "ПРОЕКТАНТ" Then acAttRef.TextString = Zapis(17)
                    '    If acAttRef.Tag = "FILE_PATH" Then acAttRef.TextString = Zapis(18)
                    '    If acAttRef.Tag = "Ном.заявление" Then acAttRef.TextString = Zapis(19)
                    '    If acAttRef.Tag = "Дата_заявление" Then acAttRef.TextString = Zapis(20)
                    '    If acAttRef.Tag = "SAP" Then acAttRef.TextString = Zapis(21)
                    '    If acAttRef.Tag = "Дружество" Then acAttRef.TextString = Zapis(22)
                    '    If acAttRef.Tag = "брой_листове" Then acAttRef.TextString = layoutCount.ToString

                    '    If acAttRef.Tag = "ТОЧКА_3" Then acAttRef.TextString = Zapis(23)

                    '    If acAttRef.Tag = "ИМЕ" Then acAttRef.TextString = Zapis(24)
                    '    If acAttRef.Tag = "АДРЕС" Then acAttRef.TextString = Zapis(25)
                    '    If acAttRef.Tag = "ПАРТИДА" Then acAttRef.TextString = Zapis(26)

                    'Next
                    For Each objID As ObjectId In attCol
                        Dim dbObj As DBObject = actrans.GetObject(objID, OpenMode.ForWrite)
                        Dim acAttRef As AttributeReference = CType(dbObj, AttributeReference)
                        ' Попълване от речника, ако има стойност за този таг
                        If Data.ContainsKey(acAttRef.Tag) Then
                            acAttRef.TextString = Data(acAttRef.Tag)
                        ElseIf acAttRef.Tag = "брой_листове" Then
                            acAttRef.TextString = layoutCount.ToString()
                        End If
                    Next
                Next
                Dim ptBasePointRes As PromptPointResult
                Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
                ' Prompt for the start point
                pPtOpts.Message = vbLf & "Изберете къде да поставя подписите: "
                ptBasePointRes = acDoc.Editor.GetPoint(pPtOpts)
                Dim ptBasePoint As Point3d = ptBasePointRes.Value
                'Set_Signature(Zapis) ' клонира блокове от един DWG файл в друг
                Dim Index As Integer = 0
                For i As Integer = 0 To 18
                    If acBlkTbl.Has(Zapis(i)) Then
                        cu.InsertText(Zapis(i), New Point3d(ptBasePoint.X, ptBasePoint.Y - Index * 15, 0), "podpisi", 3, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
                        cu.InsertBlock(Zapis(i), New Point3d(ptBasePoint.X + 30, ptBasePoint.Y - Index * 15, 0), "podpisi", New Scale3d(1, 1, 1))
                        Index += 1
                    End If
                Next
                ' Приключване на транзакцията
                actrans.Commit()
            End Using
            SaveProjectDataToExcel()
        Catch ex As Exception
            MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
        End Try
    End Sub
    Private Sub SaveProjectDataToExcel()
        Dim nameExcel As String = "\\MONIKA\Monika\_НАСТРОЙКИ\Обекти.xlsx"
        Dim excel_App As excel.Application = Nothing
        Dim excel_Workbook As excel.Workbook = Nothing
        Dim wsObekri As excel.Worksheet = Nothing
        Dim ExcelColumnToTag As New Dictionary(Of String, String) From {
            {"Обект", "ОБЕКТ"},
            {"Дата", "ДАТА"},
            {"Местоположение", "МЕСТОПОЛОЖЕНИЕ"},
            {"Възложител", "ВЪЗЛОЖИТЕЛ"},
            {"Собственик", "СОБСТВЕНИК"},
            {"Фаза", "ФАЗА"},
            {"Архитект", "АРХИТЕКТ"},
            {"Конструктор", "КОНСТРУКТОР"},
            {"Технология", "ТЕХНОЛОГИЯ"},
            {"ВиК", "ВИК"},
            {"ОВ", "ОВ"},
            {"Геодезия", "ГЕОДЕЗИЯ"},
            {"ВП", "ВП"},
            {"ЕЕФ", "ЕЕФ"},
            {"ПБ", "ПБ"},
            {"ПБЗ", "ПБЗ"},
            {"ПУСО", "ПУСО"},
            {"Проектант", "ПРОЕКТАНТ"},
            {"Път", "FILE_PATH"}
        }
        Try
            ' 1. Стартиране на Excel "тихо"
            excel_App = New excel.Application
            excel_App.Visible = False
            excel_App.DisplayAlerts = False
            ' 2. Отваряне на работната книга
            excel_Workbook = excel_App.Workbooks.Open(nameExcel)
            wsObekri = CType(excel_Workbook.Sheets("Обекти"), excel.Worksheet)
            ' 3. Проверка дали таблицата "ОБЕКТИ" съществува
            If wsObekri.ListObjects.Count = 0 Then
                Throw New Exception("Няма дефинирани таблици в листа 'Обекти'.")
            End If
            Dim tblObekti As excel.ListObject = wsObekri.ListObjects("ОБЕКТИ")
            ' --- ПРОВЕРКА ЗА ДУБЛИРАНЕ ---
            Dim isDuplicate As Boolean = False
            Try
                ' Търсим съвпадение
                Dim duplicateRow As Object = excel_App.WorksheetFunction.Match(Zapis(0), tblObekti.ListColumns("Обект").Range, 0)
                isDuplicate = True ' Ако горният ред не даде грешка, значи има дубликат
            Catch
                ' Ако Match гръмне, значи е нов обект
                isDuplicate = False
            End Try
            If isDuplicate Then
                Dim result As MsgBoxResult = MsgBox("Обектът '" & Zapis(0) & "' вече съществува. Искате ли да го запишете отново?",
                                                    MsgBoxStyle.YesNo + MsgBoxStyle.Question, "Дублиран запис")
                ' Ако потребителят избере "НЕ", затваряме и излизаме чисто през Finally
                If result = MsgBoxResult.No Then Return
            End If
            ' 5. Добавяне на нов ред в края на таблицата
            ' Добавяме новия запис най-отгоре, за да са новите обекти винаги на видно място
            Dim newRow As excel.ListRow = tblObekti.ListRows.Add(1)
            ' Попълване на "Платено" (ако колоната съществува)
            If tblObekti.ListColumns.Contains("Платено") Then
                newRow.Range.Cells(1, tblObekti.ListColumns("Платено").Index).Value = "НЕ Е ПЛАТЕНО"
            End If

            ' Попълване на останалите колони според мапинга
            For Each kvp As KeyValuePair(Of String, String) In ExcelColumnToTag
                Dim excelColName As String = kvp.Key      ' напр. "Обект"
                Dim dataTag As String = kvp.Value         ' напр. "ОБЕКТ"

                If tblObekti.ListColumns.Contains(excelColName) AndAlso Data.ContainsKey(dataTag) Then
                    newRow.Range.Cells(1, tblObekti.ListColumns(excelColName).Index).Value = Data(dataTag)
                End If
            Next
            'newRow.Range.Cells(1, 1).Value = "НЕ"
            '' Основни данни — по име на колона
            'newRow.Range.Cells(1, tblObekti.ListColumns("Платено").Index).Value = "НЕ Е ПЛАТЕНО"
            'newRow.Range.Cells(1, tblObekti.ListColumns("Обект").Index).Value = Zapis(0)
            'newRow.Range.Cells(1, tblObekti.ListColumns("Дата").Index).Value = Zapis(1)
            'newRow.Range.Cells(1, tblObekti.ListColumns("Местоположение").Index).Value = Zapis(2)
            'newRow.Range.Cells(1, tblObekti.ListColumns("Възложител").Index).Value = Zapis(3)
            'newRow.Range.Cells(1, tblObekti.ListColumns("Собственик").Index).Value = Zapis(4)
            'newRow.Range.Cells(1, tblObekti.ListColumns("Фаза").Index).Value = Zapis(5)
            'newRow.Range.Cells(1, tblObekti.ListColumns("Архитект").Index).Value = Zapis(6)
            'newRow.Range.Cells(1, tblObekti.ListColumns("Конструктор").Index).Value = Zapis(7)
            'newRow.Range.Cells(1, tblObekti.ListColumns("Технология").Index).Value = Zapis(8)
            'newRow.Range.Cells(1, tblObekti.ListColumns("ВиК").Index).Value = Zapis(9)
            'newRow.Range.Cells(1, tblObekti.ListColumns("ОВ").Index).Value = Zapis(10)
            'newRow.Range.Cells(1, tblObekti.ListColumns("Геодезия").Index).Value = Zapis(11)
            'newRow.Range.Cells(1, tblObekti.ListColumns("ВП").Index).Value = Zapis(12)
            'newRow.Range.Cells(1, tblObekti.ListColumns("ЕЕФ").Index).Value = Zapis(13)
            'newRow.Range.Cells(1, tblObekti.ListColumns("ПБ").Index).Value = Zapis(14)
            'newRow.Range.Cells(1, tblObekti.ListColumns("ПБЗ").Index).Value = Zapis(15)
            'newRow.Range.Cells(1, tblObekti.ListColumns("ПУСО").Index).Value = Zapis(16)
            'newRow.Range.Cells(1, tblObekti.ListColumns("Проектант").Index).Value = Zapis(17)
            'newRow.Range.Cells(1, tblObekti.ListColumns("Път").Index).Value = Zapis(18)
        Catch ex As Exception
            MsgBox("Грешка при запис в Excel: " & ex.Message, MsgBoxStyle.Critical)
        Finally
            ' 6. Запис
            excel_Workbook.Save()
            ' 7. Освобождаване на ресурсите
            If Not wsObekri Is Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wsObekri)
                wsObekri = Nothing
            End If
            If Not excel_Workbook Is Nothing Then
                excel_Workbook.Close(SaveChanges:=False) ' вече сме запазили
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel_Workbook)
                excel_Workbook = Nothing
            End If
            If Not excel_App Is Nothing Then
                excel_App.Quit()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel_App)
                excel_App = Nothing
            End If
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub
End Class
Class SurroundingClass
    <CommandMethod("MyCmd")>
    Public Shared Sub MyCmd()
        Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        Dim blkName As String = PickBlock(ed)
    End Sub
    Private Shared Function PickBlock(ByVal ed As Editor) As String
        Dim blkName As String = ""
        Dim opt As PromptEntityOptions = New PromptEntityOptions(vbLf & "Pick a block:")
        opt.SetRejectMessage("Bad pick: must be a block")
        opt.AddAllowedClass(GetType(BlockReference), True)
        Dim res As PromptEntityResult = ed.GetEntity(opt)

        If res.Status = PromptStatus.OK Then
            Using tran As Transaction = res.ObjectId.Database.TransactionManager.StartTransaction()
                Dim blk As BlockReference = TryCast(tran.GetObject(res.ObjectId, OpenMode.ForRead), BlockReference)
                If blk IsNot Nothing Then
                    If blk.IsDynamicBlock Then
                        Dim br As BlockTableRecord = TryCast(tran.GetObject(blk.DynamicBlockTableRecord, OpenMode.ForRead), BlockTableRecord)
                        blkName = br.Name
                    Else
                        blkName = blk.Name
                    End If
                End If
                tran.Commit()
            End Using
        End If
        Return blkName
    End Function
End Class

