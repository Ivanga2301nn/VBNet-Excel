Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports Autodesk.AutoCAD.DatabaseServices.Filters
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.TrackBar
Public Class Insert_Signature
    Dim cu As CommonUtil = New CommonUtil()
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
        Dim File_Path As String = Path.GetDirectoryName(name_file)
        Dim File_name As String = Path.GetFileName(name_file)
        Dim Path_Name As String = Mid(File_Path, InStrRev(File_Path, "\") + 1, Len(File_Path))

        Dim Zapis(26) As String
        Zapis(0) = cu.GetObjects_TEXT("Изберете Наименование на ОБЕКТА", vbFalse)
        If Zapis(0) = "  #####  " Then
            MsgBox("Няма избран текст Наименование на ОБЕКТА: ")
            Exit Sub
        End If
        Zapis(1) = cu.GetObjects_TEXT("Изберете Местоположение на ОБЕКТА", vbFalse)
        Zapis(2) = cu.GetObjects_TEXT("Изберете ВЪЗЛОЖИТЕЛ на проекта", vbFalse)
        Zapis(3) = cu.GetObjects_TEXT("Изберете СОСТВЕНИК на обекта", vbFalse)
        Zapis(4) = cu.GetObjects_TEXT("Изберете ФАЗА на проекта", vbFalse)
        Zapis(5) = cu.GetObjects_TEXT("Изберете ДАТА на проекта", vbFalse)
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
        Zapis(17) = cu.GetObjects_TEXT("Изберете ПРОЕКТАНТ")
        Zapis(17) = IIf(Zapis(17) = "  #####  ", "инж. М.Тонкова-Генчева", Zapis(17))
        Zapis(18) = File_Path
        Zapis(19) = cu.GetObjects_TEXT("Изберете Номер на заявлене")
        Zapis(20) = cu.GetObjects_TEXT("Изберете Дата на изготвяне")
        Zapis(21) = cu.GetObjects_TEXT("Изберете съгласувал SAP номер")
        Zapis(22) = cu.GetObjects_TEXT("Изберете дружество")
        Zapis(22) = IIf(Zapis(22) = "  #####  ", Chr(34) + "Електроразпределителни мрежи " + Chr(34) + "Запад" + Chr(34) + " ЕАД", Zapis(22))

        Zapis(23) = cu.GetObjects_TEXT("Изберете текст съдържащ описанието по т.3 от становището")

        Zapis(24) = cu.GetObjects_TEXT("Изберете Име")
        Zapis(25) = cu.GetObjects_TEXT("Изберете Адрес")
        Zapis(26) = cu.GetObjects_TEXT("Изберете Партида")

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
                    For Each objID As ObjectId In attCol
                        ' Получаване на атрибута
                        Dim dbObj As DBObject = actrans.GetObject(objID, OpenMode.ForWrite)
                        Dim acAttRef As AttributeReference = dbObj
                        ' Проверка тагът на атрибута и промяна на текста на атрибута
                        If acAttRef.Tag = "ОБЕКТ" Then acAttRef.TextString = Zapis(0)
                        If acAttRef.Tag = "МЕСТОПОЛОЖЕНИЕ" Then acAttRef.TextString = Zapis(1)
                        If acAttRef.Tag = "ВЪЗЛОЖИТЕЛ" Then acAttRef.TextString = Zapis(2)
                        If acAttRef.Tag = "СОБСТВЕНИК" Then acAttRef.TextString = Zapis(3)
                        If acAttRef.Tag = "ФАЗА" Then acAttRef.TextString = Zapis(4)
                        If acAttRef.Tag = "ДАТА" Then acAttRef.TextString = Zapis(5)
                        If acAttRef.Tag = "АРХИТЕКТ" Then acAttRef.TextString = Zapis(6)
                        If acAttRef.Tag = "КОНСТРУКТОР" Then acAttRef.TextString = Zapis(7)
                        If acAttRef.Tag = "ТЕХНОЛОГИЯ" Then acAttRef.TextString = Zapis(8)
                        If acAttRef.Tag = "ВИК" Then acAttRef.TextString = Zapis(9)
                        If acAttRef.Tag = "ОВ" Then acAttRef.TextString = Zapis(10)
                        If acAttRef.Tag = "ГЕОДЕЗИЯ" Then acAttRef.TextString = Zapis(11)
                        If acAttRef.Tag = "ВП" Then acAttRef.TextString = Zapis(12)
                        If acAttRef.Tag = "ЕЕФ" Then acAttRef.TextString = Zapis(13)
                        If acAttRef.Tag = "ПБ" Then acAttRef.TextString = Zapis(14)
                        If acAttRef.Tag = "ПБЗ" Then acAttRef.TextString = Zapis(15)
                        If acAttRef.Tag = "ПУСО" Then acAttRef.TextString = Zapis(16)
                        If acAttRef.Tag = "ПРОЕКТАНТ" Then acAttRef.TextString = Zapis(17)
                        If acAttRef.Tag = "FILE_PATH" Then acAttRef.TextString = Zapis(18)
                        If acAttRef.Tag = "Ном.заявление" Then acAttRef.TextString = Zapis(19)
                        If acAttRef.Tag = "Дата_заявление" Then acAttRef.TextString = Zapis(20)
                        If acAttRef.Tag = "SAP" Then acAttRef.TextString = Zapis(21)
                        If acAttRef.Tag = "Дружество" Then acAttRef.TextString = Zapis(22)
                        If acAttRef.Tag = "брой_листове" Then acAttRef.TextString = layoutCount.ToString

                        If acAttRef.Tag = "ТОЧКА_3" Then acAttRef.TextString = Zapis(23)

                        If acAttRef.Tag = "ИМЕ" Then acAttRef.TextString = Zapis(24)
                        If acAttRef.Tag = "АДРЕС" Then acAttRef.TextString = Zapis(25)
                        If acAttRef.Tag = "ПАРТИДА" Then acAttRef.TextString = Zapis(26)

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
        Catch ex As Exception
            MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
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

