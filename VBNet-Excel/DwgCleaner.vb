Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Colors
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime

Public Class DwgCleaner
    ' Координати за пълно изтриване
    Private ReadOnly xMin As Double = -184432.7953
    Private ReadOnly xMax As Double = 19797.1499
    Private ReadOnly yMin As Double = -58162.2524
    Private ReadOnly yMax As Double = 126580.8506

    ''' <summary>
    ''' Главна процедура за почистване на DWG файл.
    ''' Изпълнява серия от стъпки за премахване на излишни листове, обекти, атрибути, разбиване на блокове,
    ''' оптимизация на чертежа чрез OVERKILL, AUDIT и PURGE и финално записване.
    ''' </summary>
    <CommandMethod("DwgCleaner")>
    <CommandMethod("ДВГСЛЕАНЕР")>
    Public Sub RunCleaner()
        ' Вземаме текущия активен документ
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Dim filePath As String = db.Filename
        ' --- Проверка: Файлът трябва да е в папка "документация" ---
        If filePath.IndexOf("документация", StringComparison.OrdinalIgnoreCase) = -1 Then
            ed.WriteMessage(vbLf & "ОТКАЗ: Командата е разрешена само за файлове в папка 'документация'!")
            Return
        End If
        Try
            ' ===============================
            ' СТЪПКА 1: Изтриване на листове "настройки"
            ' ===============================
            ed.WriteMessage(vbLf & "1. Премахване на излишни листове...")
            DeleteSettingsLayouts(doc)
            ' ===============================
            ' СТЪПКА 2: Почистване на обекти в ModelSpace по координати
            ' ===============================
            ed.WriteMessage(vbLf & "2. Почистване на обекти извън работната зона...")
            WipeModelSpaceByArea(doc)
            ' ===============================
            ' СТЪПКА 3: Изчистване съдържанието на динамични блокове "Качване"
            ' ===============================
            ed.WriteMessage(vbLf & "3. Изчистване съдържанието на блокове 'Качване'...")
            ClearAttributesInDynamicBlocks(doc, "Качване")
            FindMylniq(doc)
            ' ===============================
            ' СТЪПКА 4: Native BURST (разбиване на блокове)
            ' ===============================
            NativeBurst(doc)
            ExplodeAllArrays(doc)
            NativeBurst(doc)
            ' ===============================
            ' СТЪПКА 5: OVERKILL (оптимизация на геометрията)
            ' ===============================
            ed.WriteMessage(vbLf & "5. Изпълнение на OVERKILL...")
            Try
                ' Избираме всичко и изпълняваме Overkill
                ed.Command("-OVERKILL", "_All", "", "")
                ed.WriteMessage(vbLf & "Overkill приключи.")
            Catch ex As System.Exception
                ed.WriteMessage(vbLf & "(!) OVERKILL грешка: " & ex.Message)
            End Try
            ' ===============================
            ' СТЪПКА 6: AUDIT (проверка и поправка на базата данни)
            ' ===============================
            ed.WriteMessage(vbLf & "6. Проверка на файла (Audit)...")
            Try
                ed.Command("._AUDIT", "_Yes")
            Catch
                ed.WriteMessage(vbLf & "(!) Грешка при Audit.")
            End Try
            ' ===============================
            ' СТЪПКА 7: PURGE (пълно почистване на неизползвани елементи)
            ' ===============================
            ed.WriteMessage(vbLf & "7. Пълно почистване на неизползвани слоеве и блокове...")
            ' Повтаряме няколко пъти заради вложени зависимости
            For i As Integer = 1 To 4
                ed.Command("-PURGE", "_All", "*", "_No")
            Next
            ' Почистваме и RegApps (често източник на тежки файлове)
            ed.Command("-PURGE", "_Reg", "*", "_No")
            ' ===============================
            ' СТЪПКА 8: Финален запис на файла
            ' ===============================
            db.SaveAs(db.Filename, True, DwgVersion.Current, Nothing)
            ed.WriteMessage(vbLf & "8. Файлът е успешно записан.")
            ' ===============================
            ' СТЪПКА 9: Финален изглед на чертежа
            ' ===============================
            ed.Command("_.ZOOM", "_E")
            ' Крайно съобщение за успешна процедура
            ed.WriteMessage(vbLf & "--- [УСПЕХ] Процедурата 'DwgCleaner' приключи! ---")
        Catch ex As System.Exception
            ' Грешка в главния цикъл
            ed.WriteMessage(vbLf & "Критична грешка в главния цикъл: " & ex.Message)
        End Try
    End Sub
    ''' <summary>
    ''' Реализира Native BURST поведение:
    ''' - Превръща атрибутите в текст (без TAG)
    ''' - Експлодира геометрията на блока
    ''' - Наследява слой и цвят коректно
    ''' - Изтрива оригиналния блок
    ''' </summary>
    ''' <param name="doc">Текущият AutoCAD документ</param>
    Private Sub NativeBurst(doc As Document)
        ' 1. Списък с имена на блокове, които НЕ трябва да бъдат разбивани (Skip List)
        Dim protectedBlocks As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
        "s_c60_circ_break",
        "s_ct_cont_no",
        "s_dpnn_vigi_circ_break",
        "s_GV2",
        "s_i_ng_switch_disconn",
        "s_id_res_circ_break",
        "s_in_ins_inv_disconn",
        "s_min",
        "s_ng125_circ_break",
        "s_ns100_motor_fixed",
        "s_switch_light_sens",
        "s_tesys_cont_no",
        "s_tl",
        "s_vigi_res",
        "Мълниезащита вертикално"
        }
        ' Вземаме базата данни и редактора на текущия документ
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        ' Създаваме транзакция за безопасна работа с обекти
        Using tr As Transaction = db.TransactionManager.StartTransaction()
            ' Отваряме текущото пространство за писане (ModelSpace или PaperSpace)
            Dim btrCurrent As BlockTableRecord = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite)
            ' Създаваме филтър за селекция само на блокови референции (INSERT)
            Dim filter As New SelectionFilter({New TypedValue(0, "INSERT")})
            Dim selRes As PromptSelectionResult = ed.SelectAll(filter)
            ' Ако има блокове за обработка
            If selRes.Status = PromptStatus.OK Then
                Dim count As Integer = 0
                ' Обхождаме всички избрани блокови референции
                For Each id As ObjectId In selRes.Value.GetObjectIds()
                    If id.IsErased Then Continue For
                    Dim br As BlockReference = tr.GetObject(id, OpenMode.ForWrite)
                    ' --- НОВАТА ПРОВЕРКА ТУК ---
                    ' Вземаме името на блока (поддържа и динамични блокове)
                    Dim blockName As String = If(br.IsDynamicBlock,
                    DirectCast(tr.GetObject(br.DynamicBlockTableRecord, OpenMode.ForRead), BlockTableRecord).Name,
                    br.Name)
                    ' Ако името е в списъка, прескачаме този блок
                    If protectedBlocks.Contains(blockName) Then
                        Continue For
                    End If
                    ' ---------------------------
                    ' Запазваме оригиналния слой и цвят на блока
                    Dim blockLayer As String = br.Layer
                    Dim blockColor As Color = br.Color
                    ' ===============================
                    ' СТЪПКА 1: Превръщаме атрибутите на блока в DBText
                    ' ===============================
                    For Each attId As ObjectId In br.AttributeCollection
                        Dim attRef As AttributeReference = tr.GetObject(attId, OpenMode.ForRead)
                        ' Пропускаме празни атрибути
                        If Not String.IsNullOrWhiteSpace(attRef.TextString) Then
                            Dim newText As New DBText()
                            newText.SetDatabaseDefaults()
                            ' Копираме свойствата на атрибута
                            newText.TextString = attRef.TextString
                            newText.Position = attRef.Position
                            newText.Height = attRef.Height
                            newText.Rotation = attRef.Rotation
                            newText.TextStyleId = attRef.TextStyleId
                            ' Ако атрибутът е в слой "0", наследява слоя и цвета на блока
                            If attRef.Layer = "0" Then
                                newText.Layer = br.Layer
                                newText.Color = br.Color
                            Else
                                ' Иначе запазваме оригиналния слой и цвят на атрибута
                                newText.Layer = attRef.Layer
                                newText.Color = attRef.Color
                            End If
                            ' Добавяме новия текст в текущото пространство
                            btrCurrent.AppendEntity(newText)
                            tr.AddNewlyCreatedDBObject(newText, True)
                        End If
                    Next
                    ' ===============================
                    ' СТЪПКА 2: Explode (разбиване) на геометрията на блока
                    ' ===============================
                    Dim explodedObjects As New DBObjectCollection()
                    br.Explode(explodedObjects)
                    For Each obj As DBObject In explodedObjects
                        ' Пропускаме атрибутите, защото вече са конвертирани в текст
                        If TypeOf obj Is AttributeReference OrElse TypeOf obj Is AttributeDefinition Then
                            Continue For
                        End If
                        Dim ent As Entity = DirectCast(obj, Entity)
                        ' Наследяване на слой и цвят, ако е необходимо
                        If ent.Layer = "0" Then
                            ent.Layer = blockLayer
                        End If
                        If ent.Color.ColorMethod = ColorMethod.ByBlock Then
                            ent.Color = blockColor
                        End If
                        ' Добавяме обекта в текущото пространство
                        btrCurrent.AppendEntity(ent)
                        tr.AddNewlyCreatedDBObject(ent, True)
                    Next
                    ' ===============================
                    ' СТЪПКА 3: Изтриваме оригиналния блок
                    ' ===============================
                    br.Erase()
                    count += 1
                Next
                ' Съобщение в редактора за брой обработени блокове
                ed.WriteMessage(vbLf & "Native BURST: Обработени " & count & " блока.")
            End If
            ' Потвърждаваме всички промени
            tr.Commit()
        End Using
    End Sub
    ''' <summary>
    ''' Изтрива всички Layout-и, съдържащи "настройки" в името, с изключение на "Model".
    ''' </summary>
    ''' <param name="doc">Текущият AutoCAD документ</param>
    Private Sub DeleteSettingsLayouts(doc As Document)
        ' Вземаме базата данни и редактора на текущия документ
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        ' Създаваме транзакция за безопасна работа с обекти
        Using tr As Transaction = db.TransactionManager.StartTransaction()
            ' Отваряме LayoutDictionary за четене
            Dim layoutDict As DBDictionary = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead)
            Dim deletedCount As Integer = 0
            ' Събираме имената на Layout-и, които трябва да изтрием
            Dim toDelete As New List(Of String)
            For Each entry As DictionaryEntry In layoutDict
                Dim name As String = entry.Key.ToString()
                ' Ако името съдържа "настройки" и не е "Model"
                If name.ToLower().Contains("настройки") AndAlso name.ToLower() <> "model" Then
                    toDelete.Add(name)
                End If
            Next
            ' Ако има Layout-и за изтриване
            If toDelete.Count > 0 Then
                Dim layMgr As LayoutManager = LayoutManager.Current
                For Each name As String In toDelete
                    ' Изтриваме Layout-а
                    layMgr.DeleteLayout(name)
                    deletedCount += 1
                Next
                ' Потвърждаваме транзакцията
                tr.Commit()
                ' Записваме файла след изтриването на листовете
                db.SaveAs(db.Filename, True, DwgVersion.Current, Nothing)
                ' Извеждаме съобщение с броя изтрити Layout-и
                ed.WriteMessage(vbLf & "Изтрити листове: " & deletedCount & ". Файлът е записан.")
            Else
                ' Ако няма какво да се изтрива, прекратяваме транзакцията
                tr.Abort()
            End If
        End Using
    End Sub
    ''' <summary>
    ''' Изтрива всички обекти в Model Space, чиито центрови точки попадат в зададената зона (xMin, xMax, yMin, yMax).
    ''' </summary>
    ''' <param name="doc">Текущият AutoCAD документ</param>
    Private Sub WipeModelSpaceByArea(doc As Document)
        ' Вземаме базата данни и редактора на текущия документ
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        ' Създаваме транзакция за безопасна работа с обекти
        Using tr As Transaction = db.TransactionManager.StartTransaction()
            ' Отваряме BlockTable и ModelSpace за четене/писане
            Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim ms As BlockTableRecord = tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            Dim deletedCount As Integer = 0
            ' Обхождаме всички обекти в ModelSpace
            For Each id As ObjectId In ms
                Dim ent As Entity = tr.GetObject(id, OpenMode.ForRead)
                Try
                    ' Опитваме се да вземем геометричните граници на обекта
                    Dim ext As Extents3d = ent.GeometricExtents
                    ' Изчисляваме центъра на обекта
                    Dim midX As Double = (ext.MinPoint.X + ext.MaxPoint.X) / 2
                    Dim midY As Double = (ext.MinPoint.Y + ext.MaxPoint.Y) / 2
                    ' Ако центърът попада в зададения правоъгълник
                    If midX >= xMin And midX <= xMax And midY >= yMin And midY <= yMax Then
                        ' Отваряме обекта за писане
                        ent.UpgradeOpen()
                        ' Изтриваме обекта и увеличаваме броя
                        ent.Erase(True)
                        deletedCount += 1
                    End If
                Catch
                    ' Ако обектът няма GeometricExtents или възникне грешка, го пропускаме
                    Continue For
                End Try
            Next
            ' Потвърждаваме промените в транзакцията
            tr.Commit()
            ' Ако има изтрити обекти, записваме документа и извеждаме съобщение
            If deletedCount > 0 Then
                db.SaveAs(db.Filename, True, DwgVersion.Current, Nothing)
                ed.WriteMessage(vbLf & "Изтрити обекти от Model Space: " & deletedCount & ". Файлът е записан.")
            End If
        End Using
    End Sub
    ''' <summary>
    ''' Изчиства съдържанието на всички атрибути в динамични или обикновени блокове с дадено име.
    ''' </summary>
    ''' <param name="doc">Текущият AutoCAD документ</param>
    ''' <param name="targetName">Името на блока, чийто атрибути ще бъдат изчистени (пример: "Качване")</param>
    Private Sub ClearAttributesInDynamicBlocks(doc As Document, targetName As String)
        ' Вземаме базата данни и редактора на текущия документ
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        ' Създаваме транзакция за безопасна работа с обекти
        Using tr As Transaction = db.TransactionManager.StartTransaction()
            ' Вземаме BlockTable и ModelSpace за четене/писане
            Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim ms As BlockTableRecord = tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            Dim count As Integer = 0
            ' Обхождаме всички обекти в ModelSpace
            For Each id As ObjectId In ms
                Dim ent As Entity = tr.GetObject(id, OpenMode.ForRead)
                ' Проверяваме дали обектът е референция към блок
                If TypeOf ent Is BlockReference Then
                    Dim br As BlockReference = DirectCast(ent, BlockReference)
                    ' Вземаме ефективното име на блока
                    Dim effectiveName As String = ""
                    If br.IsDynamicBlock Then
                        ' Ако е динамичен блок, вземаме името на дефиницията
                        Dim btr As BlockTableRecord = tr.GetObject(br.DynamicBlockTableRecord, OpenMode.ForRead)
                        effectiveName = btr.Name
                    Else
                        ' Ако е обикновен блок
                        effectiveName = br.Name
                    End If
                    ' Ако името съвпада с целевото име (без да се взема предвид главни/малки букви)
                    If effectiveName.Equals(targetName, StringComparison.OrdinalIgnoreCase) Then
                        ' Обхождаме всички атрибути на блока
                        For Each attId As ObjectId In br.AttributeCollection
                            Dim attRef As AttributeReference = tr.GetObject(attId, OpenMode.ForWrite)
                            attRef.TextString = "" ' Изчистваме съдържанието на атрибута
                        Next
                        count += 1
                    End If
                End If
            Next
            ' Потвърждаваме промените
            tr.Commit()
            ' Извеждаме съобщение с броя на обработените блокове
            ed.WriteMessage(vbLf & "Изчистени атрибути в динамичен блок '" & targetName & "': " & count)
        End Using
    End Sub
    ''' <summary>
    ''' Разбива всички масиви (BlockReference масиви) в ModelSpace
    ''' без да пипа реалните блокове и атрибути.
    ''' </summary>
    Public Sub ExplodeAllArrays(doc As Document)
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        ' Стартираме транзакция за безопасна работа с обекти
        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Dim bt As BlockTable = CType(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
            Dim ms As BlockTableRecord = CType(tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)
            Dim idsToErase As New List(Of ObjectId)
            ' Обхождаме всички обекти в ModelSpace
            For Each id As ObjectId In ms
                Dim ent As Entity = TryCast(tr.GetObject(id, OpenMode.ForRead), Entity)
                If ent Is Nothing Then Continue For
                ' Проверка дали е BlockReference
                If TypeOf ent Is BlockReference Then
                    Dim br As BlockReference = CType(ent, BlockReference)
                    ' Проверяваме само дали е DynamicBlock
                    If br.IsDynamicBlock Then
                        Dim objs As New DBObjectCollection()
                        br.Explode(objs) ' Разбиваме масива / DynamicBlock
                        ' Добавяме всички обекти от Explode в ModelSpace
                        For Each o As DBObject In objs
                            Dim e As Entity = TryCast(o, Entity)
                            If e IsNot Nothing Then
                                ms.AppendEntity(e)
                                tr.AddNewlyCreatedDBObject(e, True)
                            End If
                        Next
                        ' Добавяме оригиналния BlockReference за изтриване
                        idsToErase.Add(br.ObjectId)
                    End If
                End If
            Next
            ' Изтриваме оригиналните BlockReference масиви
            For Each id As ObjectId In idsToErase
                Dim e As Entity = TryCast(tr.GetObject(id, OpenMode.ForWrite), Entity)
                If e IsNot Nothing Then
                    e.Erase()
                End If
            Next
            tr.Commit()
        End Using
        ed.WriteMessage(vbLf & "Всички масиви са разбити успешно (BlockReference запазени).")
    End Sub
    ''' <summary>
    ''' Търси в текущия чертеж текстове и блокове, свързани с мълниезащита.
    ''' При намиране на конкретен текст го заменя с нов MText,
    ''' като използва параметри и атрибути от съответния блок „Мълниезащита вертикално“.
    ''' </summary>
    ''' <param name="doc">Текущият AutoCAD документ</param>
    Private Sub FindMylniq(doc As Document)
        ' Списъци за съхранение на намерените обекти (ID-та)
        Dim mylniqTextIds As New List(Of ObjectId)
        Dim mylniqBlockIds As New List(Of ObjectId)
        ' Референции към базата данни и Editor-а
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        ' Променливи за стойности, извлечени от блока
        Dim valTip As String = ""
        Dim valH As String = ""
        Dim valKategoria As String = ""
        Dim valRa As String = ""
        Try
            Using tr As Transaction = db.TransactionManager.StartTransaction()
                ' Създаваме филтър за INSERT (блокове), MTEXT и TEXT
                Dim filter As New SelectionFilter({
                    New TypedValue(0, "INSERT,MTEXT,TEXT")
                })
                ' Избираме всички обекти в чертежа, които отговарят на филтъра
                Dim selRes As PromptSelectionResult = ed.SelectAll(filter)
                If selRes.Status = PromptStatus.OK Then
                    ' Търсена уникална фраза в текстовете
                    Dim searchPhrase As String =
                        "За защита от мълнии да се монтира мълниеприемник с изпреварващо действие"
                    ' Обхождаме всички намерени обекти
                    For Each id As ObjectId In selRes.Value.GetObjectIds()
                        Dim ent As Entity = tr.GetObject(id, OpenMode.ForRead)

                        ' === 1. Обработка на текстове (MText и DBText) ===
                        If TypeOf ent Is MText OrElse TypeOf ent Is DBText Then
                            ' Извличаме съдържанието според типа текст
                            Dim content As String = If(TypeOf ent Is MText,
                                DirectCast(ent, MText).Contents,
                                DirectCast(ent, DBText).TextString)
                            ' Ако текстът съдържа търсената фраза, запомняме ID-то
                            If content.IndexOf(searchPhrase, StringComparison.OrdinalIgnoreCase) >= 0 Then
                                mylniqTextIds.Add(id)
                            End If
                            ' === 2. Обработка на блокове ===
                        ElseIf TypeOf ent Is BlockReference Then
                            Dim br As BlockReference = DirectCast(ent, BlockReference)
                            Dim btr As BlockTableRecord =
                                tr.GetObject(br.DynamicBlockTableRecord, OpenMode.ForRead)
                            ' Проверяваме името на блока
                            If btr.Name.Trim().Equals("Мълниезащита вертикално",
                                                       StringComparison.OrdinalIgnoreCase) Then
                                mylniqBlockIds.Add(id)
                            End If
                        End If
                    Next
                End If
                ' === 3. Извличане на атрибути и динамични параметри от блока ===
                If mylniqBlockIds.Count > 0 Then
                    Dim brMyl As BlockReference =
                        tr.GetObject(mylniqBlockIds(0), OpenMode.ForRead)
                    ' Четене на атрибутите H и RA
                    For Each attId As ObjectId In brMyl.AttributeCollection
                        Dim attRef As AttributeReference =
                            tr.GetObject(attId, OpenMode.ForRead)
                        If attRef.Tag.ToUpper() = "H" Then valH = attRef.TextString
                        If attRef.Tag.ToUpper() = "RA" Then valRa = attRef.TextString
                    Next

                    ' Четене на динамичните параметри (ако блокът е динамичен)
                    If brMyl.IsDynamicBlock Then
                        For Each prop As DynamicBlockReferenceProperty _
                            In brMyl.DynamicBlockReferencePropertyCollection
                            If prop.PropertyName.ToUpper() = "КАТЕГОРИЯ" Then valKategoria = prop.Value.ToString()
                            If prop.PropertyName.ToUpper() = "ТИП" Then valTip = prop.Value.ToString()
                        Next
                    End If
                End If

                ' === 4. Замяна на стария текст с нов ===
                If mylniqTextIds.Count > 0 Then
                    ' Отваряме стария текст за редакция
                    Using oldEnt As Entity =
                        tr.GetObject(mylniqTextIds(0), OpenMode.ForWrite)
                        Dim oldMt As MText = TryCast(oldEnt, MText)
                        If oldMt IsNot Nothing Then
                            ' А. Запазваме параметрите на стария текст
                            Dim layer As String = oldMt.Layer
                            Dim position As Point3d = oldMt.Location
                            Dim textStyle As ObjectId = oldMt.TextStyleId
                            Dim textHeight As Double = oldMt.TextHeight
                            ' Б. Създаваме нов MText със същите параметри
                            Dim newMt As New MText()
                            newMt.Layer = layer
                            newMt.Location = position
                            newMt.TextStyleId = textStyle
                            newMt.TextHeight = textHeight
                            ' В. Сглобяване на финалния текст
                            Dim finalText As String =
                                "{\LЗабележки:\P}" &
                                "       1. За защита от мълнии да се монтира мълниеприемник с изпреварващо действие PREVECTRON®3, Millenium модел " & valTip & ", или подобен.\P" &
                                "       2. Фактическата височина на монтажа на мълниеприемника h над повърхнината, която трябва да бъде защитавана да бъде " & valH & " m.\P" &
                                "       3. За присъединяване на мълниеприемника към мълниезащитния прът да се използва детайл съгласно спецификация на производителя.\P" &
                                "       4. Радиус на защита за ниво на защита " & valKategoria & "при h(m) = " & valH & " m е Rз = " & valRa & " m\P" &
                                "       5. Мълниезащитните отводи да се изпълнят от екструдиран проводник Ф8мм.\P" &
                                "       6. Минимален радиус на огъване на мълниезащитните отводи R 200.\P" &
                                "       7. Токоотвода да се постави на вертикална противопожарна ивица с ширина 0,50m, с клас по реакция на огън А2."

                            newMt.Contents = finalText
                            newMt.Width = 0
                            ' Г. Добавяме новия текст в чертежа
                            Dim btr As BlockTableRecord =
                                tr.GetObject(oldMt.OwnerId, OpenMode.ForWrite)
                            btr.AppendEntity(newMt)
                            tr.AddNewlyCreatedDBObject(newMt, True)
                            ' Д. Изтриваме стария текст
                            oldMt.Erase(True)
                            ' Потвърждаваме промените
                            tr.Commit()
                            ' Обновяваме екрана
                            Application.DocumentManager.MdiActiveDocument.Editor.Regen()
                        End If
                    End Using
                End If
            End Using
        Catch ex As Exception
            ' Изписване на грешка в командния ред
            ed.WriteMessage(vbLf & "Грешка в FindMylniq: " & ex.Message)
        End Try
    End Sub
End Class
