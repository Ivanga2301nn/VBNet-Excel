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
Imports iTextSharp.text.pdf.parser
Imports System.Collections
Imports System.Collections.Generic

Public Class DwgCleaner
    ' Координати за пълно изтриване
    Private ReadOnly xMin As Double = -184432.7953
    Private ReadOnly xMax As Double = 19797.1499
    Private ReadOnly yMin As Double = -58162.2524
    Private ReadOnly yMax As Double = 126580.8506
    ' В класа, преди всички методи
    Private sw As IO.StreamWriter
    ' Константа за път до ErrorLog файл
    Private Const ERROR_LOG_PATH As String = "\\MONIKA\Monika\_НАСТРОЙКИ\Нова папка\ErrorLog.txt"

    ''' <summary>
    ''' Главна процедура за почистване на DWG файл.
    ''' Изпълнява серия от стъпки за премахване на излишни листове, обекти, атрибути, разбиване на блокове,
    ''' оптимизация на чертежа чрез OVERKILL, AUDIT и PURGE и финално записване.
    ''' Основна входна процедура за обработка на DWG файл.
    ''' Определя дали текущият чертеж се намира в папка "документация".
    ''' - Ако НЕ е → стартира масова обработка на всички файлове в папката (Batch режим).
    ''' - Ако Е → стартира обработка само на текущия файл.
    ''' Създава лог файл DwgCleaner.txt в папката на DWG файла
    ''' и проверява за наличие на грешки в ErrorLog.txt.
    ''' </summary>
    <CommandMethod("DwgCleaner")>
    <CommandMethod("ДВГСЛЕАНЕР")>
    Public Sub ProcessFile()
        ' Вземаме текущия активен AutoCAD документ
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        ' Пълен път до текущия DWG файл
        Dim filePath As String = db.Filename
        ' Папката, в която се намира DWG файлът
        Dim dwgFolder As String = IO.Path.GetDirectoryName(filePath)
        ' Път до основния лог файл (в папката на DWG)
        Dim logFile As String = IO.Path.Combine(dwgFolder, "DwgCleaner.txt")
        ' Създаваме StreamWriter за логване (презаписва файла)
        sw = New IO.StreamWriter(logFile, False)
        sw.AutoFlush = True
        sw.WriteLine("===============================")
        sw.WriteLine($"--- КОДА СЕ Стартиране от ФАЙЛ -> {filePath}")
        sw.WriteLine("===============================")
        ' Път до файл за грешки (централен ErrorLog)
        Dim errorLogFile As String = ERROR_LOG_PATH
        ' Ако има стар ErrorLog – изтриваме го, за да започнем начисто
        If IO.File.Exists(errorLogFile) Then
            IO.File.Delete(errorLogFile)
        End If
        ' --- ЛОГИКА ЗА ИЗБОР НА РЕЖИМ НА РАБОТА ---
        ' Проверяваме дали DWG файлът се намира в папка "документация"
        If filePath.IndexOf("документация", StringComparison.OrdinalIgnoreCase) = -1 Then
            ' Файлът НЕ е в "документация":
            ' Стартираме масова обработка на всички DWG файлове в папката
            BatchCleaner(dwgFolder)
        Else
            ' Файлът Е в "документация":
            ' Стартираме обработка само на текущия DWG файл
            RunCleaner(doc)
            ' ===============================
            ' СТЪПКА 8: Финален запис на файла
            ' ===============================
            db.SaveAs(db.Filename, True, DwgVersion.Current, Nothing)
            sw.WriteLine("8. Файлът е успешно записан.")
            ' Крайно съобщение за успешна процедура
            sw.WriteLine("===============================")
            sw.WriteLine("         ФАЙЛ ОБРАБОТЕН        ")
            sw.WriteLine(filePath)
            sw.WriteLine("Дата/час: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
            sw.WriteLine("===============================")
            sw.WriteLine("--- [УСПЕХ] Процедурата 'DwgCleaner' приключи! ---")
        End If
        ' --- ЗАТВАРЯНЕ И ПОЧИСТВАНЕ НА ЛОГ ФАЙЛА ---
        If sw IsNot Nothing Then
            sw.Close()
            sw.Dispose()
        End If
        ' --- ПРОВЕРКА ЗА ГРЕШКИ ---
        ' Ако ErrorLog.txt съществува и не е празен – уведомяваме потребителя
        If IO.File.Exists(errorLogFile) Then
            If New IO.FileInfo(errorLogFile).Length > 0 Then
                Application.ShowAlertDialog("Има записани грешки в ErrorLog.txt!")
            End If
        End If
    End Sub
    Public Sub RunCleaner(doc As Document)
        ' Вземаме пълния път на файла от Document
        Dim filePath As String = doc.Name
        sw.WriteLine("===============================")
        sw.WriteLine("         ОБРАБОТВАМ ФАЙЛ       ")
        sw.WriteLine(filePath)
        sw.WriteLine("===============================")
        Try
            ' ===============================
            ' СТЪПКА 1: Изтриване на листове "настройки"
            ' ===============================
            DeleteSettingsLayouts(doc)
            ' ===============================
            ' СТЪПКА 2: Почистване на обекти в ModelSpace по координати
            ' ===============================
            WipeModelSpaceByArea(doc)
            ' ===============================
            ' СТЪПКА 3: Изчистване съдържанието на динамични блокове "Качване"
            ' ===============================
            ClearAttributesInDynamicBlocks(doc, "Качване")
            ' ===============================
            ' СТЪПКА 4: Изчистване съдържанието на динамични блокове "Качване"
            ' ===============================
            FindMylniq(doc)
            ' ===============================
            ' СТЪПКА 5: Native BURST (разбиване на блокове)
            ' ===============================
            sw.WriteLine("5: Native BURST (разбиване на блокове) ...")
            NativeBurst(doc)
            ExplodeAllArrays(doc)
            NativeBurst(doc)
            ' ===============================
            ' СТЪПКА 6: Bind на всички Xref-и
            ' ===============================
            Using docLock As DocumentLock = doc.LockDocument()
                NativeBind(doc.Database)
            End Using
            ' ===============================
            ' СТЪПКА 7: PURGE (пълно почистване на неизползвани елементи)
            ' ===============================
            NativePurge(doc)
        Catch ex As System.Exception
            ' Грешка в главния цикъл
            sw.WriteLine("Критична грешка в главния цикъл: " & ex.Message)
            ' Ако даден файл е зает (например отворения в момента), 
            ' той ще бъде прескочен и ще продължи със следващия.
            ' Логика за записване на грешката в текстов файл
            Dim logPath As String = ERROR_LOG_PATH
            ' Използваме 'Using', за да сме сигурни, че файлът се отваря и затваря правилно
            Using swError As New IO.StreamWriter(logPath, True)
                swError.WriteLine("========================================")
                swError.WriteLine("Дата/час: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
                swError.WriteLine("Файл: " & filePath)
                swError.WriteLine("Грешка: " & ex.Message)
                swError.WriteLine("Source: " & ex.Source)
                swError.WriteLine("HResult: " & ex.HResult.ToString())
                swError.WriteLine("StackTrace: ")
                swError.WriteLine(ex.StackTrace)
                ' Извличане на първия ред от StackTrace (както си го замислил)
                Dim lines() As String = ex.StackTrace.Split({vbCrLf}, StringSplitOptions.RemoveEmptyEntries)
                If lines.Length > 0 Then
                    swError.WriteLine("Ред: " & lines(0).Trim())
                End If
                swError.WriteLine("========================================")
            End Using
        End Try
    End Sub
    ''' <summary>
    ''' Изпълнява "Native BURST" върху блокове в текущото пространство.
    ''' Разбива блокове, конвертира атрибути в DBText, наследява слой и цвят
    ''' и пропуска защитени и Xref блокове.
    ''' </summary>
    ''' <param name="doc">Активният AutoCAD документ</param>
    Private Sub NativeBurst(doc As Document)
        Dim db As Database = doc.Database
        sw.WriteLine("5: Native BURST (разбиване на блокове) ...")

        ' Списък с имена на блокове, които НЕ трябва да бъдат разбивани
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

        Try
            Using tr As Transaction = db.TransactionManager.StartTransaction()
                Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                Dim btrCurrent As BlockTableRecord = tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                'Dim btrCurrent As BlockTableRecord = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite)
                Dim count As Integer = 0

                For Each id As ObjectId In btrCurrent
                    If id.IsErased Then Continue For
                    Dim ent As Entity = TryCast(tr.GetObject(id, OpenMode.ForRead), Entity)
                    If ent Is Nothing Then Continue For
                    If Not TypeOf ent Is BlockReference Then Continue For
                    ' Филтър по слой: само "EL*"
                    If Not ent.Layer.StartsWith("EL", StringComparison.OrdinalIgnoreCase) Then Continue For
                    Dim br As BlockReference = DirectCast(ent, BlockReference)
                    ' Пропускаме Xref блокове (освен ако не са unloaded — но те нямат геометрия)
                    Dim btr As BlockTableRecord = TryCast(tr.GetObject(br.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)
                    If btr?.IsFromExternalReference AndAlso Not btr.IsUnloaded Then
                        sw.WriteLine("Пропуснат Xref: " & br.Name)
                        Continue For
                    End If
                    ' Вземаме истинското име на блока (поддържа динамични блокове)
                    Dim blockName As String = If(
                    br.IsDynamicBlock,
                    DirectCast(tr.GetObject(br.DynamicBlockTableRecord, OpenMode.ForRead), BlockTableRecord).Name,
                    br.Name
                )
                    If protectedBlocks.Contains(blockName) Then Continue For
                    ' Запазваме оригиналния слой и цвят на блока
                    Dim blockLayer As String = br.Layer
                    Dim blockColor As Color = br.Color
                    ' ===============================
                    ' СТЪПКА 1: Конвертиране на атрибути в DBText
                    ' ===============================
                    For Each attId As ObjectId In br.AttributeCollection
                        Dim attRef As AttributeReference = tr.GetObject(attId, OpenMode.ForRead)
                        If String.IsNullOrWhiteSpace(attRef.TextString) Then Continue For
                        Dim newText As New DBText()
                        newText.SetDatabaseDefaults(db)
                        newText.TextString = attRef.TextString
                        newText.Position = attRef.Position
                        newText.Height = attRef.Height
                        newText.Rotation = attRef.Rotation
                        newText.TextStyleId = attRef.TextStyleId
                        ' Наследяване на слой и цвят
                        If attRef.Layer = "0" Then
                            newText.Layer = blockLayer
                            newText.Color = blockColor
                        Else
                            newText.Layer = attRef.Layer
                            newText.Color = attRef.Color
                        End If
                        btrCurrent.AppendEntity(newText)
                        tr.AddNewlyCreatedDBObject(newText, True)
                    Next
                    '===============================
                    'СТЪПКА 2: Explode и добавяне на геометрията
                    '===============================
                    Dim explodedObjects As New DBObjectCollection()
                    br.Explode(explodedObjects)
                    For Each obj As DBObject In explodedObjects
                        Dim subEnt As Entity = TryCast(obj, Entity)
                        If subEnt Is Nothing Then Continue For
                        ' Пропускаме атрибути
                        If TypeOf subEnt Is AttributeReference OrElse TypeOf subEnt Is AttributeDefinition Then
                            Continue For
                        End If
                        ' Наследяване на слой и цвят
                        If subEnt.Layer = "0" Then
                            subEnt.Layer = blockLayer
                        End If
                        If subEnt.Color.ColorMethod = ColorMethod.ByBlock Then
                            subEnt.Color = blockColor
                        End If
                        ' Добавяме в текущото пространство
                        btrCurrent.AppendEntity(subEnt)
                        tr.AddNewlyCreatedDBObject(subEnt, True)
                    Next
                    ' ===============================
                    ' СТЪПКА 3: Изтриване на оригиналния блок
                    ' ===============================
                    br.UpgradeOpen()
                    br.Erase()
                    count += 1
                Next
                sw.WriteLine($"Native BURST: Обработени {count} блока.")
                tr.Commit()
            End Using
        Catch ex As Exception
            SaveError(ex, db.Filename)
        End Try
    End Sub
    ''' <summary>
    ''' Изтрива всички Layout-и, съдържащи "настройки" в името, с изключение на "Model".
    ''' </summary>
    ''' <param name="doc">Текущият AutoCAD документ</param>
    Private Sub DeleteSettingsLayouts(doc As Document)
        sw.WriteLine("1. Премахване на излишни листове...")
        ' Взимаме базата данни от документа
        Dim db As Database = doc.Database
        Try
            Using tr As Transaction = db.TransactionManager.StartTransaction()
                Dim layoutDict As DBDictionary = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead)
                Dim layMgr As LayoutManager = LayoutManager.Current
                Dim toDelete As New List(Of String)
                Dim notDeleteCount As Integer = 0
                ' 1. Първо преброяваме какво имаме
                For Each entry As DictionaryEntry In layoutDict
                    Dim name As String = entry.Key.ToString()
                    If name.ToLower() <> "model" Then
                        ' Проверяваме дали името съдържа "настройки"
                        If name.ToLower().Contains("настройки") Then
                            toDelete.Add(name)
                        Else
                            ' Броим листовете, които ще останат (напр. ако вече има Layout1)
                            notDeleteCount += 1
                        End If
                    End If
                Next
                ' 2. Превключваме на Model, за да можем да трием безопасно
                layMgr.CurrentLayout = "Model"
                ' 3. ПРОВЕРКА: Ако всички листове са за триене (както на снимката)
                If notDeleteCount = 0 Then
                    ' Проверяваме дали случайно "Layout1" вече не съществува
                    If Not layoutDict.Contains("Layout1") Then
                        layMgr.CreateLayout("Layout1")
                        sw.WriteLine("Създаден нов 'Layout1', за да не остане базата празна.")
                    End If
                    notDeleteCount = 1 ' Вече имаме един лист, който ще остане
                End If
                ' 4. Трием само ако имаме поне един лист, който ще оцелее
                If toDelete.Count > 0 And notDeleteCount > 0 Then
                    Dim deletedCount As Integer = 0
                    For Each name As String In toDelete
                        ' Допълнителна защита: AutoCAD изисква поне 1 Layout + Model (общо 2 в речника)
                        If layoutDict.Count > 2 Then
                            Try
                                layMgr.DeleteLayout(name)
                                deletedCount += 1
                            Catch ex As System.Exception
                                sw.WriteLine("Грешка при триене на " & name & ": " & ex.Message)
                            End Try
                        End If
                    Next
                    ' ВАЖНО: Тук е мястото за актуализация на връзките, ако е нужно
                    ' Преди записа на файла
                    tr.Commit()
                    sw.WriteLine("Изтрити листове: " & deletedCount)
                Else
                    ' Ако нищо не е променено
                    tr.Abort()
                End If
            End Using
        Catch ex As Exception
            SaveError(ex, db.Filename)
        End Try
    End Sub
    ''' <summary>
    ''' Изтрива всички обекти в Model Space, чиито центрови точки попадат в зададената зона (xMin, xMax, yMin, yMax).
    ''' </summary>
    ''' <param name="doc">Текущият AutoCAD документ</param>
    Private Sub WipeModelSpaceByArea(doc As Document)
        Dim db As Database = doc.Database
        sw.WriteLine("2. Почистване на обекти извън работната зона...")
        Try
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
                    sw.WriteLine("Изтрити обекти от Model Space: " & deletedCount & ". Файлът е записан.")
                End If
            End Using
        Catch ex As Exception
            SaveError(ex, db.Filename)
        End Try
    End Sub
    ''' <summary>
    ''' Изчиства съдържанието на всички атрибути в динамични или обикновени блокове с дадено име.
    ''' </summary>
    ''' <param name="doc">Текущият AutoCAD документ</param>
    ''' <param name="targetName">Името на блока, чийто атрибути ще бъдат изчистени (пример: "Качване")</param>
    Private Sub ClearAttributesInDynamicBlocks(doc As Document, targetName As String)
        Dim db As Database = doc.Database
        sw.WriteLine("3. Изчистване съдържанието на блокове 'Качване'...")
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
            sw.WriteLine("Изчистени атрибути в динамичен блок '" & targetName & "': " & count)
        End Using
    End Sub
    ''' <summary>
    ''' Разбива всички масиви (BlockReference масиви) в ModelSpace
    ''' без да пипа реалните блокове и атрибути.
    ''' </summary>
    Public Sub ExplodeAllArrays(doc As Document)
        Dim db As Database = doc.Database
        Try
            sw.WriteLine("6: Разбиване на блокове) ...")
            ' Стартираме транзакция за безопасна работа с обекти
            Using tr As Transaction = db.TransactionManager.StartTransaction()
                Dim bt As BlockTable = CType(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                Dim ms As BlockTableRecord = CType(tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)
                Dim idsToErase As New List(Of ObjectId)
                ' Обхождаме всички обекти в ModelSpace
                Dim count As Integer = 0
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
                            count += 1
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
                sw.WriteLine("ExplodeAllArrays: Обработени " & count & " масива.")
            End Using
        Catch ex As Exception
            SaveError(ex, db.Filename)
        End Try
    End Sub
    ''' <summary>
    ''' Търси в текущия чертеж текстове и блокове, свързани с мълниезащита.
    ''' При намиране на конкретен текст го заменя с нов MText,
    ''' като използва параметри и атрибути от съответния блок „Мълниезащита вертикално“.
    ''' </summary>
    ''' <param name="doc">Текущият AutoCAD документ</param>
    Private Sub FindMylniq(doc As Document)
        Dim db As Database = doc.Database
        sw.WriteLine("3. Изчистване съдържанието на блокове 'МЪЛНИЯ'...")
        ' Списъци за съхранение на намерените обекти (ID-та)
        Dim mylniqTextIds As New List(Of ObjectId)
        Dim mylniqBlockIds As New List(Of ObjectId)
        ' Променливи за стойности, извлечени от блока
        Dim valTip As String = ""
        Dim valH As String = ""
        Dim valKategoria As String = ""
        Dim valRa As String = ""
        Try
            Using tr As Transaction = db.TransactionManager.StartTransaction()
                ' Взимаме ModelSpace (или текущото пространство)
                Dim btrCurrent As BlockTableRecord = tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead)
                ' Обхождаме всички обекти
                For Each id As ObjectId In btrCurrent
                    If id.IsErased Then Continue For
                    Dim ent As Entity = TryCast(tr.GetObject(id, OpenMode.ForRead), Entity)
                    If ent Is Nothing Then Continue For
                    ' --- Точен текст ---
                    If TypeOf ent Is MText OrElse TypeOf ent Is DBText Then
                        Dim content As String = If(TypeOf ent Is MText,
                                                   DirectCast(ent, MText).Contents,
                                                   DirectCast(ent, DBText).TextString)

                        ' Търсим **точната фраза**
                        If content.IndexOf("За защита от мълнии да се монтира мълниеприемник с изпреварващо действие", StringComparison.OrdinalIgnoreCase) >= 0 Then
                            mylniqTextIds.Add(id)
                        End If
                        ' --- Точен блок ---
                    ElseIf TypeOf ent Is BlockReference Then
                        Dim br As BlockReference = DirectCast(ent, BlockReference)
                        Dim btr As BlockTableRecord = TryCast(tr.GetObject(br.DynamicBlockTableRecord, OpenMode.ForRead), BlockTableRecord)

                        ' Пропускаме Xref
                        If btr.IsFromExternalReference AndAlso Not btr.IsUnloaded Then Continue For
                        ' Търсим **точния блок**
                        If btr.Name.Trim().Equals("Мълниезащита вертикално", StringComparison.OrdinalIgnoreCase) Then
                            mylniqBlockIds.Add(id)
                        End If
                    End If
                Next
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
            sw.WriteLine("Грешка в FindMylniq: " & ex.Message)
        End Try
    End Sub
    ''' <summary>
    ''' Дълбоко почистване на неизползвани блокове, слоеве, линии и стилове.
    ''' </summary>
    Private Sub NativePurge(doc As Document)
        Dim db As Database = doc.Database
        Try
            Dim changed As Boolean = True
            sw.WriteLine("--- Пълно почистване на неизползвани слоеве и блокове...")
            ' Въртим цикъла докато спрем да намираме излишни неща (заради вложените зависимости)
            Dim count As Integer = 0
            While changed
                changed = False
                Using tr As Transaction = db.TransactionManager.StartTransaction()
                    ' Колекция за всички кандидати за триене
                    Dim idsToCheck As New ObjectIdCollection()
                    ' 1. Добавяме Блоковете
                    Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                    For Each id As ObjectId In bt
                        idsToCheck.Add(id)
                    Next
                    ' 2. Добавяме Слоевете
                    Dim lt As LayerTable = tr.GetObject(db.LayerTableId, OpenMode.ForRead)
                    For Each id As ObjectId In lt
                        idsToCheck.Add(id)
                    Next
                    ' 3. Добавяме Текстовите стилове
                    Dim tst As TextStyleTable = tr.GetObject(db.TextStyleTableId, OpenMode.ForRead)
                    For Each id As ObjectId In tst
                        idsToCheck.Add(id)
                    Next
                    ' 4. Добавяме Linetypes (Типове линии)
                    Dim ltt As LinetypeTable = tr.GetObject(db.LinetypeTableId, OpenMode.ForRead)
                    For Each id As ObjectId In ltt
                        idsToCheck.Add(id)
                    Next
                    ' 5. Добавяме DimStyles (Размерни стилове)
                    Dim dst As DimStyleTable = tr.GetObject(db.DimStyleTableId, OpenMode.ForRead)
                    For Each id As ObjectId In dst
                        idsToCheck.Add(id)
                    Next
                    ' --- МАГИЯТА: db.Purge ---
                    ' Този метод премахва от списъка всичко, което СЕ ПОЛЗВА.
                    ' В idsToCheck остават само ненужните.
                    db.Purge(idsToCheck)
                    ' Ако има останали обекти, ги трием
                    If idsToCheck.Count > 0 Then
                        For Each id As ObjectId In idsToCheck
                            Dim obj As DBObject = tr.GetObject(id, OpenMode.ForWrite)
                            obj.Erase()
                        Next
                        changed = True ' Намерихме нещо, значи въртим цикъла пак
                    End If
                    tr.Commit()
                End Using
                count += 1
            End While
            sw.WriteLine("ExplodeAllArrays: Обработени " & count & " масива.")
        Catch ex As Exception
            SaveError(ex, db.Filename)
        End Try
    End Sub
    ''' <summary>
    ''' Вгражда всички прикачени Xref-ове в чертежа като локални блокове
    ''' чрез класически Bind с префикси.
    ''' </summary>
    ''' <param name="db">Активната база данни на AutoCAD документа</param>
    Private Sub NativeBind(db As Database)
        ' Записваме информация за текущия файл в лог файла.
        sw.WriteLine($"---Native BIND на Xref-ове за файл: {IO.Path.GetFileName(db.Filename)}")
        Try
            ' Колекция с ObjectId на всички Xref BlockTableRecord-и.
            Dim xrefsCollection As New ObjectIdCollection()
            ' Обхождаме BlockTable и събираме всички заредени Xref-и.
            Using tr As Transaction = db.TransactionManager.StartTransaction()
                Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                For Each btrId As ObjectId In bt
                    Dim btr As BlockTableRecord = tr.GetObject(btrId, OpenMode.ForRead)
                    ' Проверяваме дали записът е Xref и дали не е Unloaded.
                    If btr.IsFromExternalReference AndAlso Not btr.IsUnloaded Then
                        xrefsCollection.Add(btrId)
                    End If
                Next
                tr.Commit()
            End Using
            ' Ако има намерени Xref-ове, изпълняваме Bind директно върху базата.
            If xrefsCollection.Count > 0 Then
                Try
                    ' False = класически Bind с префикси (Bind, не Insert).
                    db.BindXrefs(xrefsCollection, False)
                    sw.WriteLine($"--- Успешно вградени {xrefsCollection.Count} Xref-а.")
                Catch ex As Exception
                    sw.WriteLine("--- Грешка при изпълнение на BindXrefs.")
                    SaveError(ex, db.Filename)
                End Try
            Else
                sw.WriteLine("--- Няма Xref-ове за вграждане.")
            End If
        Catch ex As Exception
            SaveError(ex, db.Filename)
        End Try
    End Sub
    ''' <summary>
    ''' Batch обработка на DWG файлове в папка.
    ''' Всеки файл се отваря скрито, изпълнява се NativeBind,
    ''' и резултатът се записва в поддиректория "Документация".
    ''' </summary>
    ''' <param name="folderPath">Път до папката с DWG файлове</param>
    Public Sub BatchCleaner(folderPath As String)
        ' Всички DWG файлове в подадената папка.
        Dim dwgFiles() As String = IO.Directory.GetFiles(folderPath, "*.dwg")
        ' Път до папката "Документация".
        Dim subFolderPath As String = IO.Path.Combine(folderPath, "Документация")
        ' Създаваме папката "Документация", ако не съществува.
        If Not IO.Directory.Exists(subFolderPath) Then
            IO.Directory.CreateDirectory(subFolderPath)
        End If
        ' Обработваме последователно всички DWG файлове.
        For Each dwgPath In dwgFiles
            Dim fileName As String = IO.Path.GetFileName(dwgPath)
            Dim newSavePath As String = IO.Path.Combine(subFolderPath, fileName)
            Try
                ' ===============================
                ' СТЪПКА 1: Тихо отваряне на файла
                ' ===============================
                Dim doc As Document = Nothing
                Using lock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument()
                    ' Отваряме DWG файла без визуален интерфейс.
                    doc = Application.DocumentManager.Open(dwgPath, False)
                End Using
                ' ===============================
                ' СТЪПКА 2: Bind и запис
                ' ===============================
                Using docLock As DocumentLock = doc.LockDocument()
                    ' Изпълняваме NativeBind върху отворения документ.
                    NativeBind(doc.Database)
                    ' Записваме файла в папка "Документация".
                    doc.Database.SaveAs(newSavePath, DwgVersion.Current)
                End Using
                ' Затваряме документа без допълнително записване.
                doc.CloseAndDiscard()
                sw.WriteLine($"Обработен: {fileName}")
            Catch ex As Exception
                SaveError(ex, dwgPath)
            End Try
        Next
        ' Тук може да се добави автоматично отваряне на обработените файлове.
        OpenProcessedFiles(subFolderPath)
    End Sub
    ' ================================
    ' OpenProcessedFiles – отваряне на копията и пускане на RunCleaner
    ' ================================
    Private Sub OpenProcessedFiles(targetFolder As String)
        Dim filesToOpen() As String = IO.Directory.GetFiles(targetFolder, "*.dwg")
        For Each filePath In filesToOpen
            Dim doc As Document = Nothing
            Try
                ' Отваряме DWG като жив документ
                doc = Application.DocumentManager.Open(filePath, False)
                Application.DocumentManager.MdiActiveDocument = doc
                ' Заключваме документа за безопасност
                Using doc.LockDocument()
                    RunCleaner(doc)  ' Тук всички Native операции са безопасни
                End Using
                ' Записваме и затваряме файла
                ' CloseAndSave автоматично записва промените и затваря документа
                sw.WriteLine($"--- ОПИТВАМ СЕ ДА ЗАПИША ФАЙЛ -> {filePath}")
                ' === ЗАПИСВАМЕ КОПИЕ БЕЗ ДА ЗАСЯГАМЕ ОРИГИНАЛА ===
                doc.Database.SaveAs(filePath, DwgVersion.Current)
                doc.CloseAndDiscard()
            Catch ex As Exception
                SaveError(ex, filePath)
                ' При грешка опитваме да затворим файла без запис, за да не остане отворен
                If doc IsNot Nothing Then
                    Try
                        If Not doc.IsDisposed Then
                            doc.CloseAndDiscard()
                        End If
                    Catch
                        ' Ако не може да се затвори, поне опитахме
                    End Try
                End If
            End Try
        Next
    End Sub
    ''' <summary>
    ''' Записва информация за грешка в централен ErrorLog файл.
    ''' Използва се при обработка на изключения в различни методи на класа.
    ''' </summary>
    ''' <param name="ex">Обектът изключение, съдържащ информация за грешката</param>
    ''' <param name="filePath">Пътят до DWG файла, при обработката на който е възникнала грешката</param>
    Private Sub SaveError(ex As Exception, filePath As String)
        ' Определяне на пътя до лог файла за грешки
        ' Използва се константата ERROR_LOG_PATH, дефинирана в началото на класа
        Dim logPath As String = ERROR_LOG_PATH
        ' Отваряне на StreamWriter за записване в лог файла
        ' Параметърът True означава append mode - новите записи се добавят в края на файла
        ' Using statement гарантира автоматично затваряне и освобождаване на ресурсите
        Using swError As New IO.StreamWriter(logPath, True)
            ' ========================================
            ' ИЗВЛЕЧВАНЕ НА ДЕТАЙЛНА ИНФОРМАЦИЯ ЗА ГРЕШКАТА
            ' ========================================
            ' Създаване на StackTrace обект от изключението
            ' Параметърът True указва да се зарежда информация от .pdb файла (debug symbols)
            ' Това позволява да се получи точния номер на реда, където е възникнала грешката
            Dim st As New System.Diagnostics.StackTrace(ex, True)
            ' Вземане на първия фрейм от стека на извикванията
            ' GetFrame(0) връща фрейма, където е хвърлено изключението
            ' Това е най-важният фрейм, защото показва точното място на грешката
            Dim frame As System.Diagnostics.StackFrame = st.GetFrame(0)
            ' Извличане на номера на реда от изходния код
            ' GetFileLineNumber() връща номера на реда, ако .pdb файлът е наличен
            ' Ако няма .pdb файл, връща 0
            Dim line As Integer = frame.GetFileLineNumber()
            ' Извличане на името на метода, в който е възникнала грешката
            ' GetMethod().Name връща името на метода като низ
            Dim methodName As String = frame.GetMethod().Name
            ' ========================================
            ' ЗАПИСВАНЕ НА ИНФОРМАЦИЯТА В ЛОГ ФАЙЛА
            ' ========================================
            ' Разделителна линия за по-добра четимост в лог файла
            swError.WriteLine("========================================")
            ' Записване на датата и часа на възникване на грешката
            ' Форматът е yyyy-MM-dd HH:mm:ss за стандартизиран вид
            swError.WriteLine("Дата/час: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
            ' Записване на пътя до DWG файла, който се обработваше при грешката
            ' Това помага да се идентифицира проблемния файл
            swError.WriteLine("Файл (DWG): " & filePath)
            ' Записване на съобщението за грешка от изключението
            ' Това е основното описание на проблема
            swError.WriteLine("Грешка: " & ex.Message)
            ' Проверка дали е наличен номер на реда (т.е. дали има .pdb файл)
            If line > 0 Then
                ' Ако има .pdb файл, записваме точния номер на реда и името на метода
                ' Това е много полезно за бързо намиране на проблемния код
                swError.WriteLine("ГРЕШКА В КОДА НА РЕД: " & line)
                swError.WriteLine("МЕТОД: " & methodName)
            Else
                ' Ако няма .pdb файл, записваме предупреждение
                ' Без .pdb файл не може да се определи точният ред на грешката
                swError.WriteLine("Ред: Не е открит (Увери се, че .pdb файлът е в папката на AutoCAD)")
            End If
            ' Записване на Source - името на приложението или обекта, който е причинил грешката
            ' Полезно за идентифициране дали грешката идва от AutoCAD API, .NET Framework и т.н.
            swError.WriteLine("Source: " & ex.Source)
            ' Записване на пълния StackTrace
            ' Това показва целия път на извикванията от началото до мястото на грешката
            ' Много полезно за проследяване на сложни проблеми
            swError.WriteLine("Full StackTrace: ")
            swError.WriteLine(ex.StackTrace)
            ' Затваряща разделителна линия
            swError.WriteLine("========================================")
        End Using
        ' Тук StreamWriter автоматично се затваря и освобождава благодарение на Using statement
    End Sub
End Class