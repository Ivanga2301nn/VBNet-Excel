Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports Excel = Microsoft.Office.Interop.Excel
Imports Forms = System.Windows.Forms
Imports Word = Microsoft.Office.Interop.Word


''' <summary>
''' Клас за създаване на финална документационна папка и копиране на файлове от текущия проект.
''' </summary>
Public Class Dokmentaci
    Private zapis As New Dictionary(Of String, String)
    ''' <summary>
    ''' Команда за AutoCAD "Dokmentaciq".
    ''' Създава папка "Документация", изчиства я ако съществува и копира готовите файлове.
    ''' След това генерира два PDF файла от Word документа "Обяснителна записка.docx".
    ''' </summary>
    <CommandMethod("Dokumentaciq")>
    <CommandMethod("Документация")>
    Public Sub Dokmentaciq()

        Dim doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

        ' Проверка дали текущият DWG файл е записан
        If String.IsNullOrEmpty(doc.Name) Then
            doc.Editor.WriteMessage(vbLf & "Файлът трябва да е записан, за да се създаде папка Документация.")
            Exit Sub
        End If

        If Not zapis.Any() Then
            ' Тук кодът се изпълнява, когато речникът е ПРАЗЕН
            FillInsertSignatureAttributes(zapis)
        End If


        ' Взимаме пътя на текущия DWG файл
        Dim dwgPath = Path.GetDirectoryName(doc.Name)
        ' Създаваме пътя към папката "Документация" под текущата папка
        Dim dirPath = Path.Combine(dwgPath, "Документация")

        ' Ако папката не съществува – създаваме я
        If Directory.Exists(dirPath) = False Then
            Directory.CreateDirectory(dirPath)
        Else
            ' Ако съществува – изтриваме всички файлове вътре
            DeleteAllFiles(dirPath)
        End If

        Dim f1 = "Обяснителна записка.docx"
        Dim f2 = "Количествена сметка.xlsx"
        Dim f3 = "Block.dwg"
        Dim f4 = "Графична част.pdf"
        Dim f5 = "Светлотехнически.pdf"


        Dim openFileDialog As New Forms.OpenFileDialog()
        ' Задаваме текущата папка на приложението като начална
        openFileDialog.InitialDirectory = dwgPath
        openFileDialog.Title = "Моля, изберете файлoве за копиране - СЪДЪРЖАЩИ ЧЕРТЕЖИТЕ "
        ' Филтър за типовете файлове, които търсим
        openFileDialog.Filter = "AutoCAD & Office Files|*.dwg"
        'openFileDialog.Filter = "AutoCAD & Office Files|*.dwg;*.docx;*.xlsx;*.pdf|All files (*.*)|*.*"
        openFileDialog.Multiselect = True ' ТОВА ПОЗВОЛЯВА МНОЖЕСТВЕН ИЗБОР
        ' 3. Показваме диалога
        If openFileDialog.ShowDialog() = Forms.DialogResult.OK Then
            ' 3. Цикъл за копиране на всеки избран файл
            For Each sourceFile As String In openFileDialog.FileNames
                Dim fileNameOnly As String = System.IO.Path.GetFileName(sourceFile)
                ' Копиране на файловете в папката "Документация"
                CopyFile(dwgPath, dirPath, fileNameOnly, fileNameOnly, doc)
            Next
        End If
        openFileDialog.Title = "Моля, изберете файл за копиране - СЪДЪРЖАЩ ОБЯСНИТЕЛНАТА ЗАПИСКА"
        ' Филтър за типовете файлове, които търсим
        openFileDialog.Multiselect = False ' ТОВА ВЕЧЕ ЗАБРАНЯВА МНОЖЕСТВЕН ИЗБОР
        openFileDialog.Filter = "ОБЯСНИТЕЛНА ЗАПИСКА|*.docx"
        If openFileDialog.ShowDialog() = Forms.DialogResult.OK Then
            ' Вземаме пълния път на избрания файл
            Dim sourceFilePath As String = openFileDialog.FileName
            Dim fileNameOnly As String = System.IO.Path.GetFileName(sourceFilePath)
            CopyFile(dwgPath, dirPath, fileNameOnly, f1, doc)
        End If

        openFileDialog.Title = "Моля, изберете файл за копиране - СЪДЪРЖАЩ КОЛИЧЕСТВЕНАТА СМЕТКА"
        ' Филтър за типовете файлове, които търсим
        openFileDialog.Multiselect = False ' ТОВА ВЕЧЕ ЗАБРАНЯВА МНОЖЕСТВЕН ИЗБОР
        openFileDialog.Filter = "КОЛИЧЕСТВЕНА СМЕТКА|*.xlsx"
        If openFileDialog.ShowDialog() = Forms.DialogResult.OK Then
            ' Вземаме пълния път на избрания файл
            Dim sourceFilePath As String = openFileDialog.FileName
            Dim fileNameOnly As String = System.IO.Path.GetFileName(sourceFilePath)
            CopyFile(dwgPath, dirPath, fileNameOnly, f2, doc)
        End If

        openFileDialog.Title = "Моля, изберете файл за копиране - СЪДЪРЖАЩ СТАНОВИЩЕТО"
        ' Филтър за типовете файлове, които търсим
        openFileDialog.Multiselect = True ' ТОВА ПОЗВОЛЯВА МНОЖЕСТВЕН ИЗБОР
        openFileDialog.Filter = "Изображения|*.pdf,*.jpg;*.jpeg;*.png;*.bmp;*.tif;*.tiff|All files (*.*)|*.*"
        ' 3. Показваме диалога
        If openFileDialog.ShowDialog() = Forms.DialogResult.OK Then
            Dim selectedFiles As String() = openFileDialog.FileNames
            ' Проверка и обработка
            Dim validFiles As List(Of String) = ValidateAndGetFiles(selectedFiles, dirPath)
            ' Ако са PDF файлове, можеш да ги обработиш тук
            If Path.GetExtension(validFiles(0)).ToLower() = ".pdf" Then
                MessageBox.Show("Всички файлове са PDF. Може да се обработват.")
                ' MergePdfFiles(validFiles, "C:\Result\Merged.pdf")
            End If
        End If


        Exit Sub
        '' Копиране на файловете в папката "Документация"
        'CopyFile(dwgPath, dirPath, f1, doc)
        'CopyFile(dwgPath, dirPath, f2, doc)
        'CopyFile(dwgPath, dirPath, f3, doc)
        'CopyFile(dwgPath, dirPath, f4, doc)
        'CopyFile(dwgPath, dirPath, f5, doc)

        CopyStanovishteFile(dwgPath, dirPath, doc)

        ' Генериране на два PDF файла от Word документа
        ProcessWordFile(Path.Combine(dirPath, f1), doc)
        ' Генериране на PDF от Excel файла "KS__.xlsx"
        ProcessExcelFile(Path.Combine(dirPath, f2), doc)
        ' Експортираме останалите Layouts в един PDF
        MergeProjectPDFs(dirPath, doc)
    End Sub
    ''' <summary>
    ''' Проверява файловете и създава PDF, ако са само изображения.
    ''' Ако файловете са смесени PDF и изображения, показва съобщение и връща Nothing.
    ''' </summary>
    ''' <param name="files">Масив с файлови пътища</param>
    ''' <param name="outputPdfPath">Път за създаване на PDF от изображения</param>
    ''' <returns>Списък с валидни файлове или Nothing при смесени типове</returns>
    Public Function ValidateAndGetFiles(files As IEnumerable(Of String), outputPdfPath As String) As List(Of String)
        Dim pdfFiles As New List(Of String)
        Dim imageFiles As New List(Of String)
        Dim validFiles As New List(Of String)
        For Each f As String In files
            Dim ext As String = Path.GetExtension(f).ToLower()
            If ext = ".pdf" Then
                pdfFiles.Add(f)
            ElseIf ext = ".jpg" OrElse ext = ".jpeg" OrElse ext = ".png" OrElse ext = ".bmp" OrElse ext = ".tif" OrElse ext = ".tiff" Then
                imageFiles.Add(f)
            Else
                ' Непознат формат → пропускаме
            End If
        Next
        ' Ако има смесени типове
        If pdfFiles.Count > 0 AndAlso imageFiles.Count > 0 Then
            MessageBox.Show("Не може да се обработват PDF и картинки заедно. Моля изберете само PDF или само изображения.", "Невалиден избор", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return Nothing
        End If
        ' Всички файлове са PDF
        If pdfFiles.Count > 0 Then
            validFiles.AddRange(pdfFiles)
            Return validFiles
        End If
        ' Всички файлове са изображения → създаваме PDF
        If imageFiles.Count > 0 Then
            validFiles.AddRange(imageFiles)
            ' Създаваме PDF от изображения
            ConvertImagesToSinglePdf_iTextSharp(validFiles, outputPdfPath)
            MessageBox.Show("PDF файлът е създаден: " & outputPdfPath)
            Return validFiles
        End If
        Return Nothing
    End Function
    ''' <summary>
    ''' Конвертира множество изображения в един PDF файл, използвайки iTextSharp.
    ''' </summary>
    ''' <param name="imageFiles">Колекция от пътища към изображения (.jpg, .jpeg, .png, .bmp, .tif, .tiff)</param>
    ''' <param name="pdfPath">
    ''' Път до PDF файла за запис. 
    ''' Ако се подаде директория, автоматично се създава файл "CombinedImages.pdf" в нея.
    ''' </param>
    Public Sub ConvertImagesToSinglePdf_iTextSharp(imageFiles As IEnumerable(Of String), pdfPath As String)
        ' Сортираме изображенията по азбучен ред, за да бъдат добавени в PDF в правилната последователност.
        Dim sortedFiles As List(Of String) = imageFiles.ToList()
        sortedFiles.Sort()
        ' Създаваме нов PDF документ.
        Dim pdfDoc As New iTextSharp.text.Document()
        ' Ако pdfPath сочи към директория, добавяме подразбиращо се име на файла.
        If Directory.Exists(pdfPath) Then
            pdfPath = Path.Combine(pdfPath, "CombinedImages.pdf")
        End If
        ' Създаваме поток за запис на PDF и PdfWriter за iTextSharp.
        Using stream As New FileStream(pdfPath, FileMode.Create, FileAccess.Write, FileShare.None)
            Dim writer As PdfWriter = PdfWriter.GetInstance(pdfDoc, stream)
            pdfDoc.Open()
            ' Обхождаме всяко изображение в сортирания списък.
            For Each imgPath As String In sortedFiles
                ' Пропускаме файлове, които не съществуват.
                If Not File.Exists(imgPath) Then Continue For
                Dim ext As String = Path.GetExtension(imgPath).ToLower()
                ' Допустими формати за изображения.
                If ext <> ".jpg" AndAlso ext <> ".jpeg" AndAlso ext <> ".png" AndAlso ext <> ".bmp" AndAlso ext <> ".tif" AndAlso ext <> ".tiff" Then
                    Continue For
                End If
                ' Зареждаме изображението като iTextSharp Image.
                Dim pdfImg As iTextSharp.text.Image = iTextSharp.text.Image.GetInstance(imgPath)
                ' Настройваме размера на страницата според размерите на изображението.
                pdfDoc.SetPageSize(New iTextSharp.text.Rectangle(pdfImg.Width, pdfImg.Height))
                pdfDoc.NewPage()
                ' Поставяме изображението на страницата с абсолютни координати и скалиране.
                pdfImg.SetAbsolutePosition(0, 0)
                pdfImg.ScaleAbsolute(pdfImg.Width, pdfImg.Height)
                ' Добавяме изображението към PDF документа.
                pdfDoc.Add(pdfImg)
            Next
            ' Затваряме PDF документа след добавяне на всички изображения.
            pdfDoc.Close()
        End Using
    End Sub


    ''' <summary>
    ''' Търси първия файл в dwgPath, който съдържа "Становище" в името, и го копира в dirPath с ново име.
    ''' Новото име се генерира от стойностите на ключовете "SAP", "Ном.заявление" и "Дата_заявление" в речника.
    ''' </summary>
    ''' <param name="dwgPath">Папката, в която се търси файлът</param>
    ''' <param name="dirPath">Целевата папка за копиране</param>
    ''' <param name="doc">Активният документ на AutoCAD за съобщения</param>
    Private Sub CopyStanovishteFile(dwgPath As String, dirPath As String, doc As Autodesk.AutoCAD.ApplicationServices.Document)
        ' Генериране на новото име
        Dim sap As String = If(zapis.ContainsKey("SAP"), zapis("SAP"), "нето")
        Dim nomZayav As String = If(zapis.ContainsKey("Ном.заявление"), zapis("Ном.заявление"), "нето")
        Dim dataZayav As String = If(zapis.ContainsKey("Дата_заявление"), zapis("Дата_заявление"), "нето")
        Dim newName As String = String.Join("_", "Становище", sap, nomZayav, dataZayav) & ".pdf"
        ' Търсим първия файл, който съдържа "Становище"
        Dim srcFile As String = Directory.GetFiles(dwgPath).FirstOrDefault(Function(f) Path.GetFileName(f).Contains("Становище"))

        ' Проверка дали файлът е намерен
        If String.IsNullOrEmpty(srcFile) Then
            doc.Editor.WriteMessage(vbLf & "Не е намерен файл съдържащ 'Становище' в папка: " & dwgPath)
            Exit Sub
        End If
        ' Сглобяване на пълния път до целевия файл
        Dim dstFile As String = Path.Combine(dirPath, newName)
        ' Копиране на файла (презаписваме, ако съществува)
        File.Copy(srcFile, dstFile, True)
        ' Информационно съобщение за успешно копиране
        doc.Editor.WriteMessage(vbLf & "Копиран файл: " & Path.GetFileName(srcFile) & " -> " & newName)
    End Sub
    ''' <summary>
    ''' Взема всички атрибути от всички блокови референции на блока "Insert_Signature"
    ''' и ги записва в речника 'zapis'.
    ''' Ключът е Tag на атрибута, стойността е TextString на атрибута.
    ''' </summary>
    ''' <param name="zapis">Речникът за съхраняване на данните (ключ: Tag, стойност: TextString)</param>
    Private Sub FillInsertSignatureAttributes(zapis As Dictionary(Of String, String))
        ' Получаване на активния документ
        Dim acDoc As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ' Получаване на базата данни на активния документ
        Dim acCurDb As Database = acDoc.Database
        ' Започване на транзакция
        Using actrans As Transaction = acDoc.TransactionManager.StartTransaction()
            ' Отваряне на таблицата с блокове
            Dim acBlkTbl As BlockTable = actrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            ' Проверка дали блокът съществува
            If Not acBlkTbl.Has("Insert_Signature") Then Exit Sub
            ' Получаване на ID на записа на блока "Insert_Signature"
            Dim blkRecId As ObjectId = acBlkTbl("Insert_Signature")
            ' Получаване на записа на блока
            Dim acBlkTblRec As BlockTableRecord = actrans.GetObject(blkRecId, OpenMode.ForRead)
            ' Обхождане на всички блокови референции
            For Each blkRefId As ObjectId In acBlkTblRec.GetBlockReferenceIds(True, True)
                ' Получаване на конкретната блокова референция
                Dim acBlkRef As BlockReference = actrans.GetObject(blkRefId, OpenMode.ForWrite)
                ' Получаване на колекцията от атрибути
                Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                ' Обхождане на всички атрибути
                For Each objID As ObjectId In attCol
                    ' Получаване на атрибута
                    Dim dbObj As DBObject = actrans.GetObject(objID, OpenMode.ForWrite)
                    Dim acAttRef As AttributeReference = dbObj
                    ' Добавяне на стойността в речника
                    ' Ключ: Tag на атрибута
                    ' Стойност: TextString на атрибута
                    zapis.Add(acAttRef.Tag, acAttRef.TextString)
                Next
            Next
            actrans.Commit()
        End Using
    End Sub
    ''' <summary>
    ''' Събира и обединява всички PDF документи за проект "Част електро" в един финален PDF файл.
    ''' 
    ''' Процедурата:
    ''' 1. Чете името на проектанта от блока "Insert_Signature" в активния чертеж.
    ''' 2. Открива последната (по година) папка с удостоверения.
    ''' 3. Избира под-папка според проектанта.
    ''' 4. Събира необходимите PDF файлове в определен ред.
    ''' 5. Обединява ги в един PDF чрез MergePDFs().
    ''' </summary>
    ''' <param name="Path_Doc">
    ''' Път до папката с проектната документация,
    ''' където се намират обяснителни записки, графична част и др.
    ''' </param>
    Public Sub MergeProjectPDFs(Path_Doc As String, mainDoc As Autodesk.AutoCAD.ApplicationServices.Document)
        ' Променлива, в която ще запишем името на проектанта
        Dim projectant As String = ""
        If zapis.ContainsKey("ПРОЕКТАНТ") Then projectant = zapis("ПРОЕКТАНТ")
        ' Основна мрежова папка с удостоверения
        Dim basePath As String = "\\MONIKA\Monika\_НАСТРОЙКИ\Udostoqwereniq"
        ' Взимаме всички папки (години)
        Dim yearDirs() As String
        yearDirs = Directory.GetDirectories(basePath)
        ' Определяме последната (най-голяма) година
        Dim lastYearDir As String = ""
        Dim maxYear As Integer = 0
        For Each dir As String In Directory.GetDirectories(basePath)
            Dim year As Integer
            If Integer.TryParse(System.IO.Path.GetFileName(dir), year) Then
                If year > maxYear Then
                    maxYear = year
                    lastYearDir = dir
                End If
            End If
        Next
        ' Определяме папката на проектанта според името
        Dim targetSubFolder As String = ""
        Dim subDirs() As String = Directory.GetDirectories(lastYearDir)
        Select Case True
            Case projectant.ToLower().Contains("тонкова")
                targetSubFolder =
                System.IO.Path.Combine(lastYearDir, "МОНИКА")
            Case projectant.ToLower().Contains("василев")
                targetSubFolder = System.IO.Path.Combine(lastYearDir, "ЕВГЕНИ")
            Case projectant.ToLower().Contains("иванова")
                targetSubFolder =
                System.IO.Path.Combine(lastYearDir, "ИВАН")
        End Select
        ' Взимаме всички файлове в папката на проектанта
        Dim filesInFolder() As String = Directory.GetFiles(targetSubFolder)
        ' Търсим PDF файл, съдържащ "ppp" в името
        Dim pppFile As String = filesInFolder.FirstOrDefault(Function(f) Path.GetFileName(f).ToLower().Contains("ppp"))
        ' Търсим първия файл в Path_Doc, съдържащ "Становище"
        Dim stanovishteFile As String =
            Directory.GetFiles(Path_Doc).FirstOrDefault(Function(f) Path.GetFileName(f).Contains("Становище"))
        ' Списък с PDF документи за обединяване
        Dim Dokuments As New List(Of String)
        ' Добавяме файловете в точния ред, ако съществуват
        If System.IO.File.Exists(Path.Combine(Path_Doc, "Обяснителна записка 1.pdf")) Then
            Dokuments.Add(Path.Combine(Path_Doc, "Обяснителна записка 1.pdf"))
            mainDoc.Editor.WriteMessage(vbLf & "Добавен файл: Обяснителна записка 1.pdf")
        End If

        If System.IO.File.Exists(Path.Combine(Path_Doc, pppFile)) Then
            Dokuments.Add(Path.Combine(Path_Doc, pppFile))
            mainDoc.Editor.WriteMessage(vbLf & "Добавен файл: " & pppFile)
        End If

        If System.IO.File.Exists(Path.Combine(targetSubFolder, "Застраховка.pdf")) Then
            Dokuments.Add(Path.Combine(targetSubFolder, "Застраховка.pdf"))
            mainDoc.Editor.WriteMessage(vbLf & "Добавен файл: " & targetSubFolder & "\Застраховка.pdf")
        End If

        If Not String.IsNullOrEmpty(stanovishteFile) Then
            Dokuments.Add(stanovishteFile)
            mainDoc.Editor.WriteMessage(vbLf & "Добавен файл: " & Path.GetFileName(stanovishteFile))
        End If

        If System.IO.File.Exists(Path.Combine(Path_Doc, "Обяснителна записка 2.pdf")) Then
            Dokuments.Add(Path.Combine(Path_Doc, "Обяснителна записка 2.pdf"))
            mainDoc.Editor.WriteMessage(vbLf & "Добавен файл: Обяснителна записка 2.pdf")
        End If

        If System.IO.File.Exists(Path.Combine(Path_Doc, "Светлотехнически.pdf")) Then
            Dokuments.Add(Path.Combine(Path_Doc, "Светлотехнически.pdf"))
            mainDoc.Editor.WriteMessage(vbLf & "Добавен файл: Светлотехнически.pdf")
        End If

        If System.IO.File.Exists(Path.Combine(Path_Doc, "Обяснителна записка 3.pdf")) Then
            Dokuments.Add(Path.Combine(Path_Doc, "Обяснителна записка 3.pdf"))
            mainDoc.Editor.WriteMessage(vbLf & "Добавен файл: Обяснителна записка 3.pdf")
        End If

        If System.IO.File.Exists(Path.Combine(Path_Doc, "Количествена_сметка.pdf")) Then
            Dokuments.Add(Path.Combine(Path_Doc, "Количествена_сметка.pdf"))
            mainDoc.Editor.WriteMessage(vbLf & "Добавен файл: Количествена_сметка.pdf")
        End If

        If System.IO.File.Exists(Path.Combine(Path_Doc, "Графична част.pdf")) Then
            Dokuments.Add(Path.Combine(Path_Doc, "Графична част.pdf"))
            mainDoc.Editor.WriteMessage(vbLf & "Добавен файл: Графична част.pdf")
        End If

        MergePDFs(Path.Combine(Path_Doc, "Част електро.pdf"), Dokuments)
        mainDoc.Editor.WriteMessage(vbLf & "Създаден файл: Част електро.pdf")
    End Sub
    ''' <summary>
    ''' Обединява няколко PDF файла в един.
    ''' </summary>
    ''' <param name="outputFile">Път и име на крайния обединен PDF файл</param>
    ''' <param name="inputFiles">Списък с пълни пътища към PDF файловете, които ще се обединят</param>
    Public Sub MergePDFs(outputFile As String, inputFiles As List(Of String))
        Try
            ' --------------------------------------------------------------
            ' FileStream: обект за работа с физическия файл на диска.
            ' FileMode.Create: създава нов файл, ако не съществува, 
            ' или презаписва стар файл със същото име.
            ' Това е “потокът”, в който ще се записва крайният PDF.
            ' --------------------------------------------------------------
            Using pdfStream As New System.IO.FileStream(outputFile, System.IO.FileMode.Create)
                ' --------------------------------------------------------------
                ' Document: обектът на iTextSharp, който представлява PDF документа в паметта.
                ' Това е контейнерът, в който ще се добавят всички страници.
                ' --------------------------------------------------------------
                Dim pdfContainer As New iTextSharp.text.Document()
                ' --------------------------------------------------------------
                ' PdfCopy: "двигателят", който копира страници от други PDF файлове в нашия PDF контейнер.
                ' Параметри:
                ' - pdfContainer: контейнерът, който създаваме
                ' - pdfStream: потокът, където ще се запише крайният PDF
                ' --------------------------------------------------------------
                Dim pdfEngine As New iTextSharp.text.pdf.PdfCopy(pdfContainer, pdfStream)
                ' --------------------------------------------------------------
                ' Отваряме контейнера за писане. 
                ' Всички операции с pdfEngine трябва да се извършват след това.
                ' --------------------------------------------------------------
                pdfContainer.Open()
                ' --------------------------------------------------------------
                ' Обхождаме списъка с PDF файловете, които трябва да се обединят
                ' --------------------------------------------------------------
                For Each filePath In inputFiles
                    ' Проверяваме дали файлът съществува, за да избегнем грешки
                    If System.IO.File.Exists(filePath) Then
                        Try
                            ' ----------------------------------------------------------
                            ' PdfReader: обект, който чете съдържанието на текущия PDF файл
                            ' Параметър: filePath – пълният път към PDF файла
                            ' ----------------------------------------------------------
                            Dim pdfSource As New iTextSharp.text.pdf.PdfReader(filePath)
                            ' ----------------------------------------------------------
                            ' Обхождаме всички страници на текущия PDF файл
                            ' NumberOfPages връща броя на страниците в pdfSource
                            ' ----------------------------------------------------------
                            For i As Integer = 1 To pdfSource.NumberOfPages
                                ' Взимаме конкретната страница от pdfSource
                                ' PdfImportedPage е представяне на страница, което pdfEngine може да добави
                                Dim importedPage As iTextSharp.text.pdf.PdfImportedPage = pdfEngine.GetImportedPage(pdfSource, i)
                                ' Добавяме страницата към новия PDF документ
                                pdfEngine.AddPage(importedPage)
                            Next
                            ' ----------------------------------------------------------
                            ' Затваряме PdfReader за текущия файл, за да освободим ресурсите
                            ' ----------------------------------------------------------
                            pdfSource.Close()
                        Catch ex As Exception
                            ' Ако има проблем с четенето на PDF файла, показваме съобщение
                            MsgBox("Грешка при четене на PDF: " & filePath & vbCrLf & ex.Message)
                        End Try
                    End If
                Next
                ' --------------------------------------------------------------
                ' Затваряме финалния PDF контейнер
                ' Всички добавени страници се записват физически във файла
                ' --------------------------------------------------------------
                pdfContainer.Close()
            End Using
        Catch ex As Exception
            ' Ако възникне грешка при създаването на крайния PDF, показваме съобщение
            MsgBox("Грешка при генериране на финалния PDF: " & ex.Message)
        End Try
    End Sub
    ''' <summary>
    ''' Копира файл от изходната папка към целевата папка.
    ''' Копира файл от една директория в друга.
    ''' Процедурата сглобява пълните пътища до изходния и целевия файл,
    ''' проверява дали изходният файл съществува и ако да – го копира,
    ''' като презаписва съществуващ файл със същото име.
    ''' Всички съобщения за успех или грешка се извеждат
    ''' в командния ред на AutoCAD.
    ''' </summary>
    ''' <param name="dwgPath">
    ''' Път до директорията, от която ще се копира файлът.
    ''' </param>
    ''' <param name="dirPath">
    ''' Път до директорията, в която ще бъде копиран файлът.
    ''' </param>
    ''' <param name="fn">
    ''' Име на файла (включително разширението), който трябва да бъде копиран.
    ''' </param>
    ''' <param name="doc">
    ''' Активният AutoCAD документ, използван за извеждане
    ''' на съобщения към потребителя.
    ''' </param>
    Private Sub CopyFile(dwgPath As String, dirPath As String, FileName As String, newFile As String, doc As Autodesk.AutoCAD.ApplicationServices.Document)
        ' Сглобяване на пълния път до изходния файл
        Dim src = Path.Combine(dwgPath, FileName)
        ' Сглобяване на пълния път до целевия файл
        Dim dst = Path.Combine(dirPath, newFile)
        ' Проверка дали изходният файл съществува
        If File.Exists(src) = False Then
            ' Ако файлът липсва, извеждаме съобщение в командния ред на AutoCAD
            doc.Editor.WriteMessage(vbLf & "Липсва файл: " & FileName)
            ' Прекратяване на процедурата, защото няма какво да копираме
            Exit Sub
        End If
        ' Копиране на файла в целевата директория
        ' Третият параметър True означава, че файлът ще бъде презаписан,
        ' ако вече съществува
        File.Copy(src, dst, True)
        ' Информационно съобщение за успешно копиране
        doc.Editor.WriteMessage(vbLf & "Копиран файл: " & newFile)
    End Sub
    ''' <summary>
    ''' Изтрива всички файлове в дадена папка, без да изтрива самата папка.
    ''' </summary>
    Private Sub DeleteAllFiles(dirPath As String)
        Dim files = Directory.GetFiles(dirPath)
        For Each f In files
            Try
                File.Delete(f)
            Catch
                ' Игнорираме грешките при изтриване, например файл в употреба
            End Try
        Next
    End Sub
    ''' <summary>
    ''' Процедура за обработка на Word документ и експортиране на PDF файлове.
    ''' Логиката включва:
    ''' 1. Експортиране на първите две страници като "Обяснителна записка 1.pdf".
    ''' 2. Проверка за наличие на файл в папката, започващ с "свет".
    ''' 3. Ако няма такъв файл, експортира останалите страници като "Обяснителна записка 2.pdf".
    ''' 4. Ако има такъв файл, търси параграф с текст "по безопасност, хигиена на труда и пожарна безопасност"
    '''    и създава два допълнителни PDF файла:
    '''      - От страница 3 до параграфа - 1 → "Обяснителна записка 2.pdf"
    '''      - От параграфа до края → "Обяснителна записка 3.pdf"
    ''' </summary>
    ''' <param name="filePath">Пълният път до Word документа</param>
    ''' <param name="doc">Активният AutoCAD документ за извеждане на съобщения</param>
    Private Sub ProcessWordFile(filePath As String, doc As Autodesk.AutoCAD.ApplicationServices.Document)
        ' Декларация на обекти за Word
        Dim wordApp As Word.Application = Nothing
        Dim wordDoc As Word.Document = Nothing
        Try
            ' Стартираме нов екземпляр на Word
            wordApp = New Word.Application
            wordApp.Visible = False ' Word е невидим за потребителя
            ' Отваряне на Word документа
            wordDoc = wordApp.Documents.Open(filePath)
            ' Вземаме папката, където се намира документа
            Dim folderPath As String = System.IO.Path.GetDirectoryName(filePath)
            ' --- Експортиране на първи PDF ---
            ' Винаги първи PDF: страници 1 и 2
            Dim pdf1Path As String = System.IO.Path.Combine(folderPath, "Обяснителна записка 1.pdf")
            wordDoc.ExportAsFixedFormat(pdf1Path, Word.WdExportFormat.wdExportFormatPDF,
                                    OpenAfterExport:=False,
                                    OptimizeFor:=Word.WdExportOptimizeFor.wdExportOptimizeForPrint,
                                    Range:=Word.WdExportRange.wdExportFromTo,
                                    From:=1, To:=2)
            ' --- Проверка за файл започващ с "свет" ---
            ' Преглеждаме всички файлове в папката и проверяваме дали името започва с "свет" (главни/малки букви)
            Dim fileExists As Boolean = System.IO.Directory.GetFiles(folderPath).Any(Function(f) System.IO.Path.GetFileName(f).ToLower().StartsWith("свет"))
            ' --- Ако няма такъв файл ---
            If Not fileExists Then
                doc.Editor.WriteMessage(vbLf & "Няма файл, започващ с 'свет'. Създаваме втори PDF...")
                ' Вземаме общия брой страници в документа
                Dim lastPage As Integer = wordDoc.ComputeStatistics(Word.WdStatistic.wdStatisticPages)

                ' Втори PDF: от страница 3 до последната
                If lastPage >= 3 Then
                    Dim pdf2Path As String = System.IO.Path.Combine(folderPath, "Обяснителна записка 2.pdf")
                    wordDoc.ExportAsFixedFormat(pdf2Path, Word.WdExportFormat.wdExportFormatPDF,
                                            OpenAfterExport:=False,
                                            OptimizeFor:=Word.WdExportOptimizeFor.wdExportOptimizeForPrint,
                                            Range:=Word.WdExportRange.wdExportFromTo,
                                            From:=3, To:=lastPage)
                End If
                ' --- Ако има такъв файл ---
            Else
                doc.Editor.WriteMessage(vbLf & "Има файл, започващ с 'свет'. Търсим параграф за безопасност...")
                ' Търсим параграф, съдържащ текста за безопасност
                Dim targetParagraph As Word.Range = Nothing
                For Each p As Word.Paragraph In wordDoc.Paragraphs
                    If p.Range.Text.ToLower().Contains("по безопасност, хигиена на труда и пожарна безопасност") Then
                        targetParagraph = p.Range
                        Exit For ' Спираме търсенето, след като намерим първия параграф
                    End If
                Next
                Dim targetPage As Integer = 0
                If targetParagraph IsNot Nothing Then
                    ' Вземаме страницата, на която се намира параграфът
                    targetPage = targetParagraph.Information(Word.WdInformation.wdActiveEndAdjustedPageNumber)
                    doc.Editor.WriteMessage(vbLf & "Параграфът е намерен на страница: " & targetPage)
                    ' Общ брой страници
                    Dim lastPage As Integer = wordDoc.ComputeStatistics(Word.WdStatistic.wdStatisticPages)
                    ' --- Втори PDF: от страница 3 до targetPage - 1 ---
                    If targetPage > 3 Then
                        Dim pdf2Path As String = System.IO.Path.Combine(folderPath, "Обяснителна записка 2.pdf")
                        wordDoc.ExportAsFixedFormat(pdf2Path, Word.WdExportFormat.wdExportFormatPDF,
                                                OpenAfterExport:=False,
                                                OptimizeFor:=Word.WdExportOptimizeFor.wdExportOptimizeForPrint,
                                                Range:=Word.WdExportRange.wdExportFromTo,
                                                From:=3, To:=targetPage - 1)
                    End If
                    ' --- Трети PDF: от targetPage до края ---
                    If targetPage <= lastPage Then
                        Dim pdf3Path As String = System.IO.Path.Combine(folderPath, "Обяснителна записка 3.pdf")
                        wordDoc.ExportAsFixedFormat(pdf3Path, Word.WdExportFormat.wdExportFormatPDF,
                                                OpenAfterExport:=False,
                                                OptimizeFor:=Word.WdExportOptimizeFor.wdExportOptimizeForPrint,
                                                Range:=Word.WdExportRange.wdExportFromTo,
                                                From:=targetPage, To:=lastPage)
                    End If
                Else
                    ' Ако параграфът не е намерен
                    doc.Editor.WriteMessage(vbLf & "Параграфът не е намерен.")
                End If
            End If
        Catch ex As Exception
            ' Ако има грешка при обработката, се извежда съобщение в AutoCAD
            doc.Editor.WriteMessage(vbLf & "Грешка при експортиране на PDF: " & ex.Message)
        Finally
            ' Затваряне на документа и на Word, за да се освободи паметта
            If wordDoc IsNot Nothing Then wordDoc.Close(False)
            If wordApp IsNot Nothing Then wordApp.Quit(False)
        End Try
    End Sub
    ''' <summary>
    ''' Обработва Excel файла KS__.xlsx:
    ''' 1. Изтрива всички листове, с изключение на "Количествена сметка"
    ''' 2. Записва файла (в случай, че искате да запазите промените)
    ''' 3. Записва листа "Количествена сметка" като PDF
    ''' </summary>
    ''' <param name="filePath">Пътят към Excel файла</param>
    ''' <param name="doc">Текущият AutoCAD документ за писане на съобщения</param>
    Private Sub ProcessExcelFile(filePath As String, doc As Autodesk.AutoCAD.ApplicationServices.Document)
        Dim excelApp As Excel.Application = Nothing
        Dim excelBook As Excel.Workbook = Nothing
        Dim excelBooks As Excel.Workbooks = Nothing
        Try
            ' Стартиране на Excel невидимо
            excelApp = New Excel.Application
            excelApp.Visible = False
            excelApp.DisplayAlerts = False ' Изключваме предупреждения при изтриване на листове

            ' Отваряне на Excel файла
            excelBook = excelApp.Workbooks.Open(filePath)

            ' Изтриване на всички листове с изключение на "Количествена сметка"
            For i As Integer = excelBook.Sheets.Count To 1 Step -1
                Dim ws As Excel.Worksheet = excelBook.Sheets(i)
                If ws.Name <> "Количествена сметка" Then
                    ws.Delete()
                End If
            Next

            ' Записваме Excel файла, за да се запазят направените промени
            excelBook.Save()

            ' Път към PDF файла
            Dim folderPath As String = Path.GetDirectoryName(filePath)
            Dim pdfPath As String = Path.Combine(folderPath, "Количествена_сметка.pdf")

            ' Записваме листа "Количествена сметка" като PDF
            Dim wsSheet As Excel.Worksheet = excelBook.Sheets("Количествена сметка")
            wsSheet.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, pdfPath)

            ' Съобщение за успешно експортиране
            doc.Editor.WriteMessage(vbLf & "Excel файлът е обработен и PDF създаден: " & pdfPath)

        Catch ex As Exception
            ' Писане на грешка в командния ред на AutoCAD
            doc.Editor.WriteMessage(vbLf & "Грешка при обработка на Excel файла: " & ex.Message)
        Finally
            ' ВАЖНО: Правилно затваряне и освобождаване на ресурсите
            If excelBook IsNot Nothing Then
                excelBook.Close(False)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook)
            End If

            If excelBooks IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBooks)
            End If

            If excelApp IsNot Nothing Then
                excelApp.Quit()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp)
            End If

            ' Финално почистване на Garbage Collector-а
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try

    End Sub

    ''' <summary>
    ''' Функция за интерактивно питане на потребителя с опции Да/Не.
    ''' </summary>
    ''' <param name="mainDoc">Документът, в който се извежда въпроса</param>
    ''' <param name="question">Текстът на въпроса</param>
    ''' <returns>True, ако потребителят избере "Да", False при "Не" или по подразбиране</returns>
    Private Function AskYesNo(mainDoc As Autodesk.AutoCAD.ApplicationServices.Document,
                          question As String) As Boolean
        Dim ed = mainDoc.Editor
        ' Създаваме опциите за отговор с ключови думи Да/Не
        Dim opts As New Autodesk.AutoCAD.EditorInput.PromptKeywordOptions(vbLf & question & " [Да/Не]:", "Да Не")
        opts.AllowNone = True  ' Позволяваме да натисне Enter за подразбиране
        ' Получаваме отговора
        Dim res As Autodesk.AutoCAD.EditorInput.PromptResult = ed.GetKeywords(opts)
        ' Ако статусът е ОК и е избрано "Да" -> връщаме True
        If res.Status = Autodesk.AutoCAD.EditorInput.PromptStatus.OK AndAlso res.StringResult = "Да" Then
            Return True
        End If
        ' Всичко останало -> False (по подразбиране)
        Return False
    End Function
    ''' <summary>
    ''' Изтрива всички Layouts, чиито имена започват с "Настройки".
    ''' Работи върху отворено копие на DWG файла и извежда съобщения в основния документ.
    ''' </summary>
    ''' <param name="copiedDoc">DWG копието, върху което се работи</param>
    ''' <param name="mainDoc">Основният документ за извеждане на съобщения</param>
    Private Sub DeleteLayout(copiedDoc As Autodesk.AutoCAD.ApplicationServices.Document, mainDoc As Autodesk.AutoCAD.ApplicationServices.Document)
        Try
            Dim db = copiedDoc.Database

            ' Стартираме транзакция
            Using trans = db.TransactionManager.StartTransaction()
                Dim layoutDict As Autodesk.AutoCAD.DatabaseServices.DBDictionary = trans.GetObject(db.LayoutDictionaryId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                ' Списък с Layouts за изтриване
                Dim toDelete As New List(Of Autodesk.AutoCAD.DatabaseServices.ObjectId)
                ' Обхождаме LayoutDictionary
                Dim layoutEnumerator As IDictionaryEnumerator = layoutDict.GetEnumerator()
                While layoutEnumerator.MoveNext()
                    Dim layout As Autodesk.AutoCAD.DatabaseServices.Layout = trans.GetObject(DirectCast(layoutEnumerator.Value, Autodesk.AutoCAD.DatabaseServices.ObjectId), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                    ' Ако Layout започва с "Настройки", добавяме за изтриване
                    If layout.LayoutName.StartsWith("Настройки") Then
                        toDelete.Add(layout.Id)
                    End If
                End While
                ' Изтриваме избраните Layouts
                For Each id In toDelete
                    trans.GetObject(id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite).Erase()
                Next
                trans.Commit()
            End Using
            ' Съобщение в основния документ
            mainDoc.Editor.WriteMessage(vbLf & "DWG копието е обработено: Layouts ""Настройки"" са изтрити.")
        Catch ex As Exception
            mainDoc.Editor.WriteMessage(vbLf & "Грешка при изтриване на Layouts ""Настройки"": " & ex.Message)
        End Try
    End Sub
End Class
