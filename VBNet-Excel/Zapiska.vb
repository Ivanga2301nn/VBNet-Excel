Imports System.IO
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Runtime
Imports Word = Microsoft.Office.Interop.Word
Imports Autodesk.AutoCAD.EditorInput
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions

Imports System.Globalization

Imports VBNet_Excel.Form_ExcelUtilForm
Imports Forms = System.Windows.Forms
Imports excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices.ComTypes
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.TrackBar
Imports System.Drawing
Imports System.Drawing.Text
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox
Imports Microsoft.SqlServer.Server
Imports System.Security.Cryptography
Imports Microsoft.Office.Core
Imports System.Diagnostics
Imports System.Runtime.ConstrainedExecution
Public Class Zapiska
    Dim PI As Double = 3.1415926535897931
    Dim cu As CommonUtil = New CommonUtil()
    Dim ShSet As SheetSet = New SheetSet()
    Dim Кавички As String = Chr(34)
    Dim wordApp As New Word.Application
    Dim picList As New List(Of PIC)()
    Structure srtSheetSet
        Dim NameSubset As String
        Dim NameSheet As String
        Dim NumberSheet As String
    End Structure
    Structure srtCable
        Dim Тип As String
        Dim Тръба As String
        Dim Дължина As Double
        Dim Материал As Boolean
        Dim Делта_U As Double
        Dim count As Integer
    End Structure
    Structure PIC
        Dim ZN As String
        Dim NOM As String
        Dim AD As String
        Dim Tablo As String
        Dim Visibility As String
        Dim CountS As Integer
    End Structure
    <CommandMethod("Zapiska")>
    <CommandMethod("Записка")>
    Public Sub New_zapiska()
        Dim wordDoc As Word.Document
        Try
            Dim fullName As String = Application.DocumentManager.MdiActiveDocument.Name
            Dim filePath As String = Path.GetDirectoryName(fullName)
            Dim Path_Name As String = Mid(filePath, InStrRev(filePath, "\") + 1, Len(filePath))
            Dim fileName As String = filePath + "\" + "Обяснителна записка.docx"
            Dim Text As String = ""
            Dim File_DST As String = filePath + "\" + Path_Name + ".dst"
            wordDoc = OpenWordDocument(fileName, wordApp)
            Dim dicSignature As Dictionary(Of String, String)
            If wordDoc Is Nothing Then
                Exit Sub
            End If
            CreateCustomStyles(wordDoc)
            ' ApplyCustomStyles(wordDoc)
            '
            ' Извлича данните от блок "Insert_Signature"
            '
            dicSignature = addArrayBlock()

            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database

            Челен_Лист(wordDoc, dicSignature)
            СЪДЪРЖАНИЕ(wordDoc, dicSignature)
            ОБЯСНИТЕЛНА_ЗАПИСКА(wordDoc, dicSignature)
            ПРИСЪЕДИНЯВАНЕ(wordDoc, acDoc, acCurDb, dicSignature)
            Външно_Захранване(wordDoc, dicSignature)
            РАЗПРЕДЕЛИТЕЛНИ_ТАБЛА(wordDoc, acDoc, acCurDb, dicSignature)
            ОСВЕТИТEЛЕНА(wordDoc, acDoc, acCurDb)
            КОНТАКТИ(wordDoc)
            Заземителна(wordDoc, dicSignature)
            Мълниезащита(wordDoc, acDoc, acCurDb)
            Слаботокова(wordDoc, acDoc, acCurDb)
            Заключение(wordDoc, dicSignature)
            BHTB(wordDoc, dicSignature)
            POIS(wordDoc, dicSignature)
            MsgBox("Завърших записката.")
        Catch ex As Exception
            MsgBox("Error:  " & ex.Message)
        Finally
            If wordDoc IsNot Nothing Then
                Try
                    wordDoc.Save()
                    wordDoc.Close(False)
                Catch ex As Exception
                    MsgBox("Error while closing document: " & ex.Message)
                Finally
                    Marshal.ReleaseComObject(wordDoc)
                    wordDoc = Nothing
                End Try
            End If
            If wordApp IsNot Nothing Then
                ' Затваряме всички отворени документи в Word приложението
                For Each doc As Word.Document In wordApp.Documents
                    doc.Close(False)
                    Marshal.ReleaseComObject(doc)
                Next
                ' Затваряме Word приложението
                Try
                    wordApp.Quit(False)
                    Marshal.ReleaseComObject(wordApp)
                    wordApp = Nothing
                Catch ex As Exception
                    MsgBox("Error while quitting Word: " & ex.Message)
                End Try
                ' Прекратяване на процеса на Word
                For Each proc As Process In Process.GetProcessesByName("WINWORD")
                    proc.Kill()
                Next
            End If
            ' Извикайте събирача на боклука, за да освободите незабавно ресурсите
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub
    ' РАЗПРЕДЕЛИТЕЛНИ ТАБЛА.
    Private Sub РАЗПРЕДЕЛИТЕЛНИ_ТАБЛА(wordDoc As Word.Document,
                                      acDoc As Document,
                                      acCurDb As Database,
                                      dicObekt As Dictionary(Of String, String))
        Dim Text As String
        FormatParagraph(wordDoc, "РАЗПРЕДЕЛИТЕЛНИ ТАБЛА", wordApp)

        Text = "В помещение " + Кавички + dicObekt("GRT_Помещение") + Кавички
        Text += IIf(dicObekt("GRT_Кота") = "  #####  ", "", $" на кота {dicObekt("GRT_Кота")}")
        Text += " ще се монтира главното разпределителното табло "
        Text += Кавички + dicObekt("GRT_Name") + Кавички + ", изпълнено съгласно приложената схема в проекта."
        Text += TextTablo(dicObekt("GRT_Text"))
        AddParagraph(wordDoc, Text)
        AddParagraph(wordDoc, $"Главното разпределително табло {Кавички + dicObekt("GRT_Name") + Кавички } ще се заземи чрез свързване на заземителната му шина към комплектен заземител, състоящ се от два броя, свързани помежду си, поцинковани колове 63/63/6мм, всеки с дължина 1,5м.")
        Dim pDouOpts = New PromptDoubleOptions("")
        With pDouOpts
            .Keywords.Add("Да")
            .Keywords.Add("Не")
            .Keywords.Default = "Не"
            .Message = vbCrLf & "Жалаете ли да добавите още табла "
            .AllowZero = False
            .AllowNegative = False
        End With
        Dim boTabla As Boolean = False
        Do While True
            Dim pKeyRes = acDoc.Editor.GetDouble(pDouOpts)
            If pKeyRes.StringResult = "Не" Then Exit Do
            Dim СРТ_Име = cu.GetObjects_TEXT("Изберете текст съдържаш името на силовото разпределително табло")
            Dim СРТ_Кабел_Тип As String = cu.GetObjects_TEXT("Изберете текст съдържаш типа на кабела захранващ табло " & СРТ_Име)
            Dim СРТ_Кабел_Сечение As String = cu.GetObjects_TEXT("Изберете текст съдържаш сечението на кабела захранващ табло " & СРТ_Име)
            Dim СРТ_Текст As String = cu.GetObjects_TEXT("Изберете текст съдържаш описание на табло " & СРТ_Име, False).Replace(vbCrLf, "; ").Replace(vbCr, "; ").Replace(vbLf, "; ").ToLower()
            Dim СРТ_Помещение = cu.GetObjects_TEXT("Изберете текст съдържаш помещението в което се намира на табло " & СРТ_Име)
            Dim СРТ_Кота = cu.GetObjects_TEXT("Изберете текст съдържаш котата на която се намира на табло " & СРТ_Име)

            Text = TabloNa4alo(dicObekt("GRT_Name"),
                               СРТ_Помещение,
                               СРТ_Кота,
                               СРТ_Име,
                               СРТ_Кабел_Тип,
                               СРТ_Кабел_Сечение)
            Dim templates As String() = {"Таблото ще се изпълни съгласно приложената в проекта схема.",
                "Изпълнението на таблото ще бъде в съответствие с приложената в проекта схема.",
                "Таблото ще бъде изпълнено според приложената в проекта схема.",
                "Таблото ще бъде реализирано съгласно схемата, приложена в проекта.",
                "Таблото ще се изработи съгласно схемата, приложена в проекта.",
                "Таблото ще бъде изпълнено в съответствие с приложената в проекта схема."
            }
            ' Избор на случаен шаблон
            Dim random As New Random()
            Dim templateIndex As Integer = random.Next(0, templates.Length)
            Text += " " & templates(templateIndex)

            Text += TextTablo(СРТ_Текст)
            AddParagraph(wordDoc, Text)
            boTabla = True
        Loop
        Text = IIf(boTabla, "В електрическите табла", "В електрическото табло") + " ще бъде разположена предпазна апаратура за всички консуматори."
        Text += " Защитата на електрическите съоръжения и кабелите към тях ще се осъществява чрез автоматични прекъсвачи."
        Text += " Изборът на автоматичните прекъсвачи е съобразен с токовете на късо съединение, като са спазени изискванията за селективност."
        Text += " Изводите към контактите с общо предназначение и електрическите бойлери ще бъдат допълнително защитени с дефектнотокова защита."
        AddParagraph(wordDoc, Text)
        Text = IIf(boTabla, "Електрическите табла", "Електрическото табло")
        Text += " трябва да "
        Text += IIf(boTabla, "бъдат изпълнени", "бъдe изпълненo")
        Text += " в съответствие с изискванията на БДС EN 61439-1."
        Text += " Апаратурата и тоководещите части трябва да бъдат монтирани зад защитни капаци."
        Text += " Достъпът до палците и ръкохватките на комутационните апарати се осигурява посредством отвори в защитните капаци."
        Text += " При смяна на типа апаратура трябва да се преизчисли схемата."
        Text += " При смяна на номиналния ток на защитната апаратура трябва да се преизчисли сечението на кабелите."
        AddParagraph(wordDoc, Text)
    End Sub
    ' ЗАХРАНВАНЕ НА ОБЕКТА С ЕЛЕКТРИЧЕСКА ЕНЕРГИЯ
    Private Sub Външно_Захранване(wordDoc As Word.Document, dicObekt As Dictionary(Of String, String))

        Dim stackTrace As New StackTrace()
        Dim callingMethod As StackFrame = stackTrace.GetFrame(1)
        Dim methodName As String = callingMethod.GetMethod().Name
        Dim PV_Записка As Boolean = False
        If methodName = "PV_New_zapiska" Then
            PV_Записка = True
        Else
            PV_Записка = False
        End If

        Dim text As String = ""

        Dim ss_Kabeli = cu.GetObjects("LINE", "Изберете КАБЕЛИТЕ в чертеж за външно захранване на сградата:")
        Dim ss_Tabla = cu.GetObjects("INSERT", "Изберете БЛОКОВЕТЕ в чертеж за външно захранване на сградата:")
        If ss_Tabla Is Nothing Then
            MsgBox("Няма маркиран нито един блок.")
            Exit Sub
        End If
        If ss_Kabeli Is Nothing Then
            MsgBox("Няма маркиран нито едина линия.")
            Exit Sub
        End If

        Dim Kabel(50, 2) As String
        Kabel = cu.GET_LINE_TYPE_KABEL(Kabel, ss_Kabeli, vbFalse)
        Dim arrКабели As New List(Of srtCable)()

        Dim count_FTP As Integer = 0
        Dim count_коакс As Integer = 0

        For i = LBound(Kabel) To UBound(Kabel)
            If Kabel(i, 0) = "" Then Exit For
            If InStr(Kabel(i, 0), "поц.шина") > 0 Then Continue For
            If InStr(Kabel(i, 0), "ПВ-A2") > 0 Then Continue For
            If InStr(Kabel(i, 0), "AlMgSi") > 0 Then Continue For
            If InStr(Kabel(i, 0), "ELEKTRO") > 0 Then Continue For
            If InStr(Kabel(i, 0), "H1Z2Z2") > 0 Then Continue For
            If InStr(Kabel(i, 0), "Оптичен") > 0 Then Continue For

            If InStr(Kabel(i, 0), "FTP") > 0 Then
                count_FTP = 1
                Continue For
            End If

            If InStr(Kabel(i, 0), "коакс") > 0 Then
                count_коакс = 1
                Continue For
            End If

            ' Проверка дали вече съществува запис със същия тип и тръба в Kabeli
            Dim index As Integer = arrКабели.FindIndex(Function(k) k.Тип = Kabel(i, 0) AndAlso k.Тръба = Kabel(i, 1))

            If index = -1 Then
                ' Ако не съществува, създаваме нов запис
                Dim кабел As New srtCable
                кабел.Тип = Kabel(i, 0)
                кабел.Тръба = Kabel(i, 1)
                кабел.Дължина = Double.Parse(Kabel(i, 2)) / 100
                кабел.count = 1

                If (InStr(Kabel(i, 1), "PVC") > 0 Or (InStr(Kabel(i, 1), "HDPE") > 0) And InStr(Kabel(i, 0), "x10mm") > 0) Then
                    кабел.Тип = Kabel(i, 0).Replace("САВТ", "СВТ")
                Else
                    кабел.Тип = Kabel(i, 0)
                End If

                If кабел.Тип = "САВТ" Then
                    кабел.Материал = vbTrue
                End If
                arrКабели.Add(кабел)
            Else
                Dim кабел As New srtCable
                кабел = arrКабели(index)
                кабел.Дължина += Double.Parse(Kabel(i, 2)) / 100
                arrКабели(index) = кабел
            End If
        Next

        Dim count_HDPE As Integer = 0
        Dim count_PVC As Integer = 0

        Dim count_траншея_80 As Integer = 0
        Dim count_траншея_110 As Integer = 0
        Dim Delta_U As Double
        If Not PV_Записка Then
            For i As Integer = 0 To arrКабели.Count - 1
                Dim кабел As srtCable = arrКабели(i)
                кабел.Делта_U = cu.CalculateVoltageDropAC(кабел.Дължина, Val(dicObekt("GRT_Pприс")), кабел.Тип)
                arrКабели(i) = кабел
            Next
            For Each кабел In arrКабели
                Delta_U += кабел.Делта_U
            Next
        End If

        Dim Zazemitel(100) As strZazeml
        Zazemitel = cu.GET_Zazemlenie(ss_Tabla)
        Dim конзола As Integer = 0
        Dim Опъвач As Integer = 0
        Dim Захранващо_табло As String = ""
        Dim Електромерно_табло_монтаж As String = ""
        Dim Електромерно_табло_Име As String = ""
        Dim Съществуващо_табло As String = ""
        Dim Стълб_нов As String = ""
        Dim брой_пилони_същ As Integer = 0
        Dim брой_пилони_нов As Integer = 0
        Dim брой_стълбове_същ As Integer = 0
        Dim брой_стълбове_нов_835 As Integer = 0
        Dim брой_стълбове_нов_590 As Integer = 0
        Dim брой_стълбове_нов_250 As Integer = 0
        Dim брой_стълбове_нов As Integer = 0
        Dim брой_траншей As Integer = 0
        Dim брой_Репер As Integer = 0

        For Each contact In Zazemitel
            If contact.count = 0 Then Exit For
            Select Case contact.blVisibility
                Case "Надпокривна конзола"
                    конзола += contact.count
                Case "Опъвач_НЕрегулируем", "Опъвач_Регулируем"
                    Опъвач += contact.count
                Case "Сечение-траншея"
                Case "Сечение"
                Case "Репер"
                    брой_Репер += contact.count
                Case "Заземител-БЕЗ контролна клема"
                Case "Заземител-СЪС контролна клема"
                Case "Пилон-съществуващ СЪС Заземител"
                    брой_пилони_същ += contact.count
                Case "Пилон-съществуващ БЕЗ Заземител"
                    брой_пилони_същ += contact.count
                Case "Пилон-нов СЪС Заземител"
                    брой_пилони_нов += contact.count
                Case "Пилон-нов БЕЗ Заземител"
                    брой_пилони_нов += contact.count
                Case "Стълб-нов НЦ 835"
                    брой_стълбове_нов_835 += contact.count
                Case "Стълб-нов НЦ 590"
                    брой_стълбове_нов_590 += contact.count
                Case "Стълб-нов НЦ 250"
                    брой_стълбове_нов_250 += contact.count
                Case "Стълб-нов"
                    брой_стълбове_нов += contact.count
                Case "Стълб-съществуващ"
                    брой_стълбове_същ += contact.count
                Case "Табло ГРТ - СЪС Заземител"
                Case "Табло ГРТ - БЕЗ Заземител"
                Case "Табло същ. със заземител", "Табло същ. без заземител"
                    Захранващо_табло = "съществуващо"
                Case "Табло електромерно стоящо"
                    Захранващо_табло = "електромерно"
                    Електромерно_табло_монтаж = "на фундамент"
                    Електромерно_табло_Име = contact.blТАБЛО
                Case "Табло електромерно на СЪЩЕСТУВАЩ стълб"
                    Захранващо_табло = "електромерно"
                    Електромерно_табло_монтаж = "съществуващ стълб"
                    Електромерно_табло_Име = contact.blТАБЛО
                Case "Табло електромерно на НОВ стълб"
                    Захранващо_табло = "електромерно"
                    Електромерно_табло_монтаж = "новомонтиран стълб"
                    Електромерно_табло_Име = contact.blТАБЛО
                Case "Табло електромерно на НОВ пилoн"
                    Захранващо_табло = "електромерно"
                    Електромерно_табло_монтаж = "нов пилон с височина H=7,5m"
                    Електромерно_табло_Име = contact.blТАБЛО
                Case "Табло електромерно на СЪЩЕСТУВАЩ пилoн"
                    Захранващо_табло = "електромерно"
                    Електромерно_табло_монтаж = "съществуващ пилон"
                    Електромерно_табло_Име = contact.blТАБЛО
                Case Else
                    If contact.blName = "Траншея" Then
                        брой_траншей += contact.count
                    End If
            End Select
        Next contact

        Електромерно_табло_монтаж = IIf(Електромерно_табло_монтаж = "", "", " монтирано на " + Електромерно_табло_монтаж)
        If PV_Записка Then
            FormatParagraph(wordDoc, "ВРЪЗКАТА МЕЖДУ ИНВЕРТОРИТЕ И ТАБЛА НН", wordApp, level:=2)
        Else
            FormatParagraph(wordDoc, "ЗАХРАНВАНЕ НА ОБЕКТА С ЕЛЕКТРИЧЕСКА ЕНЕРГИЯ", wordApp)
            text = "Съгласно изискванията на Наредба № 3 от 9 юни 2004 г. за устройството на електрическите уредби и електропроводните линии,"
            text += " обектът е трета категория по отношение на непрекъснатост на електрозахранването и се захранва от един източник."
            text += " В проекта не е предвидено резервно захранване с електрическа енергия."
            AddParagraph(wordDoc, text, False)
        End If
        Dim кабели As String = IIf(arrКабели.Count = 1,
                           If(IsNumeric(arrКабели(0).Тип), "", $"кабел тип {arrКабели(0).Тип}"),
                           $"кабели тип {String.Join(", ",
                                                     arrКабели.Take(arrКабели.Count - 1).
                                                     Where(Function(c) Not IsNumeric(c.Тип)).
                                                     Select(Function(c) c.Тип)) &
                                         If(IsNumeric(arrКабели.Last().Тип), "", " и " & arrКабели.Last().Тип)}")

        Dim message As String = IIf(arrКабели.Count = 1, кабели, $"последователно свързани {кабели}")
        If PV_Записка Then
            text = $"Връзката между инверторите и табла НН трафопоста ще се осъществи с {кабели}."
        Else
            text = $"Обектът ще се захрани с електрическа енергия от {Захранващо_табло} табло {Кавички & Електромерно_табло_Име & Кавички} {Електромерно_табло_монтаж}."
            text += $" Връзката между {Захранващо_табло}то табло {Кавички & Електромерно_табло_Име & Кавички} и главното разпределително табло на обекта {Кавички & dicObekt("GRT_Name") & Кавички} ще се изпълни с {message}."

        End If
        AddParagraph(wordDoc, text, False)
        Dim indexAlR = arrКабели.FindIndex(Function(cable) cable.Тип.Contains("Al/R"))
        If indexAlR >= 0 Then
            Dim cable_AlR As String
            Dim конзола_text As String = ""
            If конзола > 0 Then
                конзола_text = "надпокривна конзола, монтирана на покрива"
            ElseIf Опъвач > 0 Then
                конзола_text = "кука с дюбел ф12, монтирана на фасадата"
            End If
            cable_AlR = arrКабели(indexAlR).Тип
            Dim indexblVis = Array.FindIndex(Zazemitel, Function(f) f.blVisibility = "Табло електромерно на стълб")
            text = $"Кабел тип {cable_AlR} ще бъде изтеглен въздушно, в имота на възложителя от {Електромерно_табло_монтаж} {Захранващо_табло} табло {Кавички & Електромерно_табло_Име & Кавички} до"
            text += $" {конзола_text} на сградата на указаното в проекта място."
            Dim Нов_пилон_text As String = " В трасето на въздушната линия ще"
            Select Case брой_пилони_нов
                Case = 0
                    Нов_пилон_text = ""
                Case = 1
                    Нов_пилон_text += " бъде монтиран един нов пилон с височина H=7,5m, място на монтаж е указано в проекта."
                    Нов_пилон_text += $" Новомонтираният пилон ще бъде заземен с един заземителен кол 63/63/6mm, с дължина 1,5m."
                Case > 1
                    Нов_пилон_text += $" бъдат монтирани {брой_пилони_нов} броя нови пилони с височина H=7,5m. В графичната част на проекта са указани местата, където ще се монтират пилоните."
                    Нов_пилон_text += $" Новомонтираните пилони ще бъдат заземени с по един заземителен кол 63/63/6mm, с дължина 1,5m."
            End Select
            text += $"{Нов_пилон_text} В двата края на въздушния участък ще се монтират опъвачи - нерегулируем при стълба и регулируем - от страната на сградата."
            AddParagraph(wordDoc, text, False)
            If (count_FTP + count_коакс) > 0 Then
                text = $"Не се допуска монтаж на слаботоковите кабели по кабел тип {cable_AlR}."
                text += " При невъзможност да се избере друго трасе слаботоковите кабели да се изтеглят на отделно носещо въже монтирано на разстояние 1.5м под силовия кабел."
                text += ""
                AddParagraph(wordDoc, text, False)
            End If
        End If
        'Dim indexHDPE = arrКабели.FindIndex(Function(cable) cable.Тръба.StartsWith("HDPE"))
        Dim indexHDPE = arrКабели.FindIndex(Function(cable) cable.Тръба.Contains("HDPE"))
        If indexHDPE > 0 Then
            text = ""
            'Dim match As Match = Regex.Match(Input, pattern)
            'text += $"{cable_HDPE}, изтеглен в тръба {Line.Linetype.ToUpper()}, положена в изкоп 80х50 см."
            AddParagraph(wordDoc, text, False)
            text = ""
            Select Case брой_траншей
                Case 0
                    text = "Захранващият кабел ще се положи в кабелна траншея."
                Case 1
                    text = "В графичната част на проекта е приложено сечениe на кабелната траншея."
                    text += " На него е показана дълбочината на полагане на тръбите по цялото протежение на изкопа."
                Case > 1
                    text = $"В графичната част на проекта са приложени {брой_траншей} броя сечения на кабелната траншея."
                    text += " На тях е показана дълбочината на полагане на тръбите по цялото протежение на изкопа."
            End Select

            If count_траншея_110 > 0 Then
                text += " При пресичане на пътища дълбочината на полагане ще е 1,10m, а в останалата част ще е 0,80m от кота терен."
            Else
                text += " Дълбочината на полагане ще е 0,80m от кота терен."
            End If
            text += " Не се допуска кабелите да се полагат по-близо от 0,6m от основите на сградите"
            IIf(PV_Записка, " и вертикалните елементи на конструкцията", "")
            text += "."
            AddParagraph(wordDoc, text, False)
            text = "Всички извивки на кабела трябва да бъдат с радиус не по-малък от 15 пъти външния диаметър на същия."
            text += " Полагането на тръбата ще стане чрез поставяне на под нея и отгоре й слой пясък или пресята пръст с дебелина не по-малка от 10cm."
            text += " Върху горния слой пясък се нанася добре уплътнена пръст до достигане на дълбочина 0,35m от кота терен."
            text += $" На тази дълбочина се поставя специална полиетиленова лента, сигнализираща за наличието на кабел под напрежение с надпис {Кавички}Внимание електрически кабел{Кавички}."
            text += " Ширината на лентата е 200 mm, а цветът е жълт."
            text += " Върху лентата отново се нанася пръст и се уплътнява до нивото на терена."
            text += " След засипване на изкопа да се възстановява настилката на терена в първоначалния ѝ вид, преди започването на изкопните работи."
            Select Case брой_Репер
                Case = 0
                    text += ""
                Case = 1
                    text += " На чупката на кабелния изкоп ще се постави реперно колче."
                Case > 1
                    text += $" На чупките на кабелния изкоп ще се поставят {брой_Репер} броя реперни колчета."
            End Select
            AddParagraph(wordDoc, text, False)
            If (count_HDPE + count_коакс) > 0 Then
                text = $"Изречение за полагане на слаботоковите тръби"
                AddParagraph(wordDoc, text, False)
            End If
            text = "При полагане на кабела в изкоп да се спазват минималните допустими вертикални и хоризонтални отстояния между кабели НН и други подземни комуникации съгласно изискванията на Наредба № 3 от 9 юни 2004 г. за устройството на електрическите уредби и електропроводните линии."
            AddParagraph(wordDoc, text, False)
        End If
        Dim indexPVC = arrКабели.FindIndex(Function(cable) cable.Тръба.StartsWith("PVC"))
        If indexPVC > 0 Then
            'text += " Кабел тип "
            'text += cab_SWT
            'text += " ще се изтегли в тръба "
            'text += cab_SWT_Tipe
            'text += ", положена скрито под мазилката."
            text += " Връзката между кабелите ще се осъществи чрез маншон за кабели, изолиран, биметален."
            AddParagraph(wordDoc, text, False)
        End If
        If Not PV_Записка Then
            text = $"За присъединената мощност Pприс.= {dicObekt("GRT_Pприс")}kW и {message} падът на напрежение ще бъде {String.Format("{0:0.000}", Delta_U)}%, което удовлетворява изискването за допустим пад на напрежение."
            AddParagraph(wordDoc, text, False)
        End If

    End Sub
    ' Заземителна
    Private Sub Заземителна(wordDoc As Word.Document,
                            dicObekt As Dictionary(Of String, String)
                            )

        Dim Text As String = ""
        FormatParagraph(wordDoc, "ЗАЗЕМИТЕЛНА ИНСТАЛАЦИЯ", wordApp)
        Dim stackTrace As New StackTrace()
        Dim callingMethod As StackFrame = stackTrace.GetFrame(1)
        Dim methodName As String = callingMethod.GetMethod().Name

        If methodName = "PV_New_zapiska" Then
            Text = "Инверторите, DC таблата и конструкцията на ФЕЦ"
        Else
            Text = "Главното разпределително табло " + Кавички + dicObekt("GRT_Name") + Кавички
        End If
        AddParagraph(wordDoc, Text + " ще се " +
                        IIf(methodName = "PV_New_zapiska", "заземят ", "заземи ") +
                        "чрез свързване на заземителната " +
                         IIf(methodName = "PV_New_zapiska", "им ", "му ") +
                        "шина към комплектен заземител, състоящ се от два броя, свързани помежду си, поцинковани колове 63/63/6mm, всеки с дължина 1,5m.")
        If methodName = "PV_New_zapiska" Then
            Text = "инверторите, DC таблата и конструкцията на ФЕЦ"
        Else
            Text = "главното разпределително табло " + Кавички + dicObekt("GRT_Name") + Кавички
        End If
        AddParagraph(wordDoc, "Връзката между " +
                            Text +
                            " и комплектния заземител ще се осъществи чрез проводник тип ПВ 1х16mm² изтеглен в PVC тръба ф16мм.")
        Text = "Преходното съпротивление на заземителя не трябва да надвишава " +
                IIf(methodName = "PV_New_zapiska", "10Ω ", "30Ω ") +
                "при най - неблагоприятните годишни условия, в противен случай да се увеличи броят на заземителните колове."
        Text += " При по-високо специфично съпротивление на почвата е необходимо увеличаване броя на заземителните колове."
        AddParagraph(wordDoc, Text, False)
        Text = "Горният край на вертикалните заземители е на 0,8 метра под повърхността на терена."
        Text += " Разстоянието между съседни вертикални заземители като правило не трябва да е по-малко от два пъти тяхната дължина."
        Text += " При констатирано измятане или огъване на заземителния кол при неговото набиване, същият да се замени с нов."
        AddParagraph(wordDoc, Text, False)
        Text = "Допуска се замяна на два броя, свързани помежду си, поцинковани колове 63/63/6mm, всеки с дължина 1,5m с един брой заземителен кол при условие, че долния край на кола е на дълбочина не по-малка от 3,0m от терена."
        AddParagraph(wordDoc, Text, False)
        Text = "Всички връзки на заземителните шини под повърхността на терена да бъдат изпълнени чрез заварка."
        Text += " Болтови съединения в земята не се допускат."
        Text += " Дължината на шева на заварката да бъде не по-малко от двойната широчина на заваряваните ленти и не по-малко от десет пъти диаметъра на проводника при кръгли сечения."
        Text += " След заварката мястото да се обработи цинк съдържаща боя и асфалт-лак за предпазване от корозия."
        Text += " По време на изпълнението на заварките стриктно да се спазват предписанията по отношение на безопасност при работа и постигане на максимално качество на заварките."
        Text += " Забранява се последователно свързване на няколко подлежащи на заземяване части, съоръжения и конструкции."
        AddParagraph(wordDoc, Text, False)
        Text = "На височина 1.2м над терена ще се монтират контролно-ревизионна кутия с разглобяема клема."
        Text += " Връзката между отводите и заземлението ще стане със сертифицирана контролна клема."
        Text += " Кутиите са разположени на подходящи за целта места, показани на чертежа."
        Text += " В тях се свързват токоотводите на мълниезащитата и заземителната инсталация."
        AddParagraph(wordDoc, Text, False)

    End Sub
    ' Заключение
    Private Sub Заключение(wordDoc As Word.Document, dicObekt As Dictionary(Of String, String))
        FormatParagraph(wordDoc, "ЗАКЛЮЧЕНИЕ", wordApp)
        AddParagraph(wordDoc, "При изпълнението на електро-монтажните работи да се спазят изискванията на действащите нормативни актове, както и на всички изменения и допълнения към тях, влезли в сила към момента на завършване на дейностите по изграждане на инсталациите предвидени в проекта.")
        AddParagraph(wordDoc, "Продуктите, които се влагат в предвидените в проекта електрически инсталации, трябва да съответстват на европейските технически спецификации при спазване изискванията на Регламент № 305/2011 на Европейския парламент и на Съвета за определяне на хармонизирани условия за предлагане на пазара на строителни продукти и за отмяна на Директива 89/106/ЕИО на Съвета и чл. 5, ал. 1 от НСИСОССП – придружен с маркировка СЕ и с прилагане на декларация за експлоатационните показатели на продукта и указания за прилагане, изготвени на български език.")
        AddParagraph(wordDoc, "Навсякъде в проекта, където са посочени изрично конкретни продукти с конкретни търговски марки следва да разбира и оферира не задължително посочените продукти, а равностойни, еквивалентни, със същите или по-добри параметри от посочените, като се спазват всички изисквания на действащите нормативни документи и съответстват на решението на проектанта и избраната технология!")
        AddParagraph(wordDoc, "При откриване на кабели в процеса на разкопаване, които не са показани в чертежа на ситуацията работата да бъде незабавно прекратена и да се търси съдействието на проектанта.")
        AddParagraph(wordDoc, "Електрическата инсталация да се въведе в експлоатация след издаване на Сертификат за контрол в т.ч. протоколи за:       ")
        Dim stackTrace As New StackTrace()
        Dim callingMethod As StackFrame = stackTrace.GetFrame(1)
        Dim methodName As String = callingMethod.GetMethod().Name

        Dim Мълния_блок As String = ""
        Dim Мълния_Коефициент As Double = 100

        If methodName = "PV_New_zapiska" Then
            Мълния_блок = "Мълниезащита вертикално_PV"
            Мълния_Коефициент = 1000
        Else
            Мълния_блок = "Мълниезащита вертикално"
            Мълния_Коефициент = 100
        End If

        Dim Paragraphs = New List(Of String) From {
            "Контрол на съпротивление на заземителна уредба;",
            "Контрол на съпротивление на мълниезащитна заземителна уредба;"
        }

        If methodName <> "PV_New_zapiska" Then
            Paragraphs.Add("Контрол на функционална годност на защитни прекъсвачи за токове с нулева последователност (RCD) в схеми TN и TT;")
            Paragraphs.Add("Контрол на импедансът Zs на контур 'фаза-защитен проводник;")
        End If
        Paragraphs.Add("Контрол изолационно съпротивление на захранващи кабели.")

        'Dim listTemplate As Word.ListTemplate
        Dim ListTemplate = wordDoc.ListTemplates.Add()
        ' Конфигуриране на нивото на списъка
        Dim listLevel As Word.ListLevel = ListTemplate.ListLevels(1)
        With ListTemplate.ListLevels(1)
            .NumberFormat = ChrW(&H2022).ToString() ' Символ за точка (•)
            .TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab ' Добавяне на табулация след точката
            .NumberStyle = Word.WdListNumberStyle.wdListNumberStyleBullet ' Стил на номериране - точка
            .NumberPosition = wordApp.InchesToPoints(1.25 / 2.54) ' Позиция на точката (1.5 см)
            .Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft ' Подравняване на нивото на списъка
            .TextPosition = wordApp.InchesToPoints(1.5 / 2.54) ' Позиция на текста след точката (1.75 см)
            .TabPosition = wordApp.InchesToPoints(1.5 / 2.54) ' Позиция на табулацията (1.75 см)
            .ResetOnHigher = 0 ' Нулиране на номерацията на по-високи нива
            .StartAt = 1 ' Начална позиция на списъка
            .LinkedStyle = "" ' Свързан стил (ако има такъв)
        End With

        ' Добавяне на всеки параграф като елемент от списъка с точки
        For i As Integer = 0 To Paragraphs.Count - 1
            ' Създаване на нов параграф в документа
            Dim para = wordDoc.Content.Paragraphs.Add()
            ' Задаване на текст за параграфа
            para.Range.Text = Paragraphs(i)
            ' Прилагане на шаблон за списък с точки към параграфа
            para.Range.ListFormat.ApplyListTemplateWithLevel(ListTemplate, ContinuePreviousList:=False)
            ' Добавяне на нов параграф след текущия
            para.Range.InsertParagraphAfter()
        Next
        wordDoc.Paragraphs(wordDoc.Paragraphs.Count).Range.ListFormat.RemoveNumbers()
        AddParagraph(wordDoc, "Към проекта е приложена специална обяснителна записка за мероприятията за безопасна работа и противопожарна защита.")
        AddПодпис(wordDoc, dicObekt("ПРОЕКТАНТ"))
    End Sub
    ' Събира данни и вмъква Слаботоковата инсталация
    Private Sub Слаботокова(wordDoc As Word.Document,
                             acDoc As Document,
                             acCurDb As Database)

        Dim Слабо_Име As String = "Сл.табло"
        Dim Слабо_Блок = cu.GetObjects("INSERT", "Изберете блок който съдържа името на слабтоковото табло ", False)
        If Слабо_Блок Is Nothing Then
            MsgBox("Няма маркиран нито един блок съдържащ името на слабтоковото табло.")
            Exit Sub
        End If
        Dim text As String = ""
        Try
            FormatParagraph(wordDoc, "СЛАБОТОКОВИ ИНСТАЛЦИИ", wordApp)
            Dim blkRecId As ObjectId
            Using trans As Transaction = acDoc.TransactionManager.StartTransaction()
                blkRecId = Слабо_Блок(0).ObjectId
                Dim acBlkTbl As BlockTable = trans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim acBlkRef As BlockReference =
                        DirectCast(trans.GetObject(blkRecId, OpenMode.ForRead), BlockReference)
                Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                For Each objID As ObjectId In attCol
                    Dim dbObj As DBObject = trans.GetObject(objID, OpenMode.ForRead)
                    Dim acAttRef As AttributeReference = dbObj
                    If acAttRef.Tag = "ТАБЛО" Then Слабо_Име = acAttRef.TextString
                Next
                trans.Commit()
            End Using
        Catch ex As Exception
            MsgBox("Възникна грешка " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
        End Try
        Dim Слабо_Помещение = cu.GetObjects_TEXT("Изберете текст съдържаш помещението в което се намира на табло " & Слабо_Име)
        Dim Слабо_Кота = cu.GetObjects_TEXT("Изберете текст съдържаш котата на която се намира на табло " & Слабо_Име)
        text = "Връзката с доставчика услугите ще се осъществи на границата на имота."
        text += " На място указано от доставчика ще се монтира допълнително табло доставено от доставчика на услугите."
        text += " От това табло ще се изтеглят по един брой кабели тип FTP 4х2х24AWG и RG 6/64 изтеглени в две тръби HDPE ф63mm."
        text += " Трасето по което ще се положат тръбите е указано в графичната част."
        text += " При необходимост и по указания на доставчика на услугите кабелите могат да се подменят, като се използват положените тръби."
        text += " Не се допуска изтеглянето на силови и слаботокови кабели в една тръба."
        AddParagraph(wordDoc, text)
        text = "В обекта се предвижда да се монтира слаботоково табло " + Кавички + Слабо_Име + Кавички
        text += IIf(Слабо_Помещение = "  #####  ", ".", ", монтирано в помещение " + Кавички + Слабо_Помещение + Кавички)
        text += IIf(Слабо_Кота = "  #####  ", ".", " на кота " + Слабо_Кота + ".")
        text += " Размерът и типът на слаботоковото табло, както и броя и типът на активното и пасивното оборудване поместени в него не са обект на настоящия проект."
        text += " Те следва да бъдат специфицирани от фирмата доставчик на услугите."
        text += " Приблизителните размери на слаботоково табло за вграден монтаж за два реда са ширина – 350mm; височина – 450mm и дълбочина – 100mm."
        text += " В таблото трябва да се монтират минимум два броя еднофазни контакта за захранване на активното оборудване."
        AddParagraph(wordDoc, text)
        AddParagraph(wordDoc, "На местата, където би могло да възникне нужда от достъп до интернет мрежата ще се монтира розетки тип RJ45. Всяка розетка тип RJ45 се свързва със слаботоковото табло чрез самостоятелна кабелна линия изпълнена с кабел тип FTP 4х2х24AWG.")
        'AddParagraph(wordDoc, "До всеки телевизор ще се монтира телевизионна розетка. Всяка телевизионна розетка се свързва със слаботоковото табло чрез самостоятелна кабелна линия изпълнена с кабел тип RG 6/64.")
        AddParagraph(wordDoc, "Слаботоковите розетки следва да бъдат инсталирани в общ стенен блок заедно с електрическите контакти.")
        AddParagraph(wordDoc, "Кабелите ще бъдат положени скрито в гофрирани тръби ф16мм по трасета, указани в графичната част към проекта.")
        AddParagraph(wordDoc, "При изпълнение на слаботоковите инсталации да се спазват изискванията на Наредба за правилата и нормите за проектиране, разполагане и демонтаж на електронни съобщителни мрежи.")
        text = " При полагане на кабелите за слаботоковите инсталации да се спазват изискванията за дистанциране на информационните кабели от електрическите."
        text += " Слаботоковите кабели може да преминават успоредно на отоплителна или газова инсталация на разстояние не по-малко от 0,30 m."
        text += " Слаботоковите кабели може да преминават успоредно на електрическата инсталация на разстояние не по-малко от 0,20 m."
        AddParagraph(wordDoc, text)
        AddParagraph(wordDoc, "Допуска се изтегляне на съобщителен и силов кабел в една и съща тръба или кабелен канал в следните случаи:")
        AddParagraph(wordDoc, "1. съобщителният кабел е оптичен или коаксиален кабел;", FirstLine:=50)
        AddParagraph(wordDoc, "2. съобщителният кабел е екраниран кабел;", FirstLine:=50)
        AddParagraph(wordDoc, "3. съобщителният кабел не е екраниран, но кабелният канал е с разделител.", FirstLine:=50)
    End Sub
    ' Събира данни и вмъква МЪЛНИЕЗАЩИТА
    Private Sub Мълниезащита(wordDoc As Word.Document,
                             acDoc As Document,
                             acCurDb As Database)
        Dim Мълния = cu.GetObjects("INSERT", "Изберете блок който съдържа нълниеприемника на обекта ", False)
        Dim text As String = ""
        If Мълния Is Nothing Then
            MsgBox("Няма маркиран нито един блок за мълниезащита.")
            FormatParagraph(wordDoc, "МЪЛНИЕЗАЩИТНА ИНСТАЛЦИЯ", wordApp)
            text = "Съгласно изискванията на Наредба № 4 от 22 декември 2010 г. за мълниезащита на сгради, външни съоръжения и открити пространства (Обн. ДВ, бр. 6 от 18 януари 2011 г.), обектът е класифициран в III категория на мълниезащита."
            text += " В съответствие с тази категоризация, мерките за защита от мълнии трябва да отговарят на нормативните изисквания за обекти от този тип."
            AddParagraph(wordDoc, text)
            text = "В настоящия проект не се предвижда изграждането на нова мълниезащитна инсталация, тъй като съществуващата мълниезащитна инсталация на обекта отговаря на всички изисквания на Наредба № 4."
            text += "Тя е проектирана в съответствие с действащата нормативна уредба и осигурява необходимото ниво на защита срещу пряко попадение на мълнии."
            text += " Милниезащитния радиус на системата е такъв, че покрива целия обект, осигурявайки защита на сградите и съоръженията."
            AddParagraph(wordDoc, text)
            text = "Изчисленията на съществуващата инсталация потвърждават нейната ефективност и съответствие с нормативните изисквания."
            text += " Изпълнението на допълнителни мерки не е необходимо, тъй като съществуващата инсталация покрива всички изисквания за мълниезащита за III категория на защита."
            AddParagraph(wordDoc, text)
            Exit Sub
        End If
        Dim Мълния_Категория As String = ""
        Dim Мълния_защита As String = ""
        Dim Мълния_Тип As String = ""
        Dim Мълния_H As String = ""
        Dim Мълния_Rz As String = ""
        Dim Hкъща As String = ""


        Dim stackTrace As New StackTrace()
        Dim callingMethod As StackFrame = stackTrace.GetFrame(1)
        Dim methodName As String = callingMethod.GetMethod().Name

        Dim Мълния_блок As String = ""
        Dim Мълния_Коефициент As Double = 100
        If methodName = "PV_New_zapiska" Then
            Мълния_блок = "Мълниезащита вертикално_PV"
            Мълния_Коефициент = 1000
        Else
            Мълния_блок = "Мълниезащита вертикално"
            Мълния_Коефициент = 100
        End If
        Try
            Using actrans As Transaction = acDoc.TransactionManager.StartTransaction()
                ' Започване на транзакция
                Dim acBlkTbl As BlockTable = actrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                ' Получаване на ID на записа на блока "Insert_Signature" в таблицата на блоковете
                Dim blkRecId As ObjectId = acBlkTbl(Мълния_блок)
                ' Получаване на записа на блока
                Dim acBlkTblRec As BlockTableRecord = actrans.GetObject(blkRecId, OpenMode.ForRead)
                ' Обхождане на всички блокови референции за блока "Insert_Signature"
                For Each blkRefId As ObjectId In acBlkTblRec.GetBlockReferenceIds(True, True)
                    ' Получаване на блоковата референция
                    Dim acBlkRef_ref As BlockReference = actrans.GetObject(blkRefId, OpenMode.ForWrite)
                    ' Получаване на колекцията от атрибути на блоковата референция
                    Dim attCol As AttributeCollection = acBlkRef_ref.AttributeCollection
                    ' Обхождане на всички атрибути
                    For Each objID As ObjectId In attCol
                    Next
                    Dim props_ref As DynamicBlockReferencePropertyCollection = acBlkRef_ref.DynamicBlockReferencePropertyCollection
                    For Each prop As DynamicBlockReferenceProperty In props_ref
                        'This is where you change states based on input
                        If prop.PropertyName = "Категория" Then Мълния_защита = prop.Value
                        If prop.PropertyName = "Тип" Then Мълния_Тип = prop.Value
                        If prop.PropertyName = "Hm" Then Мълния_H = String.Format("{0#,##0.00}", prop.Value / Мълния_Коефициент)
                        If prop.PropertyName = "Hкъща" Then Hкъща = String.Format("{0#,##0.00}", prop.Value / Мълния_Коефициент)
                        If prop.PropertyName = "Rb" Then Мълния_Rz = String.Format("{0#,##0}", prop.Value / Мълния_Коефициент)
                    Next
                Next

                blkRecId = Мълния(0).ObjectId
                Dim acBlkRef As BlockReference =
                    DirectCast(actrans.GetObject(blkRecId, OpenMode.ForWrite), BlockReference)
                Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection
                For Each prop As DynamicBlockReferenceProperty In props
                    'This is where you change states based on input
                    If prop.PropertyName = "Категория" Then Мълния_защита = prop.Value
                    If prop.PropertyName = "Тип" Then Мълния_Тип = prop.Value
                    If prop.PropertyName = "Hm" Then Мълния_H = String.Format("{0:#,##0.00}", prop.Value / Мълния_Коефициент)
                    If prop.PropertyName = "Hкъща" Then Hкъща = String.Format("{0:#,##0.00}", prop.Value / Мълния_Коефициент)
                    If prop.PropertyName = "Rb" Then Мълния_Rz = String.Format("{0:#,##0}", prop.Value / Мълния_Коефициент)
                Next
                actrans.Commit()
            End Using
        Catch ex As Exception
            MsgBox("Възникна грешка " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
        End Try
        ' Извличане на текста преди първото ','
        Dim римска_цифра As String = Мълния_защита.Split(",")(0)
        ' Определяне на категорията въз основа на получената стойност
        Select Case римска_цифра
            Case "I" : Мълния_Категория = "първа"
            Case "II" : Мълния_Категория = "втора"
            Case "III" : Мълния_Категория = "трета"
            Case "IV" : Мълния_Категория = "четвърта"
        End Select
        FormatParagraph(wordDoc, "МЪЛНИЕЗАЩИТНА ИНСТАЛЦИЯ", wordApp)
        text = "Съгласно изискванията на Наредба № 4 от 22 декември 2010 г. за мълниезащитата на сгради, външни съоръжения и открити пространства,"
        text += $" Обн. ДВ., бр. 6 от 18 януари 2011 г. обекта е {Мълния_Категория} категория на мълниезащита."
        AddParagraph(wordDoc, text)
        text = $"Мълниезащитата на обекта ще се изпълни с мълниеприемник с изпреварващо действие PREVECTRON®3, Millenium модел {Мълния_Тип} или подобен."
        text += $" Мълниеприемникът ще бъде монтиран на покрива на сградата, върху носеща мачта с височина {Мълния_H}m над защитаваната повърхност."
        text += $" Мълниезащитната мачта ще се монтира на височина {Hкъща}m от кота терен."
        text += $" При тези параметри на мълниеприемника ще се реализира защитен радиус Rз={Мълния_Rz}m, при което ще се защитят всички части на обекта."
        text += $" За присъединяване на мълниеприемника към мълниезащитния прът да се използва детайл по спецификация на производителя."
        AddParagraph(wordDoc, text)
        text = $"Допуска се замяната на мълниеприемника с модел на друг производител, ако са изпълнени следните условия"
        text += $" радиусът на защита за ниво на защита {Мълния_защита} при h(m) = {Мълния_H}m е не по-малко от Rз = {Мълния_Rz}m."
        AddParagraph(wordDoc, text)
        text = "Мълниезащитните отводи ще се изпълнят от екструдиран проводник тип AlMgSi ф8мм."
        text += " При полагане на отводите минималния радиус на огъване да не е по-малък от R 200мм."
        text += " Токоотвода да се постави на вертикална противопожарна ивица с ширина 0,50m, с клас по реакция на огън А2."
        AddParagraph(wordDoc, text)
        text = "Заземителят ще се изпълни от два броя, свързани помежду си, поцинковани колове 63/63/6mm, всеки с дължина 1,5m."
        text += " Преходното съпротивление на заземителя не трябва да надвишава 10Ω, в противен случай да се увеличи броят на заземителните колове."
        text += " При изпълнение на заземителя да се спазват изискванията, описани по-горе."
        text = "Допуска се замяна на два броя, свързани помежду си, поцинковани колове 63/63/6mm, всеки с дължина 1,5m с един брой заземителен кол при условие, че долния край на кола е на дълбочина не по-малка от 3,0m от терена."
        AddParagraph(wordDoc, text)
    End Sub
    ' ЕЛЕКТРИЧЕСКА ОСВЕТИТЕЛНА.
    Private Sub ОСВЕТИТEЛЕНА(wordDoc As Word.Document,
                             acDoc As Document,
                             acCurDb As Database)

        Dim Text As String
        FormatParagraph(wordDoc, "ОСВЕТИТEЛЕНА ИНСТАЛАЦИЯ", wordApp)
        Text = "Електрическата инсталация за захранване на осветителните тела ще се изпълни с кабели, тип СВТ3х1,5mm², изтеглени в PVC тръби."
        AddParagraph(wordDoc, Text)
        Text = "Осветителните тела са съобразени с характера на помещенията."
        Text += " Броят, видът и мястото на монтиране осигуряват необходимата осветеност и зрителен комфорт."
        Text += " Във всички помещения с нормална опасност за поражение от електрически ток ще се монтират осветителни тела със степен на защита IP-21."
        Text += " В помещенията с повишена опасност за поражение от електрически ток степента на защита на осветителните тела ще бъде IP–54."
        AddParagraph(wordDoc, Text)
        Text = "Включването на осветлението в помещенията ще се извършва ръчно с ключове или автоматично чрез датчици за движение."
        AddParagraph(wordDoc, Text)
        Text = "Местоположението на осветителните тела и ключовете в помещенията, е посочена в графичната част към проекта, съобразено с архитектурното обзавеждане."
        AddParagraph(wordDoc, Text)

        Dim pDouOpts = New PromptDoubleOptions("")
        With pDouOpts
            .Keywords.Add("Да")
            .Keywords.Add("Не")
            .Keywords.Default = "Не"
            .Message = vbCrLf & "Жалаете ли да добавите текст за Евакуационното осветление?"
            .AllowZero = False
            .AllowNegative = False
        End With
        Dim boTabla As Boolean = False
        Dim pKeyRes = acDoc.Editor.GetDouble(pDouOpts)
        If pKeyRes.StringResult = "Да" Then
            AddParagraph(wordDoc, "Евакуационното осветление", True)
            Text = "Проектът за евакуационно осветление отговаря на изискванията на чл. 55 от Наредба № Iз-1971 от 29 октомври 2009 г. за строително-технически правила и норми за осигуряване на безопасност при пожар."
            Text += $" Той съответства на раздел II {Кавички + "Захранване на аварийно и евакуационно осветление" + Кавички} от глава 40 {Кавички + "Вътрешни и външни осветителни уредби" + Кавички} на Наредба № 3 от 9 юни 2004 г. за устройството на електрическите уредби и електропроводните линии."
            AddParagraph(wordDoc, Text)
            Text = "Осветеността на евакуационния път по осевата линия на пода ще бъде не по-малка от 1 lx."
            Text += " За осветление на пътищата за евакуация са предвидени светодиодни осветителни тела с мощност 4W."
            Text += " Те ще бъдат с вградени акумулаторни батерии, осигуряващи минимална продължителност на работа не по-малка от един час."
            Text += $" Върху тях ще има надпис {Кавички + "ИЗХОД" + Кавички} или стрелка, указваща посоката на пътя на евакуацията."
            Text += $" Над всички врати по пътищата за евакуация са предвидени осветителни тела надпис {Кавички + "ИЗХОД" + Кавички}."
            Text += " Местоположението на осветителните тела е посочена в графичната част към проекта."
            AddParagraph(wordDoc, Text)
            Text = "Захранването на инсталацията за евакуационно осветление ще се осъществи от автоматичен прекъсвач с номинален ток 10А."
            Text += "За изпълнението й ще се използват кабели тип СВТ3х1,5mm², изтеглени в PVC тръби."
            AddParagraph(wordDoc, Text)
        End If
    End Sub
    ' ЕЛЕКТРИЧЕСКА ИНСТАЛАЦИЯ КОНТАКТИ.
    Private Sub КОНТАКТИ(wordDoc As Word.Document)
        Dim Text As String
        FormatParagraph(wordDoc, "ЕЛЕКТРИЧЕСКА ИНСТАЛАЦИЯ КОНТАКТИ", wordApp)
        Text = "Електрическата инсталация за захранване на контактите и ел. бойлера ще се изпълни с кабели тип СВТ3х2,5mm², изтеглени в PVC тръби."
        Text += "Кухненската печка ще се захрани с кабел тип СВТ3х4,0mm², изтеглен в PVC тръба."

        AddParagraph(wordDoc, Text)
        Text = $"Всички еднофазни контакти ще бъдат тип {Кавички}Шуко{Кавички}."
        Text += " Местоположението и височината на монтаж е посочена в графичната част към проекта, съобразено с архитектурното обзавеждане."
        AddParagraph(wordDoc, Text)
        Text = "Степента на защита на контактите и разклонителните кутии в помещенията с нормална опасност за поражение от електрически ток среда ще бъде IP-32."
        Text += " В помещенията с повишена опасност за поражение от електрически ток ще бъде IP–54."
        AddParagraph(wordDoc, Text)
    End Sub
    ' СЪДЪРЖАНИЕ
    Private Sub СЪДЪРЖАНИЕ(wordDoc As Word.Document, dicObekt As Dictionary(Of String, String))
        With wordDoc.Content.Paragraphs.Add
            .Range.Text = "СЪДЪРЖАНИЕ"
            .Range.Font.Name = "Cambria"
            .Range.Font.Size = 20
            .Range.Font.Bold = True
            .Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            .Format.SpaceAfter = 24
            .Range.InsertParagraphAfter()
        End With
        Dim paragraphs As New List(Of String) From {
            "Челен лист",
            "Съдържание",
            "Копие на Удостоверение за проектантска правоспособност",
            $"Копие на Застрахователна полица {Кавички}Професионална отговорност в проектирането и строителството{Кавички}",
            dicObekt("Text_SAP"),
            "Обяснителна записка"}
        Dim stackTrace As New StackTrace()
        Dim callingMethod As StackFrame = stackTrace.GetFrame(1)
        Dim methodName As String = callingMethod.GetMethod().Name
        Dim PV_Записка As Boolean = False
        If methodName = "PV_New_zapiska" Then
            paragraphs.AddRange({"Обяснителна записка пожароизвестителна система и система за звукова сигнализация"})
        Else
            paragraphs.AddRange({"Обяснителна записка по БХТПБ"})
        End If
        paragraphs.AddRange({
            "Обяснителна записка по техника на безопасност по време на строителството",
            "Количествена сметка",
            "Графична част"
        })
        ' Добавяне на параграфите към документа
        For i As Integer = 0 To paragraphs.Count - 1
            AddParagraph(wordDoc, paragraphs(i))
        Next
        ' Създаване на нов шаблон за списък
        Dim listTemplate As Word.ListTemplate
        listTemplate = wordDoc.ListTemplates.Add() ' Създава нов шаблон за списък с нива
        ' Настройка на шаблона за списък
        With listTemplate.ListLevels(1)
            .NumberFormat = "%1."                                           ' Задава формат на номера
            .TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab     ' Задава символ след номера
            .NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic   ' Задава стил на номера
            .Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft     ' Задава подравняване на номера
            .ResetOnHigherOld = 0                                           ' Задава рестартиране на номера при по-високо ниво
            .StartAt = 1                                                    ' Задава начален номер
            .LinkedStyle = ""                                               ' Задава свързан стил
        End With
        ' Превръщане на параграфите в номериран списък
        Dim start As Integer = wordDoc.Paragraphs.Count - paragraphs.Count
        Dim end_ As Integer = wordDoc.Paragraphs.Count
        For i As Integer = start To end_
            Dim para As Word.Paragraph = wordDoc.Paragraphs(i)
            para.Range.ListFormat.ApplyListTemplateWithLevel(
                ListTemplate:=listTemplate,                                             ' Прилага предварително дефинирания шаблон за списък
                ContinuePreviousList:=True,                                             ' Продължава предишния списък
                ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList,                     ' Прилага се към целия списък
                DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)   ' Използва се поведението на списъка от Word 2010
        Next
        wordDoc.Paragraphs(wordDoc.Paragraphs.Count).Range.ListFormat.RemoveNumbers()
        wordDoc.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak)
    End Sub
    ' Челен лист
    Private Sub Челен_Лист(wordDoc As Word.Document, Obekt As Dictionary(Of String, String))
        Try
            ' Вмъкване на таблицата след последния параграф
            Dim Table = wordDoc.Tables.Add(wordDoc.Paragraphs.Last.Range, 2, 4)
            ' Форматиране на таблицата
            With Table
                .Borders.Enable = False
                .Rows.Alignment = Word.WdRowAlignment.wdAlignRowRight
                With .Columns(1)
                    ' Задаване на ширината на първата колона на 4.5 см и центриране на текста
                    .Width = wordDoc.Application.CentimetersToPoints(4.5)
                    .Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop
                End With
                With .Columns(2)
                    ' Задаване на ширината на втората колона на 13.5 см и подравняване на текста вляво
                    .Width = wordDoc.Application.CentimetersToPoints(5.5)
                    .Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop
                End With
                With .Columns(3)
                    ' Задаване на ширината на втората колона на 13.5 см и подравняване на текста вляво
                    .Width = wordDoc.Application.CentimetersToPoints(3.5)
                    .Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop
                End With
                With .Columns(4)
                    ' Задаване на ширината на втората колона на 13.5 см и подравняване на текста вляво
                    .Width = wordDoc.Application.CentimetersToPoints(4.5)
                    .Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop
                End With
            End With
            ' Добавяне на нов ред
            Dim newRow = Table.Rows(1)
            Dim stackTrace As New StackTrace()
            Dim callingMethod As StackFrame = stackTrace.GetFrame(1)
            Dim methodName As String = callingMethod.GetMethod().Name
            Dim PV_Записка As Boolean = False
            If methodName = "PV_New_zapiska" Or
                methodName = "ПОЖАРНА" Then
                SetCellFormat(newRow.Cells(1), "ПРОЕКТНО РЕШЕНИЕ", "Cambria", True, 24, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter, Height:=1.5, border:=False)
            Else
                SetCellFormat(newRow.Cells(1), "ИНВЕСТИЦИОНЕН ПРОЕКТ", "Cambria", True, 24, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter, Height:=1.5, border:=False)
            End If
            newRow = Table.Rows(2)
            SetCellFormat(newRow.Cells(1), "ВЪЗЛОЖИТЕЛ:", "Cambria", True, 12, Word.WdParagraphAlignment.wdAlignParagraphRight, Word.WdCellVerticalAlignment.wdCellAlignVerticalTop, Height:=1.0, border:=False)
            SetCellFormat(newRow.Cells(2), Obekt("ВЪЗЛОЖИТЕЛ"), "Cambria", False, 12, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdCellVerticalAlignment.wdCellAlignVerticalTop, Height:=1.0, border:=False)
            newRow = Table.Rows.Add
            SetCellFormat(newRow.Cells(1), "ПОДПИС:", "Cambria", True, 12, Word.WdParagraphAlignment.wdAlignParagraphRight, Word.WdCellVerticalAlignment.wdCellAlignVerticalTop)
            newRow = Table.Rows.Add
            newRow = Table.Rows.Add
            newRow = Table.Rows.Add
            SetCellFormat(newRow.Cells(1), "ОБЕКТ:", "Cambria", True, 12, Word.WdParagraphAlignment.wdAlignParagraphRight, Word.WdCellVerticalAlignment.wdCellAlignVerticalTop, Height:=1.5, border:=False)
            SetCellFormat(newRow.Cells(2), Obekt("ОБЕКТ"), "Cambria", False, 12, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdCellVerticalAlignment.wdCellAlignVerticalTop, Height:=1.5, border:=False)
            newRow = Table.Rows.Add
            If Obekt("ОБЕКТ") <> Obekt("МЕСТОПОЛОЖЕНИЕ") Then
                SetCellFormat(newRow.Cells(1), "МЕСТОПОЛОЖЕНИЕ:", "Cambria", True, 12, Word.WdParagraphAlignment.wdAlignParagraphRight, Word.WdCellVerticalAlignment.wdCellAlignVerticalTop, Height:=1.5, border:=False)
                SetCellFormat(newRow.Cells(2), Obekt("МЕСТОПОЛОЖЕНИЕ"), "Cambria", False, 12, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdCellVerticalAlignment.wdCellAlignVerticalTop, Height:=1.5, border:=False)
            End If
            newRow = Table.Rows.Add
            SetCellFormat(newRow.Cells(1), "ЧАСТ:", "Cambria", True, 12, Word.WdParagraphAlignment.wdAlignParagraphRight, Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom, Height:=1.5, border:=False)
            If methodName = "ПОЖАРНА" Then
                SetCellFormat(newRow.Cells(2), "ПОЖАРНА БЕЗОПАСНОСТ", "Cambria", False, 12, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom, Height:=1.5, border:=False)
            Else
                SetCellFormat(newRow.Cells(2), "ЕЛЕКТРО", "Cambria", False, 12, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom, Height:=1.5, border:=False)
            End If
            newRow = Table.Rows.Add
            SetCellFormat(newRow.Cells(1), "ФАЗА:", "Cambria", True, 12, Word.WdParagraphAlignment.wdAlignParagraphRight, Word.WdCellVerticalAlignment.wdCellAlignVerticalTop)
            SetCellFormat(newRow.Cells(2), Obekt("ФАЗА"), "Cambria", False, 12, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdCellVerticalAlignment.wdCellAlignVerticalTop)
            newRow = Table.Rows.Add
            SetCellFormat(newRow.Cells(1), "ДАТА:", "Cambria", True, 12, Word.WdParagraphAlignment.wdAlignParagraphRight, Word.WdCellVerticalAlignment.wdCellAlignVerticalTop, Height:=1.5, border:=False)
            SetCellFormat(newRow.Cells(2), Obekt("ДАТА"), "Cambria", False, 12, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdCellVerticalAlignment.wdCellAlignVerticalTop, Height:=1.5, border:=False)
            newRow = Table.Rows.Add
            SetCellFormat(newRow.Cells(1), "ЧАСТ:", "Cambria", True, 12, Word.WdParagraphAlignment.wdAlignParagraphRight, Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom)
            SetCellFormat(newRow.Cells(2), "СЪГЛАСУВАЛИ:", "Cambria", True, 12, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom)

            If Not Obekt("АРХИТЕКТ").Contains("##") Then
                newRow = Table.Rows.Add
                SetCellFormat(newRow.Cells(1), "АРХИТЕКТУРА",
                              "Cambria", True, 12, Word.WdParagraphAlignment.wdAlignParagraphRight, Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom)
                SetCellFormat(newRow.Cells(2), Obekt("АРХИТЕКТ").Replace("&", " ").Trim(),
                              "Cambria", False, 12, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom)

            End If
            If Not Obekt("КОНСТРУКТОР").Contains("##") Then
                newRow = Table.Rows.Add
                SetCellFormat(newRow.Cells(1), "КОНСТРУКТОР",
                              "Cambria", True, 12, Word.WdParagraphAlignment.wdAlignParagraphRight, Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom)
                SetCellFormat(newRow.Cells(2), Obekt("КОНСТРУКТОР").Replace("&", " ").Trim(),
                              "Cambria", False, 12, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom)
                SetCellFormat(newRow.Cells(3), " ",
                              "Cambria", False, 12, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom, Height:=1, border:=True)
            End If

            If Not Obekt("ТЕХНОЛОГИЯ").Contains("##") Then
                newRow = Table.Rows.Add
                SetCellFormat(newRow.Cells(1), "ТЕХНОЛОГИЯ",
                              "Cambria", True, 12, Word.WdParagraphAlignment.wdAlignParagraphRight, Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom)
                SetCellFormat(newRow.Cells(2), Obekt("ТЕХНОЛОГИЯ").Replace("&", " ").Trim(),
                              "Cambria", False, 12, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom)
                SetCellFormat(newRow.Cells(3), " ",
              "Cambria", False, 12, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom, Height:=1, border:=True)
            End If

            If Not Obekt("ВИК").Contains("##") Then
                newRow = Table.Rows.Add
                SetCellFormat(newRow.Cells(1), "ВиК",
                              "Cambria", True, 12, Word.WdParagraphAlignment.wdAlignParagraphRight, Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom)
                SetCellFormat(newRow.Cells(2), Obekt("ВИК").Replace("&", " ").Trim(),
                              "Cambria", False, 12, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom)
                SetCellFormat(newRow.Cells(3), " ",
              "Cambria", False, 12, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom, Height:=1, border:=True)
            End If
            If Not Obekt("ОВ").Contains("##") Then
                newRow = Table.Rows.Add
                SetCellFormat(newRow.Cells(1), "ОВ",
                              "Cambria", True, 12, Word.WdParagraphAlignment.wdAlignParagraphRight, Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom)
                SetCellFormat(newRow.Cells(2), Obekt("ОВ").Replace("&", " ").Trim(),
                              "Cambria", False, 12, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom)
                SetCellFormat(newRow.Cells(3), " ",
              "Cambria", False, 12, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom, Height:=1, border:=True)
            End If
            If Not Obekt("ГЕОДЕЗИЯ").Contains("##") Then
                newRow = Table.Rows.Add
                SetCellFormat(newRow.Cells(1), "ГЕОДЕЗИЯ",
                              "Cambria", True, 12, Word.WdParagraphAlignment.wdAlignParagraphRight, Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom)
                SetCellFormat(newRow.Cells(2), Obekt("ГЕОДЕЗИЯ").Replace("&", " ").Trim(),
                              "Cambria", False, 12, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom)
                SetCellFormat(newRow.Cells(3), " ",
              "Cambria", False, 12, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom, Height:=1, border:=True)
            End If
            If Not Obekt("ВП").Contains("##") Then
                newRow = Table.Rows.Add
                SetCellFormat(newRow.Cells(1), "ВП",
                              "Cambria", True, 12, Word.WdParagraphAlignment.wdAlignParagraphRight, Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom)
                SetCellFormat(newRow.Cells(2), Obekt("ВП").Replace("&", " ").Trim(),
                              "Cambria", False, 12, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom)
                SetCellFormat(newRow.Cells(3), " ",
              "Cambria", False, 12, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom, Height:=1, border:=True)
            End If
            If Not Obekt("ЕЕФ").Contains("##") Then
                newRow = Table.Rows.Add
                SetCellFormat(newRow.Cells(1), "ЕЕФ",
                              "Cambria", True, 12, Word.WdParagraphAlignment.wdAlignParagraphRight, Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom)
                SetCellFormat(newRow.Cells(2), Obekt("ЕЕФ").Replace("&", " ").Trim(),
                              "Cambria", False, 12, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom)
                SetCellFormat(newRow.Cells(3), " ",
              "Cambria", False, 12, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom, Height:=1, border:=True)
            End If
            If Not Obekt("ПБ").Contains("##") Then
                newRow = Table.Rows.Add
                SetCellFormat(newRow.Cells(1), "ПБ",
                              "Cambria", True, 12, Word.WdParagraphAlignment.wdAlignParagraphRight, Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom)
                SetCellFormat(newRow.Cells(2), Obekt("ПБ").Replace("&", " ").Trim(),
                              "Cambria", False, 12, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom)
                SetCellFormat(newRow.Cells(3), " ",
              "Cambria", False, 12, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom, Height:=1, border:=True)
            End If

            If Not Obekt("ПБЗ").Contains("##") Then
                newRow = Table.Rows.Add
                SetCellFormat(newRow.Cells(1), "ПБЗ",
                              "Cambria", True, 12, Word.WdParagraphAlignment.wdAlignParagraphRight, Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom)
                SetCellFormat(newRow.Cells(2), Obekt("ПБЗ").Replace("&", " ").Trim(),
                              "Cambria", False, 12, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom)
                SetCellFormat(newRow.Cells(3), " ",
              "Cambria", False, 12, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom, Height:=1, border:=True)
            End If
            If Not Obekt("ПУСО").Contains("##") Then
                newRow = Table.Rows.Add
                SetCellFormat(newRow.Cells(1), "ПУСО",
                              "Cambria", True, 12, Word.WdParagraphAlignment.wdAlignParagraphRight, Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom)
                SetCellFormat(newRow.Cells(2), Obekt("ПУСО").Replace("&", " ").Trim(),
                              "Cambria", False, 12, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom)
                SetCellFormat(newRow.Cells(3), " ",
              "Cambria", False, 12, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom, Height:=1, border:=True)
            End If

            newRow = Table.Rows.Add
            newRow.Cells(2).Merge(newRow.Cells(4))
            SetCellFormat(newRow.Cells(2), "ПРОЕКТАНТ: ______________________________", "Cambria", False, 12,
                          Word.WdParagraphAlignment.wdAlignParagraphRight, Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom, Height:=5, border:=False)
            newRow = Table.Rows.Add
            SetCellFormat(newRow.Cells(2), "/" + Obekt("ПРОЕКТАНТ") + "/", "Cambria", False, 12, Word.WdParagraphAlignment.wdAlignParagraphRight, Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom)
            Table.Cell(1, 1).Merge(Table.Cell(1, 4))
            Table.Cell(2, 2).Merge(Table.Cell(2, 4))
            Table.Cell(6, 2).Merge(Table.Cell(6, 4))
            Table.Cell(7, 2).Merge(Table.Cell(7, 4))
            With Table.Cell(1, 1)
                .Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            End With
            With Table.Cell(6, 2)
                .Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
            End With
            With Table.Cell(7, 2)
                .Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
            End With
            With Table.Cell(2, 2)
                .Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
            End With
            With Table.Cell(3, 2)
                With .Borders(Word.WdBorderType.wdBorderBottom)
                    .LineStyle = Word.WdLineStyle.wdLineStyleSingle
                    .LineWidth = Word.WdLineWidth.wdLineWidth025pt
                    .Color = Word.WdColor.wdColorBlack
                End With
            End With
        Catch ex As Exception
            MsgBox("Error: " & ex.Message)
        End Try
        wordDoc.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak)
    End Sub
    'ПРИСЪЕДИНЯВАНЕ НА ОБЕКТА КЪМ ЕЛЕКТРОРАЗПРЕДЕЛИТЕЛНАТА МРЕЖА.
    Private Sub ПРИСЪЕДИНЯВАНЕ(wordDoc As Word.Document,
                               acDoc As Document,
                               acCurDb As Database,
                               dicObekt As Dictionary(Of String, String)
                               )
        Dim text As String
        FormatParagraph(wordDoc, "ПРИСЪЕДИНЯВАНЕ НА ОБЕКТА КЪМ ЕЛЕКТРОРАЗПРЕДЕЛИТЕЛНАТА МРЕЖА", wordApp)
        If Not dicObekt("SAP").Contains("##") Then
            Dim Stylb = cu.GetObjects("INSERT", "Изберете блок който съдържа текстовете за стълба", False)
            If Stylb Is Nothing Then
                MsgBox("Няма маркиран нито един блок.")
                Exit Sub
            End If
            Dim rezult As String = ""
            Try
                Dim blkRecId As ObjectId
                Using trans As Transaction = acDoc.TransactionManager.StartTransaction()
                    blkRecId = Stylb(0).ObjectId
                    Dim acBlkTbl As BlockTable = trans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                    Dim acBlkRef As BlockReference =
                            DirectCast(trans.GetObject(blkRecId, OpenMode.ForRead), BlockReference)

                    Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                    For br = 0 To 10
                        Dim ddd As String = "NA4IN_" + br.ToString
                        For Each objID As ObjectId In attCol
                            Dim dbObj As DBObject = trans.GetObject(objID, OpenMode.ForRead)
                            Dim acAttRef As AttributeReference = dbObj
                            If acAttRef.Tag = ddd AndAlso acAttRef.TextString <> "" Then
                                rezult += acAttRef.TextString + ", "
                            End If
                        Next
                    Next
                    trans.Commit()
                End Using
            Catch ex As Exception
                MsgBox("Възникна грешка " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
            End Try
            text = $"Съгласно становище SAP № {dicObekt("SAP")} присъединяването на обекта към електроразпределителната мрежа на "
            text += $" {dicObekt("Дружество")} ще се извърши към {rezult.Substring(0, rezult.Length - 2)}."
            text += $" Съгласно същото становище предоставената мощност е {dicObekt("GRT_Pпред")}."
            text += $" След реализиране на проекта присъединената мощност ще бъде Pприс.={dicObekt("GRT_Pприс")}kW, при което е реализиран {dicObekt("GRT_Ke")}."
            AddParagraph(wordDoc, text)
            AddParagraph(wordDoc, dicObekt("ТОЧКА_3"))
            AddParagraph(wordDoc, "Присъединяването на обекта към електроразпределителната мрежа ще се извърши след съгласуване на настоящия проект и сключване на окончателен договор за присъединяване.")
        Else
            text = "Съгласно УДОСТОВЕРЕНИЕ от "
            text += dicObekt("Дружество")
            text += " изх. № "
            text += dicObekt("Ном.заявление")
            text += " / "
            text += dicObekt("Дата_заявление")
            text += ", че "
            text += dicObekt("ИМЕ")
            text += " e клиент на "
            text += dicObekt("Дружество")
            text += ", в качеството си на страна по договор за продажба на електрическа енергия за обект с адрес: "
            text += dicObekt("АДРЕС")
            text += " с клиентски номер: "
            text += dicObekt("ПАРТИДА")
            text += ", поземления имот е захранен чрез електрическо отклонение до съществуващ стълб, на който е монтирано електромерното табло за имота. "
            AddParagraph(wordDoc, text)
            text = $" Предоставената мощност е {dicObekt("GRT_Pпред")}."
            text += $" След реализиране на проекта присъединената мощност ще бъде Pприс.= {dicObekt("GRT_Pприс")} kW, при което е реализиран {dicObekt("GRT_Ke")}."
            AddParagraph(wordDoc, text)
        End If
    End Sub
    ' Начало на ОБЯСНИТЕЛНА ЗАПИСКА
    Private Sub ОБЯСНИТЕЛНА_ЗАПИСКА(wordDoc As Word.Document, dicObekt As Dictionary(Of String, String))
        With wordDoc.Content.Paragraphs.Add
            .Range.Text = "ОБЯСНИТЕЛНА ЗАПИСКА"
            .Range.Font.Name = "Cambria"
            .Range.Font.Size = 20
            .Range.Font.Bold = True
            .Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            .Range.InsertParagraphAfter()
        End With
        Dim text As String = " "
        AddParagraph(wordDoc, text, False)
        text = $"ОБЕКТ: {dicObekt("ОБЕКТ")}"
        AddParagraph(wordDoc, text, True)
        If dicObekt("ОБЕКТ") <> dicObekt("МЕСТОПОЛОЖЕНИЕ") Then
            text = $"МЕСТОПОЛОЖЕНИЕ: {dicObekt("МЕСТОПОЛОЖЕНИЕ")}"
            AddParagraph(wordDoc, text, True)
        End If
        AddParagraph(wordDoc, " ", True)

        Dim Части As String = "     "
        Части += IIf(dicObekt("АРХИТЕКТ") = "  #####  ", "", "Архитектурна, ")
        Части += IIf(dicObekt("КОНСТРУКТОР") = "  #####  ", "", "Конструктивна, ")
        Части += IIf(dicObekt("ТЕХНОЛОГИЯ") = "  #####  ", "", "Технологична, ")
        Части += IIf(dicObekt("ВИК") = "  #####  ", "", "ВИК, ")
        Части += IIf(dicObekt("ОВ") = "  #####  ", "", "ОВ, ")
        Части += IIf(dicObekt("ГЕОДЕЗИЯ") = "  #####  ", "", "Геодезия, ")
        Части += IIf(dicObekt("ВП") = "  #####  ", "", "ВП, ")
        Части += IIf(dicObekt("ЕЕФ") = "  #####  ", "", "ЕЕФ, ")
        Части += IIf(dicObekt("ПБ") = "  #####  ", "", "ПБ, ")
        Части += IIf(dicObekt("ПБЗ") = "  #####  ", "", "ПБЗ, ")
        Части += IIf(dicObekt("ПУСО") = "  #####  ", "", "ПУСО, ")
        FormatParagraph(wordDoc, "ОБЩА ЧАСТ", wordApp)
        Части = Части.Substring(0, Части.Length - 2)
        text = "Настоящият проект се разработи по искане на Възложителя "
        text += dicObekt("ВЪЗЛОЖИТЕЛ")
        text += " на основание разработени проекти по части "
        text += Части
        text += " и на основание "
        text += dicObekt("Text_SAP_дълъг")

        AddParagraph(wordDoc, text)
        AddParagraph(wordDoc, "При разработване на проекта са спазени изискванията на :")
        AddParagraph(wordDoc, "1. Наредба № 3 от 9 юни 2004 г. за устройството на електрическите уредби и електропроводните линии, Обн. ДВ., бр. 90 от 13 октомври 2004 г. и бр. 91 от 14 октомври 2004 г.")
        AddParagraph(wordDoc, "2. Наредба № 16-116 от 8 февруари 2008 г. за техническата експлоатация на енергийните съоръжения, Обн. ДВ., бр. 26 от 7 март 2008 г.")
        AddParagraph(wordDoc, "3. Наредба № 1 от 27 май 2010 г. за проектиране, изграждане и поддържане на електрически уредби НН в сгради, Обн. ДВ., бр. 46 от 18 юни 2010 г.")
        AddParagraph(wordDoc, "4. Наредба № Iз-1971 от 29 октомври 2009 г. за строително-технически правила и норми за осигуряване на безопасност при пожар, Обн. ДВ., бр. 96 от 4 декември 2009 г.")
        AddParagraph(wordDoc, "5. Наредба № 4 от 22 декември 2010 г. за мълниезащитата на сгради, външни съоръжения и открити пространства, Обн. ДВ., бр. 6 от 18 януари 2011 г.")

    End Sub
    Private Function TextTablo(text As String) As String
        Dim rezult As String = ""
        ' Раздели текста на отделни компоненти
        Dim components As String() = text.Split(New String() {"; "}, StringSplitOptions.RemoveEmptyEntries)
        ' Декларирай променливи за различните части на текста

        Dim шкаф As String = ""
        Dim монтаж As String = ""
        Dim модули As String = ""
        Dim врата As String = ""
        Dim защита As String = ""
        Dim ширина As String = ""
        Dim височина As String = ""
        Dim дълбочина As String = ""
        Dim размери As String = ""

        ' Обработи всяка част от текста и я присвой към съответната променлива
        For Each component As String In components
            If component.Contains("брой модули:") Then модули = component.Replace("брой модули:", "").Trim()
            If component.Contains("врата:") Then врата = component.Replace("врата:", "").Trim()
            If component.Contains("степен на защита:") Then защита = component.Replace("степен на защита:", "").Trim().ToUpper()
            If component.Contains("монтаж") Then монтаж = component
            If component.Contains("шкаф") Then шкаф = component
            If component.Contains("ш:") Then ширина = component.Replace("ш:", "").Trim().ToUpper()
            If component.Contains("в:") Then височина = component.Replace("в:", "").Trim().ToUpper()
            If component.Contains("д:") Then дълбочина = component.Replace("д:", "").Trim().ToUpper().Replace(";", "")
        Next
        If Not монтаж.Contains("монтаж") Then
            If шкаф.Contains("стоящ") Then
                монтаж = "стоящ монтаж"
            Else
                монтаж = "вграден монтаж"
            End If
        End If
        If Not String.IsNullOrEmpty(ширина) AndAlso Not String.IsNullOrEmpty(височина) AndAlso Not String.IsNullOrEmpty(дълбочина) Then
            размери = $"с размери - ширина: {ширина}mm, височина: {височина}mm и дълбочина: {дълбочина}mm"
        Else
            If модули.Contains("(") Then
                Dim parts() As String = модули.Split(" (")
                размери = $"състоящо се от {parts(0)} броя модули, разположени на {parts(1).Replace(")", "").Replace("(", "")} реда"
            Else
                размери = $"състоящо се от {модули} броя модули"
            End If
        End If
        ' Различни шаблони за резултата
        Dim templates As String() = {
        $"То ще бъде {шкаф} за {монтаж}, {размери}, {врата} врата и степен на защита {защита}.",
        $"Изпълнението на таблото ще бъде {шкаф} за {монтаж}, {размери}, {врата} врата и степен на защита {защита}.",
        $"Това табло ще бъде изпълнено като {шкаф} за {монтаж}, {размери}, {врата} врата и ще има степен на защита {защита}.",
        $"Предвидено е таблото да бъде {шкаф} за {монтаж}, {размери}, {врата} врата със степен на защита {защита}.",
        $"Таблото да бъде {шкаф} за {монтаж}, {размери}, {врата} врата със степен на защита {защита}."
    }
        ' Избор на случаен шаблон
        Dim random As New Random()
        Dim templateIndex As Integer = random.Next(0, templates.Length)
        rezult = " " & templates(templateIndex)
        Return rezult
    End Function
    Private Function TabloNa4alo(Име_ГРТ As String,
                                 Помещение As String,
                                 Кота As String,
                                 Име As String,
                                 Кабел_тип As String,
                                 Кабел_сечение As String) As String
        ' Деклариране на променлива rezult, която ще съхранява крайния резултат
        Dim rezult As String = ""
        ' Обграждане на входните параметри с кавички
        Име_ГРТ = Кавички + Име_ГРТ + Кавички
        Помещение = Кавички + Помещение + Кавички
        Име = Кавички + Име + Кавички
        ' Проверка на стойността на параметъра Кота
        If Кота = "  #####  " Then
            Кота = ""  ' Ако Кота е с дадена стойност, се задава празен низ
        Else
            Кота = "на кота" + Кота  ' В противен случай, се добавя текста "на кота" пред стойността на Кота
        End If
        ' Дефиниране на масив от шаблони за първата част на резултата
        Dim templates_First As String() = {
        $"В проекта е предвидено в помещение {Помещение} {Кота} да се монтира силово разпределително табло {Име}.",
        $"В помещение {Помещение} {Кота} ще бъде инсталирано силово разпределително табло {Име}.",
        $"Проектът предвижда монтаж на силово разпределително табло {Име} в помещение {Помещение} {Кота}.",
        $"Според проекта, в помещение {Помещение} {Кота} ще се монтира силово разпределително табло {Име}.",
        $"Съгласно проекта, в помещение {Помещение} {Кота} е предвидено да се монтира силово разпределително табло {Име}.",
        $"Монтажът на силово разпределително табло {Име} ще се извърши в помещение {Помещение} {Кота}.",
        $"В помещение {Помещение} {Кота} е предвидено за монтаж на силово разпределително табло {Име}.",
        $"В проекта се предвижда монтаж на силово разпределително табло {Име} в помещение {Помещение} {Кота}."
    }
        ' Инициализация на обект от тип Random за случайно избиране на шаблон
        Dim random As New Random()
        ' Избиране на случаен индекс от масива templates_First
        Dim templateIndex As Integer = random.Next(0, templates_First.Length)
        ' Присвояване на избрания шаблон към променливата rezult
        rezult = templates_First(templateIndex)
        ' Дефиниране на масив от шаблони за втората част на резултата
        Dim templates_Second As String() = {
        $"Таблото ще се захрани от главното разпределително табло {Име_ГРТ}, чрез кабел тип {Кабел_тип}{Кабел_сечение}mm².",
        $"Захранването на таблото ще бъде от главното разпределително табло {Име_ГРТ}, използвайки кабел тип {Кабел_тип}{Кабел_сечение}mm².",
        $"Таблото ще бъде захранено от главното разпределително табло {Име_ГРТ} посредством кабел тип {Кабел_тип}{Кабел_сечение}mm².",
        $"Чрез кабел тип {Кабел_тип}{Кабел_сечение}mm² ще осигури връзка между таблото и главното разпределително табло {Име_ГРТ}.",
        $"Таблото ще бъде захранвано чрез кабел тип {Кабел_тип}{Кабел_сечение}mm² от главното разпределително табло {Име_ГРТ}."
    }
        ' Избиране на случаен индекс от масива templates_Second
        templateIndex = random.Next(0, templates_Second.Length)
        ' Добавяне на избрания шаблон към променливата rezult
        rezult += " " & templates_Second(templateIndex)
        ' Връщане на крайния резултат
        Return rezult
    End Function
    Public Function addArrayBlock() As Dictionary(Of String, String)
        Dim zapis As New Dictionary(Of String, String)
        ' Получаване на активния документ
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        ' Получаване на базата данни на активния документ
        Dim acCurDb As Database = acDoc.Database
        ' Започване на транзакция
        Using actrans As Transaction = acDoc.TransactionManager.StartTransaction()
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
                    zapis.Add(acAttRef.Tag, acAttRef.TextString)
                Next
            Next
        End Using
        Dim stackTrace As New StackTrace()
        Dim callingMethod As StackFrame = stackTrace.GetFrame(1)
        Dim methodName As String = callingMethod.GetMethod().Name

        If methodName <> "PV_New_zapiska" Then
            zapis.Add("GRT_Name", cu.GetObjects_TEXT("Изберете текст съдържаш името на главното разпределително табло"))
            zapis.Add("GRT_Pпред", cu.GetObjects_TEXT("Изберете текст съдържаш ПРЕДОСТАВЕНАТА мощност"))
            zapis.Add("GRT_Ke", cu.GetObjects_TEXT("Изберете текст съдържаш Ке"))
            zapis.Add("GRT_Pприс", cu.GetObjects_TEXT("Изберете текст съдържаш ПРИСЪЕДИНЕНАТА мощност"))
            zapis.Add("GRT_Text", cu.GetObjects_TEXT("Изберете текст съдържаш описание на табло " & zapis("GRT_Name"), False).Replace(vbCrLf, "; ").Replace(vbCr, "; ").Replace(vbLf, "; ").ToLower())
            zapis.Add("GRT_Помещение", cu.GetObjects_TEXT("Изберете текст съдържаш помещението в което се намира на табло " & zapis("GRT_Name")))
            zapis.Add("GRT_Кота", cu.GetObjects_TEXT("Изберете текст съдържаш котата на която се намира на табло " & zapis("GRT_Name")))
        End If
        Dim Text_SAP As String = "Копие на"
        Dim Text_SAP_дълъг As String = ""

        If Not zapis("SAP").Contains("##") Then
            Text_SAP += " Становище № " + zapis("SAP") + " за условията и начина за присъединяване на клиенти към електрическата мрежа"
            Text_SAP_дълъг = "Становище "
            Text_SAP_дълъг += "SAP № "
            Text_SAP_дълъг += zapis("SAP")
            Text_SAP_дълъг += " за условията и начина за присъединяване на клиенти към електрическата мрежа, номер на заявление "
            Text_SAP_дълъг += zapis("Ном.заявление")
            Text_SAP_дълъг += ", Дата на издаване:     "
            Text_SAP_дълъг += zapis("Дата_заявление")
            Text_SAP_дълъг += " издадено от "
            Text_SAP_дълъг += zapis("Дружество")
            Text_SAP_дълъг += ". Копие от становището се прилага към проекта."
        Else
            Text_SAP = "УДОСТОВЕРЕНИЕ от "
            Text_SAP += zapis("Дружество")
            Text_SAP += " изх. № "
            Text_SAP += zapis("Ном.заявление")
            Text_SAP += " / "
            Text_SAP += zapis("Дата_заявление")

            Text_SAP_дълъг = "УДОСТОВЕРЕНИЕ от "
            Text_SAP_дълъг += zapis("Дружество")
            Text_SAP_дълъг += " изх. № "
            Text_SAP_дълъг += zapis("Ном.заявление")
            Text_SAP_дълъг += " / "
            Text_SAP_дълъг += zapis("Дата_заявление")
            Text_SAP_дълъг += ", че "
            Text_SAP_дълъг += zapis("ИМЕ")
            Text_SAP_дълъг += " e клиент на "
            Text_SAP_дълъг += zapis("Дружество")
            Text_SAP_дълъг += ", в качеството си на страна по договор за продажба на електрическа енергия за обект с адрес: "
            Text_SAP_дълъг += zapis("АДРЕС")
            Text_SAP_дълъг += " с клиентски номер: "
            Text_SAP_дълъг += zapis("ПАРТИДА")
            Text_SAP_дълъг += ". Копие от удостоверението се прилага към проекта."
        End If
        zapis.Add("Text_SAP", Text_SAP)
        zapis.Add("Text_SAP_дълъг", Text_SAP_дълъг)
        Return zapis
    End Function
    Private Function addArrayDSD(File_DST As String,
                                 wordDoc As Word.Document
                                 ) As Dictionary(Of String, String)
        Dim zapis As New Dictionary(Of String, String)
        ' Получаване на препратка към обекта Sheet Set Manager
        Dim sheetSetManager As IAcSmSheetSetMgr = New AcSmSheetSetMgr
        ' Отваряне на файл за набор от листове
        Dim sheetSetDatabase As AcSmDatabase = sheetSetManager.OpenDatabase(File_DST, False)
        Dim sheetSet As AcSmSheetSet = sheetSetDatabase.GetSheetSet()
        ' Set the name and description of the sheet set
        If ShSet.LockDatabase(sheetSetDatabase, True) = False Then
            ' Display error message
            MsgBox("Sheet set не може да бъде отворен за четене.")
            Return zapis
        End If

        ' Получаване на обектите в набора от листове
        Dim enumerator As IAcSmEnumPersist = sheetSetDatabase.GetEnumerator()
        Dim item As IAcSmPersist
        Dim itemSheet As Object
        ' Деклариране на списък
        Dim arrSheet As New List(Of srtSheetSet)()
        ' Вземане на първия елемент от изброимия обект
        item = enumerator.Next()
        ' Изпълнение на цикъл докато има елементи в изброимия обект
        Do While Not item Is Nothing
            Dim sheet As IAcSmSheet = Nothing
            ' Проверка дали типът на текущия елемент е "AcSmSubset"
            If item.GetTypeName() = "AcSmSubset" Then
                ' Преобразуване на елемента в IAcSmSubset
                Dim subset As IAcSmSubset = item
                ' Получаване на изброим обект за листовете в текущия поднабор
                Dim enumSheets = subset.GetSheetEnumerator()
                ' Вземане на първия лист в текущия поднабор
                itemSheet = enumSheets.Next()
                ' Цикъл докато има листове в текущия поднабор
                Do While Not itemSheet Is Nothing
                    ' Преобразуване на текущия елемент в IAcSmSheet
                    sheet = itemSheet
                    ' Добавяне на информация за текущия лист в списъка
                    arrSheet.Add(New srtSheetSet() With {
                         .NameSubset = subset.GetName(),                                            ' Вземане на името на поднабора
                         .NameSheet = sheet.GetName().Substring(sheet.GetName().IndexOf(" ") + 1),  ' Вземане на името на листа след първото интервално разстояние
                         .NumberSheet = sheet.GetNumber()})                                         ' Вземане на номера на листа
                    ' Вземане на следващия лист в поднабора
                    itemSheet = enumSheets.Next()
                Loop
            End If
            ' Вземане на следващия елемент в основния набор от листове
            item = enumerator.Next()
        Loop
        ' Сортиране на списъка по номера на листовете (след преобразуване в число) и записването му обратно в arrSheet
        arrSheet = arrSheet.OrderBy(Function(x) Integer.Parse(x.NumberSheet)).ToList()

        ' Създаване на таблица с два реда
        Dim currentPos As Integer = wordDoc.Content.End - 1
        Dim table As Word.Table = wordDoc.Tables.Add(wordDoc.Range(currentPos, currentPos), 1, 2)
        Dim subsetNames As New HashSet(Of String)()

        ' Задаване на заглавията на колоните
        SetCellFormat(table.Cell(1, 1), "Номер на чертеж", "Cambria", False, 12, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter)
        SetCellFormat(table.Cell(1, 2), "Съдържание на чертежа", "Cambria", False, 12, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter)

        ' Сортиране на таблицата по първата колона
        With table
            .Borders.Enable = False
            .Rows.Alignment = Word.WdRowAlignment.wdAlignRowRight
            With .Columns(1)
                ' Задаване на ширината на първата колона на 2.25 см и центриране на текста
                .Width = wordApp.CentimetersToPoints(2.25)
                .Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop
            End With
            With .Columns(2)
                ' Задаване на ширината на втората колона на 12.5 см и подравняване на текста вляво
                .Width = wordApp.CentimetersToPoints(12.5)
                .Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop
            End With
        End With

        ' Обхождане на всеки елемент в списъка arrSheet
        For Each indexSet As srtSheetSet In arrSheet
            ' Добавяне на нов ред в таблицата
            Dim row As Word.Row = table.Rows.Add()
            ' Проверка дали името на поднабора е ново (не е добавено до момента)
            If subsetNames.Add(indexSet.NameSubset) Then
                If row.Cells.Count > 1 Then    ' Ако името на поднабора е ново и редът има повече от една клетка
                    ' Създаване на нов ред преди обединяването на клетките
                    Dim newRow As Word.Row = table.Rows.Add()
                    ' Задаване на името на поднабора в първата клетка на текущия ред и обединяване на клетките
                    With row.Cells(1)
                        .Merge(row.Cells(2))
                        .Range.Text = indexSet.NameSubset
                        SetCellFormat(row.Cells(1), indexSet.NameSubset, "Cambria", True, 12, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter)
                    End With
                    ' Задаване на номера на чертежа в първата клетка на новия ред
                    SetCellFormat(newRow.Cells(1), indexSet.NumberSheet, "Cambria", False, 12, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter)
                    ' Задаване на името на чертежа във втората клетка на новия ред
                    SetCellFormat(newRow.Cells(2), indexSet.NameSheet, "Cambria", False, 12, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter)
                End If
            Else                ' Ако името на поднабора вече е добавено
                ' Задаване на номера на чертежа в първата клетка на новия ред
                SetCellFormat(row.Cells(1), indexSet.NumberSheet, "Cambria", False, 12, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter)
                ' Задаване на името на чертежа във втората клетка на новия ред
                SetCellFormat(row.Cells(2), indexSet.NameSheet, "Cambria", False, 12, Word.WdParagraphAlignment.wdAlignParagraphLeft, Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter)
            End If
        Next

        'Dim Obekt As String = IIf(String.IsNullOrEmpty(GetCustomProperty(SheetSet, "Обект")) OrElse GetCustomProperty(SheetSet, "Обект") = "#####", "", GetCustomProperty(SheetSet, "Обект"))
        'Dim Място As String = IIf(String.IsNullOrEmpty(GetCustomProperty(SheetSet, "Местоположение")) OrElse GetCustomProperty(SheetSet, "Местоположение") = "#####", "", GetCustomProperty(SheetSet, "Местоположение"))
        'Части += IIf(String.IsNullOrEmpty(GetCustomProperty(SheetSet, "Архитект")) OrElse GetCustomProperty(SheetSet, "Архитект") = "#####", "", Кавички & "Архитектурна" & Кавички & ", ")
        'Части += IIf(String.IsNullOrEmpty(GetCustomProperty(sheetSet, "Конструктор")) OrElse GetCustomProperty(sheetSet, "Конструктор") = "#####", "", Кавички & "Конструктивна" & Кавички & ", ")
        'Части += IIf(String.IsNullOrEmpty(GetCustomProperty(sheetSet, "Технология")) OrElse GetCustomProperty(sheetSet, "Технология") = "#####", "", Кавички & "Технологична" & Кавички & ", ")
        'Части += IIf(String.IsNullOrEmpty(GetCustomProperty(sheetSet, "ВиК")) OrElse GetCustomProperty(sheetSet, "ВиК") = "#####", "", Кавички & "ВиК" & Кавички & ", ")
        'Части += IIf(String.IsNullOrEmpty(GetCustomProperty(sheetSet, "ОВ")) OrElse GetCustomProperty(sheetSet, "ОВ") = "#####", "", Кавички & "ОВ" & Кавички & ", ")
        'Части += IIf(String.IsNullOrEmpty(GetCustomProperty(sheetSet, "Геодезия")) OrElse GetCustomProperty(sheetSet, "Геодезия") = "#####", "", Кавички & "Геодезия" & Кавички & ", ")
        'Части += IIf(String.IsNullOrEmpty(GetCustomProperty(sheetSet, "ВП")) OrElse GetCustomProperty(sheetSet, "ВП") = "#####", "", Кавички & "ВП" & Кавички & ", ")
        'Части += IIf(String.IsNullOrEmpty(GetCustomProperty(sheetSet, "ЕЕФ")) OrElse GetCustomProperty(sheetSet, "ЕЕФ") = "#####", "", Кавички & "ЕЕФ" & Кавички & ", ")
        'Части += IIf(String.IsNullOrEmpty(GetCustomProperty(sheetSet, "ПБ")) OrElse GetCustomProperty(sheetSet, "ПБ") = "#####", "", Кавички & "ПБ" & Кавички & ", ")
        'Части += IIf(String.IsNullOrEmpty(GetCustomProperty(sheetSet, "ПБЗ")) OrElse GetCustomProperty(sheetSet, "ПБЗ") = "#####", "", Кавички & "ПБЗ" & Кавички & ", ")
        'Части += IIf(String.IsNullOrEmpty(GetCustomProperty(sheetSet, "ПБЗ")) OrElse GetCustomProperty(sheetSet, "ПБЗ") = "#####", "", Кавички & "ПБЗ" & Кавички & ", ")
        'Части += IIf(String.IsNullOrEmpty(GetCustomProperty(sheetSet, "ПУСО")) OrElse GetCustomProperty(sheetSet, "ПУСО") = "#####", "", Кавички & "ПУСО" & Кавички & ", ")
        'Proektant = IIf(String.IsNullOrEmpty(GetCustomProperty(SheetSet, "Проектант")) OrElse GetCustomProperty(SheetSet, "Проектант") = "#####", "", "/" & GetCustomProperty(SheetSet, "Проектант") & "/")
        'Text += IIf(String.IsNullOrEmpty(GetCustomProperty(SheetSet, "Възложител")) OrElse GetCustomProperty(SheetSet, "Възложител") = "#####", "", GetCustomProperty(SheetSet, "Възложител"))


        Dim aaa = "ОБЕКТ
МЕСТОПОЛОЖЕНИЕ
ВЪЗЛОЖИТЕЛ
СОБСТВЕНИК
ФАЗА
ДАТА
АРХИТЕКТ
КОНСТРУКТОР
ТЕХНОЛОГИЯ
ВИК
ОВ
ГЕОДЕЗИЯ
ВП
ЕЕФ
ПБ
ПБЗ
ПУСО
ПРОЕКТАНТ
Ном.заявление
Дата_заявление
SAP
Дружество
брой_листове
"
        ShSet.LockDatabase(sheetSetDatabase, False)
        Return zapis
    End Function
    Private Sub AddParagraph1(ByVal wordDoc As Word.Document, ByVal text As String, Optional Bold As Boolean = False, Optional FirstLine As Double = 35.4375)
        With wordDoc.Content.Paragraphs.Add
            .Range.Text = text
            .Range.Font.Name = "Cambria"
            .Range.Font.Size = 12
            .Range.Font.Bold = Bold
            .Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
            .Format.SpaceBefore = 0 ' Разредка преди: 0 пкт
            .Format.SpaceAfter = 6 ' Разредка след: 6 пкт
            .Format.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle ' Редова разрядка: единична
            .Range.ParagraphFormat.LeftIndent = 0
            .Range.ParagraphFormat.RightIndent = 0
            .Range.ParagraphFormat.FirstLineIndent = FirstLine
            .Range.InsertParagraphAfter()
        End With
    End Sub
    ' Форматира пафаграфа за римска цифра и текст
    Sub FormatParagraph(wordDoc As Word.Document,               ' Документът, в който се добавя и форматира параграфът.
                        text As String,                         ' Текстът, който ще бъде добавен и форматиран в новия параграф.
                        wordApp As Word.Application,            ' Обектът на приложението Word, използван за достъп до функции на Word.
                        Optional level As Integer = 1,          ' Нивото на номериране (от 1 до 9). По подразбиране е 1.
                        Optional resetLevel As Boolean = False) ' Булев параметър, който указва дали номерирането на първото ниво трябва да се ресетира от 1.
        Try
            ' Създаване на нов параграф в документа
            Dim para As Word.Paragraph
            para = wordDoc.Content.Paragraphs.Add()  ' Добавяне на нов параграф в края на документа
            With para
                .Range.Text = text  ' Задаване на текста на параграфа
                .Range.Font.Name = "Cambria"  ' Задаване на шрифта
                .Range.Font.Size = 12  ' Задаване на размера на шрифта
                .Range.Font.Bold = True  ' Насечен текст
                .Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify  ' Подравняване на текста
                .Format.SpaceBefore = 6  ' Разстояние преди параграфа
                .Format.SpaceAfter = 6  ' Разстояние след параграфа
                .Format.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle  ' Единично междуредие
                .Range.ParagraphFormat.LeftIndent = 0  ' Ляв отстъп на параграфа
                .Range.ParagraphFormat.RightIndent = 0  ' Десен отстъп на параграфа
                .Range.ParagraphFormat.FirstLineIndent = 35.4375  ' Отстъп на първия ред
            End With

            ' Намиране на съществуващ шаблон за номериране с римски и арабски цифри или създаване на нов, ако не съществува
            Dim listTemplate As Word.ListTemplate
            listTemplate = Nothing
            For Each lt As Word.ListTemplate In wordDoc.ListTemplates
                If lt.Name = "RomanAndArabicNumbering" Then
                    listTemplate = lt
                    Exit For
                End If
            Next

            If listTemplate Is Nothing Then
                listTemplate = wordDoc.ListTemplates.Add(True, "RomanAndArabicNumbering")
                ' Първо ниво с римски цифри
                With listTemplate.ListLevels(1)
                    .NumberFormat = "%1."  ' Формат на римски цифри с точка след тях
                    .TrailingCharacter = Word.WdTrailingCharacter.wdTrailingSpace
                    .NumberStyle = Word.WdListNumberStyle.wdListNumberStyleUppercaseRoman  ' Стил за римски цифри
                    .NumberPosition = wordApp.InchesToPoints(0.5)  ' Позиция на номера
                    .Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft
                    .TextPosition = wordApp.InchesToPoints(0.75)
                    .TabPosition = wordApp.InchesToPoints(0.75)
                    .ResetOnHigher = 0
                    .LinkedStyle = ""
                End With

                ' Нива от 2 до 9 с арабски цифри, показващи предходните нива
                For i = 2 To 9
                    With listTemplate.ListLevels(i)
                        .NumberFormat = "%1.%2" & If(i > 2, String.Concat(Enumerable.Range(3, i - 2).Select(Function(j) ".%" & j)), "") & "."  ' Формат за показване на предходни нива
                        .TrailingCharacter = Word.WdTrailingCharacter.wdTrailingSpace
                        .NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic  ' Стил за арабски цифри
                        .NumberPosition = wordApp.InchesToPoints(0.5 + (i - 1) * 0.25)
                        .Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft
                        .TextPosition = wordApp.InchesToPoints(0.75 + (i - 1) * 0.25)
                        .TabPosition = wordApp.InchesToPoints(0.75 + (i - 1) * 0.25)
                        .ResetOnHigher = i - 1
                        .StartAt = 1
                        .LinkedStyle = ""
                    End With
                Next
            End If
            ' Принудително рестартиране на първото ниво, ако resetLevel е True
            If resetLevel Then
                para.Range.ListFormat.ListLevelNumber = 1
                para.Range.ListFormat.ApplyListTemplateWithLevel(
                        ListTemplate:=listTemplate,
                        ContinuePreviousList:=False,
                        ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList,
                        DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)
                para.Range.ListFormat.ListTemplate.ListLevels(1).StartAt = 1
            End If
            ' Прилагане на шаблон за номериране с римски и арабски цифри и продължаване на номерирането
            para.Range.ListFormat.ApplyListTemplate(listTemplate, ContinuePreviousList:=True)
            ' Задаване на конкретно ниво на списъка въз основа на параметъра (1-во до 9-то ниво)
            para.Range.ListFormat.ListLevelNumber = level
            ' Принудително обновяване на списъка и форматирането
            ' Приложи шаблона отново, за да се увериш, че форматирането е правилно зададено
            para.Range.ListFormat.ApplyListTemplateWithLevel(
                ListTemplate:=listTemplate,
                ContinuePreviousList:=True,
                ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList,
                DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)
            ' Добавяне на нов параграф след приложената номерация
            para.Range.InsertParagraphAfter()

        Catch ex As Exception
            MsgBox("Error opening or activating Word document: " & ex.Message)
        End Try
    End Sub
    ''' <summary>
    ''' Добавя нов, напълно стандартизиран и форматиран абзац в Word документ.
    '''
    ''' Основна цел на процедурата:
    ''' ------------------------------------------
    ''' Да осигури ЕДИНЕН стил на всички текстови блокове в проекта,
    ''' без да разчитаме на шаблони, стилове или ръчно форматиране.
    '''
    ''' Какво прави процедурата стъпка по стъпка:
    '''
    ''' 1) Създава нов параграф в края на документа.
    '''
    ''' 2) Премахва всякаква автоматична номерация или списък
    '''    (важно при добавяне след таблици, списъци или булети).
    '''
    ''' 3) Прилага стандартизирано форматиране:
    '''    - Шрифт: Cambria
    '''    - Размер: 12 pt
    '''    - Подравняване: двустранно (justify)
    '''    - Междуредие: единично
    '''    - Без разстояние преди/след параграфа
    '''    - Контролирани отстъпи (ляв, десен и първи ред)
    '''
    ''' 4) Поддържа интелигентно удебеляване:
    '''    Ако makeBoldUntilColon = True и текстът съдържа ':',
    '''    всичко ДО и ВКЛЮЧИТЕЛНО двоеточието става Bold,
    '''    а останалата част след него остава нормална или Bold
    '''    според параметъра Bold.
    '''
    '''    Това е особено полезно за заглавни части от типа:
    '''    "Основание: съгласно чл. 5 от закона..."
    '''
    ''' 5) Винаги минава текста през CorrectText(),
    '''    което означава, че тук може да се прави:
    '''    - корекция на кавички
    '''    - корекция на тирета
    '''    - корекция на интервали
    '''    - замяна на грешни символи
    '''
    ''' 6) В края автоматично добавя празен ред (нов параграф),
    '''    за да може следващият текст да започне коректно.
    '''
    ''' Кога да се използва тази процедура:
    ''' ------------------------------------------
    ''' - При автоматично генериране на текст в Word
    ''' - При масова обработка на документи
    ''' - При създаване на проектни доклади
    ''' - При попълване на шаблонни текстове
    ''' - При автоматизирано сглобяване на документация
    '''
    ''' Това е ключова помощна процедура за изграждане на
    ''' чисти, подредени и еднакво форматирани документи.
    ''' </summary>
    '''
    ''' <param name="wordDoc">
    ''' Активният Word документ, в който ще се добави параграфът.
    ''' </param>
    '''
    ''' <param name="text">
    ''' Текстът на абзаца. Ако е празен, процедурата не прави нищо.
    ''' </param>
    '''
    ''' <param name="Bold">
    ''' Определя дали целият текст (или частта след двоеточието)
    ''' да бъде удебелен. По подразбиране False.
    ''' </param>
    '''
    ''' <param name="FirstLine">
    ''' Отстъп на първия ред (висящ отстъп).
    ''' Стойността 35.4375 е типична за професионални документи.
    ''' </param>
    '''
    ''' <param name="makeBoldUntilColon">
    ''' Ако е True и текстът съдържа ':',
    ''' всичко до двоеточието става Bold,
    ''' а останалата част след него остава според параметъра Bold.
    ''' </param>
    Private Sub AddParagraph(ByVal wordDoc As Word.Document,
                         ByVal text As String,
                         Optional Bold As Boolean = False,
                         Optional FirstLine As Double = 35.4375,
                         Optional makeBoldUntilColon As Boolean = False)
        ' ---------------------------------------------------------
        ' 1) Ако няма текст – няма смисъл да правим нищо
        ' ---------------------------------------------------------
        If text = "" Then Exit Sub
        ' ---------------------------------------------------------
        ' 2) Дефинираме стандарта за форматиране на документа
        '    (можеш да ги изнесеш като глобални настройки по-късно)
        ' ---------------------------------------------------------
        Dim fontName As String = "Cambria"
        Dim fontSize As Integer = 12
        Dim fontBold As Boolean = Bold
        Dim paragraphAlignment As Word.WdParagraphAlignment =
        Word.WdParagraphAlignment.wdAlignParagraphJustify
        Dim spaceBefore As Single = 0
        Dim spaceAfter As Single = 0
        Dim lineSpacing As Word.WdLineSpacing =
        Word.WdLineSpacing.wdLineSpaceSingle
        Dim leftIndent As Single = 0
        Dim rightIndent As Single = 0
        Dim firstLineIndent As Single = FirstLine

        ' Премахваме всякаква номерация/списък, ако случайно е активна
        wordDoc.Paragraphs(wordDoc.Paragraphs.Count).Range.ListFormat.RemoveNumbers()

        Dim paragraphs = text.Split(
                        New String() {vbCrLf},
                        StringSplitOptions.RemoveEmptyEntries)
        For Each singleParagraph In paragraphs
            ' ---------------------------------------------------------
            ' 3) Създаваме нов параграф в края на документа
            ' ---------------------------------------------------------
            Dim parag = wordDoc.Content.Paragraphs.Add
            wordDoc.Paragraphs(wordDoc.Paragraphs.Count).Range.ListFormat.RemoveNumbers()
            text = singleParagraph
            With parag
                ' ---------------------------------------------------------
                ' 4) Проверяваме дали ще правим специално Bold до ':'
                ' ---------------------------------------------------------
                If makeBoldUntilColon AndAlso text.Contains(":") Then
                    ' Намерим позицията на първото двоеточие
                    Dim colonIndex As Integer = text.IndexOf(":") + 1
                    ' Разделяме текста на две части:
                    Dim beforeColon As String = text.Substring(0, colonIndex)
                    Dim afterColon As String = text.Substring(colonIndex)
                    ' ---- Първа част (до ':') → винаги Bold ----
                    .Range.Text = CorrectText(beforeColon)
                    .Range.Font.Name = fontName
                    .Range.Font.Size = fontSize
                    .Range.Font.Bold = True
                    ' ---- Втора част (след ':') ----
                    Dim afterRange As Word.Range = .Range.Duplicate
                    afterRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                    afterRange.Text = afterColon
                    afterRange.Font.Name = fontName
                    afterRange.Font.Size = fontSize
                    afterRange.Font.Bold = fontBold
                Else
                    ' ---------------------------------------------------------
                    ' 5) Обикновен случай – целият текст с еднакъв формат
                    ' ---------------------------------------------------------
                    .Range.Text = CorrectText(text)
                    .Range.Font.Name = fontName
                    .Range.Font.Size = fontSize
                    .Range.Font.Bold = fontBold
                End If
                ' ---------------------------------------------------------
                ' 6) Общи настройки за параграфа
                ' ---------------------------------------------------------
                .Format.Alignment = paragraphAlignment
                .Format.SpaceBefore = spaceBefore
                .Format.SpaceAfter = spaceAfter
                .Format.LineSpacingRule = lineSpacing
                .Range.ParagraphFormat.LeftIndent = leftIndent
                .Range.ParagraphFormat.RightIndent = rightIndent
                .Range.ParagraphFormat.FirstLineIndent = firstLineIndent
                ' Добавяме празен ред след абзаца
                .Range.InsertParagraphAfter()
            End With
        Next
    End Sub
    Private Sub AddПодпис(ByVal wordDoc As Word.Document, Name As String)
        With wordDoc.Content.Paragraphs.Add
            .Range.Text = "Проектант: ______________________________"
            .Range.Font.Name = "Cambria"
            .Range.Font.Size = 12
            .Range.Font.Bold = False
            .Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
            .Format.SpaceBefore = 108 ' Разредка преди: 108 пкт
            .Format.SpaceAfter = 0 ' Разредка след: 0 пкт
            .Format.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle ' Редова разрядка: единична
            .Format.FirstLineIndent = 35.4375 ' Отстъп: специален -> 1,25 см
            .Range.InsertParagraphAfter()
        End With
        With wordDoc.Content.Paragraphs.Add
            .Range.Text = "/ " + Name + " /"
            .Range.Font.Name = "Cambria"
            .Range.Font.Size = 12
            .Range.Font.Bold = False
            .Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
            .Format.SpaceBefore = 0 ' Разредка преди: 0 пкт
            .Format.SpaceAfter = 6  ' Разредка след: 6 пкт
            .Format.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle ' Редова разрядка: единична
            .Format.FirstLineIndent = 35.4375 ' Отстъп: специален -> 1,25 см
            .Range.InsertParagraphAfter()
        End With
    End Sub
    ' ОБЯСНИТЕЛНА ЗАПИСКА по безопасност, хигиена на труда и пожарна безопасност
    Private Sub BHTB(wordDoc As Word.Document, dicObekt As Dictionary(Of String, String))
        wordDoc.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak)
        With wordDoc.Content.Paragraphs.Add
            .Range.Text = "ОБЯСНИТЕЛНА ЗАПИСКА"
            .Range.Font.Name = "Cambria"
            .Range.Font.Size = 20
            .Range.Font.Bold = True
            .Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            .Range.InsertParagraphAfter()
        End With
        With wordDoc.Content.Paragraphs.Add
            .Range.Text = "по безопасност, хигиена на труда и пожарна безопасност"
            .Range.Font.Name = "Cambria"
            .Range.Font.Size = 14
            .Range.Font.Bold = True
            .Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            .Range.InsertParagraphAfter()
        End With
        Dim text As String = " "
        AddParagraph(wordDoc, text, False)
        text = $"ОБЕКТ: {dicObekt("ОБЕКТ")}"
        AddParagraph(wordDoc, text, True)
        If dicObekt("ОБЕКТ") <> dicObekt("МЕСТОПОЛОЖЕНИЕ") Then
            text = $"МЕСТОПОЛОЖЕНИЕ: {dicObekt("МЕСТОПОЛОЖЕНИЕ")}"
            AddParagraph(wordDoc, text, True)
        End If
        AddParagraph(wordDoc, " ", True)

        AddParagraph(wordDoc, "I. Обща част", True, 0)
        text = "Този раздел БХТПБ към част " + Кавички + "ЕЛЕКТРО" + Кавички + " се разработи въз основа на Инструкция №1 за обема и съдържанието на тази част към проекта."
        AddParagraph(wordDoc, text, False)
        AddParagraph(wordDoc, "При разработване на проекта са спазени изискванията на :", False)
        AddParagraph(wordDoc, "1. Наредба № 3 от 9 юни 2004 г. за устройството на електрическите уредби и електропроводните линии, Обн. ДВ., бр. 90 от 13 октомври 2004 г. и бр. 91 от 14 октомври 2004 г.")
        AddParagraph(wordDoc, "2. Наредба № 16-116 от 8 февруари 2008 г. за техническата експлоатация на енергийните съоръжения, Обн. ДВ., бр. 26 от 7 март 2008 г.")
        AddParagraph(wordDoc, "3. Наредба № 1 от 27 май 2010 г. за проектиране, изграждане и поддържане на електрически уредби НН в сгради, Обн. ДВ., бр. 46 от 18 юни 2010 г.")
        AddParagraph(wordDoc, "4. Наредба № Iз-1971 от 29 октомври 2009 г. за строително-технически правила и норми за осигуряване на безопасност при пожар, Обн. ДВ., бр. 96 от 4 декември 2009 г.")
        AddParagraph(wordDoc, "5. Наредба № 4 от 22 декември 2010 г. за мълниезащитата на сгради, външни съоръжения и открити пространства, Обн. ДВ., бр. 6 от 18 януари 2011 г.")
        AddParagraph(wordDoc, "6. Правилник за безопасност и здраве при работа в електрически уредби на електрически и топлофикационни централи и по електрически мрежи")
        AddParagraph(wordDoc, "II. Предвидени са следните мероприятия по БХТПБ съгласно номенклатурата на факторите", True, 0)
        AddParagraph(wordDoc, "код 1.Електрообезопасяване", True)
        AddParagraph(wordDoc, "1. Осигурена е III-та категория по сигурност на електрозахранване.", False)
        AddParagraph(wordDoc, "2. Избрана е система на ел. захранване с директно заземен център на трансформатора на страна НН (380/220V) и е предвиден четирижилен (с нулево жило) захранващ кабел НН.")
        AddParagraph(wordDoc, "3. Корпусите на всички електрически съоръжения и апарати са свързани към заземителния проводник на електрическата мрежа, който не е защитен никъде по протежението си с предпазни апарати и има сигурни метални връзки.")
        AddParagraph(wordDoc, "4. Съобразно т. 2 е предвидено повторно заземление, което се осъществява чрез свързване заземителната шина на главното разпределително табло към комплектен заземител, състоящ се от 2бр. поц. колове 63/63/6 мм, с дължина 1,5м.", False)
        AddParagraph(wordDoc, "5. Срещу авария на електрическите съоръжения и захранващите линии в ел. таблата са предвидени:")
        AddParagraph(wordDoc, "а) за защита от къси съединения и претоварване - автоматични предпазители и прекъсвачи", False, 50)
        AddParagraph(wordDoc, "б) Към автоматичните прекъсвачи на изводите на ел. таблото към контактите с общо предназначение, както и към ел. бойлерите ще се монтират модули дефектно токова защита осигуряващи защита на потребителите срещу непряк контакт; допълнителна защита срещу пряк контакт и защита на електрическите инсталации срещу пробив в изолацията и пожар.", False, 50)
        AddParagraph(wordDoc, "Защитата на сградата от преки попадения на мълния се осъществява чрез мълниеприемник с изпреварващо действие, монтиран на покрива на сградата. Заземление ще се осъществи с помощта на 1 брой мълниезащитен отвод от екструдиран проводник AlMgSi ф8мм, свързан към заземително устройство, състоящо се от 2 броя заземителни колове 63/63/6 мм с дължина 1,5м. Преходното съпротивление на заземителите не трябва да надвишава 10 ома, в противен случай да се увеличи броя на заземителните колове.", False)
        AddParagraph(wordDoc, "код 4. Изкуствено осветление", True)
        AddParagraph(wordDoc, "Във всички помещения на сградата е осигурена нормална осветеност съгласно нормите на МХП, като са спазени изискванията на Наредба №49 за изкуствено осветление на сградите.", False)
        AddParagraph(wordDoc, "код 9. Пожарна безопасност", True)
        AddParagraph(wordDoc, "За предотвратяване възникването на пожар или взрив електрическите съоръжения и ел. инсталации са съобразно изискванията на Наредба №3 и Наредба N Iз–1971/29.10.2009 год.", False)
        AddParagraph(wordDoc, "III. Специфични изисквания за ел. съоръженията", True, 0)
        text = "1. Контрол за преходното съпротивление на заземителната уредба трябва да се извършва не по-рядко от един път в годината в сухите летни месеци, съгласно изискванията на нормативните актове."
        AddParagraph(wordDoc, text, False)
        text = "2.Изисквания към ел. осветителната инсталация:"
        AddParagraph(wordDoc, text, False)
        text = "2.1. Не по-рядко от един път в годината да се проверява състоянието на осветителната уредба/наличие на стъкла, решетки, мрежи, изправност на уплътненията на осветителните тела със специално предназначение."
        AddParagraph(wordDoc, text, False, 50)
        text = "2.2. Почистването на осветителните тела, подмяната на изгорели лампи и ремонт на инсталацията да се извършва при изключено напрежение."
        AddParagraph(wordDoc, text, False, 50)
        text = "2.3. Всички работи на височина да се извършват от правоспособен персонал с използването на обезопасени помощни средства."
        AddParagraph(wordDoc, text, False, 50)
        AddПодпис(wordDoc, dicObekt("ПРОЕКТАНТ"))
    End Sub
    ' ОБЯСНИТЕЛНА ЗАПИСКА по техника на безопасност по време на строителството
    Private Sub POIS(wordDoc As Word.Document, dicObekt As Dictionary(Of String, String))
        Dim stackTrace As New StackTrace()
        Dim callingMethod As StackFrame = stackTrace.GetFrame(1)
        Dim methodName As String = callingMethod.GetMethod().Name
        Dim PV_Записка As Boolean = False
        If methodName = "PV_New_zapiska" Then
            PV_Записка = True
        Else
            PV_Записка = False
        End If

        wordDoc.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak)
        With wordDoc.Content.Paragraphs.Add
            .Range.Text = "ОБЯСНИТЕЛНА ЗАПИСКА"
            .Range.Font.Name = "Cambria"
            .Range.Font.Size = 20
            .Range.Font.Bold = True
            .Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            .Range.InsertParagraphAfter()
        End With
        With wordDoc.Content.Paragraphs.Add
            .Range.Text = "по техника на безопасност по време на строителството"
            .Range.Font.Name = "Cambria"
            .Range.Font.Size = 14
            .Range.Font.Bold = True
            .Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            .Range.InsertParagraphAfter()
        End With
        Dim text As String = " "
        AddParagraph(wordDoc, text, False)
        text = $"ОБЕКТ: {dicObekt("ОБЕКТ")}"
        AddParagraph(wordDoc, text, True)
        If dicObekt("ОБЕКТ") <> dicObekt("МЕСТОПОЛОЖЕНИЕ") Then
            text = $"МЕСТОПОЛОЖЕНИЕ: {dicObekt("МЕСТОПОЛОЖЕНИЕ")}"
            AddParagraph(wordDoc, text, True)
        End If
        AddParagraph(wordDoc, " ", True)
        AddParagraph(wordDoc, "I. Общи изисквания за безопасност при строително-монтажни работи на електрически инсталации", True)
        AddParagraph(wordDoc, "1. На строителната площадка не се допускат лица които:", False)
        AddParagraph(wordDoc, "- не са навършили 18 години;", False, 50)
        AddParagraph(wordDoc, "- не са преминали предварителен медицински преглед;", False, 50)
        AddParagraph(wordDoc, "- не са преминали първоначален инструктаж;", False, 50)
        AddParagraph(wordDoc, "- не са запознати с условията за възникване на аварии и пожари.", False, 50)
        AddParagraph(wordDoc, "2. Лицата, работещи на строителната площадка са длъжни:", False)
        AddParagraph(wordDoc, "- да се явяват на работа в състояние, което позволява да изпълняват възложената им задачи;", False, 50)
        AddParagraph(wordDoc, "- да се грижат за здравето и безопасността си, както и за здравето и безопасността и на други лица, пряко засегнати от тяхната дейност, в съответствие с квалификацията им и дадените им инструкции;", False, 50)
        AddParagraph(wordDoc, "- да работят само с техника и съоръжения, за които имат документ за правоспособност;", False, 50)
        AddParagraph(wordDoc, "- да работят само с изправна техника, съоръжения и инструменти, а при неизправност да сигнализират незабавно прекия си ръководител;", False, 50)
        AddParagraph(wordDoc, "- да проверяват преди започване на работа изправността на работното оборудване;", False, 50)
        AddParagraph(wordDoc, "- да поддържат ред и чистота на работното място;", False, 50)
        AddParagraph(wordDoc, "- да ползват предоставените им лични предпазни средства и работно облекло, съгласно тяхното предназначение;", False, 50)
        AddParagraph(wordDoc, "- да поддържат и повишават знанията и квалификацията си по безопасност и хигиена на труда и противопожарна охрана;", False, 50)
        AddParagraph(wordDoc, "- да оказват долекарска помощ на пострадалия при трудова злополука или при други увреждания.", False, 50)
        AddParagraph(wordDoc, "3. Работници трябва да носят подходящи ЛПС, включително каски, ръкавици, очила и обувки със защитно бомбе.", False)
        AddParagraph(wordDoc, "4. Работниците трябва да преминат подходящо обучение за безопасност преди започване на работа на строителната площадка. Това обучение трябва да обхваща теми като първа помощ, разпознаване на опасности и процедури за реагиране при извънредни ситуации.", False)
        AddParagraph(wordDoc, "5. Пушенето, храненето и пиенето са забранени в рамките на определената работна зона.", False)
        AddParagraph(wordDoc, "6. Използването на инструменти и оборудване трябва да бъде в съответствие с инструкциите на производителя.", False)
        AddParagraph(wordDoc, "7. Работните зони трябва да бъдат добре осветени и свободни от препятствия, за да се предотврати спъване или падане. Всички отпадъци и материали трябва да се съхраняват правилно, за да се избегне създаването на опасни ситуации.", False)
        AddParagraph(wordDoc, "8. Временното главното ел. табло трябва да бъде заземено.", False)
        AddParagraph(wordDoc, "9. При използване на фургони за офисни и битови нужди да трябва да бъдат заземени.", False)
        AddParagraph(wordDoc, "10. Съпротивлението на заземлението да бъде доказано със сертификат от специализиран орган за контрол.", False)
        AddParagraph(wordDoc, "II. Обучение и инструктаж на работниците при строително-монтажни работи на електрически инсталации", True)
        AddParagraph(wordDoc, "Съгласно Наредба № РД-07-2 от 16.12.2009 г. за условията и реда за провеждане на периодично обучение и инструктаж на работниците и служителите по правилата за осигуряване на здравословни и безопасни условия на труд, трябва да се осигури подходящо обучение и инструктажи на работещите на строителната площадка.")
        'AddParagraph(wordDoc, "1. Първоначално обучение:", True)
        AddParagraph(wordDoc, "Преди да започне работа на строителната площадка, всеки работник трябва да премине първоначално обучение по безопасност и здраве при работа.", False)
        'AddParagraph(wordDoc, "Това обучение трябва да покрива следните теми:", False)
        'AddParagraph(wordDoc, "- Основни нормативни изисквания за безопасност и здраве при работа на строителни обекти;", False, 50)
        'AddParagraph(wordDoc, "- Мерки за осигуряване на безопасност и здраве при работа на строителни обекти;", False, 50)
        'AddParagraph(wordDoc, "- Правила за пожарна безопасност;", False, 50)
        'AddParagraph(wordDoc, "- Правила за използване на лични предпазни средства и специално работно облекло;", False, 50)
        'AddParagraph(wordDoc, "- Управление на отпадъците и опасни вещества на строителната площадка;", False, 50)
        'AddParagraph(wordDoc, "- Управление на риска на строителната площадка;", False, 50)
        'AddParagraph(wordDoc, "- Специфични рискове и мерки за безопасност, свързани с вида работа, изпълнявана от работника;", False, 50)
        'AddParagraph(wordDoc, "2. Периодични инструктажи:", True)
        AddParagraph(wordDoc, "Да провеждат периодични инструктажи въз основа на оценката на риска, но не по-малко от един три месеца.", False)
        'AddParagraph(wordDoc, "Тези инструктажи трябва да обхващат:", False)
        'AddParagraph(wordDoc, "- Нови законови изисквания, насоки и добри практики за безопасност и здраве при работа;", False, 50)
        'AddParagraph(wordDoc, "- Промени в строителния обект, работните процеси или използваното оборудване, които могат да повлияят на безопасността и здравето;", False, 50)
        'AddParagraph(wordDoc, "- Резултати от оценката на риска и предприетите коригиращи действия;", False, 50)
        'AddParagraph(wordDoc, "- Специфични рискове и мерки за безопасност, свързани с вида работа, изпълнявана от работника.", False, 50)
        'AddParagraph(wordDoc, "3. Ежедневен инструктаж:", True)
        AddParagraph(wordDoc, "Ежедневните инструктажи са кратки и целят преглед на мерките за безопасност и обсъждане на потенциални рискове.", False)
        AddParagraph(wordDoc, "Видовете инструктажи да се документират съгласно изискванията на Наредба № РД-07-2 от 16.12.2009 г.", False)
        AddParagraph(wordDoc, "III. Специфични изисквания за безопасност при строително-монтажни работи на електрически инсталации", True)
        AddParagraph(wordDoc, "При строително-монтажни работи на електрически инсталации да се спазват изискванията на:", False)
        AddParagraph(wordDoc, "ПРАВИЛНИК за безопасност и здраве при работа по електрообзавеждането с напрежение до 1000V", False, 50)
        AddParagraph(wordDoc, "ПРАВИЛНИК за безопасност и здраве при работа в електрически уредби на електрически и топлофикационни централи и по електрически мрежи", False, 50)
        AddParagraph(wordDoc, "Работниците трябва да избягват контакт с мокри или влажни повърхности.", False)
        AddParagraph(wordDoc, "Всички открити проводници трябва да бъдат ясно маркирани.", False)
        AddParagraph(wordDoc, "Работниците трябва да избягват използването на удължители или разклонители, ако е възможно, и никога да не ги претоварват.", False)
        AddParagraph(wordDoc, "Всички кабели и проводници трябва да бъдат здраво закрепени и защитени от механично увреждане.", False)
        'AddParagraph(wordDoc, "Работниците трябва да избягват използването на удължители или разклонители и никога да не ги претоварват.", False)
        AddParagraph(wordDoc, "Към строително - монтажни работи на електрически съоръжения и уредби се пристъпва само след изключване на напрежението, поставяне на предупредителна табела " + Кавички + "Не включвай - работят хора" + Кавички + " и след проверката за отсъствие на напрежение.", False)
        'AddParagraph(wordDoc, "Използване на изолирани инструменти и оборудване, които са тествани и сертифицирани за работа под напрежение.", False)
        'AddParagraph(wordDoc, "Проверки на всички електрически кабели, устройства и инструменти за повреди или дефекти преди и след всяка смяна.", False)
        AddParagraph(wordDoc, "Поставяне на ясни предупредителни знаци в зоните с високо напрежение и опасни работни зони. Поставяне на ясни предупредителни знаци в зоните с високо напрежение и опасни работни зони. Ограждане на работните зони, за да се предотврати несанкциониран достъп.", False)
        AddParagraph(wordDoc, "При изпълнение на ел. инсталациите и свързващите ги елементи да се съблюдава за възможни включвания и попадане под напрежение.", False)
        AddParagraph(wordDoc, "При изпълнение на ел. инсталациите се забранява изтегляне на кабели и проводници в неукрепени тръби и разклонителни кутии, когато в тях има други проводници под напрежение.", False)
        AddParagraph(wordDoc, "Пробиването на отвори в стени и междуетажни плочи, както и изтеглянето на проводници в хоризонтална посока да става от скелета или платформи.", False)
        AddParagraph(wordDoc, "Забранява се използването на случайни предмети и нестандартни стълби за извършване на монтажни работи по ел. инсталациите.", False)
        AddParagraph(wordDoc, "Задължително е при включване на електрическите двигатели да бъдат предупредени всички работници чрез звуков сигнал и устно и да бъдат отстранени от машините, съоръженията и т.н.", False)
        AddParagraph(wordDoc, "Забранено е да се правят ремонти, подменят предпазители и други операции в трафопостове, по електрическите табла и останалите електрически уредби и съоръжения, когато са подложени на въздействието на лоши атмосферни влияния - дъжд, сняг, буря и гръмотевици.", False)
        'AddParagraph(wordDoc, "IV. Процедури за реагиране при извънредни ситуации", True)
        'AddParagraph(wordDoc, "1. При поражение от електрически ток", True)
        'AddParagraph(wordDoc, "- Незабавно премахнете лицето от източника на енергия. Ако е възможно, изключете директно източника на енергия. Ако не можете да стигнете до превключвателя, използвайте непроводим предмет (като дървен стол или изолираща щанга), за да отделите лицето от източника на енергия.", False, 50)
        'AddParagraph(wordDoc, "- Осигурете първа помощ: Ако лицето все още е в контакт с електрическия ток, използвайте суха, непроводима материя (като гумени ръкавици или изолираща щанга), за да го освободите. След като лицето е освободено, проверете жизнените му показатели.", False, 50)
        'AddParagraph(wordDoc, "- Обадете се за помощ: Свържете се с местните служби за спешна помощ веднага щом ситуацията позволява. Дайте информация за вида на електрическото напрежение (високо или ниско напрежение) и дали лицето все още е в контакт с източника на енергия.", False, 50)
        'AddParagraph(wordDoc, "2. При възникване на пожар:", True)
        'AddParagraph(wordDoc, "В случай на пожар работещите на строителната площадка трябва да направят следното:", False, 50)
        'AddParagraph(wordDoc, "- При пожар е незабавното напускане на сградата. Затворете вратите след себе си, за да забавите разпространението на огъня.", False, 50)
        'AddParagraph(wordDoc, "- След като излезете от сградата, извикайте местните служби за спешна помощ и информирайте оператора за местоположението на пожара и естеството.", False, 50)
        'AddParagraph(wordDoc, "- Ако пожарът е малък и е на безопасно разстояние, използвайте пожарогасителя. Уверете се, че знаете как да използвате пожарогасителя, преди да го използвате.", False, 50)
        'AddParagraph(wordDoc, "- Стойте далеч от района на пожара, за да избегнете вдишване на дим.", False, 50)
        'AddParagraph(wordDoc, "3. Наранявания или заболявания:", True)
        'AddParagraph(wordDoc, "Нараняванията или заболяванията могат да възникнат по различни начини на строителна площадка, от леки порязвания и натъртвания до по-тежки наранявания или здравословни проблеми. Ето какво трябва да направите в случай на нараняване или заболяване: ", False)
        'AddParagraph(wordDoc, "- Прекратете дейността Ако сте ранени или болни, спрете незабавно всяка дейност, която може да влоши състоянието ви.", False, 50)
        'AddParagraph(wordDoc, "- Осигурете първа помощ: Ако нараняването е леко (напр. порязване или натъртване), почистете и превържете раната, за да предотвратите инфекция. Ако нараняването е тежко или имате съмнения, свържете се с местните служби за спешна помощ.", False, 50)
        'AddParagraph(wordDoc, "- Докладвайте инцидента: Съобщете за инцидента на вашия супервайзор или отговорник за безопасността. Те ще документират инцидента и ще гарантират, че получавате подходяща медицинска помощ, ако е необходимо.", False, 50)
        'AddParagraph(wordDoc, "- Върнете се на работа: След като получите разрешение от медицински специалист, върнете се на работа. Уверете се, че сте взели всички необходими предпазни мерки, за да предотвратите повторно нараняване или заболяване.", False, 50)

        If PV_Записка Then
            AddParagraph(wordDoc, "IV. Специфични изисквания за безопасност при строително-монтажни работи по фотоволтаични централи", True)
            AddParagraph(wordDoc, "Строително-монтажните работи по фотоволтаичните централи се извършват без изключване на напрежението от две лице с четвърта квалификационна група при следните условия:", False)
            AddParagraph(wordDoc, "1. лицето да е преминало обучение за работа под напрежение над 1000 V;", False)
            AddParagraph(wordDoc, "2. да се използват само специализирани инструменти за работа под напрежение над 1000 V, вкл. диелектрични кърпи, капачки, щипки и накрайници;", False)
            AddParagraph(wordDoc, "3. да се работи със защитни очила/щит за лице, диелектрични ръкавици и обувки;", False)
            AddParagraph(wordDoc, "4. инструментите и защитните средства да са изпитани и преди работа да са прегледани за механични увреждания на изолацията;", False)
            AddParagraph(wordDoc, "5. работното облекло да е с дълги ръкави и да не се работи с навити ръкави.", False)
            AddParagraph(wordDoc, "V. Заключение", True)
        Else
            AddParagraph(wordDoc, "IV. Заключение", True)
        End If
        AddParagraph(wordDoc, "При извършване на ел.монтажни работи и изпитване на готови електрически инсталации да се вземат предпазни мерки за защита на работещите, както и на други лица, намиращи се на строежа, от попадане под напрежение и поражения от електрически ток.", False)
        AddParagraph(wordDoc, "Инсталации, в частност връзки в електроинсталации, заварки и укрепвания на тръби, фасонни части, отоплителни тела и др., които се изпълняват едновременно с други видове СМР, се монтират с повишено внимание и под непосредствено наблюдение на ръководителя на обекта или упълномощен от него бригадир.", False)
        AddParagraph(wordDoc, "При извършване на огневи работи задължително се спазват всички изисквания от НАРЕДБА № I-209 от 22.11.2004г.", False)
        AddПодпис(wordDoc, dicObekt("ПРОЕКТАНТ"))
    End Sub
    Private Function GetCustomProperty(owner As IAcSmPersist, propertyName As String) As String
        ' Create a reference to the Custom Property Bag
        Dim customPropertyBag As AcSmCustomPropertyBag
        If owner.GetTypeName() = "AcSmSheet" Then
            Dim sheet As AcSmSheet = CType(owner, AcSmSheet)
            customPropertyBag = sheet.GetCustomPropertyBag()
        ElseIf owner.GetTypeName() = "AcSmSheetSet" Then
            Dim sheetSet As AcSmSheetSet = CType(owner, AcSmSheetSet)
            customPropertyBag = sheetSet.GetCustomPropertyBag()
        Else
            Throw New InvalidCastException("The owner object is not of type AcSmSheet or AcSmSheetSet.")
        End If
        ' Get the property
        Dim customPropertyValue As AcSmCustomPropertyValue = customPropertyBag.GetProperty(propertyName)
        If customPropertyValue IsNot Nothing Then
            ' Return the value of the property
            Return customPropertyValue.GetValue().ToString
        Else
            Return Nothing
        End If
    End Function
    ' Декларация на функцията OpenWordDocument
    Public Function OpenWordDocument(ByVal fileName As String, ByVal wordApp As Word.Application) As Word.Document
        Dim wordDoc As Word.Document = Nothing
        Try
            ' Проверка дали wordApp е Nothing
            If wordApp Is Nothing Then
                ' Ако wordApp е Nothing, създайте нова инстанция на Word.Application
                wordApp = New Word.Application()
            End If
            ' Проверка дали файлът е отворен
            For Each doc As Word.Document In wordApp.Documents
                If doc.FullName = fileName Then
                    wordDoc = doc
                    wordDoc.Activate()
                    Exit For
                End If
            Next
            ' Ако файлът не е отворен, отворете го
            If wordDoc Is Nothing Then
                If File.Exists(fileName) Then
                    wordDoc = wordApp.Documents.Open(fileName)
                Else
                    wordDoc = wordApp.Documents.Add()
                    wordDoc.SaveAs(fileName)
                End If
                ' Активиране на документа
                If Not wordDoc.ActiveWindow Is Nothing Then
                    wordDoc.Activate()
                End If
            End If
            ' Изтриване на съдържанието, ако е необходимо
            wordDoc.Content.Delete()
            wordApp.Visible = True
        Catch ex As Exception
            MsgBox("Error opening or activating Word document: " & ex.Message)
        End Try
        Return wordDoc
    End Function
    ' Функция за задаване на форматирането на клетките
    Sub SetCellFormat(cell As Word.Cell,
                      text As String,
                      fontName As String,
                      bold As Boolean,
                      size As Integer,
                      alignment As Word.WdParagraphAlignment,
                      verticalAlignment As Word.WdCellVerticalAlignment,
                      Optional border As Boolean = False,
                      Optional Height As Double = 0.1)
        Try
            With cell
                .Range.Text = CorrectText(text)
                .Range.Font.Name = fontName
                .Range.Font.Bold = bold
                .Range.Font.Size = size
                .VerticalAlignment = verticalAlignment ' Вертикално центриране
                .Height = wordApp.InchesToPoints(Height / 2.54)
                With .Range.ParagraphFormat
                    .Alignment = alignment ' Хоризонтално центриране
                    .SpaceBefore = 0 ' Нула интервал преди параграфа
                    .SpaceAfter = 0  ' Нула интервал след параграфа
                End With
                If border Then
                    With .Borders(Word.WdBorderType.wdBorderBottom)
                        .LineStyle = Word.WdLineStyle.wdLineStyleSingle
                        .LineWidth = Word.WdLineWidth.wdLineWidth025pt
                        .Color = Word.WdColor.wdColorBlack
                    End With
                    With .Borders(Word.WdBorderType.wdBorderTop)
                        .LineStyle = Word.WdLineStyle.wdLineStyleSingle
                        .LineWidth = Word.WdLineWidth.wdLineWidth025pt
                        .Color = Word.WdColor.wdColorBlack
                    End With
                Else
                    'With .Borders(Word.WdBorderType.wdBorderBottom)
                    '    .LineStyle = Word.WdLineStyle.wdLineStyleNone
                    'End With
                End If
            End With
        Catch ex As Exception
            MsgBox("Error opening or activating Word document: " & ex.Message)
        End Try
    End Sub
    ' Функция за поправка на текст
    Public Function CorrectText(Text As String) As String
        Do
            Dim originalText As String = Text
            Text = Text.Replace("..", ".")
            Text = Text.Replace("  ", " ")
            Text = Text.Replace(",,", ",")
            Text = Text.Replace(" ,", ",")
            Text = Text.Replace(",.", ".")
            Text = Text.Replace(vbCrLf, " ")
            ' Прекъсване на цикъла, ако няма повече замени
            If Text = originalText Then Exit Do
        Loop
        ' Дефиниране на регулярен израз за търсене на точки след главни букви
        Dim pattern As String = "\.\p{L}"
        Dim regex As New Regex(pattern)
        ' Намиране на съвпадения
        Dim matches As MatchCollection = regex.Matches(Text)
        For Each match As Match In matches
            Text = Text.Insert(match.Index + 1, " ") ' Добавя интервал след точката
        Next
        Return Text
    End Function
    Sub CreateCustomStyles(ByVal doc As Word.Document)
        ' Проверка дали стилът вече съществува
        If Not StyleExists(doc, "CustomTitle") Then
            ' Създаване на нов стил за заглавие
            Dim titleStyle As Word.Style = doc.Styles.Add("CustomTitle", Word.WdStyleType.wdStyleTypeParagraph)
            titleStyle.Font.Name = "Arial"
            titleStyle.Font.Size = 24
            titleStyle.Font.Bold = True
            titleStyle.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        End If

        If Not StyleExists(doc, "CustomSubtitle") Then
            ' Създаване на нов стил за подзаглавие
            Dim subtitleStyle As Word.Style = doc.Styles.Add("CustomSubtitle", Word.WdStyleType.wdStyleTypeParagraph)
            subtitleStyle.Font.Name = "Arial"
            subtitleStyle.Font.Size = 18
            subtitleStyle.Font.Italic = True
            subtitleStyle.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
        End If

        If Not StyleExists(doc, "CustomNormalText") Then
            ' Създаване на нов стил за нормален текст
            Dim normalTextStyle As Word.Style = doc.Styles.Add("CustomNormalText", Word.WdStyleType.wdStyleTypeParagraph)
            normalTextStyle.Font.Name = "Calibri"
            normalTextStyle.Font.Size = 12
            normalTextStyle.Font.Bold = False
            normalTextStyle.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
        End If
    End Sub
    Function StyleExists(ByVal doc As Word.Document, ByVal styleName As String) As Boolean
        For Each style As Word.Style In doc.Styles
            If style.NameLocal = styleName Then
                Return True
            End If
        Next
        Return False
    End Function
    <CommandMethod("PVZapiska")>
    Public Sub PV_New_zapiska()
        Dim fullName As String = Application.DocumentManager.MdiActiveDocument.Name
        Dim filePath As String = Path.GetDirectoryName(fullName)
        Dim Path_Name As String = Mid(filePath, InStrRev(filePath, "\") + 1, Len(filePath))
        Dim fileName As String = filePath + "\" + "Обяснителна записка_PV.docx"
        Dim Text As String = ""
        Dim File_DST As String = filePath + "\" + Path_Name + ".dst"
        Dim wordDoc As Word.Document = OpenWordDocument(fileName, wordApp)
        Dim dicSignature As Dictionary(Of String, String)
        Try
            If wordDoc Is Nothing Then
                Exit Sub
            End If
            CreateCustomStyles(wordDoc)
            ' Извлича данните от блок "Insert_Signature"
            dicSignature = addArrayBlock()

            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database

            Челен_Лист(wordDoc, dicSignature)
            СЪДЪРЖАНИЕ(wordDoc, dicSignature)
            ОБЯСНИТЕЛНА_ЗАПИСКА_PV(wordDoc, acDoc, acCurDb, dicSignature)
            ЗАПИСКА_Батерия(wordDoc, acDoc, acCurDb, dicSignature)
            Външно_Захранване(wordDoc, dicSignature)
            Заземителна(wordDoc, dicSignature)
            Мълниезащита(wordDoc, acDoc, acCurDb)
            Заключение(wordDoc, dicSignature)
            Записка_ПИЦ(wordDoc, acDoc, acCurDb, dicSignature)
            POIS(wordDoc, dicSignature)
            If MsgBox(Title:="Завърших записката.", Buttons:=MsgBoxStyle.YesNo, Prompt:="Да създам ли записка ПОЖАРНА БЕЗОПАСНОСТ") = MsgBoxResult.Yes Then

                wordDoc.Save()
                wordDoc.Close(False)

                fileName = filePath + "\" + "Обяснителна записка_PV_Пожарна.docx"

                wordDoc = OpenWordDocument(fileName, wordApp)
                ПОЖАРНА(wordDoc, acDoc, acCurDb, dicSignature)
                MsgBox("Завърших записка ПОЖАРНА БЕЗОПАСНОСТ.")
            End If
        Catch ex As Exception
            MsgBox("Error:  " & ex.Message)
        Finally
            If wordDoc IsNot Nothing Then
                Try
                    wordDoc.Save()
                    wordDoc.Close(False)
                Catch ex As Exception
                    MsgBox("Error while closing document: " & ex.Message)
                Finally
                    Marshal.ReleaseComObject(wordDoc)
                    wordDoc = Nothing
                End Try
            End If
            If wordApp IsNot Nothing Then
                ' Затваряме всички отворени документи в Word приложението
                For Each doc As Word.Document In wordApp.Documents
                    doc.Close(False)
                    Marshal.ReleaseComObject(doc)
                Next
                ' Затваряме Word приложението
                Try
                    wordApp.Quit(False)
                    Marshal.ReleaseComObject(wordApp)
                    wordApp = Nothing
                Catch ex As Exception
                    MsgBox("Error while quitting Word: " & ex.Message)
                End Try
                ' Прекратяване на процеса на Word
                For Each proc As Process In Process.GetProcessesByName("WINWORD")
                    proc.Kill()
                Next
            End If
            ' Извикайте събирача на боклука, за да освободите незабавно ресурсите
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub
    '
    ' ОБЯСНИТЕЛНА ЗАПИСКА ЗА ПОЖАРНА
    '
    Private Sub ПОЖАРНА(wordDoc As Word.Document,
                        acDoc As Document,
                        acCurDb As Database,
                        dicObekt As Dictionary(Of String, String))

        Dim pDouOpts = New PromptDoubleOptions("")
        wordDoc.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak)
        Челен_Лист(wordDoc, dicObekt)
        СЪДЪРЖАНИЕ(wordDoc, dicObekt)
        Dim stackTrace As New StackTrace()
        Dim callingMethod As StackFrame = stackTrace.GetFrame(1)
        Dim methodName As String = callingMethod.GetMethod().Name
        Dim PV_Записка As Boolean = False
        Dim boТРАФО As Boolean = False
        If methodName = "PV_New_zapiska" Then
            PV_Записка = True
        Else
            PV_Записка = False
        End If
        With wordDoc.Content.Paragraphs.Add
            .Range.Text = "ОБЯСНИТЕЛНА ЗАПИСКА"
            .Range.Font.Name = "Cambria"
            .Range.Font.Size = 20
            .Range.Font.Bold = True
            .Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            .Range.InsertParagraphAfter()
        End With
        With wordDoc.Content.Paragraphs.Add
            .Range.Text = "ПОЖАРНА БЕЗОПАСНОСТ"
            .Range.Font.Name = "Cambria"
            .Range.Font.Size = 14
            .Range.Font.Bold = True
            .Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            .Range.InsertParagraphAfter()
        End With
        Dim text As String = " "
        AddParagraph(wordDoc, text, False)
        text = $"ОБЕКТ: {dicObekt("ОБЕКТ")}"
        AddParagraph(wordDoc, text, True)
        If dicObekt("ОБЕКТ") <> dicObekt("МЕСТОПОЛОЖЕНИЕ") Then
            text = $"МЕСТОПОЛОЖЕНИЕ: {dicObekt("МЕСТОПОЛОЖЕНИЕ")}"
            AddParagraph(wordDoc, text, True)
        End If

        FormatParagraph(wordDoc, "ОБЩА ЧАСТ", wordApp, level:=1, resetLevel:=True)
        AddParagraph(wordDoc, dicObekt("Продажба"), False)
        text = "Обяснителната записка е съставена съгласно изискванията на НАРЕДБА № 1з-1971 за строително-техническите правила и норми за осигуряване на безопасност при пожар от 29.10.2009 година."
        text += " При разработването на раздела ПОЖАРНА БЕЗОПАСНОСТ са спазени изискванията на Приложение №3 към чл.4 ал.1 на Наредбата за обхват и съдържание на част"
        text += $" {Кавички}пожарна безопасност на инвестиционния проект{Кавички}."
        AddParagraph(wordDoc, text, False)
        AddParagraph(wordDoc, $"При разработката на част {Кавички}пожарна безопасност{Кавички} са спазени изискванията на следните нормативни документи:")
        AddParagraph(wordDoc, "- Наредба № Із-1971 за строително-техническите правила и норми за осигуряване на безопасност при пожар (ДВ бр. 96/2009 год.)")
        AddParagraph(wordDoc, "– Наредба № 3 за устройството на електрическите уредби и електропроводните линии (ДВ бр. 90 и 91/2004 год.);")
        AddParagraph(wordDoc, "– Наредба № 16-116 за техническа експлоатация на енергообзавеждането (ДВ. бр.26/2008 год.);")
        AddParagraph(wordDoc, "– Наредба № 9 от 9.06.2004 г. за техническата експлоатация на електрически централи и мрежи (ДВ бр. 72/2004 год.);")
        AddParagraph(wordDoc, "– Наредба № 2 за минималните изисквания за осигуряване на здравословни и безопасни условия на труд при извършване на строително-монтажни работи (ДВ бр. 37/2004 год.).")
        text = $"Възложителят {dicObekt("ВЪЗЛОЖИТЕЛ")} в границите на имота ще въведе в експлоатация необходимата електрическа уредба на генераторно напрежение със следните показатели:"
        AddParagraph(wordDoc, text, False)
        AddParagraph(wordDoc, dicObekt("Площ"))
        AddParagraph(wordDoc, dicObekt("Тегло"))
        AddParagraph(wordDoc, dicObekt("Височина"))
        AddParagraph(wordDoc, dicObekt("Инсталирана"))
        AddParagraph(wordDoc, dicObekt("Изходна"))
        AddParagraph(wordDoc, dicObekt("Панели"))
        AddParagraph(wordDoc, dicObekt("Инвертори"))

        FormatParagraph(wordDoc, "ПАСИВНИ МЕРКИ ЗА ПОЖАРНА БЕЗОПАСНОСТ", wordApp, level:=1)
        FormatParagraph(wordDoc, "Проектни обемно-планировъчни и функционални показатели на строежа", wordApp, level:=2)
        With text
            AddParagraph(wordDoc, dicObekt("Земя"), False)
            pDouOpts = New PromptDoubleOptions("")
            With pDouOpts
                .Keywords.Add("Да")
                .Keywords.Add("Не")
                .Keywords.Default = "Да"
                .Message = vbCrLf & "В проекта има ли предвидено изграждането на трафопост?"
                .AllowZero = False
                .AllowNegative = False
            End With

            Dim pKeyRes = acDoc.Editor.GetDouble(pDouOpts)
            If pKeyRes.StringResult = "Да" Then
                boТРАФО = True
            Else
                boТРАФО = False
            End If
            Dim strOilInLiters As String = ""
            Dim strOilInKg As String = ""
            Dim strМощТрафо As String = ""

            If boТРАФО Then
                Dim oilData As New Dictionary(Of String, Integer) From {
                    {"100", 140}, {"160", 180}, {"200", 210},
                    {"250", 240}, {"315", 270}, {"400", 310},
                    {"500", 350}, {"630", 420}, {"800", 480},
                    {"1000", 540}, {"1250", 650}, {"1600", 740},
                    {"2000", 900}, {"2500", 1050}, {"3150", 1170},
                    {"сух", 0} ' За сух трансформатор и двете стойности са 0
                    }

                ' Извеждаме наличните стойности на мощността
                Dim promptOptions As New PromptKeywordOptions("Каква е мощността на трансформатора?")
                With promptOptions
                    ' Добавяме ключовете в promptOptions
                    For Each key As String In oilData.Keys
                        .Keywords.Add(key)
                    Next
                    .AllowNone = False ' Забраняваме вход без избор
                    ' Задаваме стойност по подразбиране
                    .Keywords.Default = "1000" ' Стойност по подразбиране
                End With
                ' Получаваме резултата от избора на потребителя
                Dim result = acDoc.Editor.GetKeywords(promptOptions)
                ' Проверяваме дали потребителят е избрал валидни данни
                If result.Status = PromptStatus.OK Then
                    'Получаваме избраната от потребителя мощност
                    strМощТрафо = result.StringResult.Trim().ToLower() ' Съхраняваме избраната мощност
                    ' Проверяваме дали входът съществува в речника
                    If oilData.ContainsKey(strМощТрафо) Then
                        ' Вземаме информацията за маслото
                        Dim oilInKg As Integer = oilData(strМощТрафо)
                        ' Изчисляваме количеството масло в литри
                        Dim density As Double = 0.895 ' Плътност на маслото в кг/л
                        Dim oilInLiters As Integer = Math.Round(oilInKg / density)
                        ' Присвояваме стойностите на низовите променливи
                        strOilInKg = oilInKg.ToString() ' Стойностите за килограми вече са цели числа
                        strOilInLiters = Math.Round(oilInLiters).ToString() ' Закръгляме стойността за литри до цяло число
                    End If
                End If
            End If
            text = "За нуждите на фотоволтаична електрическа централа е предвидено изграждането на фотоволтаичните модули (панели)"
            text += ", кабелни връзки между отделните фотоволтаични панели и връзката им с инверторите"
            text += ", кабели ниско напрежение"
            text += If(boТРАФО, ", кабели средно напрежение", "")
            text += If(boТРАФО, ", бетонов комплектен трансформаторен пост предназначен за електрозахранване на фотоволтаични електрически централи /БКТП/", "")
            text += "."
            AddParagraph(wordDoc, text)
            If boТРАФО Then
                text = "Силовият трансформатор е разположен в отделно трансформаторно помещение-килия, изолирано от другите помещения."
                If strМощТрафо = "сух" Then
                    text += " В пректното решение е предвидено да се използва сух трансформатор."
                Else
                    text += $" Трансформаторът e с мощност {strМощТрафо} kVА, маслен, херметичен, съдържащ до {strOilInKg}кг /{strOilInLiters}л/ трансформаторно масло."
                End If
                text += $" В килията пред трансформатора се монтира метална предпазна врата-решетка."
                text += $" Отварянето ѝ и достъпът да трансформатора е възможен само ако е заземен разединителя в КРУ {Кавички}Защита трафо{Кавички}."
                AddParagraph(wordDoc, text)
            End If
        End With
        FormatParagraph(wordDoc, "Клас и подклас на функционална пожарна опасност", wordApp, level:=2)
        With text
            AddParagraph(wordDoc, "Съгласно чл. 8. (1) и таблица 1 на Наредба № Із-1971 за строително-техническите правила и норми за осигуряване на безопасност при пожар (ДВ бр. 96/2009 год.) стоежа се класифицира като: ", Bold:=False)
            AddParagraph(wordDoc, "• Клас на функционална пожарна опасност - Ф5;", Bold:=True)
            AddParagraph(wordDoc, "• Подклас                                                                     - Ф5.1;", Bold:=True)
            AddParagraph(wordDoc, "• Категория по пожарна опасност                    - Ф5Д.", Bold:=True)

            text = "За нуждите на фотоволтаичната електрическа централа не се предвижда съхранение, използване и/или произвеждане на горими материали."
            text += " Високоенергиен източник на запалване може за възникне само в аварийна ситуаця - неправилни действия от страна на поддържащия персонал, възникване на аварийна ситуация в други части на обекта."
            text += " Поради това вероятността за възникване на високоенергиен източник на запалване е минимална."
            AddParagraph(wordDoc, text, Bold:=False)
            text = $"Съгласно чл. 237 на Наредба № Із-1971 за строително-техническите правила и норми за осигуряване на безопасност при пожар (ДВ бр. 96/2009 год.) стоежа се класифицира като:"
            AddParagraph(wordDoc, text, Bold:=False)
            AddParagraph(wordDoc, $"•  първа група - {Кавички}Нормална пожарна опасност{Кавички}.", Bold:=True)
        End With
        FormatParagraph(wordDoc, "Степен на огнеустойчивост на строежа и на конструктивните му елементи", wordApp, level:=2)
        With text
            AddParagraph(wordDoc, "• Фотоволтаична електроцентрала-панели", Bold:=True)
            AddParagraph(wordDoc, "• Кабел ниско напрежение ", Bold:=True)
            If boТРАФО Then
                AddParagraph(wordDoc, "• Кабел средно напрежение", Bold:=True)
                AddParagraph(wordDoc, "• Разпределителна уредба средно напрежение", Bold:=True)
                AddParagraph(wordDoc, "• Разпределителна уредба ниско напрежение", Bold:=True)
                AddParagraph(wordDoc, "• Силов трансформатор", Bold:=True)
            End If
        End With
        FormatParagraph(wordDoc, "Огнеустойчивост на обслужващи и вентилационни шахти:", wordApp, level:=2)
        AddParagraph(wordDoc, "• В проектно решение не се предвижда изграждане на обслужващи и вентилационни шахти.", Bold:=False)
        FormatParagraph(wordDoc, "Огнеустойчивост на пожарозащитните прегради", wordApp, level:=2)
        AddParagraph(wordDoc, "• В проектно решение не се предвижда изграждане на пожарозащитните прегради.", Bold:=False)
        FormatParagraph(wordDoc, "Проектна огнеустойчивост и клас по реакция на огън на огнезащитаваните конструктивни елементи на сградата:", wordApp, level:=2)
        AddParagraph(wordDoc, "• В проектно решение не се предвижда изграждане на огнезащитавани конструктивни елементи на сградата.", Bold:=False)
        FormatParagraph(wordDoc, "Пътища за противопожарни цели", wordApp, level:=2)
        With text
            text = "До входа на фотоволтаична електрическа централа е предвидено да има път, който може да се използва за противопожарни цели."
            text += " Пътя е с достатъчна ширина за да може по него да преминава противопожарна техника необходима за пожарогасителни нужди."
            text += " На територията на ФЕЦ няма изградени площадки и стълби за пожарогасителни и аварийно-спасителни дейности."

            If boТРАФО Then
                text += " Предвиденият в проекта бетонов комплектен трансформаторен пост е ситуиран на границата на имота."
                text += " До границата на имота им изграден път, чрез който ще се осигури достъп за провеждане на пожарогасителни и аварийно-спасителни дейности."
            End If
            AddParagraph(wordDoc, text, False)
        End With
        FormatParagraph(wordDoc, "Класове по реакция на огън на продуктите за конструктивни елементи", wordApp, level:=2)
        Select Case dicObekt("Монтаж")
            Case "Земя"
                text = "В проекта е предвидено цялата фотоволтаична електрическа централа да се изгради на открит терен."
                text += " Поради тази причина не се предвиждат мерки за предотвратяване на разпространението на горенето между етажите при пожар в сградата"
            Case "Покрив"
                FormatParagraph(wordDoc, "Мерки за предотвратяване на разпространението на горенето между етажите при пожар в сградата", wordApp, level:=2)
                text = " В проекта е предвидено цялата фотоволтаична електрическа централа да се изгради върху съществуваща покривна конструкция."
                text += " Достъпът до откритата част на ФЕЦ ще се осъществява от съществуващата сграда."
                text += " При възникване на необходимост от евакуация на техническия екип ще се използват евакуационните маршрути на съществуващата сграда."
            Case "Двете"
                FormatParagraph(wordDoc, "Мерки за предотвратяване на разпространението на горенето между етажите при пожар в сградата", wordApp, level:=2)
                text = " В проекта е предвидено част фотоволтаична електрическа централа да се изгради на терена, а друга част да се изгради върху съществуваща покривна конструкция."
                text += " В проекта за изградена на открит терен част е предвидено достатъчно разстояние между конструкциите, на които са монтирани панелите (по-голямо от 4 м), което ще позволи нормалната евакуация от зоната на конструкциите."
                text += " Достъпът до частта, изградена върху покривната конструкция, до откритата част на ФЕЦ, ще се осъществява от съществуващата сграда."
                text += " При възникване на необходимост от евакуация на техническия екип ще се използват евакуационните маршрути на съществуващата сграда."
        End Select
        AddParagraph(wordDoc, text, False)
        FormatParagraph(wordDoc, "Мерки за предотвратяване на разпространението на горенето между етажите при пожар в сградата:", wordApp, level:=2)
        FormatParagraph(wordDoc, "Мерки за пожарна безопасност при проектиране на остъклени площи по цялата височина на фасадите на сгради:", wordApp, level:=2)
        FormatParagraph(wordDoc, "Мерки пожарна безопасност при проектиране на сгради с вентилируеми фасади:", wordApp, level:=2)
        FormatParagraph(wordDoc, "Мерки за предотвратяване на разпространението на горенето при пожар между пожарните сектори, разположени един над друг или един до друг:", wordApp, level:=2)
        FormatParagraph(wordDoc, "Мерки за пожарна безопасност при проектиране на отоплителни инсталации:", wordApp, level:=2)
        FormatParagraph(wordDoc, "Група опасност на помещенията, сградите, откритите съоръжения или части от тях:", wordApp, level:=2)
        FormatParagraph(wordDoc, "Осигурени условия на успешна евакуация", wordApp, level:=2)
        With text
            text = "Експлоатацията на слънчевата електроцентрала ще протича нормално без присъствието на постоянен обслужващ персонал."
            text += " Периодично, квалифициран технически екип ще извършва планови обслужвания и ремонти."
            text += " В случай на аварийни ситуации, същият екип ще осъществява необходимите ремонти и ще оказва съдействие на противопожарните екипи."
            AddParagraph(wordDoc, text, False)
            Select Case dicObekt("Монтаж")
                Case "Земя"
                    text = " В проекта е предвидено цялата фотоволтаична електрическа централа да се изгради на открит терен."
                    text += " В проекта е предвидено достатъчно разстояние между конструкциите на които са монтирани панелите (по-голямо от 4м) и това ще позволи нормалната евакуация от зоната на конструкциите."
                Case "Покрив"
                    text = " В проекта е предвидено цялата фотоволтаична електрическа централа да се изгради върху съществуваща покривна конструкция."
                    text += " Достъпът до откритата част на ФЕЦ ще се осъществява от съществуващат сграда."
                    text += " При възникване на необходимост от евакуация на техническия екип ще се използват евакуационните маршрути на съществуващата сграда."
                Case "Двете"
                    text = " В проекта е предвидено част фотоволтаична електрическа централа да се изгради на терена, а друга част да се изгради върху съществуваща покривна конструкция."
                    text += " В проекта за изградена на открит терен част е предвидено достатъчно разстояние между конструкциите, на които са монтирани панелите (по-голямо от 4 м), което ще позволи нормалната евакуация от зоната на конструкциите."
                    text += " Достъпът до частта, изградена върху покривната конструкция, до откритата част на ФЕЦ, ще се осъществява от съществуващата сграда."
                    text += " При възникване на необходимост от евакуация на техническия екип ще се използват евакуационните маршрути на съществуващата сграда."
            End Select
            text += " Липсата на постоянен технически персонал означава, че не се налагат специални мерки за евакуация площадката на централата."
            AddParagraph(wordDoc, text, False)
        End With
        FormatParagraph(wordDoc, "АКТИВНИ МЕРКИ ЗА ПОЖАРНА БЕЗОПАСНОСТ", wordApp, level:=1)
        FormatParagraph(wordDoc, "Обемно-планировъчни и функционални показатели за пожарогасителни инсталации:", wordApp, level:=2)
        text = "Съгласно т. 2.19 на приложение № 1 на Наредба № Iз-1971 от 29 октомври 2009 г. за строително-технически правила и норми за осигуряване на безопасност при пожар"
        text += " не се изисква система за пожарогасене."
        AddParagraph(wordDoc, text, False)
        FormatParagraph(wordDoc, "Обемно-планировъчни и функционални показатели за пожароизвестителни системи и системи за звукова сигнализация:", wordApp, level:=2)
        text = "Експлоатацията на слънчевата електроцентрала ще протича нормално без присъствието на постоянен обслужващ персонал."
        text += " Периодично, квалифициран технически екип ще извършва планови обслужвания и ремонти."
        AddParagraph(wordDoc, text, False)
        text = "В настоящия проект не се предвижда изграждането на пожароизвестителна система."
        AddParagraph(wordDoc, text, False)
        FormatParagraph(wordDoc, "Обемно-планировъчни и функционални показатели за системи за гласово сигнализиране:", wordApp, level:=2)
        text = "Експлоатацията на слънчевата електроцентрала ще протича нормално без присъствието на постоянен обслужващ персонал."
        text += " Периодично, квалифициран технически екип ще извършва планови обслужвания и ремонти."
        AddParagraph(wordDoc, text, False)
        If dicObekt("Монтаж") = "Земя" Then
            text = "В настоящия проект не се предвижда изграждането на системи за гласово сигнализиране."
            AddParagraph(wordDoc, text, False)
        Else
            Select Case dicObekt("Монтаж")
                Case "Покрив"
                Case "Двете"
            End Select
            AddParagraph(wordDoc, text, False)
        End If
        FormatParagraph(wordDoc, "Обемно-планировъчни и функционални показатели за вентилационни системи за отвеждане на дим и топлина", wordApp, level:=2)
        If boТРАФО Then
            text = "За зоната където са монтирани фотоволтаичните панели не се предвижда изграждането на димо-топлоотвеждаща инсталация."
            AddParagraph(wordDoc, text, False)
            text = "За димоотвеждане и топлоотвеждане на БКТП са проектирани две врати с вградени решетки разположени срещуположно."
            AddParagraph(wordDoc, text, False)
        Else
            text = "В настоящия проект не се предвижда изграждането на димо-топлоотвеждаща инсталация."
            AddParagraph(wordDoc, text, False)
        End If
        FormatParagraph(wordDoc, "Обемно-планировъчни и функционални показатели за вентилационни инсталации за предотвратяване на пожар", wordApp, level:=2)
        text = "Фотоволтаичната електроцентрала като открита площ при нормална експлоатация не се отделят горими вещества и не може да се създаде обща или локална експлозивна атмосфера."
        text += " Поради тази причина не попада под разпоредбите на чл. 66 на Наредба № Iз-1971 от 29 октомври 2009 г. за строително-технически правила и норми за осигуряване на безопасност при пожар "
        AddParagraph(wordDoc, text, False)
        text = "В настоящия проект не се предвижда изграждането на вентилационни инсталации за предотвратяване на пожар."
        AddParagraph(wordDoc, text, False)
        FormatParagraph(wordDoc, "Функционални показатели за водоснабдяване за пожарогасене:", wordApp, level:=2)
        text = "Нов водопровод за пожарогасене за фотоволтаичната централа не се предвижда, тъй като тя е от категория по пожарна безопасност Ф5Д."
        AddParagraph(wordDoc, text, False)
        text = "Фотоволтаичната централа ВИНАГИ е под напрежение и не се допуска използване на вода или други гасителни вещества, които включват вода. Възможно е използването само на газове елегазови смеси и гасителен прах."
        AddParagraph(wordDoc, text, False)
        FormatParagraph(wordDoc, "Функционални показатели за пожаротехнически средства за първоначално гасене на пожари", wordApp, level:=2)
        text = $"В приложение № 2 към чл. 3, ал. 2 – {Кавички}Пожаротехнически средства за първоначално гасене на пожари в сгради, помещения, съоръжения и инсталации, в т.ч. свободни дворни площи{Кавички}"
        text += " – за фотоволтаичната електроцентрала като открита площ не са предвидени конкретни изисквания за преносими уреди и съоръжения за първоначално пожарогасене."
        AddParagraph(wordDoc, text, False)
        text += "Независимо от това, ако се приеме, че откритата ФЕЦ се класифицира като открита разпределителна уредба, в настоящия проект за нея са предвидени два броя ръчни прахови пожарогасителя по 6 kg."
        AddParagraph(wordDoc, text, False)
        text += $"Поради специфичния характер на ФЕЦ (съоръжение от категория {Кавички}особено опасно за поражение от електрически ток{Кавички}), пожарогасителите ще бъдат разположени в специален шкаф, монтиран в близост до главното разпределително табло."
        text += " Шкафът няма да бъде заключван, с цел осигуряване на лесен и бърз достъп до пожарогасителите."
        AddParagraph(wordDoc, text, False)
        FormatParagraph(wordDoc, "Функционални показатели на аварийно евакуационно и аварийно работно осветление", wordApp, level:=2)
        text = "Експлоатацията на слънчевата електроцентрала ще протича нормално без присъствието на постоянен обслужващ персонал."
        text += " Периодично, квалифициран технически екип ще извършва планови обслужвания и ремонти."
        text += " В проекта не се предвижда работно осветление."
        text += " Поради тази причина планови обслужвания и ремонти ще се извършват само в светлата част на денонощието."
        AddParagraph(wordDoc, text, False)
        text = " Предвид спецификата на строежа не се предвижда израждане на евакуационно и аварийно работно осветление."
        AddParagraph(wordDoc, text, False)
        FormatParagraph(wordDoc, "Принципна схема на проектираните активни мерки за защита", wordApp, level:=2)
        AddParagraph(wordDoc, "Забележка: ", False)
        text = "Графичните материали за всяка от активните мерки за пожарна безопасност са елемент и се съдържат в отделните части на инвестиционния проект."
        AddParagraph(wordDoc, text, False)
        AddПодпис(wordDoc, dicObekt("ПРОЕКТАНТ"))
    End Sub
    Sub ЗАПИСКА_Батерия(wordDoc As Word.Document,
                             acDoc As Document,
                             acCurDb As Database,
                             dicObekt As Dictionary(Of String, String))
        Dim Text As String = ""
        Dim ss_Tabla = cu.GetObjects("INSERT", "Изберете БЛОКОВЕТЕ в чертеж съдържащи БАТЕРИИТЕ:")
        If ss_Tabla Is Nothing Then
            MsgBox("Няма маркиран нито един блок.")
            dicObekt.Add("СРТ_Помещение_BAT", "#####")
            dicObekt.Add("СРТ_Кота_BAT", "#####")
            Exit Sub
        End If
        Dim ss_Kabeli = cu.GetObjects("LINE", "Изберете КАБЕЛИТЕ в чертеж свързващи БАТЕРИИТЕ:")
        If ss_Kabeli Is Nothing Then
            MsgBox("Няма маркиран нито едина линия.")
            Exit Sub
        End If
        Dim ss_GRT = cu.GetObjects("INSERT", "Изберете ТАБЛОТО в чертеж към който са свързани БАТЕРИИТЕ:", allowMultiple:=False)
        Dim СРТ_Помещение_BAT = cu.GetObjects_TEXT("Изберете текст съдържаш помещението в което се намират БАТЕРИТЕ")
        dicObekt.Add("СРТ_Помещение_BAT", СРТ_Помещение_BAT)
        Dim СРТ_Кота_BAT = cu.GetObjects_TEXT("Изберете текст съдържаш котата на която се намират БАТЕРИТЕ")
        dicObekt.Add("СРТ_Кота_BAT", СРТ_Кота_BAT)

        Dim blkRecId As ObjectId = ObjectId.Null

        Dim strBattery_Name As String = ""
        Dim strBattery_Type As String = ""
        Dim strBattery_Tablo As String = ""
        Dim intBattery_Count As Integer = 0
        Dim douBattery_Power As Double = 0
        Dim douBattery_Capaci As Double = 0
        Dim douBattery_Power_1 As Double = 0
        Dim douBattery_Capaci_1 As Double = 0

        Dim strBattery_ТАБЛО As String = ""

        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                For Each sObj As SelectedObject In ss_Tabla
                    blkRecId = sObj.ObjectId
                    Dim acBlkRef As BlockReference = DirectCast(acTrans.GetObject(blkRecId, OpenMode.ForRead), BlockReference)
                    Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                    Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection

                    Dim blName As String = (CType(acBlkRef.DynamicBlockTableRecord.GetObject(OpenMode.ForRead), BlockTableRecord)).Name
                    If blName <> "Батерия" Then Continue For

                    'For Each prop As DynamicBlockReferenceProperty In props
                    '    If prop.PropertyName = "Батерия вид" Then strBattery_Name = prop.Value
                    '    If prop.PropertyName = "Батерия клетка" Then strBattery_Type = prop.Value
                    '    If prop.PropertyName = "Мощност" Then douBattery_Power += Convert.ToDouble(prop.Value)
                    '    If prop.PropertyName = "Капацитет" Then douBattery_Capaci += Convert.ToDouble(prop.Value)
                    'Next

                    intBattery_Count += 1
                    For Each objID As ObjectId In attCol
                        Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                        Dim acAttRef As AttributeReference = dbObj
                        If acAttRef.Tag = "БАТЕРИЯ_ВИД" Then strBattery_Name = acAttRef.TextString
                        If acAttRef.Tag = "БАТЕРИЯ_КЛЕТКА" Then strBattery_Type = acAttRef.TextString
                        If acAttRef.Tag = "МОЩНОСТ" Then douBattery_Power_1 = Convert.ToDouble(acAttRef.TextString)
                        If acAttRef.Tag = "КАПАЦИТЕТ" Then douBattery_Capaci_1 = Convert.ToDouble(acAttRef.TextString)
                    Next
                    douBattery_Power += douBattery_Power_1
                    douBattery_Capaci += douBattery_Capaci_1
                Next

                blkRecId = ss_GRT(0).ObjectId
                Dim acBlkRef_GRT As BlockReference = DirectCast(acTrans.GetObject(blkRecId, OpenMode.ForRead), BlockReference)
                Dim attCol_GRT As AttributeCollection = acBlkRef_GRT.AttributeCollection
                For Each objID As ObjectId In attCol_GRT
                    Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                    Dim acAttRef As AttributeReference = dbObj
                    If acAttRef.Tag = "ТАБЛО" Then strBattery_Tablo = acAttRef.TextString
                Next
            Catch ex As Exception
                MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
                acTrans.Abort()
            End Try
        End Using
        FormatParagraph(wordDoc, "ЛОКАЛНО СЪОРЪЖЕНИЕ ЗА СЪХРАНЕНИЕ НА ЕЛЕКТРИЧЕСКА ЕНЕРГИЯ", wordApp, level:=1)
        With Text
            Text = $"В проекта е предвидено да се монтират {intBattery_Count} {If(intBattery_Count > 1, "броя локални съоръжения", "брой локално съоръжения")} за съхранение на електрическа енергия тип {strBattery_Name} изградени от батерийни клетки {strBattery_Type}."
            Text += " То представлява модерна и ефективна система за съхранение на електроенергия, предназначена за различни приложения, включително интеграция с възобновяеми източници на енергия и стабилизиране на електроразпределителни мрежи."
            AddParagraph(wordDoc, Text, False)
            Text = $"{If(intBattery_Count > 1, "Предвидените ", "Предвидения ")} {intBattery_Count} {If(intBattery_Count > 1, "броя", "брой")} ЛСС съоръжения ще са с обща МОЩНОСТ {douBattery_Power}kW и ОБЩ ОПЕРАТИВЕН КАПАЦИТЕТ {douBattery_Capaci}kWh."
            AddParagraph(wordDoc, Text, False)
            Text = "Локалното съоръжение за съхранение на електрическа енергия е проектирано така, че да осигурява минимален гарантиран капацитет, позволяващ непрекъсната работа с продължителност от поне 2 часа."
            AddParagraph(wordDoc, Text, False)
            Text = $"Слсс/Pлсс = {douBattery_Capaci}/{douBattery_Power} = {douBattery_Capaci / douBattery_Power} часа"
            AddParagraph(wordDoc, Text, False)
            Text = "Където:"
            AddParagraph(wordDoc, Text, True)
            Text = "Слсс - Общ капацитет на локалното съоръжение за съхранение на електрическа енергия"
            AddParagraph(wordDoc, Text, False)
            Text = "Pлсс - Обща мощност на локалното съоръжение за съхранение на електрическа енергия"
            AddParagraph(wordDoc, Text, False)
            Text = "Локалното съоръжение за съхранение на електрическа енергия постига ефективност над 80% при c/2 и пълен цикъл на зареждане и разреждане."
            Text += " То включва инвертор с управление, конвертори, охлаждаща система, батерии и всички съпътстващи електрически консуматори."
            Text += " Не включва трансформаторната уредба в случай на присъединяване към уредба високо/средно напрежение."
            AddParagraph(wordDoc, Text, False)
        End With
        FormatParagraph(wordDoc, "ОСНОВНИ ХАРАКТЕРИСТИКИ:", wordApp, level:=2)
        With Text
            Text = $"Основни технически характеристики на предвидените в пректа локалното съоръжение за съхранение на електрическа енергия тип {strBattery_Name} са:"
            AddParagraph(wordDoc, Text, False)
            Text = $"Работно напрежение (AC)					400 V"
            AddParagraph(wordDoc, Text, False)
            Text = $"Максимален капацитет					{douBattery_Capaci_1} kWh"
            AddParagraph(wordDoc, Text, False)
            Text = $"Максимална мощност на зареждане 			≤{douBattery_Power_1} kW"
            AddParagraph(wordDoc, Text, False)
            Text = $"Максимална мощност на разреждане 			≤{douBattery_Power_1} kW"
            AddParagraph(wordDoc, Text, False)
            Text = $"Тип на клетката на батерията				{strBattery_Type}"
            AddParagraph(wordDoc, Text, False)
            Text = $"Размери (ШxДxВ) 						1350×2300×1350 mm"
            AddParagraph(wordDoc, Text, False)
            Text = $"Тегло с батериите						2500 kg"
            AddParagraph(wordDoc, Text, False)
            Text = $"Степен на защита						IP54"
            AddParagraph(wordDoc, Text, False)
            Text = $"Експлоатационен живот					6000 цикъла"
            AddParagraph(wordDoc, Text, False)
            Text = $"Минимален използваем капацитет на батерията	8% ~ 100%"
            AddParagraph(wordDoc, Text, False)
            Text = $"Работен температурен диапазон			-20 ~+50°C"
            AddParagraph(wordDoc, Text, False)
            Text = $"Охлаждане							Водно (Smart liquid cooling)"
            AddParagraph(wordDoc, Text, False)
            Text = $"Стандарти							IEC62477, IEC61000, UN38.3"
            AddParagraph(wordDoc, Text, False)
            Text = $"Пълното описание на характеристиките на {strBattery_Name} e приложено към проекта."
            AddParagraph(wordDoc, Text, False)
            Text = "Локалното съоръжение за съхранение на електрическа енергияще се разположи на указаното в проекта място на бетонова площадка."
            Text += " При монтажа на съоръженията стриктно да се спазват изискванията на завода производител разтоянието между отделните модули."
            AddParagraph(wordDoc, Text, False)
        End With
        ' ПРОСТОТИИ И ДИВОТИИ
        With Text
            Text = "Проектът ще осигури баланс на съоръжението за съхранение (balance of plant) – обхващащ всички инфраструктурни компоненти и системи, необходими за функционирането на съоръжението."
            AddParagraph(wordDoc, Text, Bold:=False)
            Text = "1. Батерийна технология: Системата използва литиево-желязо-фосфатни (LFP) батерии, които са известни с високата си безопасност, дълъг живот и устойчивост. Батериите осигуряват повече от 6000 цикъла и се очаква да функционират надеждно повече от 10 години."
            AddParagraph(wordDoc, Text, Bold:=False, makeBoldUntilColon:=True)
            Text = "2. Охлаждане: Охлаждането на системата се осъществява по въздушен път, което гарантира ефективност и по-лесна поддръжка без необходимост от сложни охладителни механизми."
            AddParagraph(wordDoc, Text, Bold:=False, makeBoldUntilColon:=True)
            Text = "3. Енергийна ефективност: Системата поддържа тотална хармонична дисторзия (THD) под 3% при номинална мощност, което гарантира минимални загуби при преобразуване на енергията."
            AddParagraph(wordDoc, Text, Bold:=False, makeBoldUntilColon:=True)
            Text = "4. Система за управление на батериите (BMS): Интелигентната система за управление на батериите следи в реално време следи тяхното състояние и предлага защити срещу презареждане, дълбок разряд, прегряване и пренапрежение. Това удължава живота на батериите и повишава надеждността."
            AddParagraph(wordDoc, Text, Bold:=False, makeBoldUntilColon:=True)
            Text = "5. Интеграция и мониторинг: Съоръженията включват усъвършенствани комуникационни интерфейси като RS485, CAN2.0 и Ethernet, които осигуряват мониторинг в реално време и позволяват интеграция с енергийни мениджмънт системи (EMS). Чрез локален екран може да се следи състоянието на модулите и да се управлява процесът на съхранение и използване на енергия."
            AddParagraph(wordDoc, Text, Bold:=False, makeBoldUntilColon:=True)
            Text = "6. Модулен дизайн: Системите са проектирани за паралелно свързване на няколко модула, което позволява лесно разширение на капацитета, ако е необходимо. Това прави съоръженията гъвкави и подходящи за различни обекти с нарастващи енергийни нужди."
            AddParagraph(wordDoc, Text, Bold:=False, makeBoldUntilColon:=True)
            Text = "7. Сигурност и защита: Системата е сертифицирана по международни стандарти за безопасност, като IEC62619 и IEC62116, и включва защити като противопожарни системи с аерозоли или специални газове (например Heptafluoropropane)."
            AddParagraph(wordDoc, Text, Bold:=False, makeBoldUntilColon:=True)
            Text = "8. Устойчивост на околната среда: Съоръженията имат степен на защита IP54, която ги прави устойчиви на различни атмосферни условия, и могат да работят в температурен диапазон от -25°C до 60°C, което ги прави подходящи както за външни, така и за вътрешни инсталации."
            AddParagraph(wordDoc, Text, Bold:=False, makeBoldUntilColon:=True)
        End With
        FormatParagraph(wordDoc, "ПРИЛОЖЕНИЯ:", wordApp, level:=2)
        ' ОЩЕ ПРОСТОТИИ И ДИВОТИИ
        With Text
            Text = $"Предвидените в проекта {intBattery_Count} {If(intBattery_Count > 1, "броя локални", "брой локално")} съоръжения за съхранение /ЛСС/ на електрическа енергия {strBattery_Name} изградени от батериини клетки {strBattery_Type}."
            Text += "Това създава възможност за ефективно използване на слънчевата енергия и оптимално управление на натоварването на електрическата мрежа."
            AddParagraph(wordDoc, Text)
            Text = "Основната функция на системата за съхранение е да акумулира излишната енергия, генерирана от фотоволтаичната инсталация през деня, когато слънчевата радиация е най-интензивна."
            Text += " Тази съхранена енергия може да бъде използвана през периодите на по-ниско производство или по-висока консумация, като вечер и през нощта, когато слънчевото захранване е минимално или липсва."
            AddParagraph(wordDoc, Text)
            Text = "Интеграцията със соларната инсталация предоставя няколко важни предимства:"
            AddParagraph(wordDoc, Text)
            Text = "• Максимално използване на възобновяеми източници: Чрез съхранение на излишната енергия през активните часове на слънчевото производство, системата увеличава ефективността на фотоволтаиците и намалява зависимостта от електрическата мрежа."
            AddParagraph(wordDoc, Text, Bold:=False, makeBoldUntilColon:=True)
            Text = "• Намаляване на пиковото натоварване: Системата може да подава съхранена енергия в моменти на върхова консумация, намалявайки натоварването на мрежата и съответно разходите за електроенергия."
            AddParagraph(wordDoc, Text, Bold:=False, makeBoldUntilColon:=True)
            Text = "• Автономност и резервно захранване: Системата осигурява енергийна независимост, като позволява продължаване на захранването дори при аварии или прекъсвания на мрежата."
            AddParagraph(wordDoc, Text, Bold:=False, makeBoldUntilColon:=True)
            Text = "• Оптимизация на енергийния баланс: Съхранената енергия може да се използва стратегически за постигане на баланс между производство и потребление, като по този начин намалява разходите за енергия."
            AddParagraph(wordDoc, Text, Bold:=False, makeBoldUntilColon:=True)
            Text = "Тази комбинация между фотоволтаична инсталация и съоръжения за съхранение на енергия осигурява оптимално управление на енергийните ресурси, намалява въглеродния отпечатък и подпомага устойчивото развитие на обекта."
            AddParagraph(wordDoc, Text)
        End With
        FormatParagraph(wordDoc, "ВРЪЗКА СЪС СИСТЕМАТА НА ВЪЗЛОИТЕЛЯ:", wordApp, level:=2)
        Dim Kabel(50, 2) As String
        Kabel = cu.GET_LINE_TYPE_KABEL(Kabel, ss_Kabeli, vbFalse)
        Dim arrКабели As New List(Of srtCable)()
        For i = LBound(Kabel) To UBound(Kabel)
            If Kabel(i, 0) = "" Then Exit For
            If InStr(Kabel(i, 0), "поц.шина") > 0 Then Continue For
            If InStr(Kabel(i, 0), "ПВ-A2") > 0 Then Continue For
            If InStr(Kabel(i, 0), "AlMgSi") > 0 Then Continue For
            If InStr(Kabel(i, 0), "ELEKTRO") > 0 Then Continue For
            If InStr(Kabel(i, 0), "H1Z2Z2") > 0 Then Continue For
            If InStr(Kabel(i, 0), "FTP") > 0 Then Continue For
            If InStr(Kabel(i, 0), "коакс") > 0 Then Continue For
            If InStr(Kabel(i, 0), "Оптичен") > 0 Then Continue For

            ' Проверка дали вече съществува запис със същия тип и тръба в Kabeli
            Dim index As Integer = arrКабели.FindIndex(Function(k) k.Тип = Kabel(i, 0) AndAlso k.Тръба = Kabel(i, 1))

            If index = -1 Then
                ' Ако не съществува, създаваме нов запис
                Dim кабел As New srtCable
                кабел.Тип = Kabel(i, 0)
                кабел.Тръба = Kabel(i, 1)
                кабел.Дължина = Double.Parse(Kabel(i, 2)) / 100
                кабел.count = 1

                If (InStr(Kabel(i, 1), "PVC") > 0 Or (InStr(Kabel(i, 1), "HDPE") > 0) And InStr(Kabel(i, 0), "x10mm") > 0) Then
                    кабел.Тип = Kabel(i, 0).Replace("САВТ", "СВТ")
                Else
                    кабел.Тип = Kabel(i, 0)
                End If

                If кабел.Тип = "САВТ" Then
                    кабел.Материал = vbTrue
                End If
                arrКабели.Add(кабел)
            Else
                Dim кабел As New srtCable
                кабел = arrКабели(index)
                кабел.Дължина += Double.Parse(Kabel(i, 2)) / 100
                arrКабели(index) = кабел
            End If
        Next
        Text = $"В проекта е предвидено връзката между табло {Кавички + strBattery_Tablo + Кавички} и {intBattery_Count} {If(intBattery_Count > 1, "броя локални", "брой локално")} съоръжения за съхранение на електрическа енергия тип {strBattery_Name}"
        Text += ", да бъде осъществена чрез"
        'силови кабели, тип САВТ 3x50+25 mm²."
        ' Проверка дали масивът съдържа само един елемент
        If arrКабели.Count = 1 Then
            ' Ако има само един елемент, добавяме директно неговия тип
            Text += " силов кабел тип "
            Text += arrКабели(0).Тип
        Else
            Text += " силови кабели тип "
            ' Цикъл за обединяване на типовете кабели, ако има повече от един
            For i As Integer = 0 To arrКабели.Count - 1
                Dim кабел As srtCable = arrКабели(i)
                ' Добавяме типа на кабела (вече съдържа "mm²")
                Text += кабел.Тип
                ' Проверяваме дали това не е последният елемент
                If i < arrКабели.Count - 2 Then
                    ' Ако не е последният, добавяме запетая
                    Text += ", "
                ElseIf i = arrКабели.Count - 2 Then
                    ' Ако е предпоследният, добавяме " и "
                    Text += " и "
                End If
            Next
        End If
        Text += "."
        dicObekt.Add("ЛСС", Text)
        AddParagraph(wordDoc, Text, False)
    End Sub
    ' ОБЯСНИТЕЛНА ЗАПИСКА ЗА PV
    Sub ОБЯСНИТЕЛНА_ЗАПИСКА_PV(wordDoc As Word.Document,
                             acDoc As Document,
                             acCurDb As Database,
                             dicObekt As Dictionary(Of String, String))

        Dim Text As String = ""
        Dim nfi As New NumberFormatInfo()
        nfi.NumberDecimalSeparator = ","
        nfi.NumberGroupSeparator = " "
        nfi.NumberGroupSizes = New Integer() {3}

        ' Получава пълното име на активния документ
        Dim fullName As String = Application.DocumentManager.MdiActiveDocument.Name
        ' Извлича пътя на файла от пълното име
        Dim filePath As String = Mid(fullName, 1, InStrRev(fullName, "\"))
        ' Извлича името на файла без разширението
        Dim fileName As String = Mid(fullName, InStrRev(fullName, "\") + 1, Len(fullName) - 6)
        ' Декларира променлива за името на Excel файла
        Dim nameExcel As String
        ' Създава нов диалогов прозорец за отваряне на файлове
        Dim OpenFileDialog1 As New Forms.OpenFileDialog()
        ' Задава началната директория на диалоговия прозорец
        OpenFileDialog1.InitialDirectory = filePath
        ' Задава филтър за файлове, които могат да се отварят (само Excel файлове)
        OpenFileDialog1.Filter = "Excel files (*.xls Or *.xlsx)|*.xls;*.xlsx"
        ' Задава предварително име на файла в диалоговия прозорец
        OpenFileDialog1.FileName = "ФВЦ.xlsx"
        ' Проверява дали потребителят е избрал файл и е натиснал OK
        If OpenFileDialog1.ShowDialog <> Windows.Forms.DialogResult.OK Then
            Exit Sub ' Ако не е натиснат OK, излиза от процедурата
        End If
        ' Запазва избраното име на файла
        nameExcel = OpenFileDialog1.FileName
        ' Проверява дали избраният Excel файл е отворен
        Dim stream As FileStream = Nothing
        Try
            stream = File.Open(nameExcel, FileMode.Open, FileAccess.ReadWrite, FileShare.None)
            stream.Close()
        Catch ex As Exception
            ' Показва съобщение, ако файлът е отворен и излиза от процедурата
            MsgBox("Отворен е файл с име : " + Chr(13) + Chr(13) +
           nameExcel + Chr(13) + Chr(13) +
           "Моля затворете го преди да продължите!")
            Exit Sub
        End Try
        ' Създава нов Excel обект
        Dim objExcel_FEC As excel.Application = New excel.Application()
        ' Отваря избрания Excel файл
        Dim excel_Workbook_FEC As excel.Workbook = objExcel_FEC.Workbooks.Open(nameExcel)
        ' Достъпва листа "Таблица" в Excel файла
        Dim wsFEC_Таблица As excel.Worksheet = excel_Workbook_FEC.Worksheets("Таблица")
        ' Достъпва листа "ФЕЦ" в Excel файла
        Dim wsFEC_ФЕЦ As excel.Worksheet = excel_Workbook_FEC.Worksheets("ФЕЦ")
        ' Декларира булева променлива за верификация
        Dim Verifi_FEC As Boolean = True
        ' Декларира променливи за конектор и групи
        Dim Kонектор As Integer = 0
        Dim Групи As Integer = 0
        ' Цикъл за проверка на стойностите в първите 6 реда на колоната 6 в листа "Таблица"
        For i = 1 To 6
            If Verifi_FEC And wsFEC_Таблица.Cells(i, 6).Value <> "OK" Then
                ' Показва съобщение, ако стойността не е "OK" и пита дали да продължи
                Dim response = MsgBox(wsFEC_Таблица.Cells(i, 1).Value + "-> Not OK" + vbCrLf + vbCrLf + "Да продължа ли?", vbYesNo)
                If response = vbNo Then
                    Exit Sub ' Ако потребителят избере "No", излиза от процедурата
                End If
            End If
        Next
        '
        With wordDoc.Content.Paragraphs.Add
            .Range.Text = "ОБЯСНИТЕЛНА ЗАПИСКА"
            .Range.Font.Name = "Cambria"
            .Range.Font.Size = 20
            .Range.Font.Bold = True
            .Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            .Range.InsertParagraphAfter()
        End With

        AddParagraph(wordDoc, Text, False)
        Text = $"ОБЕКТ: {dicObekt("ОБЕКТ")}"
        AddParagraph(wordDoc, Text, True)
        If dicObekt("ОБЕКТ") <> dicObekt("МЕСТОПОЛОЖЕНИЕ") Then
            Text = $"МЕСТОПОЛОЖЕНИЕ: {dicObekt("МЕСТОПОЛОЖЕНИЕ")}"
            AddParagraph(wordDoc, Text, True)
        End If
        AddParagraph(wordDoc, " ", True)

        Dim Части As String = "     "
        Части += IIf("Продажба" = "  #####  ", "", "Архитектурна, ")
        Части += IIf(dicObekt("КОНСТРУКТОР") = "  #####  ", "", "Конструктивна, ")
        Части += IIf(dicObekt("ТЕХНОЛОГИЯ") = "  #####  ", "", "Технологична, ")
        Части += IIf(dicObekt("ВИК") = "  #####  ", "", "ВИК, ")
        Части += IIf(dicObekt("ОВ") = "  #####  ", "", "ОВ, ")
        Части += IIf(dicObekt("ГЕОДЕЗИЯ") = "  #####  ", "", "Геодезия, ")
        Части += IIf(dicObekt("ВП") = "  #####  ", "", "ВП, ")
        Части += IIf(dicObekt("ЕЕФ") = "  #####  ", "", "ЕЕФ, ")
        Части += IIf(dicObekt("ПБ") = "  #####  ", "", "ПБ, ")
        Части += IIf(dicObekt("ПБЗ") = "  #####  ", "", "ПБЗ, ")
        Части += IIf(dicObekt("ПУСО") = "  #####  ", "", "ПУСО, ")
        FormatParagraph(wordDoc, "ОБЩА ЧАСТ", wordApp)
        Части = Части.Substring(0, Части.Length - 2)
        Text = "Настоящият проект се разработи по искане на Възложителя "
        Text += dicObekt("ВЪЗЛОЖИТЕЛ")
        Text += " за изграждане на фотоволтаична инсталация за производство на електрическа енергия от възобновяеми източници, с мощност "
        Text += String.Format("{0:0.0}", wsFEC_ФЕЦ.Cells(2, 3).Value)
        Text += "кW, "
        Dim pDouOpts_Dylbo As PromptDoubleOptions = New PromptDoubleOptions("")
        Dim Dylbo As String
        Dim pDouOpts As PromptDoubleOptions = New PromptDoubleOptions("")
        Try
            With pDouOpts
                .Keywords.Add("Продажба")
                .Keywords.Add("Собствени")
                .Keywords.Add("Двете")
                .Message = vbCrLf & "Изберете начина на продажба : "
                .AllowZero = False
                .AllowNegative = False
            End With
            Dim pKeyRes As PromptDoubleResult = acDoc.Editor.GetDouble(pDouOpts)
            If pKeyRes.Status = PromptStatus.Keyword Then
                Dylbo = pKeyRes.StringResult
            Else
                Dylbo = pKeyRes.Value.ToString()
            End If
            Select Case Dylbo
                Case "Продажба"
                    Text += "само за продажба на произведената електрическа енергия."
                Case "Собствени"
                    Text += "само за собствено потребление на произведената електрическа енергия."
                Case "Двете"
                    Text += "за собствено потребление и продажба на произведената електрическа енергия."
            End Select
        Catch ex As Exception

        End Try
        dicObekt.Add("Продажба", Text)
        AddParagraph(wordDoc, Text)
        AddParagraph(wordDoc, "При разработване на проекта са спазени изискванията на :")
        AddParagraph(wordDoc, "1. Закон за устройство на територията.")
        AddParagraph(wordDoc, "2. НАРЕДБА № 3 от 9 юни 2004 г. за устройството на електрическите уредби и електропроводните линии, Обн. ДВ., бр. 90 от 13 октомври 2004 г. и бр. 91 от 14 октомври 2004 г.")
        AddParagraph(wordDoc, "3. НАРЕДБА № Iз-1971 от 29 октомври 2009 г. за строително-технически правила и норми за осигуряване на безопасност при пожар, Обн. ДВ., бр. 96 от 4 декември 2009 г.")
        AddParagraph(wordDoc, "4. НАРЕДБА № 14 от 15.06.2005 г. за технически правила и нормативи за проектиране, изграждане и ползване на обектите и съоръженията за производство, преобразуване, пренос и разпределение на електрическа енергия.")
        AddParagraph(wordDoc, "5. ПРАВИЛНИК за безопасност и здраве при работа в електрически уредби на електрически и топлофикационни централи и по електрически мрежи (Загл. изм. - ДВ, бр.19 от 2005.")
        AddParagraph(wordDoc, "6. Наредба № 16-116 от 8 февруари 2008 г. за техническата експлоатация на енергийните съоръжения, Обн. ДВ., бр. 26 от 7 март 2008 г.")
        AddParagraph(wordDoc, "7. Наредба № 1 от 27 май 2010 г. за проектиране, изграждане и поддържане на електрически уредби НН в сгради, Обн. ДВ., бр. 46 от 18 юни 2010 г.")
        AddParagraph(wordDoc, "8. НАРЕДБА № 7 от 23.09.1999 г. за минималните изисквания за здравословни и безопасни условия на труд на работните места и при използване на работното оборудване.")
        AddParagraph(wordDoc, "9. НАРЕДБА № 16 от 09.06.2004 г. за сервитутите на енергийните обекти.")
        AddParagraph(wordDoc, "10. Наредба № 4 от 22 декември 2010 г. за мълниезащитата на сгради, външни съоръжения и открити пространства, Обн. ДВ., бр. 6 от 18 януари 2011 г.")
        AddParagraph(wordDoc, "")
        Text = "Производителят на електрическа енергия в границите на имота си ще въведе в експлоатация необходимата електрическа уредба на генераторно напрежение със следните показатели:"
        AddParagraph(wordDoc, Text, False)
        With wsFEC_Таблица
            AddParagraph(wordDoc, .Cells(1, 1).Value, False)
            AddParagraph(wordDoc, .Cells(2, 1).Value, False)
            AddParagraph(wordDoc, .Cells(3, 1).Value, False)
            AddParagraph(wordDoc, .Cells(4, 1).Value, False)
            AddParagraph(wordDoc, .Cells(5, 1).Value, False)
            AddParagraph(wordDoc, .Cells(6, 1).Value, False)
            AddParagraph(wordDoc, .Cells(7, 1).Value, False)
            dicObekt.Add("Площ", .Cells(1, 1).Value)
            dicObekt.Add("Тегло", .Cells(2, 1).Value)
            dicObekt.Add("Височина", .Cells(3, 1).Value)
            dicObekt.Add("Инсталирана", .Cells(4, 1).Value)
            dicObekt.Add("Изходна", .Cells(5, 1).Value)
            dicObekt.Add("Панели", .Cells(6, 1).Value)
            dicObekt.Add("Инвертори", .Cells(7, 1).Value)
        End With
        Dim Text_Пожарна As String = ""
        FormatParagraph(wordDoc, "ФОТОВОЛТАИЧНА ЦЕНТРАЛА", wordApp)
        Text = "Настоящият технически проект обхваща проектирането на фотоволтаична централа с обща инсталирана мощност - "
        Text += String.Format("{0:0.000}", wsFEC_ФЕЦ.Cells(2, 2).Value)
        Text += " кWр и ОБЩА НОМИНАЛНА ИЗХОДНА МОЩНОСТ: "
        Dim Pinv As Double = 0
        Dim Борой_типове_инверотри As Integer = 0
        For i = 27 To 36
            If IsNumeric(wsFEC_Таблица.Cells(i, 3).Value) Then
                Pinv += wsFEC_Таблица.Cells(i, 3).Value * wsFEC_Таблица.Cells(i, 5).Value
                Борой_типове_инверотри += 1
            End If
        Next
        Text += String.Format("{0:0.0}", Pinv)
        Text += "кW."
        AddParagraph(wordDoc, Text, False)
        Text_Пожарна = Text
        Text = "Номиналната мощност ще се реализира чрез "
        Dim values As New List(Of String)
        For i = 28 To 37
            If wsFEC_Таблица.Cells(i, 5).Value <= 0 Then
                Continue For
            End If
            Dim countValue As Integer = wsFEC_Таблица.Cells(i, 5).Value
            Dim textValue As String = String.Format("{0:0} бр. ", countValue)
            textValue += IIf(countValue = 1, "трифазен многострингов инвертор тип ", "трифазни многострингови инвертори тип ")
            textValue += wsFEC_Таблица.Cells(i, 1).Value
            values.Add(textValue)
        Next
        Dim Text_PIC As String = ""
        If values.Count > 0 Then
            If values.Count = 1 Then
                Text += values(0)
                Text_PIC = values(0)
            Else
                Text += String.Join(", ", values.Take(values.Count - 1)) & " и " & values.Last()
                Text_PIC = String.Join(", ", values.Take(values.Count - 1)) & " и " & values.Last()
            End If
        End If
        dicObekt.Add("Text_PIC", Text_PIC)
        Text += "."
        Text_Пожарна += vbCrLf + Text
        AddParagraph(wordDoc, Text, False)
        Text = "В проекта е предвидено централата да се състои от "
        values.Clear()  ' Изчистваме предишното съдържание на списъка
        For i = 14 To 22
            If IsNumeric(wsFEC_Таблица.Cells(i, 5).Value) Then
                Dim countValue As Integer = wsFEC_Таблица.Cells(i, 5).Value
                Dim textValue As String = String.Format("{0:0} бр. монокристални соларни панели тип ", countValue)
                textValue += wsFEC_Таблица.Cells(i, 1).Value
                values.Add(textValue)
            End If
        Next
        If values.Count > 0 Then
            If values.Count = 1 Then
                Text += values(0)
            Else
                Text += String.Join(", ", values.Take(values.Count - 1)) & " и " & values.Last()
            End If
        End If
        Text += ". Общата инсталирана мощност на соларните панели е "
        Text += String.Format("{0:0.000}", wsFEC_ФЕЦ.Cells(2, 2).Value) + " кWр."
        dicObekt.Add("PV_Мощност", wsFEC_ФЕЦ.Cells(2, 2).Value)
        Text_Пожарна += vbCrLf + Text
        Text += " Параметрите на соларните панели са дадени в приложение към проекта."
        AddParagraph(wordDoc, Text, False)
        Text = "След монтажа панелите ще са неподвижни, поставени на конструкции, подходящи за монтаж на панели"
        Try
            pDouOpts.Keywords.Clear()
            With pDouOpts
                .Keywords.Add("Земя")
                .Keywords.Add("Покрив")
                .Keywords.Add("Двете")
                .Message = vbCrLf & "Изберете начина на монтаж на панелите : "
                .AllowZero = False
                .AllowNegative = False
            End With

            Dim pKeyRes As PromptDoubleResult = acDoc.Editor.GetDouble(pDouOpts)
            If pKeyRes.Status = PromptStatus.Keyword Then
                Dylbo = pKeyRes.StringResult
            Else
                Dylbo = pKeyRes.Value.ToString()
            End If
            Select Case Dylbo
                Case "Земя"
                    Text += " на земен терен."
                Case "Покрив"
                    Text += " върху порив."
                Case "Двете"
                    Text += ". Част от панелите ще се монтират на земен терен, а друга част ще се монтират на покрив."
            End Select
            dicObekt.Add("Монтаж", Dylbo)
        Catch ex As Exception

        End Try
        Text_Пожарна += vbCrLf + Text
        dicObekt.Add("Земя", Text_Пожарна)
        AddParagraph(wordDoc, Text, False)
        Dim Мълния = cu.GetObjects("INSERT", "Изберете блок който e панел ", False)
        If Мълния Is Nothing Then
            MsgBox("Няма маркиран нито един блок който e панел.")
        Else
            Dim Наклон_Ширина As String = ""
            Dim Наклон_Дължина As String = ""
            Dim Азимут As Double = 0

            Try
                Using actrans As Transaction = acDoc.TransactionManager.StartTransaction()

                    Dim blkRecId = Мълния(0).ObjectId
                    Dim acBlkRef As BlockReference =
                        DirectCast(actrans.GetObject(blkRecId, OpenMode.ForWrite), BlockReference)
                    Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection
                    Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                    ' Обхождане на всички атрибути
                    For Each objID As ObjectId In attCol
                        Dim dbObj As DBObject = actrans.GetObject(objID, OpenMode.ForRead)
                        Dim acAttRef As AttributeReference = dbObj
                        'If acAttRef.Tag = "ТАБЛО" Then Азимут = acAttRef.TextString
                        Азимут = acAttRef.Rotation
                    Next
                    For Each prop As DynamicBlockReferenceProperty In props
                        'This is where you change states based on input
                        ' Разстояние_Профили Наклон_Дължина Наклон_Ширина Ширина Дължина Профил d1 ang1
                        ' Отстояние d2 d3 d4 d5 d6 d7 d8 Visibility1 d9 Проекция d11 d22 d21 d14 d13 d12 d15 d10
                        If prop.PropertyName = "Наклон_Ширина" Then Наклон_Ширина = prop.Value
                        If prop.PropertyName = "Наклон_Дължина" Then Наклон_Дължина = prop.Value
                    Next
                    actrans.Commit()
                End Using
            Catch ex As Exception
                MsgBox("Възникна грешка " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
            End Try

            Text = "Наклона на фотоволтаичните модули спрямо хоризонталната повърхност (β) е "
            Text += Наклон_Дължина
            Text += "°"

            Азимут = Азимут * 180 / PI
            Азимут = (Азимут - 90) Mod (360)
            Text += " за всичките монокристални соларни панели Азимута е "
            Text += String.Format("{0:0.0}", Азимут)
            Text += "°."
            AddParagraph(wordDoc, Text, False)
        End If
        FormatParagraph(wordDoc, "МРЕЖОВИ МНОГОСТРИНГОВИ ИНВЕРТОРИ", wordApp, level:=2)
        With Text
            Text = "Трифазните мрежови многострингови инвертори ще бъдат монтирани на местата, отразени в графичната част."
            Text += " В непосредствена близост до инверторите, ще се монтират и постояннотоковите DC табла."
            Text += " Начинът на свързване е показан на графичната част на проекта."
            AddParagraph(wordDoc, Text, False)

            Text = "Основните параметри на инверторите са както следва:"
            Text = "Тип"
            Text = "Mощност"
            Text = "Оптимално напрежение"
            Text = "Максимално напрежение"
            Text = "Максимален ток за MPPT"
            Text = "Максимален ток късо MPPT"
            Text = "Брой входове"
            Text = "Брой MPPT тракери"

            Text = "Всички параметри на инверторите са дадени в приложение към проекта."
            AddParagraph(wordDoc, Text, False)
            Text = "Като основен принцип на работа на инвертора е системата за следене на мощността."
            Text += " Тя интегрира, контролира и следи генерираната от PV генератора мощност и ако тя е достатъчна - инверторът започва да отдава електроенергия към мрежата."
            Text += " Максимално осигурената от PV панелите мощност постъпваща на входа на инвертора."
            Text += " Спрямо моментното ниво на слънчевата радиация и околната температура генерираната мощност се контролира и поддържа на максимално ниво в работната си точка от V-A характеристика."
            AddParagraph(wordDoc, Text, False)
            Text = "Системата, осигуряваща работата на инвертора в най-високата оптимална работна точка се нарича MРР Тракер, който е съставна част от инверторното устройство."
            Text += " Предвидените в проекта инвертори имат по 4/10бр. МРР тракери."
            AddParagraph(wordDoc, Text, False)
            Text = "Синхронизирането на фотоволтаичните генератори и мрежата НН-0.4кV се извършва чрез инверторите."
            Text += " Когато радиацията върху слънчевите модули падне под минималния праг, инверторите спират да функционират."
            Text += " Производителят предоставя всички необходими гаранции и сертификати, в съответствие с действащите стандарти и норми за безопасност."
            AddParagraph(wordDoc, Text, False)
            Text = "Инверторите ще се самоизключат в следните случаи:"
            AddParagraph(wordDoc, Text, False)
            Text = "Авария в електрическата мрежа: При прекъсване на електрозахранването инверторът ще се самоизключи и няма да работи в изолиран режим (островен режим). Работата му се възстановява автоматично след възстановяване на мрежовото напрежение."
            AddParagraph(wordDoc, Text, False)
            Text = "Отклонения в напрежението: Ако напрежението излезе извън номиналния диапазон, инверторът автоматично ще се самоизключи и ще възобнови работата си само при нормално мрежово напрежение."
            AddParagraph(wordDoc, Text, False)
            Text = "Отклонения в честотата: Ако честотата на мрежата се отклони от номиналния диапазон, инверторът автоматично ще се самоизключи и ще възстанови работата си при нормализиране на мрежовата честота."
            AddParagraph(wordDoc, Text, False)
            Text = "Висока температура: Инверторът разполага с естествена система за охлаждане чрез свободна конвекция. Ако температурата във вътрешността му се повиши над определена стойност, инверторът ще намали изходната мощност. Ако температурата достигне критични нива, инверторът ще се самоизключи и ще възстанови работата си при нормализиране на температурата."
            AddParagraph(wordDoc, Text, False)
        End With
        FormatParagraph(wordDoc, "СОЛАРНИ ПАНЕЛИ И СОЛАРНИ КАБЕЛИ", wordApp, level:=2)
        With Text
            Text = "Соларните панели представляват полупроводници, които преобразуват светлинната енергия в електрическа чрез фотоволтаични клетки."
            Text += " Не е необходимо тези клетки да са изложени на директна слънчева светлина, за да могат да работят – дори в облачен ден те все още са способни да генерират електричество. Мощността на фотоволтаичната клетка се измерва в киловат пик (kWp)."
            Text += " Това е скоростта, с която тя генерира енергия при върхови резултати под директна слънчева светлина през лятото."
            AddParagraph(wordDoc, Text, False)
            Text = "В проекта е предвидено централата да се състои от "
            For i = 13 To 22
                If IsNumeric(wsFEC_Таблица.Cells(i, 5).Value) Then
                    Text += IIf(IsNumeric(wsFEC_Таблица.Cells(i + 1, 5).Value), ", ", " и ")
                    Text += String.Format("{0:0}", wsFEC_Таблица.Cells(i, 5).Value) + " бр."
                    Text += " монокристални соларни панели тип "
                    Text += wsFEC_Таблица.Cells(i, 1).Value
                End If
            Next
            Text += ". Общата инсталирана мощност на соларните панели е "
            Text += String.Format("{0:0.000}", wsFEC_ФЕЦ.Cells(2, 2).Value) + " кWр."
            Text += "  Параметрите на соларните панели са дадени в приложение към проекта."
            AddParagraph(wordDoc, Text, False)
            Text = "Соларните панели са готови изделия, на които са монтирани кутии за осъществяване на електрическа връзка помежду им."
            Text += " Електрическите връзки между отделните панели, за всеки стринг, се осъществяват чрез соларни кабели с медни жила Solar Cable 1/1.5kV 1x6mm²."
            Text += " Соларни кабели се полагат открито по носещите конструкции."
            AddParagraph(wordDoc, Text, False)
            Text = "Кабелите, използвани за постояннотоково окабеляване на фотоволтаичната система, отговарят на изискванията по отношение на изолацията, устойчивостта на ултравиолетови лъчи, вода, атмосферни влияния и висока температура. "
            Text += " Сечението на кабелите е определено по икономична плътност, допустимо нагряване, термична устойчивост и допустима загуба на напрежение."
            AddParagraph(wordDoc, Text, False)
            Text = "При последователно свързване един към друг се включват плюса и минусът на два отделни модула чрез фабрично монтираните соларни кабели и конектори."
            Text += " При преминаване на по-далечни разстояния от фабричната дължина свързването става чрез допълнителни кабелни връзки Solar Cable 1/1.5kV 1x6mm² и конектори."
            Text += " Кабелите за плюса и минуса на всеки стринг да са разположени максимално близо един до друг, и така вървят до входа на DC таблото и инверторите."
            Text += " Кабелите Solar Cable 1/1.5kV 1x6mm² от всички стрингове се укрепват по конструкцията с ленти от синтетичен материал за УИП 9х180мм."
            AddParagraph(wordDoc, Text, False)
            Text = " Избраните конектори да поддържат постоянно ниско съпротивление и да не допускат натрупването на загуби в отделните връзки."
            Text += " Свързването на кабелите се осъществява бързо и сигурно, като същевременно се осигурява необходимата защита."
            Text += " Сред работните им характеристики са допустимо напрежение, допустим ток, контактно съпротивление, клас на защита от атмосферни влияния, работна температура и клас на безопасност."
            Text += " Присъединяването на кабелите към конекторите чрез кримпване."
            Text += " Същността на кримпваната връзка се изразява в едновременно запресоване на двата елемента: проводник и метална втулка."
            Text += " Кримпването се осъществява с помощта на специален инструмент. Връзката е корозионно- и виброустойчива."
            Text += " След кримпването, контактът се въвежда в контактното тяло."
            AddParagraph(wordDoc, Text, False)
            Text = "Кабелите от фотоволтаичните панели до инверторите се свързват в DC табла, монтирани в непосредствена близост до инверторите."
            Text += " DC табла да съдържат и DC прекъсвач, който дава възможност инверторът да бъде изключен от постояннотоковата страна."
            'Text += " DC таблата да се монтира възможно най-близо до модулите, за да се спести дължината на кабелите и да се осигури оптимална работа на защитата от пренапрежение."
            Text += " Кабелите от DC таблата до инверторите са оразмерени за пълния ток на късо съединение на фотоволтаичната група."
            Text += " Начинът на свързване е показан на графичната част на проекта."
            AddParagraph(wordDoc, Text, False)
        End With
        FormatParagraph(wordDoc, "КОНФИГУРАЦИЯ НА СТРИНГОВЕТЕ", wordApp, level:=2)
        With Text
            Text = "Фотоволтаичните модули преобразуват постъпващата от слънчевата радиация енергия в електрическа. Генерираната от ФВМ (фотоволтаични модули) електрическа енергия е с постояннотокови (DC) параметри: напрежение(U) и ток (I). Тя, посредством мрежа от електрически връзки, комутационни апарати, защити и кабели се подава към инвертора за преобразуване в променливотокова (АС) енергия."
            AddParagraph(wordDoc, Text, False)
            Text = "Фотоволтаичното поле се формира чрез последователно свързване на фотоволтаичните модули в стрингове, за повишаване на системното напрежение."
            Text += " Стринговете се свързват паралелно, образувайки DC клонове с цел увеличаване на изходната мощност."
            Text += " Последователното включване на модулите в стринг се извършва чрез конекторите на фабричните кабели, комплект с модулите."
            Text += " Броят на последователно свързаните модули в един стринг се определя от взаимоотношенията на техническите параметри на ФВМ и инвертора, към които се свързват."
            Text += " Начинът на свързване е показан на графичната част на проекта."
            AddParagraph(wordDoc, Text, False)
            Text = " За конкретните панели и инвертори е определена следната оптималната комбинация:"
            AddParagraph(wordDoc, Text, False)
            Dim br_Grupi As Integer = 1
            With wsFEC_ФЕЦ
                For i = 3 To 51 Step 5

                    If .Cells(8, i).Value = "(НЕ)" Then
                        Continue For
                    End If

                    Text = "ГРУПА " + String.Format("{0:0}", br_Grupi)
                    AddParagraph(wordDoc, Text, True)

                    Text = "Групата включва "
                    Text += IIf(.Cells(6, i).Value > 1, String.Format("{0:0}", .Cells(5, i).Value), "")
                    Text += IIf(.Cells(6, i).Value > 1, "бр. инвертори", "инвертор")
                    Text += " тип "
                    Text += .Cells(26, i).Value.ToString
                    Text += IIf(.Cells(6, i).Value > 1,
                                ", с номера ",
                                ", с номер ")
                    ' Проверка дали клетката съдържа нещо, което не е число
                    If Not IsNumeric(.Cells(25, i).Value) Then
                        ' Ако стойността не е число, прекратяваме цикъла
                        Text += .Cells(25, i).Value
                    Else
                        Text += .Cells(25, i).Value.ToString
                    End If
                    Text += IIf(.Cells(6, i).Value > 1,
                                ". На всеки инвертор в групата са монтирани ",
                                ". На инвертора са монтирани ")
                    Text += String.Format("{0:0}", .Cells(34, i).Value)
                    Text += " бр. панели тип "
                    Text += .Cells(8, i).Value
                    Text += IIf(.Cells(6, i).Value > 1,
                                ". Общата мощност на панелите за всеки инвертор е ",
                                ". Общата мощност на панелите за инвертора e ")
                    Text += String.Format("{0:0.000}", .Cells(35, i).Value)
                    Text += " kWp."

                    If .Cells(6, i).Value > 1 Then
                        Text += " Общата генерирана, от панелите, за група "
                        Text += String.Format("{0:0}", br_Grupi)
                        Text += " мощност е "
                        Text += String.Format("{0:0.000}", .Cells(35, i).Value)
                        Text += " kWp."
                    End If
                    Text += " Оптималното напрежение за MPPT е "
                    Text += String.Format("{0:0}", .Cells(28, i).Value)
                    Text += " V, а максималното допустимо напрежение за MPPT е "
                    Text += String.Format("{0:0}", .Cells(29, i).Value)
                    Text += " V"
                    Text += ". Според броя на панелите в един стринг, максималното напрежение при температура -15°С ще бъде "
                    Text += String.Format("{0:0}", .Cells(29, i + 3).Value)
                    Text += " V."

                    AddParagraph(wordDoc, Text, False)
                    Text = "Разпределението на панелите по стрингове е както следва: "
                    AddParagraph(wordDoc, Text, False)

                    ' Създаваме речник за съхранение на стойностите и техните броеве
                    'Dim numberCounts As New Dictionary(Of String, Integer)()
                    Dim numberCounts As New Dictionary(Of String, Tuple(Of Integer, List(Of Double), String))()

                    ' Обхождане на клетките от ред 37 до 57 в колона I (9-та колона)
                    For j As Integer = 38 To 58
                        ' Проверка дали клетката съдържа нещо, което не е число
                        If Not IsNumeric(.Cells(j, i).Value) Then
                            ' Ако стойността не е число, прекратяваме цикъла
                            Continue For
                        End If

                        ' Извличане на стойността от клетката
                        Dim cellValue As String = .Cells(j, i).Value.ToString()
                        ' Проверка дали стойността е двуцифрено число
                        If cellValue.Length = 2 Then
                            ' Проверка дали числото вече съществува в речника
                            If numberCounts.ContainsKey(cellValue) Then
                                ' Ако съществува, увеличаваме броя на срещанията с 1
                                Dim currentTuple = numberCounts(cellValue)
                                numberCounts(cellValue) = Tuple.Create(currentTuple.Item1 + 1, currentTuple.Item2, currentTuple.Item3)
                            Else
                                ' Ако не съществува, добавяме го с начална стойност 1 и стойностите от колони 10, 11 и 12
                                Dim valuesList As New List(Of Double) From {
                                    .Cells(j, i + 1).Value, ' Стойност от колона 10
                                    .Cells(j, i + 2).Value  ' Стойност от колона 11
                                }
                                Dim column12Value As String = .Cells(j, i + 3).Value.ToString() ' Стойност от колона 12 като низ
                                numberCounts(cellValue) = Tuple.Create(1,
                                                                       valuesList,
                                                                       If(column12Value.Contains("/"),
                                                                       column12Value.Split("/"c)(0),
                                                                       column12Value))
                            End If
                        End If
                    Next
                    For Each pair In numberCounts
                        Text = "- "
                        Text += pair.Value.Item1.ToString() ' Брой стрингове
                        If pair.Value.Item1 = 1 Then
                            Text += " стринг "
                        Else
                            Text += " броя стрингове "
                        End If
                        Text += "с последователно свързани "
                        Text += pair.Key ' Брой панели в стринга
                        Text += " броя панели."
                        ' Добавяне на информация от клетките 10, 11 и 12
                        Text += " Параметри на стринга: напрежение при максимална мощност Vmp = " & pair.Value.Item2(0).ToString() & " V"
                        Text += "; "
                        Text += " напрежение на отворена верига Voc = " & pair.Value.Item3 & " V"
                        Text += "; "
                        Text += " максимална мощност, която генерира един стринг Мощност = " & pair.Value.Item2(1).ToString() & " Wp"
                        Text += "."
                        AddParagraph(wordDoc, Text, False)
                    Next
                    br_Grupi += 1
                Next
            End With
            Dim texts As String() = {
                "Информация за параметрите",
                "Описание на параметрите",
                "Характеристики на параметрите",
                "Разяснение на параметрите",
                "Детайлна информация за параметрите",
                "Пояснения за параметрите"
            }
            Dim random As New Random()
            Dim randomIndex As Integer = random.Next(0, texts.Length)
            AddParagraph(wordDoc, texts(randomIndex), True)
            Text = "Vmp(Voltage at Maximum Power) : Това е напрежението, при което соларният стринг генерира максимална мощност."
            AddParagraph(wordDoc, Text, False)
            Text = "Мощност : Тази стойност показва максималната мощност (във ватове), която може да бъде генерирана от един стринг при оптимални условия."
            AddParagraph(wordDoc, Text, False)
            Text = "Voc(Open Circuit Voltage) : Напрежението на отворена верига е максималното напрежение, което се получава, когато няма товар, свързан към соларния стринг."
            AddParagraph(wordDoc, Text, False)
        End With
        Dim docManager As DocumentCollection = Application.DocumentManager
        ' Преминаваме през всеки отворен документ
        For Each doc As Document In docManager
            ' Печатаме пътя и името на файла
            Dim acfilename As String = doc.Name
            If String.IsNullOrEmpty(acfilename) Then
                acfilename = "Нов (неименуван) документ"
            End If
            Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbCrLf & "Отворен файл: " & acfilename)
        Next
        excel_Workbook_FEC.Close()
        excel_Workbook_FEC = Nothing
        FormatParagraph(wordDoc, "МОНТАЖ НА СОЛАРНИ ПАНЕЛИ, СОЛАРНИ КАБЕЛИ И МРЕЖОВИ МНОГОСТРИНГОВИ ИНВЕРТОРИ", wordApp, level:=1)
        With Text
            Dim txtZemya As String = ""
            txtZemya = "Конструкцията, върху която се монтират фотоволтаичните модули, е с клас по реакция на огън А2."
            txtZemya += " Отстоянието на фотоволтаични модули до сгради, постройки и съоръжения е не по-малко от 4 m."
            txtZemya += " В зоната на монтаж на соларните панели няма надземни газопроводи за ГГ и надземни тръбопроводи за ЛЗТ и ГТ."
            txtZemya += vbCrLf
            txtZemya += "До входа на площадката на фотоволтаичната централа е осигурен път за противопожарни цели при спазване изискванията на чл. 27, ал. 3, 7, 8 и 9 от НАРЕДБА № Iз-1971 от 29 октомври 2009 г. за строително-технически правила и норми за осигуряване на безопасност при пожар"
            txtZemya += vbCrLf
            txtZemya += "В близост до входа на фотоволтаичната централа в предвидено устройство за ръчно прекъсване на веригите за постоянен ток и за променлив ток на фотоволтаичната електрическа централа."
            txtZemya += " Конкретния начин за реализация на това изключване е отразено в графичната част на проекта."
            txtZemya += " Местоположението на устройството за ръчно прекъсване на веригите е изрично указано."
            txtZemya += vbCrLf
            txtZemya += "След въвеждане в експолатация в специално табло да се постави документация с информация за фотоволтаичната електрическа централа, определена в Наредба № 8121з-647 от 2014 г. за правилата и нормите за пожарна безопасност при експлоатация на обектите."
            Dim txtPokriv As String = ""

            txtPokriv = "Фотоволтаичните модули, монтирани върху покривите на сградите и постройките, са разположени така, че да не се нарушава целостта на покривната конструкция и изолационните слоеве." + vbCrLf
            txtPokriv += "При покриви с топлинна изолация са използвани материали с подходящ клас по реакция на огън, гарантиращи безопасност при възникване на пожар, като същевременно се запазват топлотехническите характеристики на покрива." + vbCrLf
            txtPokriv += "При покриви без външна топлинна изолация покривното покритие и хидроизолацията са също с необходимия клас по реакция на огън, съобразен с експозицията на външно огнево въздействие." + vbCrLf
            txtPokriv += "При монтаж на вградени (интегрирани) модули топлоизолацията и пароизолацията на покрива се предвиждат с необходимия клас по реакция на огън, осигуряващ безопасна експлоатация на системата." + vbCrLf
            txtPokriv += "Фотоволтаичните модули, предвидени за монтаж върху покривите на сградите и постройките, се разполагат на височина над покривното покритие не по-малка от минималното разстояние, определено от производителя, с цел осигуряване на необходимото им естествено охлаждане." + vbCrLf
            txtPokriv += "Отстоянието от фотоволтаичните модули до димни люкове, капандури, комини, шахти, куполи и други елементи, излизащи над покривната повърхност, е предвидено не по-малко от 1,0 m." + vbCrLf
            txtPokriv += "Не се допуска разполагане на фотоволтаични модули върху брандмауери и над вертикални пожарозащитни прегради, като отстоянието до такива прегради е предвидено не по-малко от 1, 0 m." + vbCrLf
            txtPokriv += "Не се предвижда монтаж на модули в огнеустойчиви фасадни участъци или на разстояние по-малко от 1, 0 m от евакуационни изходи." + vbCrLf
            If dicObekt("PV_Мощност") > 20 Then
                txtPokriv += "Модулите са разположени на минимално разстояние не по-малко от 1,0 m от краищата (контура) на покрива, като оформят отделни модулни полета с размери, ненадвишаващи 40 m х 40 m. Между отделните полета са осигурени технологични проходи с широчина не по-малка от 1,50 m, позволяващи безопасен достъп, обслужване и намеса при аварийни ситуации." + vbCrLf
            Else
                txtPokriv += "При разполагане на модулите са осигурени отстояния не по-малки от 0,50 m от краищата на покрива и не по-малки от 0,10 m от надпокривни елементи като комини, капандури и димни люкове." + vbCrLf
                txtPokriv += "Монтажната конструкция за закрепване на фотоволтаичните модули е изпълнена от материали с клас по реакция на огън не по-нисък от А2, гарантиращи устойчивост при пожар и дълготрайна експлоатация." + vbCrLf
                txtPokriv += "Инверторите на фотоволтаичната централа са предвидени за монтаж върху негорими конструкции на височина не по-малка от 0, 30 m над покрива/терена, като са спазени необходимите отстояния не по-малки от 1,0 m от конструктивни елементи и строителни продукти с по-нисък клас по реакция на огън." + vbCrLf
            End If
            txtPokriv += "Кабелните трасета за постоянен ток, преминаващи през сградата, са проектирани с повишени изисквания за пожарна безопасност. Кабелите са положени в защитни тръби или канали с минимален клас по реакция на огън А2, като алтернативно са предвидени защитни строителни продукти, осигуряващи минимална огнеустойчивост EI 30 по цялата дължина на преминаване през конструктивни елементи и помещения на сградата." + vbCrLf
            txtPokriv += "В непосредствена близост до входа на сградата е предвидено табло с документация, съдържаща информация за местоположението на фотоволтаичните модули, инверторите, прекъсващите устройства и други основни елементи на системата." + vbCrLf
            txtPokriv += "До всеки вход на сградата е предвидено поставяне на предупредителен знак, указващ наличието на фотоволтаична електрическа централа, съгласно изискванията за пожарна безопасност." + vbCrLf

            Select Case dicObekt("Монтаж")
                Case "Земя"
                    Text = "Фотоволтаичните модули ще се монтират на земен терен."
                    AddParagraph(wordDoc, txtZemya, True)
                    AddParagraph(wordDoc, txtZemya, False)
                Case "Покрив"
                    Text = "Фотоволтаичните модули ще се монтират върху покрив."
                    AddParagraph(wordDoc, Text, True)
                    AddParagraph(wordDoc, txtPokriv, False)
                Case "Двете"
                    Text = "Част от фотоволтаичните модули ще се монтират на земен терен."
                    AddParagraph(wordDoc, Text, True)
                    AddParagraph(wordDoc, txtZemya, False)
                    Text = "Част от фотоволтаичните модули ще се монтират върху покрив."
                    AddParagraph(wordDoc, Text, True)
                    AddParagraph(wordDoc, txtPokriv, False)
            End Select
        End With
    End Sub
    Sub Get_data_PIC(dicObekt As Dictionary(Of String, String),
                     acDoc As Document,
                     acCurDb As Database)
        Dim Помещение_ИНВ As String = ""
        Do
            Помещение_ИНВ = cu.GetObjects_TEXT("Изберете текст съдържащ помещението, в което се МОНТИРАТ ИНВЕРТОРИТЕ")
            If Not Помещение_ИНВ.Contains("#####") Then
                Exit Do
            End If
            Dim repeatChoice As MsgBoxResult = MsgBox(
                    Title:="Няма маркиран текст съдържащ помещението, в което се МОНТИРАТ ИНВЕРТОРИТЕ!",
                    Buttons:=MsgBoxStyle.YesNo,
                    Prompt:="Да повторя ли избора?"
                    )
            If repeatChoice = MsgBoxResult.No Then
                dicObekt.Add("Помещение_ИНВ", "#####")
                dicObekt.Add("Кота_ИНВ", "#####")
                dicObekt.Add("СРТ_Помещение_PIC", "#####")
                dicObekt.Add("СРТ_Кота_PIC", "#####")
                Exit Sub
            End If
        Loop
        dicObekt.Add("Помещение_ИНВ", Помещение_ИНВ)
        Dim Кота_ИНВ = cu.GetObjects_TEXT("Изберете текст съдържаш котата на която се МОНТИРАТ ИНВЕРТОРИТЕ")
        dicObekt.Add("Кота_ИНВ", Кота_ИНВ)
        Dim СРТ_Помещение_PIC = cu.GetObjects_TEXT("Изберете текст съдържаш помещението в което се намира на ПИЦ/ДАТЧИКА")
        dicObekt.Add("СРТ_Помещение_PIC", СРТ_Помещение_PIC)
        Dim СРТ_Кота_PIC = cu.GetObjects_TEXT("Изберете текст съдържаш котата на която се намира на ПИЦ/ДАТЧИКА")
        dicObekt.Add("СРТ_Кота_PIC", СРТ_Кота_PIC)


        Dim blkRecId As ObjectId = ObjectId.Null

        Dim ss_Tabla = cu.GetObjects("INSERT", "Изберете БЛОКОВЕТЕ в чертеж съдържащи ДАТЧИЦИТЕ ОТ ПИЦ:")
        If ss_Tabla Is Nothing Then
            MsgBox("Няма маркиран нито един блок.")
            Exit Sub
        End If
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                For Each sObj As SelectedObject In ss_Tabla
                    blkRecId = sObj.ObjectId
                    Dim acBlkRef As BlockReference = DirectCast(acTrans.GetObject(blkRecId, OpenMode.ForRead), BlockReference)

                    ' Проверка дали блокът е динамичен
                    If Not acBlkRef.IsDynamicBlock Then Continue For

                    Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                    Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection
                    Dim picItem As New PIC()

                    ' Извличане на свойството "Visibility" на динамичния блок
                    For Each prop As DynamicBlockReferenceProperty In props
                        If prop.PropertyName = "Visibility" Then picItem.Visibility = prop.Value
                    Next
                    Dim blName As String = (CType(acBlkRef.DynamicBlockTableRecord.GetObject(OpenMode.ForRead), BlockTableRecord)).Name
                    If blName <> "Табло_Ново" And
                        blName <> "Датчик_ПАБ" Then Continue For

                    Dim boVisibility As Boolean = False
                    For i As Integer = 0 To picList.Count - 1
                        If picList(i).Visibility = picItem.Visibility Then
                            boVisibility = True
                            Dim temp As PIC = picList(i) ' Копира елемента
                            temp.CountS += 1             ' Променя стойността
                            picList(i) = temp            ' Връща променения елемент обратно
                            Exit For
                        End If
                    Next
                    If boVisibility Then Continue For

                    For Each objID As ObjectId In attCol
                        Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                        Dim acAttRef As AttributeReference = dbObj
                        If acAttRef.Tag = "ТАБЛО" Then picItem.Tablo = acAttRef.TextString
                        If acAttRef.Tag = "ZN" Then picItem.ZN = acAttRef.TextString
                        If acAttRef.Tag = "NOM" Then picItem.NOM = acAttRef.TextString
                        If acAttRef.Tag = "AD" Then picItem.AD = acAttRef.TextString
                    Next
                    picItem.CountS = 1
                    ' Добавяне в списъка
                    picList.Add(picItem)

                Next
                acTrans.Commit()
            Catch ex As Exception
                MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
                acTrans.Abort()
            Finally
                If acTrans IsNot Nothing Then acTrans.Dispose()
            End Try

        End Using
    End Sub
    Sub Записка_ПИЦ(wordDoc As Word.Document,
                         acDoc As Document,
                         acCurDb As Database,
                         dicObekt As Dictionary(Of String, String))
        Dim Text As String = ""

        Get_data_PIC(dicObekt, acDoc, acCurDb)

        If dicObekt("Помещение_ИНВ").Contains("#####") Then
            Exit Sub
        End If

        wordDoc.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak)
        With wordDoc.Content.Paragraphs.Add
            .Range.Text = "ОБЯСНИТЕЛНА ЗАПИСКА"
            .Range.Font.Name = "Cambria"
            .Range.Font.Size = 20
            .Range.Font.Bold = True
            .Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            .Range.InsertParagraphAfter()
        End With
        With wordDoc.Content.Paragraphs.Add
            .Range.Text = "пожароизвестителна система и система за звукова сигнализация"
            .Range.Font.Name = "Cambria"
            .Range.Font.Size = 14
            .Range.Font.Bold = True
            .Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            .Range.InsertParagraphAfter()
        End With
        AddParagraph(wordDoc, Text, False)
        Text = $"ОБЕКТ: {dicObekt("ОБЕКТ")}"
        AddParagraph(wordDoc, Text, True)
        If dicObekt("ОБЕКТ") <> dicObekt("МЕСТОПОЛОЖЕНИЕ") Then
            Text = $"МЕСТОПОЛОЖЕНИЕ: {dicObekt("МЕСТОПОЛОЖЕНИЕ")}"
            AddParagraph(wordDoc, Text, True)
        End If
        AddParagraph(wordDoc, " ", True)
        Dim Централа As String = ""
        Dim Присъединяване As Boolean = False
        For Each item In picList
            If item.Tablo IsNot Nothing Then
                Присъединяване = True
                Exit For
            End If
        Next
        For Each item In picList
            If item.Tablo IsNot Nothing Then
                Continue For
            Else
                If item.Visibility.Contains("адресируем") Then
                    Централа = "адресируема"
                Else
                    Централа = "конвенционална"
                End If
                Exit For
            End If
        Next
        FormatParagraph(wordDoc, "ОБЩА ЧАСТ", wordApp, level:=1, resetLevel:=True)
        With Text
            AddParagraph(wordDoc, dicObekt("Продажба"), False)
            AddParagraph(wordDoc, "При разработката на пожароизвестителната система и системата за звукова сигнализация са спазени изискванията на следните нормативни документи:")
            AddParagraph(wordDoc, "- Наредба № Із-1971 за строително-техническите правила и норми за осигуряване на безопасност при пожар (ДВ бр. 96/2009 год.);")
            AddParagraph(wordDoc, $"– СД CEN/TS 54-14 {Кавички}Пожароизвестителни системи. Част 14: Указания за планиране, проектиране, монтиране, въвеждане в експлоатация, използване и поддържане{Кавички};")
            AddParagraph(wordDoc, "– Наредба № 2 за минималните изисквания за осигуряване на здравословни и безопасни условия на труд при извършване на строително-монтажни работи (ДВ бр. 37/2004 год.).")

            Text = $"За ранно откриване и известяване на евентуални признаци за възникване на пожар е предвидено да се"

            If Присъединяване Then
                Text += $" изгради нова {Централа} система за пожароизвестяване."
                AddParagraph(wordDoc, Text)
                Text = " Предвидените пожароизвестителната централа и съответните периферни устройства към нея отговарят на изискванията на СД CEN/TS 54 и трябва да притежават сертификат за съответствие."
                AddParagraph(wordDoc, Text)
            Else
                Text += " използва съществуващ "
                If Централа = "конвенционална" Then
                    Text += "конвенционален"
                Else
                    Text += "адресируем"
                End If
                Text += $" датчик монтиран в помещение {Кавички}{dicObekt("СРТ_Помещение_PIC")}{Кавички}{If(dicObekt("СРТ_Кота_PIC").Contains("#####"), "", $" на кота {dicObekt("СРТ_Кота_PIC")}")}."
                Text += " При този датчик ще се раздели съществуващият контур и ще се добавят новите периферни устройства."
                AddParagraph(wordDoc, Text)
            End If
            Text = "След консултации с възложителя е прието да не се изграждат съоръжения изпълняващи функция Е в EN 54-1 (функция предаване на сигнал за пожар), функция G в EN 54-1 (функция управление на пожарозащитни системи или на пожарозащитни устройства), функция М в EN 54-1 (функция за управление и индикация на гласовата система за пожароизвестяване)."
            Text += " Поради тази причина тези функции не са предвидени в настоящия проект."
            AddParagraph(wordDoc, Text)
            Text = "При тези консултации е прието да се реализира само функция С в EN 54-1 (функция пожароизвестяване)."
            Text += " Тази функция е предмет на настоящия проект."
            AddParagraph(wordDoc, Text)
        End With
        '
        'Описание на централта 
        '
        With Text
            If Присъединяване Then
                FormatParagraph(wordDoc, "ЦЕНТРАЛНО СЪОРЪЖЕНИЕ/CIE (ПОЖАРОИЗВЕСТИТЕЛНАТА ЦЕНТРАЛА)", wordApp, level:=1)
                AddParagraph(wordDoc, "Устройството за управление и индикация (CIE) трябва да съответства на EN 54-2 и трябва да притежава сертификат за съответствие с този стандарт.", Bold:=False)
                FormatParagraph(wordDoc, "Помещение, в които ще се монтира устройството за управление и индикация (CIE)", wordApp, level:=2)
                AddParagraph(wordDoc, "В обекта няма помещение с постоянно дежурство, поради което устройството за управление и индикация е разположено в помещение на партерния етаж, в близост до входа, който вероятно ще бъде използван от пожарната и спасителната служба.")
                AddParagraph(wordDoc, $"Помещението {Кавички}{dicObekt("СРТ_Помещение_PIC")}{Кавички}{If(dicObekt("СРТ_Кота_PIC").Contains("#####"), " ", $" на кота {dicObekt("СРТ_Кота_PIC")} ")}, в което ще се монтира {Централа} пожароизвестителната централа отговаря на следните изисквания:")
                AddParagraph(wordDoc, "• има нисък риск от пожар (малко топлинно натоварване и минимален риск от възпламеняване);")
                AddParagraph(wordDoc, "• е наблюдавано от най-малко един пожароизвестител в пожароизвестителната система;")
                AddParagraph(wordDoc, "• е чисто и сухо;")
                AddParagraph(wordDoc, "• няма риск от механична повреда на устройството;")
                AddParagraph(wordDoc, "• е достатъчно голямо и не ограничава лицата при обслужването на устройството за контрол и индикация;")
                AddParagraph(wordDoc, "• е достатъчно осветено, за да може ясно да се вижда визуалната сигнализация и")
                AddParagraph(wordDoc, "- могат лесно да се задействат елементите за управление;", FirstLine:=50)
                AddParagraph(wordDoc, "- могат лесно да бъдат прочетени всички инструкции или легенди;", FirstLine:=50)
                AddParagraph(wordDoc, "• не трябва да се предоставя допълнително изкуствено осветление;")
                AddParagraph(wordDoc, "• предвидено е аварийно осветление на пътищата за достъп, на устройството за управление и индикация и на панела за паралелна индикация;")
                AddParagraph(wordDoc, "• нивото на фоновия шум е такова, че да не възпрепятства чуваемостта на звуковата сигнализация.")
                AddParagraph(wordDoc, "Това помещение е подходящо място, откъдето може да бъде предприето първоначалното управление на пожара от персонала и/или от пожарната и спасителната служба.")

                FormatParagraph(wordDoc, "Разполагане на устройството за управление и индикация (CIE)", wordApp, level:=2)

                Text = "Устройството за управление и индикация ще бъде монтирано на стената в описаното по-горе помещение."
                Text += " Разположението му гарантира, че ъгълът на зрение осигурява лесна четливост за целия обслужващ и отговорен персонал."
                AddParagraph(wordDoc, Text)
                AddParagraph(wordDoc, "Не се предвижда устройството да се обслужва от лица със специални потребности, поради което не са включени специални мерки.")

                Text = "Височината на дисплея и индикациите ще бъде най-малко 1,4 m и най-много 1,8 m от нивото на готовия под."
                Text += " Устройството е разположено така, че всички индикации да са видими, а функциите за управление да могат да се задействат без необходимост от механични помощни средства."
                AddParagraph(wordDoc, Text)
                FormatParagraph(wordDoc, "Централното съоръжение (пожароизвестителната централа)", wordApp, level:=2)
                If Централа = "конвенционална" Then
                    Text = "Пожароизвестителната централа е проектирана да работи с конвенционални автоматични и ръчни пожароизвестители."
                    Text += "Панелът има изходи, предназначени за задействане на външно оборудване."
                    AddParagraph(wordDoc, Text, Bold:=False)
                    AddParagraph(wordDoc, "Централното съоръжение (пожароизвестителната централа) е със следните параметри:", Bold:=True)
                    AddParagraph(wordDoc, "• Пожароизвестителни линии: - 2бр.", Bold:=False)
                    AddParagraph(wordDoc, "• Максимален брой пожароизвестители в линия - 32бр.")
                    AddParagraph(wordDoc, "• Контролируеми изходи за пожар - 2бр.")
                    AddParagraph(wordDoc, "• Контролируеми изходи за повреда - 1бр.")
                    AddParagraph(wordDoc, "Функционални характеристики:", Bold:=True)
                    AddParagraph(wordDoc, "• Мониторинг на линиите за откриване на пожар и наблюдаваните изходи за условия на повреда (късо съединение и прекъсване) и автоматично нулиране на повреда;")
                    AddParagraph(wordDoc, "• Откриване на отстранен пожароизвестител и автоматично нулиране на повредата;")
                    AddParagraph(wordDoc, "• Идентифициране на задействане на ръчни пожароизвестители на линията за откриване;")
                    AddParagraph(wordDoc, "• LED индикация за състояние на пожароизвестяване и евакуация;")
                    AddParagraph(wordDoc, "• Забавяне на изходите за пожар с опционален период от време от 1 до 7 минути след откриване на състояние на пожар;")
                    AddParagraph(wordDoc, "• Възможност за сценарий;")
                    AddParagraph(wordDoc, "• Функция за евакуация, т.е. режим на работа с директно задействане на двата наблюдавани изхода;")
                    AddParagraph(wordDoc, "• Индикация на състоянието на устройството за предаване на данни по RS485 към външен оповестител;")
                    AddParagraph(wordDoc, "• Активиране/Деактивиране на интерфейса RS485;")
                    AddParagraph(wordDoc, "• Вграден зумер за пожар – еднотонален, непрекъснат, може да се изключва;")
                    AddParagraph(wordDoc, "Токозахранване:", Bold:=True)
                    AddParagraph(wordDoc, "• Мрежово захранване - 220/230VAC/0,25A;")
                    AddParagraph(wordDoc, "• Акумулаторно захранване - 2х12V DC, 7,0 Ah;")
                Else
                    Text = "Адресируемата пожароизвестителна централа е предназначена за работа с адресируеми автоматични и ръчни пожароизвестители."
                    Text += " Централата управлява адресируеми изпълнителни устройства, свързани към пожароизвестителните контури."
                    Text += " Aдресируемите изпълнителни устройства могат да бъдат захранени или от пожароизвестителния контур, или от силов контур."
                    Text += " Централата има изходи за включване на външни изпълнителни устройства."
                    AddParagraph(wordDoc, Text, Bold:=False)
                    AddParagraph(wordDoc, "Част от основните характеристики и възможности са:", Bold:=True)
                    AddParagraph(wordDoc, "• настройка на режимите на работа и параметрите на всяка пожароизвестителна зона чрез вградена клавиатура;")
                    AddParagraph(wordDoc, "• развит меню-ориентиран потребителски диалог;")
                    AddParagraph(wordDoc, "• течнокристален дисплей за визуализация в режимите на проверка и настройка на системата;")
                    AddParagraph(wordDoc, "• touch-панел към дисплея за изграждане на динамична клавиатура;")
                    AddParagraph(wordDoc, "• светодиодна индикация за сигнализиране в аварийните и екстремните ситуации;")
                    AddParagraph(wordDoc, "• архивна, енергонезависима памет за събития с указване на момента на настъпването и типа им;")
                    AddParagraph(wordDoc, "• потребителски ориентирани тестови режими;")
                    AddParagraph(wordDoc, "• вграден сериен интерфейс за връзка с други централи от същото или от по-горно ниво;")
                    AddParagraph(wordDoc, "• вграден сериен интерфейс за връзка с управляващи устройства от по-горно ниво с възможност за изграждане на връзка по телефонна линия чрез използване на стандартен модем;")
                    AddParagraph(wordDoc, "• разширяване и функционални промени на системата (предизвикани от стремеж за подобряване на противопожарната безопасност) без необходимост от преокабеляване;")
                    AddParagraph(wordDoc, "• съвместимост към разнообразен начин на проектиране на инсталацията, в рамките на предвидените ресурси на централата.")
                    AddParagraph(wordDoc, "Централното съоръжение (пожароизвестителната централа) е със следните параметри:", Bold:=True)
                    AddParagraph(wordDoc, "• Физическа конфигурация:")
                    AddParagraph(wordDoc, "  - 2 пожароизвестителни контура")
                    AddParagraph(wordDoc, "  - 1 силов контур")
                    AddParagraph(wordDoc, "  - 2 контролируеми изхода")
                    AddParagraph(wordDoc, "  - 2 релейни изхода за пожар")
                    AddParagraph(wordDoc, "  - 1 релеен изход за повреди")
                    AddParagraph(wordDoc, "• Пожароизвестителни зони:")
                    AddParagraph(wordDoc, "  - Максимален брой - 250")
                    AddParagraph(wordDoc, "  - Максимална брой пожароизвестители в зона - 60")
                    AddParagraph(wordDoc, "• Пожароизвестителни контури:")
                    AddParagraph(wordDoc, "  - Максимален брой пожароизвестители в контур - 125")
                    AddParagraph(wordDoc, "  - Вид на свързващата линия - двупроводна екранирана")
                    AddParagraph(wordDoc, "• Силов контур:")
                    AddParagraph(wordDoc, "  - Вид на свързващата линия - двупроводна")
                    AddParagraph(wordDoc, "  - Максимална консумация от контура - 1A")
                    AddParagraph(wordDoc, "• Контролируеми изходи:")
                    AddParagraph(wordDoc, "  - Тип - потенциални")
                    AddParagraph(wordDoc, "  - Електрически характеристики - (24±5)V/1A")
                    AddParagraph(wordDoc, "• Релейни изходи за пожар:")
                    AddParagraph(wordDoc, "  - Тип - безпотенциални, превключващи")
                    AddParagraph(wordDoc, "  - Електрически характеристики - 3A/125VAC; 3A/30VDC")
                    AddParagraph(wordDoc, "• Релеен изход за повреда:")
                    AddParagraph(wordDoc, "  - Тип - безпотенциален, превключващ")
                    AddParagraph(wordDoc, "  - Електрически характеристики - 3A/125VAC; 3A/30VDC")
                End If
            Else
                FormatParagraph(wordDoc, "ПРИСЪЕДИНЯВАНЕ КЪМ СЪЩЕСТВУВАЩА ПОЖАРОИЗВЕСТИТЕЛНАТА ЦЕНТРАЛА", wordApp, level:=1)
                Text = $"За нуждите на системата за пожароизвестяване ще се използва съществуващ "
                If Централа = "конвенционална" Then
                    Text += "конвенционален"
                Else
                    Text += "адресируем"
                End If
                Text += $" датчик монтиран в помещение {Кавички}{dicObekt("СРТ_Помещение_PIC")}{Кавички}{If(dicObekt("СРТ_Кота_PIC").Contains("#####"), "", $" на кота {dicObekt("СРТ_Кота_PIC")}")}."
                Text += " При този датчик ще се раздели съществуващият контур и ще се добавят новите периферни устройства."
                AddParagraph(wordDoc, Text)
                Text = "Изменението в съществуващата система да се извърши преди монтажа на предвидените в проекта инвертори."
                AddParagraph(wordDoc, Text, Bold:=True)
                Text = "След въвеждане в експлоатация на изменението да се променят съществуващите документи, предвидени в чл.9, ал.1 на НАРЕДБА № 8121з-647 ОТ 1 ОКТОМВРИ 2014 Г. За правилата и нормите за пожарна безопасност при експлоатация на обектите."
                AddParagraph(wordDoc, Text, Bold:=True)
                Text = "В процеса на интеграция на новите периферни устройства следва да се осигури непрекъснатост на работата на съществуващата системата за пожароизвестяване."
                AddParagraph(wordDoc, Text, Bold:=False)
                Text = "След приключване на дейностите да се проведат тестове и изпитвания за правилното функциониране на системата."
                Text += " Резултатите от тестовете да бъдат включени в окончателния протокол за приемане на системата."
                Text += " Персоналът, отговорен за поддръжката и експлоатацията на системата, трябва да премине инструктаж и обучение във връзка с въведените промени и устройства."
                AddParagraph(wordDoc, Text, Bold:=False)
            End If
        End With
        FormatParagraph(wordDoc, "ПЕРИФЕРНИ УСТРОСТВА - ПОЖАРОИЗВЕСТИТЕЛИ", wordApp, level:=1)
        With Text
            Dim hasCommaOrAnd As Boolean = dicObekt("Text_PIC").Contains(",") Or dicObekt("Text_PIC").Contains("и")
            Text = $"В помещение {Кавички}{dicObekt("Помещение_ИНВ")}{Кавички}{If(dicObekt("Кота_ИНВ").Contains("#####"), " ", $" на кота {dicObekt("Кота_ИНВ")} ")},"
            If hasCommaOrAnd Then
                Text += " ще се монтират: "
            Else
                Text += " ще се монтира: "
            End If
            Text += dicObekt("Text_PIC")
            Text += "."
            AddParagraph(wordDoc, Text)
            FormatParagraph(wordDoc, "Зони на откриване на признаци на пожар", wordApp, level:=2)
            Text = "След консултации с възложителя е прието да се изгради локална, автоматична пожароизвестителна система за локална защита на"
            If hasCommaOrAnd Then
                Text += " инверторите монтирани"
            Else
                Text += " инверторa монтиран"
            End If
            Text += $" в помещение {Кавички}{dicObekt("Помещение_ИНВ")}{Кавички}{If(dicObekt("Кота_ИНВ").Contains("#####"), " ", $" на кота {dicObekt("Кота_ИНВ")} ")}."
            AddParagraph(wordDoc, Text)
            Text = "Локалната защита сама по себе си може открие признаци на пожари, възникващи в рамките на защитаваната площ, но не може да служи за откриване на пожари, които са извън тази площ."
            'Text += " Поради тази причина в настоящия проект не се придвижа пълна защита на цели обект."
            AddParagraph(wordDoc, Text)
            Text = "При определяне на зоната е взето предвид вътрешното разпределение на сградата, всички възможни трудности за търсене или придвижване, осигуряването на зони на сигнализиране на тревога и наличието на всички особени опасности."
            AddParagraph(wordDoc, Text)
            Dim uniqueVisibilityList As New List(Of String)()
            '
            ' Добавяме само уникални стойности в списъка
            '
            For Each item In picList
                If Not uniqueVisibilityList.Contains(item.Visibility) Then
                    uniqueVisibilityList.Add(item.Visibility)
                End If
            Next
            FormatParagraph(wordDoc, "Избор на автоматични пожароизвестители", wordApp, level:=2)

            Text = "За ранно откриване на евентуални признаци за възникване на пожар е предвидено"
            Text += $" в помещение {Кавички}{dicObekt("Помещение_ИНВ")}{Кавички}{If(dicObekt("Кота_ИНВ").Contains("#####"), " ", $" на кота {dicObekt("Кота_ИНВ")} ")}"
            Text += "да се монтират:"
            AddParagraph(wordDoc, Text)
            Dim values As New List(Of String)

            For i = 28 To 37
                'Dim countValue As Integer = wsFEC_Таблица.Cells(i, 5).Value
                'Dim textValue As String = String.Format("{0:0} бр. ", countValue)
                'textValue += IIf(countValue = 1, "трифазен многострингов инвертор тип ", "трифазни многострингови инвертори тип ")
                'textValue += wsFEC_Таблица.Cells(i, 1).Value
                'values.Add(textValue)
            Next
            Dim Text_PIC As String = ""
            If values.Count > 0 Then
                If values.Count = 1 Then
                    Text += values(0)
                    Text_PIC = values(0)
                Else
                    Text_PIC = String.Join(", ", values.Take(values.Count - 1)) & " и " & values.Last()
                End If
            End If
            For Each item In uniqueVisibilityList
                Select Case item
                    Case "ПАБ - Термичен конвенционален диференциален",
                         "ПАБ - Термичен конвенционален",
                         "ПАБ - Термичен адресируем диференциален",
                         "ПАБ - Термичен адресируем с адаптер-7120",
                         "ПАБ - Термичен адресируем - 7101"
                        Text = "Топлинен  пожароизвестител - осигурява надеждно откриване на пожар в ранния стадий на неговото развитие, при скорост на нарастване на температурата, по-голяма от зададената или при превишаване на определена максимална температурата на охраняваната среда."
                        Text += " Температурният клас е в съответствие с Европейски стандарт EN54-5 и е програмируем от пожароизвестителна централа."
                        AddParagraph(wordDoc, Text)
                    Case "ПАБ - Димооптичен адресируем",
                         "ПАБ - Димооптичен конвенционален"
                        Text = "Оптично димните пожароизвестители - осигуряват надеждно откриване на пожар в ранния стадий на неговото развитие, по концентрацията на дим в охраняваната среда."
                        Text += " Чувствителността на дим е в съответствие с Европейски стандарт EN54-7 и е програмируем от пожароизвестителна централа."
                        Text += " Пожароизвестителя работи по усъвършенстван алгоритъм за самокомпенсация на замърсяването на оптичната камера, като сигнализира необходимостта от почистването й."
                        Text += " Бързо почистване, конструкция осигуряваща висока степен на защита от запрашаване и работа при силни въздушни течения."
                        AddParagraph(wordDoc, Text)
                    Case "ПАБ - Термичен адресируем комбиниран",
                         "ПАБ - Термичен конвенционален комбиниран"
                        Text = "Комбиниран оптично-димен и точков топлинен пожароизвестител - предназначен е за откриване на пожар в ранния стадий на неговото развитие по концентрацията на дим или при скорост на
нарастване на температурата, по - голяма от зададената или при превишаване на определена максимална температура на охраняваната среда.
Принципът на работа на оптичната част на пожароизвестителя се основава на разсейването на инфрачервени лъчи от частиците дим, попаднали
в оптичната камера. Принципът на работа на термичната му част се основава на изменение на омическото съпротивление на термистор при
промяна на околната температура. Чувствителността на дим и температурният клас се задават програмно от пожароизвестителната централа
7000M по специализирания протокол за обмен на информация UniTALK. В пожароизвестителя има вграден изолатор на късо съединение.
FD7160M се монтира на основа 7100."
                        AddParagraph(wordDoc, Text)
                    Case "ПАБ - Пламъков конвенционален"
                        Text = "Пламъчните пожароизвестители откриват излъчването от пожарите."
                        Text += " Може да бъде използвано ултравиолетово или инфрачервено излъчване или комбинация от двете."
                        Text += " Спектърът на излъчване на повечето пламтящи материали е достатъчно широколентов, за да бъде открит от който и да е пламъчен пожароизвестител, но при някои материали може да бъде необходимо да се изберат пламъчни пожароизвестители, способни да реагират на специфични части от спектъра на дължини на вълната."
                        Text += " Пламъчните пожароизвестители могат да реагират на пожарите с пламък по-бързо от топлинните или димните."
                        Text += " Поради неспособността им да откриват тлеещи пожари пламъчните пожароизвестители не биха могли да се разглеждат като общоприложими, т.е.те трябва да се използват само когато основен риск са пламъчните пожари."
                        Text += " Пламъчните пожароизвестители работят чрез осъществяване на връзка на база видимост, което означава, че не е необходимо монтирането им върху таван, но те могат да бъдат използвани само ако имат пряка линия на видимост към наблюдаваната площ."
                        Text += " Би трябвало да се вземат предпазни мерки срещу замърсяване на пожароизвестителите или срещу среди, които оказват негативно влияние на излъчването, като например: "
                        Text += "- отлагане на масло, грес или прах, стъкло при пожароизвестители с ултравиолетово излъчване, - лед, кондензация или стъкло при пожароизвестители с инфрачервено излъчване."
                        Text += " Ако при производството или други процеси може да възникват фалшиви сигнали за тревога, се препоръчва предпазливост при използването на пламъчни пожароизвестители. Например мигаща светлина, източници на радиация, заваряване и т.н."
                        Text += " Ако е вероятно пламъчните пожароизвестители да бъдат изложени на слънчева светлина, би трябвало да бъдат избирани пожароизвестители, които не са чувствителни към слънчева светлина."
                        Text += "Пламъчните пожароизвестители трябва да бъдат изпълнени, както е определено в EN 54-10, който определя различните класове за ултравиолетови и инфрачервени пожароизвестители."
                    Case "Линеен оптично димен приемник",
                         "Линеен оптично димен излъчвател"
                        Text = "Лъчевите димни пожароизвестители по принцип възприемат затъмняването на светлинен лъч и следователно са чувствителни към плътността на дима по дължина на лъча."
                        Text += "Те са особено подходящи за използване, когато димът може да бъде разсеян на голяма площ, преди да бъде открит, например за използване под високи тавани."
                        Text += " Лъчевите димни пожароизвестители трябва да съответстват на EN 54-12."
                        AddParagraph(wordDoc, Text)
                End Select
            Next
            FormatParagraph(wordDoc, "Разполагане на автоматичните пожароизвестители", wordApp, level:=2)
            Text = "Автоматичните пожароизвестители са избрани, като са отчетени предназначението и големината на защитаваният обект."
            Text += " Автоматичните пожароизвестители сa разположени така, че съответните продукти на какъвто и да е пожар в защитаваната площ да могат да достигат пожароизвестителите без прекомерно разреждане, затихване или забавяне."
            AddParagraph(wordDoc, Text)
            AddParagraph(wordDoc, "• Разстоянието между електрическите ключове/контакти и ръчните пожароизвестители трябва да бъде минимум 0,5м.")
            AddParagraph(wordDoc, "• Автоматичните пожароизвестителни датчици трябва да бъдат монтирани на разстояние минимум 0,5м от стени, трегери и стелажи и на 1м от вентилационни отвори.")
            AddParagraph(wordDoc, "• Автоматичните пожароизвестителни датчици трябва да бъдат монтирани на разстояние не по-малко от два пъти височината на съответното осветително тяло.")
            FormatParagraph(wordDoc, "Разполагане на ръчните пожароизвестители", wordApp, level:=2)
            Text = "Ръчните пожароизвестители са разположени така, че да могат лесно и бързо да бъдат задействани от всяко лице, открило пожар."
            Text += " Ръчните пожароизвестители са разположвени по евакуационните пътища, при всяка врата към евакуационни стълбища и при всеки изход, водещ на открито."
            Text += " Ръчните пожароизвестители са добре видими, ясно отличими и лесно достъпни."
            Text += " Ръчните пожароизвестители да бъдат поставени на височина между 0,9 m и 1,4 m над пода (за предпочитане 1,2 m)."
            AddParagraph(wordDoc, Text)
            AddParagraph(wordDoc, "Местоположението на периферните устройства е показана в графичната част към проекта.")
            AddParagraph(wordDoc, $"При монтажа на периферните устройства да се спазят изискванията на точка 7 {Кавички}Монтиране{Кавички} на СД CEN/TS 54-14.")
        End With
        If Присъединяване Then
            FormatParagraph(wordDoc, "ЕЛЕКТРОЗАХРАНВАНЕ НА ИНСТАЛАЦИЯТА", wordApp, level:=1)
            With Text
                Text = "Основният източник на енергия на системата е обществената електроснабдителна система."
                Text += " Не се допуска произведена за лични нужди енергия да бъде използвана за захранване на инсталацията."
                AddParagraph(wordDoc, Text)
                Text = "Централното съоръжение ще бъде свързано към мрежово захранване 220V/50Hz с кабел тип NHXH-FE180/E30 3x1,5мм² от главното ел. табло."
                Text += " Основното електрозахранване за пожароизвестителната система е снабдено със специално предназначен автоматичен прекъсвач 6А."
                Text += " Той не се използва за други цели и трябва да бъде ясно означено (например пожароизвестяване). "
                AddParagraph(wordDoc, Text)
                Text = $"Преди въвеждане в експлоатация на инсталацията трябва да бъдат взети мерки (например чрез табелка {Кавички}НЕ ИЗКЛЮЧВАЙ{Кавички} или ограничаване на достъпа) за предотвратяване на неупълномощено прекъсване на основното електрозахранване."
                AddParagraph(wordDoc, Text)
                Text = "При отпадане на основния източник на захранване с енергия е предвидено резервно захранване от две паралелно свързани акумулаторни батерии (12V) с капацитет - 7Ah, които осигуряват работа в дежурен режим - 168h."
                Text += " Капацитетът на батерията е изчислен в зависимост от тока при повреда и тока при сигнал за тревога, и от необходимото време за резерв при отпадане на захранването от мрежата."
                Text += " Той е достатъчен за захранване на системата по време на цялото вероятно прекъсване на основния източник на енергия."
                AddParagraph(wordDoc, Text)
                AddParagraph(wordDoc, "Необходимото време за резерв при отпадане на захранването от мрежата е изчислено като е взето предвид:")
                AddParagraph(wordDoc, "• времето до откриване на неизправността в електрозахранването и повикване на сервиз/ремонт;")
                AddParagraph(wordDoc, "• времето, необходимо на персонала по поддържане, за ремонт на системата и възстановяване на основното електрозахранване;")
            End With
        End If
        FormatParagraph(wordDoc, "ИЗПЪЛНЕНИЕ НА ИНСТАЛАЦИЯТА", wordApp, level:=1)
        With Text
            Text = "Периферните устройства ще се свържат помежду си и"
            If Присъединяване Then
                Text += " с централното съоръжение"
            Else
                Text += " със съществуващият датчик"
            End If
            Text += " чрез кабел тип FS 2x0,75мм²."
            Text += " Кабелът е устойчив на огън, стандартен и съответства на класификация РН 30."
            Text += " Кабелът от този тип е многожичен, двужилен, с пластмасова изолация за всяко жило, единичен неизолиран заземителен проводник, общ екран от алуминиево фолио и външна пластмасова изолация, оцветена в червено."
            Text += " Допуска се замяна на този тип кабели с друг тип, който е равностоен или еквивалентен, със същите или по-добри параметри от посочените, при спазване на всички изисквания на действащите нормативни документи."
            AddParagraph(wordDoc, Text)
            Text = "Кабелите имат подходяща защита и при пожар са в състояние да функционират 30 min."
            Text += " Това е постигнато чрез използване на кабели с класификация РН 30 по EN 50200."
            AddParagraph(wordDoc, Text)
            AddParagraph(wordDoc, "Кабелите ще се изтеглят скрито в PVC тръби ф16,0мм.")
            AddParagraph(wordDoc, "При пожар е осигурено да остават работоспособни за по-дълъг период от време следните кабели:")
            AddParagraph(wordDoc, "• връзките между основното разпределително табло за ниско напрежение и устройството за управление и индикация (ПИЦ) и всички други електрозахранващи единици, които са част от пожароизвестителната система;")
            AddParagraph(wordDoc, "• връзките между устройството за управление и индикация (ПИЦ) и всички обособени електрозахранващи единици, включително кабелите между устройствата за сигнализация за тревога и тяхното електрозахранване;")
        End With
        FormatParagraph(wordDoc, "СИСТЕМА ЗА ЗВУКОВА СИГНАЛИЗАЦИЯ", wordApp, level:=1)
        With Text
            If Присъединяване Then
                AddParagraph(wordDoc, "За звуково известяване на евентуални признаци за възникване на пожар е предвидено да се използват вътрешни и външни сирени.")
                Text = "Външната сирена е предназначена да предупреди хората, намиращи се извън сградата, за възникнал пожар и да насочи пожарните коли към обекта, в който е задействана."
                Text += " Сирената, предвидена в проекта, е за монтаж на открито с лампа и съчетава визуалния с акустичния сигнал."
                Text += " Сирената ще се монтира на фасадата на сградата на височина 3-3,5м от кота терен."
                AddParagraph(wordDoc, Text)
                AddParagraph(wordDoc, "Вътрешните сирени служат за звуково и светлинно сигнализиране на възникнали събития, регистрирани от пожароизвестителната система в закритите помещения.")
                AddParagraph(wordDoc, "Звуковите сигнализатори имат звуково ниво от 94 dBA на разстояние 1 m при хоризонтален ъгъл от 75°.")
                AddParagraph(wordDoc, "Броят и местоположението на звуковите сигнализатори удовлетворяват следите изисквания:")
                AddParagraph(wordDoc, "• Осигуреното звуково ниво е такова, че сигналът за пожарна тревога да бъде чут незабавно от фона на околния шум, но нивото на звука във всяка точка, в която обичайно има хора, не превишава стойността 118 dBA.")
                AddParagraph(wordDoc, "• Звукът, използван за целите на пожарната сигнализация, е еднакъв във всички части на обекта.")
                AddParagraph(wordDoc, "• Нивото на звука на сигнала за пожарна тревога е с 10 dBA над всеки околен шум, който може да е с продължителност повече от 30 s.")
                AddParagraph(wordDoc, "• Минимално ниво е достигнато навсякъде, където е необходимо да бъде чут звукът на сигнала за тревога.")
                AddParagraph(wordDoc, "• Звукът (тонът) на сигнала за тревога не се използва за други цели.")
                AddParagraph(wordDoc, "Местоположението на вътрешните и външните сирени е показана в графичната част към проекта.")
            Else
                AddParagraph(wordDoc, "За звуково известяване на евентуални признаци за възникване на пожар е предвидено да се използват съществуващите вътрешни и външни сирени.")
                AddParagraph(wordDoc, "В проекта не се предвижда промяна на броя и местоположението на съществуващите вътрешни и външни сирени.")
                AddParagraph(wordDoc, "В проекта не се предвижда промяна на алгоритъма на сработване на съществуващите вътрешни и външни сирени.")
            End If
        End With
        FormatParagraph(wordDoc, "ЗАКЛЮЧЕНИЕ", wordApp, level:=1)
        With Text
            Text = "При изпълнението на електро-монтажните работи да се спазят изискванията на действащите нормативни актове, както и на всички изменения и допълнения към тях, влезли в сила към момента на завършване на дейностите по изграждане на инсталациите предвидени в проекта."
            AddParagraph(wordDoc, Text)
            Text = "Навсякъде в проекта, където са посочени изрично конкретни продукти с конкретни търговски марки следва да разбира и оферира не задължително посочените продукти, а равностойни, еквивалентни, със същите или по-добри параметри от посочените, като се спазват всички изисквания на действащите нормативни документи и съответстват на решението на проектанта и избраната технология!"
            AddParagraph(wordDoc, Text)
            Text = "Ако по някаква причина проектът, се окаже несъответстващ по време на монтирането, всички необходими изменения трябва да се съгласуват с проектанта или с друго достатъчно правоспособно лице."
            Text += " Одобрените изменения, включително потвърждението на проекта, трябва да бъдат въведени в документацията."
            AddParagraph(wordDoc, Text)
            Text = "По време на конфигурацията трябва да бъде проверено дали всички периферни устройства са програмирани в съответствие със зададеното в проекта."
            AddParagraph(wordDoc, Text)
            Text = $"При въвеждане в експлоатация да се спазят изискванията на т. 9 {Кавички}Въвеждане в експлоатация, приемане и верификация{Кавички} на СД CEN/TS 54-14."
            AddParagraph(wordDoc, Text)
        End With
        AddПодпис(wordDoc, dicObekt("ПРОЕКТАНТ"))
    End Sub
End Class

'2. Обяснителна записка на фаза технически и работен проект, която съдържа:
'2.1. Пасивни мерки за пожарна безопасност:
'2.1.1. проектни обемно-планировъчни и функционални показатели на строежа,
'       в т.ч. стълбищни клетки (брой, разположение, изпълнение, осветеност), асансьорни шахти,
'       отделяне на помещения на разпределителни електрически табла, складови и производствени помещения,
'       разстояния между сградите и съоръженията; брой и размери на евакуационните изходи от сградата,
'       размери на пътищата за евакуация, определяне на изчислителното време за евакуация (когато се изисква),
'       пътища за противопожарни цели, отстояния от сгради и съоръжения на строежа до надземни и подземни инженерни проводи и др.;
'2.1.2. клас на функционална пожарна опасност;
'2.1.3. степен на огнеустойчивост на строежа и на конструктивните му елементи - проектни стойности на носимоспособността, непроницаемостта,
'       изолиращата способност и на други допълнителни критерии за определяне на огнеустойчивостта на строежа в зависимост от вида и
'       предназначението му, в т.ч. носещи стени и колони, междуетажни конструкции, фасадни и вътрешни стени, стени на евакуационните пътища,
'       стълбищни рамена, инсталационни шахти, стени на складове и производствени помещения, врати в пожарозащитните прегради;
'2.1.4. проектна огнеустойчивост на огнезащитаваните конструктивни елементи на сградата:
'2.1.4.1. огнезащита на стоманени конструктивни елементи - начини на изпълнение на покритията в зависимост от вида на сечението на стоманените
'         конструктивни елементи: отворени профили - П-профил; І-профил; L-профил; Т-профил и др.; затворени профили - 0 (правоъгълни, квадратни);
'         O (кръгли профили); D (триъгълни) и др., факторът на масивност, технологията на нанасяне на огнезащитните състави, външните (атмосферните)
'         условия, минималният брой слоеве и др.;
'2.1.5. класове по реакция на огън на продуктите за конструктивни елементи, за покрития на вътрешни (стени, тавани и подове) и външни повърхности,
'       за технологични инсталации, уредби и съоръжения (вентилационни, отоплителни, електрически и др.) в зависимост от вида на сградата и
'       предназначението на помещенията.
'2.2. Активни мерки за пожарна безопасност:
'2.2.1. обемно-планировъчни и функционални показатели за пожарогасителни инсталации в зависимост от вида и предназначението на строежа,
'       в т.ч. вид на инсталацията, площи, които подлежат на защита с пожарогасителна инсталация, изчислителни стойности на оразмеряването на инсталацията, проектни водни количества, блокировки и др.;
'2.2.2. Обемно-планировъчни и функционални показатели за пожароизвестителни инсталации в зависимост от вида и предназначението на строежа, в т.ч. вид на инсталацията, площи, които подлежат на защита с пожароизвестителна инсталация, местоположение на централата, степен на защита на оборудването, блокировки и др.;
'2.2.3. Обемно-планировъчни и функционални показатели за оповестителни инсталации в зависимост от вида и предназначението на строежа, в т.ч. площи, подлежащи на озвучаване; задействане на инсталацията и др.;
'2.2.4. Обемно-планировъчни и функционални показатели за димо-топлоотвеждащи инсталации в зависимост от вида и предназначението на строежа, в т.ч. помещения и зони, подлежащи на димо- и топлоотвеждане, определяне на незадимяемата зона в помещенията, определяне на димен участък и резервоар, кратност на въздухообмена на димо- и топлоотвеждащите инсталации, кратност на въздухообмена при аварийна вентилационна инсталация, размери и разположение на димни люкове и механични вентилатори, приточни отвори и места за подаване на чист въздух и др.;
'2.2.5. (доп. - ДВ, бр. 89 от 2014 г.) функционални показатели за водоснабдяване за пожарогасене в зависимост от вида и предназначението на строежа, в т.ч. брой на пожарните хидранти, водопровод за пожарогасене, резервоар, водоизточник (обем), засмукване и възстановяване на водните количества, инсталации за пожарогасене по време на изпълнението на строежа и др.;
'2.2.6. функционални показатели за преносими уреди и съоръжения за първоначално пожарогасене, в т.ч. вид и брой на уредите и съоръженията за помещение, за етаж или за цялата сграда;
'2.2.7. функционални показатели на евакуационно осветление в зависимост от вида и предназначението на строежа, в т.ч. минимална осветеност по пътищата за евакуация, защита от топлина на елементите на инсталацията и др.;
'2.2.8. блок-схема на проектираните активни мерки за защита (със самостоятелно задействане или управлявани от ПИС), начинът на привеждането им в действие и осигурените блокировки за съвместната работа на системите.








'Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
'    Try
'        For Each sObj As SelectedObject In ss_Tabla
'            blkRecId = sObj.ObjectId
'            Dim acBlkRef As BlockReference = DirectCast(acTrans.GetObject(blkRecId, OpenMode.ForRead), BlockReference)

'             Проверка дали блокът е динамичен
'            If Not acBlkRef.IsDynamicBlock Then Continue For

'            Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection
'            Dim picItem As New PIC()
'            picItem.Count = 0 ' Инициализация на брояча

'             Извличане на свойството "Visibility"
'            For Each prop As DynamicBlockReferenceProperty In props
'                If prop.PropertyName = "Visibility" Then
'                    picItem.Visibility = prop.Value
'                    picItem.Count += 1 ' Увеличаване на брояча при всяко "Visibility"
'                End If
'            Next

'             Ако няма Visibility, пропусни този блок
'            If String.IsNullOrEmpty(picItem.Visibility) Then Continue For

'             Проверка за уникалност на Visibility
'            Dim isUnique As Boolean = Not picList.Any(Function(p) p.Visibility = picItem.Visibility)
'            If Not isUnique Then Continue For

'             Проверка на името на блока
'            Dim blName As String = CType(acBlkRef.DynamicBlockTableRecord.GetObject(OpenMode.ForRead), BlockTableRecord).Name
'            If blName <> "Табло_Ново" AndAlso blName <> "Датчик_ПАБ" Then Continue For

'             Обработка на атрибутите
'            Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
'            For Each objID As ObjectId In attCol
'                Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
'                Dim acAttRef As AttributeReference = TryCast(dbObj, AttributeReference)
'                If acAttRef Is Nothing Then Continue For

'                Select Case acAttRef.Tag
'                    Case "ТАБЛО"
'                        picItem.Tablo = acAttRef.TextString
'                    Case "ZN"
'                        picItem.ZN = acAttRef.TextString
'                    Case "NOM"
'                        picItem.NOM = acAttRef.TextString
'                    Case "AD"
'                        picItem.AD = acAttRef.TextString
'                End Select
'            Next

'             Добавяне в списъка само ако Visibility е уникално
'            picList.Add(picItem)
'        Next

'        acTrans.Commit()

'    Catch ex As Exception
'        MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace)
'        acTrans.Abort()
'    Finally
'        If acTrans IsNot Nothing Then acTrans.Dispose()
'    End Try
'End Using
