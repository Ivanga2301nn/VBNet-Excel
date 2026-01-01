Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports System.Text
Imports System.Text.RegularExpressions
Imports VBNet_Excel.Form_ExcelUtilForm
Imports VBNet_Excel.Zapiska
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox
Imports System.Security.Policy

Public Class CommonUtil
    '
    ' можеш ли да напишеш подробни коментарии към този код, а обаснението което даваш след кода, да вмъкнеш като коментар в началото 
    '
    Dim PI As Double = 3.1415926535897931
    Structure Kabel
        Dim Layer As String
        Dim Layer_Line As String
        Dim Linetype As String
        Dim Diam As Double
        Dim Se4enie As Double
        Dim Diamet As Double
        Dim I_N As Double
    End Structure
    Structure strKabelАlign
        Dim pInsert As Point3d
        Dim Id As ObjectId
        Dim Position_X As Double
        Dim Position_Y As Double
        Dim Stypka As Double
    End Structure
    Public Structure strLine
        Dim Layer As String
        Dim Linetype As String
        Dim count As Double
    End Structure
    Public Structure strТабло
        Dim bl_Табло As String
        Dim bl_Брой As Integer
        Dim bl_ИмеБлок As String
        Dim bl_1 As String
        Dim bl_2 As String
        Dim bl_3 As String
        Dim bl_4 As String
        Dim bl_5 As String
        Dim bl_6 As String
        Dim bl_7 As String
        Dim bl_8 As String
        Dim bl_9 As String
        Dim bl_10 As String
        Dim bl_DESIGNATION As String
        Dim bl_LONGNAME As String
        Dim bl_REFNB As String
        Dim bl_SHORTNAME As String
        Dim bl_RABATY As String
        Dim bl_RABATY2 As String

    End Structure
    Public Structure str_R_Cable
        Dim сечение As String
        Dim R_Cu_20 As Double
        Dim R_Cu_50 As Double
        Dim R_Al_20 As Double
        Dim R_Al_50 As Double
    End Structure
    Public Function GetR_Cable(сечение As String) As str_R_Cable
        Dim R_Cable As New List(Of str_R_Cable)()
        Dim ind_R_Cable As str_R_Cable
        With ind_R_Cable
            .сечение = "0,0"
            .R_Al_20 = 0
            .R_Al_50 = 0
            .R_Cu_20 = 0
            .R_Cu_50 = 0
        End With
        With ind_R_Cable
            .сечение = "1,0"
            .R_Al_20 = 29
            .R_Al_50 = 30
            .R_Cu_20 = 17.8
            .R_Cu_50 = 20
        End With
        R_Cable.Add(ind_R_Cable)
        With ind_R_Cable
            .сечение = "1,5"
            .R_Al_20 = 10.33
            .R_Al_50 = 13.86
            .R_Cu_20 = 11.86
            .R_Cu_50 = 13.3
        End With
        R_Cable.Add(ind_R_Cable)
        With ind_R_Cable
            .сечение = "2,5"
            .R_Al_20 = 11.3
            .R_Al_50 = 13.2
            .R_Cu_20 = 7.12
            .R_Cu_50 = 8
        End With
        R_Cable.Add(ind_R_Cable)
        With ind_R_Cable
            .сечение = "4"
            .R_Al_20 = 7.3
            .R_Al_50 = 8.25
            .R_Cu_20 = 4.45
            .R_Cu_50 = 5
        End With
        R_Cable.Add(ind_R_Cable)
        With ind_R_Cable
            .сечение = "6"
            .R_Al_20 = 4.8
            .R_Al_50 = 5.5
            .R_Cu_20 = 2.97
            .R_Cu_50 = 3.33
        End With
        R_Cable.Add(ind_R_Cable)
        With ind_R_Cable
            .сечение = "10"
            .R_Al_20 = 2.9
            .R_Al_50 = 3.3
            .R_Cu_20 = 1.78
            .R_Cu_50 = 2
        End With
        R_Cable.Add(ind_R_Cable)
        With ind_R_Cable
            .сечение = "16"
            .R_Al_20 = 1.16
            .R_Al_50 = 1.32
            .R_Cu_20 = 0.71
            .R_Cu_50 = 0.8
        End With
        R_Cable.Add(ind_R_Cable)
        With ind_R_Cable
            .сечение = "25"
            .R_Al_20 = 1.16
            .R_Al_50 = 1.32
            .R_Cu_20 = 0.78
            .R_Cu_50 = 0.8
        End With
        R_Cable.Add(ind_R_Cable)
        With ind_R_Cable
            .сечение = "35"
            .R_Al_20 = 0.83
            .R_Al_50 = 0.94
            .R_Cu_20 = 0.51
            .R_Cu_50 = 0.57
        End With
        R_Cable.Add(ind_R_Cable)
        With ind_R_Cable
            .сечение = "35"
            .R_Al_20 = 0.83
            .R_Al_50 = 0.94
            .R_Cu_20 = 0.51
            .R_Cu_50 = 0.57
        End With
        R_Cable.Add(ind_R_Cable)
        With ind_R_Cable
            .сечение = "50"
            .R_Al_20 = 0.58
            .R_Al_50 = 0.65
            .R_Cu_20 = 0.35
            .R_Cu_50 = 0.4
        End With
        R_Cable.Add(ind_R_Cable)
        With ind_R_Cable
            .сечение = "70"
            .R_Al_20 = 0.41
            .R_Al_50 = 0.47
            .R_Cu_20 = 0.25
            .R_Cu_50 = 0.29
        End With
        R_Cable.Add(ind_R_Cable)
        With ind_R_Cable
            .сечение = "95"
            .R_Al_20 = 0.3
            .R_Al_50 = 0.35
            .R_Cu_20 = 0.19
            .R_Cu_50 = 0.21
        End With
        R_Cable.Add(ind_R_Cable)
        With ind_R_Cable
            .сечение = "120"
            .R_Al_20 = 0.21
            .R_Al_50 = 0.27
            .R_Cu_20 = 0.15
            .R_Cu_50 = 0.17
        End With
        R_Cable.Add(ind_R_Cable)
        With ind_R_Cable
            .сечение = "150"
            .R_Al_20 = 0.19
            .R_Al_50 = 0.22
            .R_Cu_20 = 0.12
            .R_Cu_50 = 0.13
        End With
        R_Cable.Add(ind_R_Cable)
        With ind_R_Cable
            .сечение = "185"
            .R_Al_20 = 0.16
            .R_Al_50 = 0.18
            .R_Cu_20 = 0.1
            .R_Cu_50 = 0.11
        End With
        R_Cable.Add(ind_R_Cable)
        With ind_R_Cable
            .сечение = "240"
            .R_Al_20 = 0.12
            .R_Al_50 = 0.14
            .R_Cu_20 = 0.07
            .R_Cu_50 = 0.08
        End With
        R_Cable.Add(ind_R_Cable)
        Dim index As Integer = R_Cable.FindIndex(Function(k) k.сечение = сечение)
        If index = -1 Then
            index = 0
        End If
        Return R_Cable(index)
    End Function
    ' Функция за изчисление на загубата на напрежение при променлив ток
    Function CalculateVoltageDropAC(Дължина As Double,                      ' Дължина на кабела
                                    Мощност As Double,                      ' Захранвана мощност
                                    Сечение As String,                      ' Сечение на кабела СВТ3х1,5mm2
                                    Optional reactance As Double = 0.1,     ' Реактивно съпротивление за един километър
                                    Optional Motor As Boolean = False       ' Ако е двигател True - КПД и cos FI да са по 0,83
                                    ) As Double
        ' Проверете дали inputString съдържа "коакс" или "FTP" поц AlMgSi 
        Dim Delta_U As Double = 0
        If Сечение.Contains("коакс") OrElse
            Сечение.Contains("FTP") OrElse
            Сечение.Contains("поц") OrElse
            Сечение.Contains("AlMgSi") OrElse
            Сечение.Contains("ELEKTRO") Then Exit Function

        If Val(Сечение) <> 0 Then Exit Function

        If Мощност = 0 Then Мощност = 20

        Dim cosPhi As Double = 1
        Dim KPD As Double = 0
        Const U380 As Double = 0.38
        Const U220 As Double = 0.22
        Dim U As Double

        Dim Inom As Double = 0
        If Motor Then
            cosPhi = 0.83
            KPD = 0.83
        Else
            cosPhi = 0.9
            KPD = 1
        End If
        ' Извлечете типа на кабела
        Dim startIndex As Integer = 0
        Dim endIndex As Integer = Сечение.IndexOf("x") - 1
        If endIndex < startIndex Then Exit Function
        Dim cableType As String = Сечение.Substring(startIndex, endIndex)

        ' Извлечете броя на жилата
        startIndex = endIndex
        endIndex = Сечение.IndexOf("x")
        Dim numberOfWires As String = Сечение.Substring(startIndex, 1)

        Select Case numberOfWires
            Case 0, 1
                Дължина = 0
                U = U220
            Case 2
                U = U220
            Case 3, 4, 5
                U = U380
        End Select

        ' Извлечете сечението на кабела поц.AlMgSi 
        startIndex = Сечение.IndexOf("x") + 1
        Dim cableSection As String = Сечение.Substring(startIndex)

        ' Премахнете "mm²" от сечението на кабела
        If cableSection.EndsWith("mm²") Then
            cableSection = cableSection.Substring(0, cableSection.Length - 3)
        End If

        Dim R = GetR_Cable(Val(cableSection).ToString)
        Dim resistance As Double = 0

        If cableType.IndexOf("САВТ") > -1 Or
            cableType.IndexOf("Al/R") > -1 Then
            resistance = R.R_Al_20
        Else
            resistance = R.R_Cu_20
        End If
        reactance = reactance
        Dim Qreak = Мощност * Math.Sqrt(1 - cosPhi * cosPhi) / cosPhi

        Delta_U = (Дължина / 1000)
        Delta_U = Delta_U * (Мощност * resistance + Qreak * reactance)
        Delta_U = Delta_U / (U * U)
        Delta_U = Delta_U * 100

        Return Delta_U / 1000
    End Function
    ' Получава всички посочени обекти (напр. Линия, MText) и връща набор за избор
    Public Function GetObjects(objType As String,
                               mesage As String,
                               Optional allowMultiple As Boolean = True,
                               Optional doc As Document = Nothing
                               ) As SelectionSet

        If doc Is Nothing Then
            doc = Application.DocumentManager.MdiActiveDocument
        End If
        Dim edt As Editor = doc.Editor
        Dim prSelOpts As PromptSelectionOptions = New PromptSelectionOptions()
        ' objType       - Тип на обектите за избор:
        ' LINE          - Линия
        ' INSERT        - Блок
        ' MTEXT         - МУЛТИ Tекст - за текст има друга функция
        ' Dim СРТ_Кота = cu.GetObjects_TEXT("......" & СРТ_Име)
        ' TEXT          - Tекст
        ' LWPOLYLINE    - Полилиния
        ' CIRCLE        - Окръжност



        '
        ' allowMultiple - True  - Избира много обекти
        '               - False - Избира само един обект
        '

        ' страница описваща филтрирането
        ' http://docs.autodesk.com/ACD/2011/ENU/filesMDG/WS1a9193826455f5ff2566ffd511ff6f8c7ca-4067.htm
        ' страница описваща как да прикачим приложението - като се завърши всичко ще потрябва
        ' http://docs.autodesk.com/ACD/2011/ENU/filesMDG/WS73099cc142f48755-5c83e7b1120018de8c0-1c12.htm

        Dim tv As TypedValue() = New TypedValue(2) {}

        tv.SetValue(New TypedValue(CInt(DxfCode.Start), objType), 0)
        tv.SetValue(New TypedValue(67, 0), 1)                               ' Изберете само от ModelSpace
        tv.SetValue(New TypedValue(CInt(DxfCode.LayerName), "EL*"), 2)      ' Изберете Всички слоеве започващи с EL

        Dim filter As SelectionFilter = New SelectionFilter(tv)

        With prSelOpts
            .MessageForAdding = vbCrLf & mesage
            .AllowDuplicates = allowMultiple
        End With

        Dim prs As PromptSelectionResult
        If allowMultiple Then
            prs = edt.GetSelection(prSelOpts, filter)
        Else
            prSelOpts.SingleOnly = True
            prs = edt.GetSelection(prSelOpts, filter)
        End If
        Dim ss As SelectionSet = prs.Value
        Return ss
    End Function
    ' Функцията GetObjects_TEXT приема входен низ като съобщение и връща нов низ,
    ' в който всички латински букви, които имат еквиваленти на кирилица, са заменени.
    Public Function GetObjects_TEXT(mesage As String,                       ' Текст който да се оправя
                                    Optional podpis As Boolean = True,      ' Да се прави ли проверка за подспис
                                    Optional doc As Document = Nothing
                                    ) As String
        If doc Is Nothing Then
            doc = Application.DocumentManager.MdiActiveDocument
        End If
        ' Инициализираме резултатния текст с празен низ
        Dim RetText As String = "  #####  "
        Try
            ' Получаваме активния документ и редактора
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim edt As Editor = acDoc.Editor
            ' Създаваме опции за избор на обект с подаденото съобщение
            Dim prEntOpts As PromptEntityOptions = New PromptEntityOptions(vbCrLf & mesage)
            ' Задаваме съобщение за отхвърляне, ако избраният обект не е текст
            prEntOpts.SetRejectMessage(vbCrLf & "Това не е текст!")
            ' Добавяме DBText и MText към разрешените класове
            prEntOpts.AddAllowedClass(GetType(DBText), True)    ' Избира само DBText
            prEntOpts.AddAllowedClass(GetType(MText), True)     ' Избира само MText
            ' Извършваме избор на обект
            Dim prEntRes As PromptEntityResult = edt.GetEntity(prEntOpts)
            ' Проверяваме дали изборът е успешен
            If prEntRes.Status = PromptStatus.OK Then
                ' Получаваме идентификатора на избрания обект
                Dim id As ObjectId = prEntRes.ObjectId
                ' Проверяваме дали идентификаторът е валиден
                If id.IsNull() OrElse Not id.IsValid() Then
                    ' Извеждаме съобщение за грешка, ако идентификаторът не е валиден
                    edt.WriteMessage(vbCrLf & "Невалиден обект.")
                End If
                ' Започваме транзакция за четене на обекта
                Using trans As Transaction = acDoc.TransactionManager.StartTransaction()
                    ' Опитваме се да получим обекта като MText
                    Dim MText As MText = TryCast(trans.GetObject(id, OpenMode.ForRead), MText)
                    ' Проверяваме дали обектът е MText
                    If MText Is Nothing Then
                        ' Ако обектът не е MText, опитваме се да го получим като DBText
                        Dim Text As DBText = TryCast(trans.GetObject(id, OpenMode.ForRead), DBText)
                        ' Записваме текста на DBText обекта в резултатния низ
                        RetText = Text.TextString
                    Else
                        ' Ако обектът е MText, записваме текста му в резултатния низ
                        RetText = MText.Text
                    End If
                End Using
                RetText = ReplaceLatinWithCyrillic(RetText)
                If podpis Then RetText = ProcessString(RetText)
            End If
            edt.WriteMessage(RetText)
        Catch ex As Exception
            ' Показваме съобщение за грешка, ако такава възникне
            MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
        End Try
        Dim aaaaaa = RetText.Count
        ' Връщаме резултатния текст, след като сме заменили латинските букви с кирилица
        Return RetText
    End Function
    Public Function ProcessString(input As String) As String
        input = Trim(input)
        ' Премахване на табулации и нови редове
        input = input.Replace(vbTab, "").Replace(vbCrLf, "").Replace(vbCr, "").Replace(vbLf, "")
        ' Проверка за "инж." и "арх." в низа и премахване на всичко преди тях
        Dim engineerIndex As Integer = input.ToLower().IndexOf("инж.")
        Dim architectIndex As Integer = input.ToLower().IndexOf("арх.")
        Dim dotIndex As Integer = input.LastIndexOf(".")

        If engineerIndex + 3 = dotIndex Then
            Dim spaceIndex As Integer = input.LastIndexOf(" ")
            If spaceIndex <> -1 Then
                input = input.Remove(spaceIndex, 1).Insert(spaceIndex, "&")
            End If
        End If

        If architectIndex + 3 = dotIndex Then
            Dim spaceIndex As Integer = input.LastIndexOf(" ")
            If spaceIndex <> -1 Then
                input = input.Remove(spaceIndex, 1).Insert(spaceIndex, "&")
            End If
        End If

        If engineerIndex >= 0 Then
            input = StrConv(input, VbStrConv.ProperCase)
            input = input.Replace("Инж.", "инж.".ToLower())
            input = input.Substring(engineerIndex)
            input = input.Replace(" ", "")
        ElseIf architectIndex >= 0 Then
            input = StrConv(input, VbStrConv.ProperCase)
            input = input.Replace("Арх.", "арх.".ToLower())
            input = input.Substring(architectIndex)
            input = input.Replace(" ", "")
        End If
        ' Премахване на повече от един интервал
        input = Regex.Replace(input, "\s+", " ")
        Return input
    End Function
    ' Функцията ReplaceLatinWithCyrillic приема входен низ и връща нов низ,
    ' в който всички латински букви, които имат еквиваленти на кирилица, са заменени.
    ' Създаваме речника веднъж и го използваме многократно
    Dim latinToCyrillicMap As New Dictionary(Of Char, Char) From {
    {"A"c, "А"c}, {"a"c, "а"c},
    {"B"c, "В"c},
    {"C"c, "С"c}, {"c"c, "с"c},
    {"E"c, "Е"c}, {"e"c, "е"c},
    {"H"c, "Н"c},
    {"K"c, "К"c}, {"k"c, "к"c},
    {"M"c, "М"c}, {"m"c, "м"c},
    {"O"c, "О"c}, {"o"c, "о"c},
    {"P"c, "Р"c}, {"p"c, "р"c},
    {"T"c, "Т"c},
    {"X"c, "Х"c}, {"x"c, "х"c},
    {"Y"c, "У"c}, {"y"c, "у"c}
}
    Function ReplaceLatinWithCyrillic(input As String) As String
        ' Инициализираме StringBuilder с начална дължина, равна на дължината на входния низ.
        Dim result As New StringBuilder(input.Length)
        For Each c As Char In input
            Dim cyrillicChar As Char
            ' Ако символът е в речника, добавяме съответния му кирилица символ към резултата.
            If latinToCyrillicMap.TryGetValue(c, cyrillicChar) Then
                result.Append(cyrillicChar)
            Else
                ' Ако символът не е в речника, добавяме го към резултата без промяна.
                result.Append(c)
            End If
        Next
        Return result.ToString()
    End Function
    Public Sub FilterSelectionSet()
        '' Get the current document editor
        Dim acDocEd As Editor = Application.DocumentManager.MdiActiveDocument.Editor
        '' Create a TypedValue array to define the filter criteria
        Dim acTypValAr(0) As TypedValue
        acTypValAr.SetValue(New TypedValue(DxfCode.Start, "CIRCLE"), 0)
        '' Assign the filter criteria to a SelectionFilter object
        Dim acSelFtr As SelectionFilter = New SelectionFilter(acTypValAr)
        '' Request for objects to be selected in the drawing area
        Dim acSSPrompt As PromptSelectionResult
        acSSPrompt = acDocEd.GetSelection(acSelFtr)
        '' If the prompt status is OK, objects were selected
        If acSSPrompt.Status = PromptStatus.OK Then
            Dim acSSet As SelectionSet = acSSPrompt.Value
            Application.ShowAlertDialog("Number of objects selected: " &
                                              acSSet.Count.ToString())
        Else
            Application.ShowAlertDialog("Number of objects selected: 0")
        End If
    End Sub
    Public Function InsertBlock(BlockName As String,    ' име на блока който ще се вмъква
                                InsertPoint As Point3d, ' точка на вмъкване
                                layer As String,        ' слой на който ще се вмъква
                                Scale As Scale3d        ' Скала на блока
                                ) As ObjectId
        ' Get the current database and start a transaction
        Dim acCurDb As Autodesk.AutoCAD.DatabaseServices.Database
        acCurDb = Application.DocumentManager.MdiActiveDocument.Database

        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            ' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim blkRecId As ObjectId = ObjectId.Null
            blkRecId = acBlkTbl(BlockName)
            Dim blkRecIdNow As ObjectId = ObjectId.Null

            If blkRecId = ObjectId.Null Then
                acTrans.Commit()
                Return blkRecIdNow
                Exit Function
            End If

            ' Insert the block into the current space
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(blkRecId, OpenMode.ForRead)

            Using acBlkRef As New BlockReference(InsertPoint, blkRecId)
                Dim acCurSpaceBlkTblRec As BlockTableRecord
                acCurSpaceBlkTblRec = acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite)
                acCurSpaceBlkTblRec.AppendEntity(acBlkRef)
                acTrans.AddNewlyCreatedDBObject(acBlkRef, True)
                blkRecIdNow = acBlkRef.ObjectId
                acBlkRef.Layer = layer
                ' Verify block table record has attribute definitions associated with it

                acBlkRef.ScaleFactors = Scale

                If Not acBlkTblRec.HasAttributeDefinitions Then
                    acTrans.Commit()
                    Return blkRecIdNow
                    Exit Function
                End If
                ' Add attributes from the block table record
                For Each objID As ObjectId In acBlkTblRec
                    Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                    If TypeOf dbObj Is AttributeDefinition Then
                        Dim acAtt As AttributeDefinition = dbObj
                        If Not acAtt.Constant Then
                            Using acAttRef As New AttributeReference
                                acAttRef.SetAttributeFromBlock(acAtt, acBlkRef.BlockTransform)
                                acAttRef.Position = acAtt.Position.TransformBy(acBlkRef.BlockTransform)
                                acAttRef.TextString = acAtt.TextString
                                acBlkRef.AttributeCollection.AppendAttribute(acAttRef)
                                acTrans.AddNewlyCreatedDBObject(acAttRef, True)
                            End Using
                        End If
                    End If
                Next
            End Using
            ' Save the new object to the database
            acTrans.Commit()
            Return blkRecIdNow
        End Using
    End Function
    Public Function line_Layer(Layer As String) As String
        Dim laye As String = ""
        Select Case Layer
            Case "EL_2x1_5"
                laye = "СВТ2x1,5mm²"
            Case "EL_2x10"
                laye = "САВТ2x10mm²"
            Case "EL_2x16"
                laye = "САВТ2x16mm²"
            Case "EL_2x25"
                laye = "САВТ2x25mm²"
            Case "EL_2x35"
                laye = "САВТ2x35mm²"
            Case "EL_2x2_5"
                laye = "СВТ2x2.5mm²"
            Case "EL_2x4"
                laye = "СВТ2x4,0mm²"
            Case "EL_2x6"
                laye = "САВТ2x6mm²"
            Case "EL_3x1_5"
                laye = "СВТ3x1,5mm²"
            Case "EL_3x10"
                laye = "СВТ3x10mm²"
            Case "EL_3x16"
                laye = "САВТ3x16mm²"
            Case "EL_3x2_5"
                laye = "СВТ3x2,5mm²"
            Case "EL_3x4"
                laye = "СВТ3x4,0mm²"
            Case "EL_3x6"
                laye = "СВТ3x6,0mm²"
            Case "EL_4x1_5"
                laye = "СВТ4x1,5mm²"
            Case "EL_4x10"
                laye = "СВТ4x10mm²"
            Case "EL_4x16"
                laye = "САВТ4x16mm²"
            Case "EL_4x2_5"
                laye = "СВТ4x2,5mm²"
            Case "EL_4x4"
                laye = "СВТ4x4,0mm²"
            Case "EL_4x6"
                laye = "СВТ4x6,0mm²"
            Case "EL_4x25"
                laye = "САВТ4x25,0mm²"
            Case "EL_4x35"
                laye = "САВТ4x35,0mm²"
            Case "EL_5x1_5"
                laye = "СВТ5x1,5mm²"
            Case "EL_3х25+16"
                laye = "САВТ3x25+16mm²"
            Case "EL_3х35+16"
                laye = "САВТ3x35+16mm²"
            Case "EL_3х50+25"
                laye = "САВТ3x50+25mm²"
            Case "EL_3х70+35"
                laye = "САВТ3x70+35mm²"
            Case "EL_3х95+50"
                laye = "САВТ3x95+50mm²"
            Case "EL_3х120+70"
                laye = "САВТ3x120+70mm²"
            Case "EL_3х150+70"
                laye = "САВТ3x150+70mm²"
            Case "EL_3х185+95"
                laye = "САВТ3x185+95mm²"
            Case "EL_3х240+120"
                laye = "САВТ3x240+120mm²"
            Case "EL_5x10"
                laye = "СВТ5x10mm²"
            Case "EL_5x16"
                laye = "СВТ5x16mm²"
            Case "EL_5x2_5"
                laye = "СВТ5x2,5mm²"
            Case "EL_5x25"
                laye = "СВТ5x25mm²"
            Case "EL_5x35"
                laye = "СВТ5x35mm²"
            Case "EL_5x4"
                laye = "СВТ5x4,0mm²"
            Case "EL_5x6"
                laye = "СВТ5x6,0mm²"
            Case "EL_5x10"
                laye = "СВТ5x10,0mm²"
            Case "EL_5x25"
                laye = "СВТ5x25,0mm²"
            Case "EL_5x35"
                laye = "СВТ5x35,0mm²"
            Case "EL_6BPL2x1_0"
                laye = "ШВПЛ-А 2х1,00mm²"
            Case "EL_HDMI"
                laye = "сиг.каб. HDMI"
            Case "EL_JC_2x05_sig", "EL_JC_2x05_zahr", "EL_JC_2x15_А", "EL_JC_2x15_Б"
                laye = "J-Y(St)YFR 2x0.80mm²"
            Case "EL_JC_2x05"
                laye = "FS 2x0,50mm²"
            Case "EL_JC_2x10"
                laye = "FS 2x1,00mm²"
            Case "EL_JC_2x75"
                laye = "FS 2x0,75mm²"
            Case "EL_JC_3x05"
                laye = "FS 3x0,50mm²"
            Case "EL_JC_3x10"
                laye = "FS 3x1,00mm²"
            Case "EL_JC_3x75"
                laye = "FS 3x0,75mm²"
            Case "EL_JC_4x05"
                laye = "FS 4x0,50mm²"
            Case "EL_JC_4x10"
                laye = "FS 4x1,00mm²"
            Case "EL_JC_4x75"
                laye = "FS 4x0,75mm²"
            Case "EL_NHXCH EF 2x15"
                laye = "NHXH FE180/Е30 2x1,5mm²"
            Case "EL_NHXCH EF 3x15"
                laye = "NHXH FE180/Е30 3x1,5mm²"
            Case "EL_NHXCH EF 3x25"
                laye = "NHXH FE180/Е30 3x2,5mm²"
            Case "EL_NHXCH EF 3x40"
                laye = "NHXH FE180/Е30 3x4,0mm²"
            Case "EL_NHXCH EF 3x60"
                laye = "NHXH FE180/Е30 3x6,0mm²"
            Case "EL_NHXCH EF 4x15"
                laye = "NHXH FE180/Е30 4x1,5mm²"
            Case "EL_NHXCH EF 4x60"
                laye = "NHXH FE180/Е30 4x6,0mm²"
            Case "EL_NHXCH EF 5x15"
                laye = "NHXH FE180/Е30 5x1,5mm²"
            Case "EL_NHXCH EF 5x25"
                laye = "NHXH FE180/Е30 5x2,5mm²"
            Case "EL_NHXCH EF 5x40"
                laye = "NHXH FE180/Е30 5x4,0mm²"
            Case "EL_NHXCH EF 5x60"
                laye = "NHXH FE180/Е30 5x6,0mm²"
            Case "EL_SOT"
                laye = "CAB/6/WH 6x25SWG"
            Case "EL_Tel"
                laye = "ПТПВ 1x2x0,5мм²"
            Case "EL_TV"
                laye = "коакс.кабел RG6/64"
            Case "EL_UTP", "EL_Video_FTP", "EL_DOMO"
                laye = "FTP 4x2x24AWG, cat. 5e"
            Case "EL_Video_RG59CU"
                laye = "RG59CU + 2x0.75mm²"
            Case "EL_PB_1x16"
                laye = "ПВ-A2 1x16mm²"
            Case "EL_PB_1x1_5"
                laye = "ПВ-A1 1x1,5mm²"
            Case "EL_UK_2x16"
                laye = "Al/R 2x16,0mm²"
            Case "EL_UK_2x25"
                laye = "Al/R 2x25,0mm²"
            Case "EL_UK_4x16"
                laye = "Al/R 4x16,0mm²"
            Case "EL_UK_4x25"
                laye = "Al/R 4x25,0mm²"
            Case "EL_UK_4x35"
                laye = "Al/R 4x35,0mm²"
            Case "EL_UK_4x50"
                laye = "Al/R 3x50,0+54,6mm²"
            Case "EL_UK_4x70"
                laye = "Al/R 3x70,0+54,6mm²"
            Case "EL_UK_4x95"
                laye = "Al/R 3x95,0+70,0mm²"
            Case "EL_UK_4x120"
                laye = "Al/R 3x120,0+70,0mm²"
            Case "EL_UK_4x150"
                laye = "Al/R 3x150,0+70,0mm²"
            Case "EL_AlMgSi Ф8мм"
                laye = "AlMgSi Ф8мм"
            Case "EL_Шина"
                laye = "поц.шина 40x4mm"
            Case "EL_3х25+16"
                laye = "САВТ3x25+16mm²"
            Case "EL_3х35+16"
                laye = "САВТ3x35+16mm²"
            Case "EL_3х50+25"
                laye = "САВТ3x50+25mm²"
            Case "EL_3х70+35"
                laye = "САВТ3x70+35mm²"
            Case "EL_3х95+50"
                laye = "САВТ3x95+50mm²"
            Case "EL_3х120+70"
                laye = "САВТ3x120+70mm²"
            Case "EL_3х150+70"
                laye = "САВТ3x150+70mm²"
            Case "EL_3х185+95"
                laye = "САВТ3x185+95mm²"
            Case "EL_3х240+120"
                laye = "САВТ3x240+120mm²"
            Case "EL_стринг",
                 "EL_стринг1", "EL_стринг2", "EL_стринг3", "EL_стринг4",
                 "EL_стринг5", "EL_стринг6", "EL_стринг7", "EL_стринг8",
                 "EL_стринг9", "EL_стринг10", "EL_стринг11", "EL_стринг12",
                 "EL_стринг13", "EL_стринг14", "EL_стринг15", "EL_стринг16",
                 "EL_стринг17", "EL_стринг18", "EL_стринг19", "EL_стринг20",
                 "EL_стринг21"
                laye = "H1Z2Z2-K 1/1.5kV 1x6мм²"
            Case "EL_САХЕКТ 35"
                laye = "3xСАХЕк(вн)П 1x35mm²"
            Case "EL_САХЕКТ 50"
                laye = "3xСАХЕк(вн)П 1x50mm²"
            Case "EL_САХЕКТ 70"
                laye = "3xСАХЕк(вн)П 1x70mm²"
            Case "EL_САХЕКТ 95"
                laye = "3xСАХЕк(вн)П 1x95mm²"
            Case "EL_САХЕКТ 120"
                laye = "3xСАХЕк(вн)П 1x120mm²"
            Case "EL_САХЕКТ 150"
                laye = "3xСАХЕк(вн)П 1x150mm²"
            Case "EL_САХЕКТ 185"
                laye = "3xСАХЕк(вн)П 1x185mm²"
            Case "EL_САХЕКТ 240"
                laye = "3xСАХЕк(вн)П 1x240mm²"
            Case "EL_САХЕКТ 300"
                laye = "3xСАХЕк(вн)П 1x300mm²"
            Case "EL_САХЕКТ 400"
                laye = "3xСАХЕк(вн)П 1x400mm²"
            Case "EL_САХЕКТ 500"
                laye = "3xСАХЕк(вн)П 1x500mm²"
            Case "EL_Траншея_80см"
                laye = "80"
            Case "EL_Траншея_110см"
                laye = "110"
            Case "EL_Траншея_130см"
                laye = "130"
            Case "EL_ТЧ_2x05"
                laye = "ТЧ 2x0,5mm²"
            Case "EL_ТЧ_2x10"
                laye = "ТЧ 2x1,0mm²"
            Case "EL_ТЧ_2x15"
                laye = "ТЧ 2x1,5mm²"
            Case "EL_OPTIC"
                laye = "Оптичен кабел"
            Case "EL_РЕЗЕРВА"
                laye = "Резервна тръба"
            Case Else
                laye = Layer
        End Select

        Return laye
    End Function
    Public Function GET_line_Type(Type_Line As String,
                                  FullName As Boolean
                                  ) As String
        '
        'FullName - vbFalse - Функцията връща само външния диаметър
        '         - vbTrue - Функцията връща и вътрешния диаметър
        '
        Dim lType As String = ""
        Select Case Type_Line
            Case "PE Ф18"
                lType = "изт. в PE тр.ф18,7/13,5mm"
            Case "PE Ф21"
                lType = "изт. в PE тр.ф21,2/16mm"
            Case "PE Ф28"
                lType = "изт. в PE тр.ф28,5/22,9mm"
            Case "PE Ф34"
                lType = "изт. в PE тр.ф34,5/28,4mm"
            Case "PE Ф46"
                lType = "изт. в PE тр.ф46,5/35,9mm"
            Case "PVC Ф16"
                lType = "изт. в PVC тр.ф16,0/11,3mm"
            Case "PVC Ф20"
                lType = "изт. в PVC тр.ф20,0/14,6mm"
            Case "PVC Ф25"
                lType = "изт. в PVC тр.ф25,0/18,5mm"
            Case "PVC Ф32"
                lType = "изт. в PVC тр.ф32,0/24,3mm"
            Case "PVC Ф40"
                lType = "изт. в PVC тр.ф40,0/31,2mm"
            Case "PVC Ф50"
                lType = "изт. в PVC тр.ф50,0/39,6mm"
            Case "PVC Ф75"
                lType = "изт. в PVC тр.ф75,0/71,4mm"
            Case "PVC Ф110"
                lType = "изт. в PVC тр.ф110/103,6mm"
            Case "PVC Ф140"
                lType = "изт. в PVC тр.ф140/134,4mm"
            Case "PVC Ф160"
                lType = "изт. в PVC тр.ф160/155,0mm"
            Case "НГPVC Ф16"
                lType = "изт. в негор.PVC тр.ф16,0/10,7mm"
            Case "НГPVC Ф20"
                lType = "изт. в негор.PVC тр.ф20,0/14,1mm"
            Case "НГPVC Ф25"
                lType = "изт. в негор.PVC тр.ф25,0/18,2mm"
            Case "НГPVC Ф32"
                lType = "изт. в негор.PVC тр.ф32,0/24,3mm"
            Case "НГPVC Ф40"
                lType = "изт. в негор.PVC тр.ф40,0/32,3mm"
            Case "НГМ Ф09"
                lType = "изт. в мет.тр.ф13,2/9,0mm"
            Case "НГМ Ф11"
                lType = "изт. в мет.тр.ф15,2/11,0mm"
            Case "НГМ Ф14"
                lType = "изт. в мет.тр.ф18,4/14,0mm"
            Case "НГМ Ф18"
                lType = "изт. в мет.тр.ф22,4/18,0mm"
            Case "НГМ Ф26"
                lType = "изт. в мет.тр.ф30,4/26,0mm"
            Case "НГМ Ф37"
                lType = "изт. в мет.тр.ф42,4/37,0mm"
            Case "Въже"
                lType = "пол. по носещо въже"
            Case "Мазилка"
                lType = "скрито под мазилката"
            Case "ПКОМ"
                lType = "открито на ПКОМ скоби"
            Case "Скара"
                lType = "пол. по кабелна скара"
            Case "Таван"
                lType = "пол. по кабелна скара"
            Case "ИЗКОП"
                lType = "пол. в изкоп"
            Case "БЕТОН"
                lType = "пол. в бет. кожух"
            Case "по конструкция"
                lType = "открито по метална конструкция"
            Case "крепежни елементи"
                lType = "крепежни елементи"
            Case "ВЪЗДУШНО"
                lType = "изт. въздушно"
            Case "HDPE Ф40"
                lType = "изт. в HDPE тр.ф40/32mm"
            Case "HDPE Ф50"
                lType = "изт. в HDPE тр.ф50/41mm"
            Case "HDPE Ф63"
                lType = "изт. в HDPE тр.ф63/53mm"
            Case "HDPE Ф75"
                lType = "изт. в HDPE тр.ф75/61mm"
            Case "HDPE Ф90"
                lType = "изт. в HDPE тр.ф90/75mm"
            Case "HDPE Ф110"
                lType = "изт. в HDPE тр.ф110/94mm"
            Case "HDPE Ф125"
                lType = "изт. в HDPE тр.ф125/108mm"
            Case "HDPE Ф140"
                lType = "изт. в HDPE тр.ф140/121mm"
            Case "HDPE Ф160"
                lType = "изт. в HDPE тр.ф160/136mm"
            Case "HDPE Ф200"
                lType = "изт. в HDPE тр.ф200/170mm"
            Case "Траншея-500"
                lType = "50"
            Case "Траншея-800"
                lType = "80"
            Case "Траншея-1000"
                lType = "100"
            Case "Траншея-1200"
                lType = "120"
            Case Else
                If Mid(Type_Line, 1, 2) = "КК" Then
                    lType = "изт. в каб.кан." + Mid(Type_Line, 4, Len(Type_Line)) + "mm"
                Else
                    lType = "ByLayer"
                End If
        End Select
        If Not FullName Then
            If InStr(lType, "/") > 0 Then
                lType = Mid(lType, 1, InStr(lType, "/") - 1) & "mm"
            End If
        End If
        Return lType
    End Function
    Public Function GET_line_Diamet(Type_Line As String) As Double
        Dim Diam As Double = -1
        Select Case Type_Line
            Case "EL_1x1_5"
                Diam = 7.2
            Case "EL_1x2_5"
                Diam = 7.6
            Case "EL_1x4"
                Diam = 8
            Case "EL_1x6"
                Diam = 8.5
            Case "EL_1x10"
                Diam = 9.3
            Case "EL_1x16"
                Diam = 10.8
            Case "EL_1x25"
                Diam = 11.9
            Case "EL_1x35"
                Diam = 13
            Case "EL_1x50"
                Diam = 14.2
            Case "EL_1x70"
                Diam = 16.7
            Case "EL_1x95"
                Diam = 18.4
            Case "EL_1x120"
                Diam = 19.9
            Case "EL_1x150"
                Diam = 21.9
            Case "EL_1x185"
                Diam = 24.3
            Case "EL_1x240"
                Diam = 27.2
            Case "EL_1x300"
                Diam = 30.3
            Case "EL_1x400"
                Diam = 33.8
            Case "EL_1x500"
                Diam = 37.9
            Case "EL_2x1_0"
                Diam = 9.5
            Case "EL_2x1_5"
                Diam = 10
            Case "EL_2x2_5"
                Diam = 10.6
            Case "EL_2x4"
                Diam = 12.5
            Case "EL_2x6"
                Diam = 13.3
            Case "EL_2x10"
                Diam = 15.2
            Case "EL_2x16"
                Diam = 18
            Case "EL_2x25"
                Diam = 21.5
            Case "EL_2x35"
                Diam = 23.8
            Case "EL_2x50"
                Diam = 27
            Case "EL_3x1_0"
                Diam = 9.7
            Case "EL_3x1_5"
                Diam = 10.2
            Case "EL_3x2_5"
                Diam = 11
            Case "EL_3x4"
                Diam = 13
            Case "EL_3x6"
                Diam = 14
            Case "EL_3x10"
                Diam = 16
            Case "EL_3x16"
                Diam = 19.5
            Case "EL_3x25"
                Diam = 22.8
            Case "EL_3x35"
                Diam = 25.3
            Case "EL_3x50"
                Diam = 28.8
            Case "EL_3x70"
                Diam = 28.8
            Case "EL_3x95"
                Diam = 33.1
            Case "EL_3x120"
                Diam = 35.9
            Case "EL_3x150"
                Diam = 39.3
            Case "EL_3x185"
                Diam = 43.8
            Case "EL_3x240"
                Diam = 49.4
            Case "EL_3х25+16"
                Diam = 24.4
            Case "EL_3х35+16"
                Diam = 27.1
            Case "EL_3х50+25"
                Diam = 30.8
            Case "EL_3х70+35"
                Diam = 31.8
            Case "EL_3х95+50"
                Diam = 36.8
            Case "EL_3х120+70"
                Diam = 40.1
            Case "EL_3х150+70"
                Diam = 44.3
            Case "EL_3х185+95"
                Diam = 49.3
            Case "EL_3х240+120"
                Diam = 55.4
            Case "EL_4x1_5"
                Diam = 11
            Case "EL_4x2_5"
                Diam = 11.9
            Case "EL_4x4"
                Diam = 14.1
            Case "EL_4x6"
                Diam = 15.4
            Case "EL_4x10"
                Diam = 17.4
            Case "EL_4x16"
                Diam = 20.6
            Case "EL_4x25"
                Diam = 25.2
            Case "EL_4x35"
                Diam = 28
            Case "EL_4x50"
                Diam = 32
            Case "EL_4x70"
                Diam = 33
            Case "EL_4x95"
                Diam = 38.1
            Case "EL_4x120"
                Diam = 41.4
            Case "EL_4x150"
                Diam = 46
            Case "EL_4x185"
                Diam = 51.1
            Case "EL_4x240"
                Diam = 58
            Case "EL_5x1_5"
                Diam = 11.8
            Case "EL_5x2_5"
                Diam = 12.8
            Case "EL_5x4"
                Diam = 15.5
            Case "EL_5x6"
                Diam = 16.8
            Case "EL_5x10"
                Diam = 19.2
            Case "EL_5x16"
                Diam = 23.2
            Case "EL_5x25"
                Diam = 27.9
            Case "EL_5x35"
                Diam = 32.4
            Case "EL_5x50"
                Diam = 36.9
            Case "EL_5x70"
                Diam = 42
            Case "EL_5x95"
                Diam = 48.7
            Case "EL_JC_2x05_sig", "EL_JC_2x05_zahr", "EL_JC_2x15_А", "EL_JC_2x15_Б",
                 "EL_JC_2x05", "EL_JC_2x10", "EL_JC_2x75",
                 "EL_JC_3x05", "EL_JC_3x10", "EL_JC_3x75",
                 "EL_JC_4x05", "EL_JC_4x10", "EL_JC_4x75",
                 "EL_SOT", "EL_Tel", "EL_TV", "EL_UTP", "EL_DOMO",
                 "EL_Video", "EL_Video_FTP", "EL_Video_RG59CU",
                 "EL_HDMI", "EL_6BPL2x1_0",
                 "EL_ТЧ_2x05", "EL_ТЧ_2x10", "EL_ТЧ_2x15",
                 "EL_OPTIC", "EL_РЕЗЕРВА"
                Diam = 7.2
            Case "EL_NHXCH EF 2x15"
                Diam = 14.5
            Case "EL_NHXCH EF 3x15"
                Diam = 15
            Case "EL_NHXCH EF 3x25"
                Diam = 16
            Case "EL_NHXCH EF 3x40"
                Diam = 17
            Case "EL_NHXCH EF 3x60"
                Diam = 17.9
            Case "EL_NHXCH EF 4x15"
                Diam = 15
            Case "EL_NHXCH EF 4x60"
                Diam = 19
            Case "EL_NHXCH EF 5x15"
                Diam = 14.5
            Case "EL_NHXCH EF 5x25"
                Diam = 16
            Case "EL_NHXCH EF 5x40"
                Diam = 16
            Case "EL_NHXCH EF 5x60"
                Diam = 19
            Case "EL_PB_1x16"
                Diam = 8.8
            Case "EL_PB_1x1_5"
                Diam = 3.3
            Case "EL_стринг",
                 "EL_стринг1", "EL_стринг2", "EL_стринг3", "EL_стринг4",
                 "EL_стринг5", "EL_стринг6", "EL_стринг7", "EL_стринг8",
                 "EL_стринг9", "EL_стринг10", "EL_стринг11", "EL_стринг12",
                 "EL_стринг13", "EL_стринг14", "EL_стринг15", "EL_стринг16",
                 "EL_стринг17", "EL_стринг18", "EL_стринг19", "EL_стринг20",
                  "EL_стринг21"
                Diam = 6.1
            Case "EL_САХЕКТ 35"
                Diam = 29
            Case "EL_САХЕКТ 50"
                Diam = 30
            Case "EL_САХЕКТ 70"
                Diam = 32
            Case "EL_САХЕКТ 95"
                Diam = 33
            Case "EL_САХЕКТ 120"
                Diam = 35
            Case "EL_САХЕКТ 150"
                Diam = 36
            Case "EL_САХЕКТ 185"
                Diam = 38
            Case "EL_САХЕКТ 240"
                Diam = 41
            Case "EL_САХЕКТ 300"
                Diam = 43
            Case "EL_САХЕКТ 400"
                Diam = 46
            Case "EL_САХЕКТ 500"
                Diam = 49
            Case "EL_Шина"
                Diam = 4
            Case Else
        End Select
        Return Diam
    End Function
    Public Function SET_line_Type(Layer_Line As String) As String
        Dim Line_Type As String = ""
        Select Case Layer_Line
            Case "EL_2x1_5", "EL_2x2_5",
                 "EL_3x1_5", "EL_3x2_5",
                 "EL_4x1_5", "EL_4x2_5",
                 "EL_5x1_5", "EL_5x2_5"
                Line_Type = "PVC Ф20"
            Case "EL_2x4", "EL_2x6",
                 "EL_3x4", "EL_3x6",
                 "EL_4x4", "EL_4x6",
                 "EL_5x4", "EL_5x6"
                Line_Type = "PVC Ф25"
            Case "EL_2x10", "EL_2x16",
                 "EL_3x10", "EL_3x16",
                 "EL_4x10", "EL_4x16",
                 "EL_5x10", "EL_5x16"
                Line_Type = "PVC Ф32"
            Case " "
                Line_Type = "PVC Ф50"
            Case "EL_5x25", "EL_5x35",
                 "EL_2x25", "EL_2x35",
                 "EL_4x25", "EL_4x35",
                 "EL_3х25+16", "EL_3х35+16",
                 "EL_3х25+16", "EL_3х35+16"
                Line_Type = "PVC Ф55"
            Case "EL_HDMI"
                Line_Type = "PVC Ф40"
            Case "EL_JC_2x05", "EL_JC_2x10", "EL_JC_2x75",
                 "EL_JC_3x05", "EL_JC_3x10", "EL_JC_3x75",
                 "EL_JC_4x05", "EL_JC_4x10", "EL_JC_4x75",
                 "EL_SOT", "EL_Tel",
                 "EL_TV", "EL_UTP", "EL_Video_FTP", "EL_Video_RG59CU",
                 "EL_6BPL2x1_0",
                 "EL_DOMO",
                 "EL_ТЧ_2x05", "EL_ТЧ_2x10", "EL_ТЧ_2x15",
                 "EL_OPTIC"
                Line_Type = "PVC Ф16"
            Case "EL_JC_2x05_sig", "EL_JC_2x05_zahr", "EL_JC_2x15_А", "EL_JC_2x15_Б"
                Line_Type = "НГPVC Ф16"
            Case "EL_NHXCH EF 2x15", "NHXH FE180/Е30 2x1,5mm²", "EL_NHXCH EF 3x15",
                 "EL_NHXCH EF 3x25", "EL_NHXCH EF 4x15"
                Line_Type = "НГМ Ф18"
            Case "EL_NHXCH EF 3x40", "EL_NHXCH EF 3x60", "EL_NHXCH EF 5x15", "EL_NHXCH EF 5x25"
                Line_Type = "НГМ Ф26"
            Case "EL_NHXCH EF 4x60", "EL_NHXCH EF 5x40", "EL_NHXCH EF 5x60"
                Line_Type = "НГМ Ф37"
            Case "EL_PB_1x16", "EL_PB_1x1_5"
                Line_Type = "PVC Ф16"
            Case "EL_UK_2x16", "EL_UK_2x25", "EL_UK_4x16", "EL_UK_4x25", "EL_UK_4x35",
                 "EL_UK_4x50", "EL_UK_4x70", "EL_UK_4x95", "EL_UK_4x120", "EL_UK_4x150"
                Line_Type = "ВЪЗДУШНО"
            Case "EL_AlMgSi Ф8мм", "EL_Шина"
                Line_Type = "крепежни елементи"
            Case "EL_стринг",
                 "EL_стринг1", "EL_стринг2", "EL_стринг3", "EL_стринг4",
                 "EL_стринг5", "EL_стринг6", "EL_стринг7", "EL_стринг8",
                 "EL_стринг9", "EL_стринг10", "EL_стринг11", "EL_стринг12",
                 "EL_стринг13", "EL_стринг14", "EL_стринг15", "EL_стринг16",
                 "EL_стринг17", "EL_стринг18", "EL_стринг19", "EL_стринг20",
                 "EL_стринг21"
                Line_Type = "по конструкция"
            Case Else
                Line_Type = "ByLayer"
        End Select
        Return Line_Type
    End Function
    ' Описание:
    ' Тази функция приема масив (`Kabel`), SelectionSet (`ss`) и булев флаг (`Length`) като входни данни.
    ' Функцията анализира избраните линии (`ss`) и категоризира типовете кабели, като записва резултатите в масива `Kabel`.
    Public Function GET_LINE_TYPE_KABEL(Kabel As Array,     ' Масив в който ще върне линиите
                                    ss As SelectionSet,     ' Избрани линии
                                    Length As Boolean       ' Какво да въща в Kabel(*, 2) ' True - Връща брой линии // False - Връща дължина                                                            
                                    ) As Array
        ' Входни данни:
        ' * `Kabel`: Масив, в който ще се съхраняват резултатите.
        ' * `ss`: SelectionSet, съдържащ избраните линии.
        ' * `Length`: Булев флаг, указващ какъв тип информация да се върне в `Kabel(*, 2)`.
        '     * `True`: Връща броя на линиите за всеки тип кабел.
        '     * `False`: Връща общата дължина за всеки тип кабел.

        ' Изходни данни:
        ' * Масив `Kabel`, съдържащ следните колони:
        '     * Колонка 0: Описание на типа кабел (напр. "EL_стринг", "EL_JC_2x15_А").
        '     * Колонка 1: Описание на типа линия (напр. "Плътна", "Щрихова").
        '     * Колонка 2: Брой линии или обща дължина, в зависимост от флага `Length`.


        ' Вземане на текущия документ
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        ' Инициализация на масив arrBlock от тип strLine с дължина 600
        Dim arrBlock(600) As strLine
        ' Инициализация на индекс за броене на уникални линии
        Dim index As Integer = 0
        ' Започване на транзакция
        Using trans As Transaction = doc.TransactionManager.StartTransaction()
            ' Обхождане на всеки избран обект
            For Each sObj As SelectedObject In ss
                ' Проверка дали обектът е от тип Line
                Dim line As Line = TryCast(trans.GetObject(sObj.ObjectId, OpenMode.ForRead), Line)
                Dim iVisib As Integer = -1
                ' Пропускане ако типът на линията е празен
                If line.Linetype = "" Then Continue For
                Dim Line_Layer As String = ""
                ' Определяне на слоя на линията
                Select Case line.Layer
                    Case "EL_стринг1", "EL_стринг2", "EL_стринг3", "EL_стринг4", "EL_стринг5", "EL_стринг6", "EL_стринг7",
                     "EL_стринг8", "EL_стринг9", "EL_стринг10", "EL_стринг11", "EL_стринг12", "EL_стринг13", "EL_стринг14",
                     "EL_стринг15", "EL_стринг16", "EL_стринг17", "EL_стринг18", "EL_стринг19", "EL_стринг20",
                     "EL_стринг21"
                        Line_Layer = "EL_стринг"
                    Case "EL_JC_2x05_sig", "EL_JC_2x05_zahr", "EL_JC_2x15_А", "EL_JC_2x15_Б"
                        Line_Layer = "EL_JC_2x15_А"
                    Case Else
                        Line_Layer = line.Layer
                End Select
                ' Проверка дали комбинацията слой-тип линия вече съществува в arrBlock
                iVisib = Array.FindIndex(arrBlock, Function(f) f.Layer = Line_Layer And f.Linetype = line.Linetype)
                If iVisib = -1 Then
                    ' Добавяне на нова комбинация ако не съществува
                    arrBlock(index).Layer = Line_Layer
                    arrBlock(index).Linetype = line.Linetype
                    arrBlock(index).count = IIf(Length, 1, line.Length)
                    index += 1
                Else
                    ' Актуализиране на дължината или броя ако комбинацията съществува
                    arrBlock(iVisib).count += IIf(Length, 1, line.Length)
                End If
            Next
            ' Завършване на транзакцията
            trans.Commit()
        End Using
        ' Проверка дали броят на уникалните линии е по-голям от допустимия размер на Kabel
        If UBound(Kabel) <= index Then
            MsgBox("Типовете кабели е по-голям от допустимия!")
        End If
        ' Запълване на масива Kabel
        For i = 0 To UBound(Kabel)
            If arrBlock(i).Layer = "" Then Exit For
            'If arrBlock(i).Linetype = "ByLayer" Then Continue For
            Kabel(i, 0) = line_Layer(arrBlock(i).Layer)
            Kabel(i, 1) = GET_line_Type(arrBlock(i).Linetype, IIf(Length, vbFalse, vbTrue))
            Kabel(i, 2) = arrBlock(i).count
        Next
        ' Връщане на резултата
        Return Kabel
    End Function
    ' Функцията GET_Zazemlenie обработва избрани обекти (блокове) в чертеж и извлича информация за тях. Функцията приема като параметър SelectedSet, който представлява набора от избрани блокове в чертежа за външно захранване. Функцията връща масив от структури strZazeml, които съдържат информация за всеки блок, включително видимост, име, броя на срещанията на блока, както и допълнителни атрибути като "ТАБЛО" и "Надпис".
    Public Function GET_Zazemlenie(SelectedSet As SelectionSet) As Array
        ' Извличане на активния документ и неговия редактор
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim edt As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database
        Dim blkRecId As ObjectId = ObjectId.Null
        Dim index As Integer = 0
        Dim arrBlock(100) As strZazeml
        ' Започване на транзакция
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                ' Обхождане на всеки избран обект в набора
                For Each sObj As SelectedObject In SelectedSet
                    blkRecId = sObj.ObjectId
                    ' Извличане на референцията към блока
                    Dim acBlkRef As BlockReference = DirectCast(acTrans.GetObject(blkRecId, OpenMode.ForRead), BlockReference)
                    Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                    Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection
                    Dim Visibility As String = ""
                    ' Извличане на свойството "Visibility" на динамичния блок
                    For Each prop As DynamicBlockReferenceProperty In props
                        If prop.PropertyName = "Visibility" Then Visibility = prop.Value
                    Next
                    Dim strТАБЛО As String = ""
                    Dim strНадпис As String = ""
                    ' Извличане на атрибутите "ТАБЛО" и "Надпис" на блока
                    For Each objID As ObjectId In attCol
                        Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                        Dim acAttRef As AttributeReference = dbObj
                        If acAttRef.Tag = "ТАБЛО" Then strТАБЛО = acAttRef.TextString
                        If acAttRef.Tag = "Надпис" Then strНадпис = acAttRef.TextString
                    Next
                    ' Извличане на името на блока
                    Dim blName As String = (CType(acBlkRef.DynamicBlockTableRecord.GetObject(OpenMode.ForRead), BlockTableRecord)).Name
                    Dim iVisib As Integer = -1
                    ' Проверка на името на блока и търсене в масива arrBlock
                    Select Case blName
                        Case "Заземление"
                            iVisib = Array.FindIndex(arrBlock, Function(f) f.blName = blName And f.blVisibility = Visibility)
                        Case "ОПЪВАЧ"
                            iVisib = Array.FindIndex(arrBlock, Function(f) f.blName = blName And f.blVisibility = Visibility)
                        Case "Траншея"
                            iVisib = Array.FindIndex(arrBlock, Function(f) f.blName = blName And f.blVisibility = Visibility)
                        Case Else
                            Continue For
                    End Select
                    ' Ако блокът не е намерен в масива, добавяне на нов запис
                    If iVisib = -1 Then
                        arrBlock(index).blVisibility = Visibility
                        arrBlock(index).blName = blName
                        arrBlock(index).count = 1
                        arrBlock(index).blТАБЛО = strТАБЛО
                        arrBlock(index).blНадпис = strНадпис
                        index += 1
                    Else
                        ' Ако блокът е намерен, увеличаване на броя на срещанията
                        arrBlock(iVisib).count = arrBlock(iVisib).count + 1
                    End If
                Next
                ' Прекратяване на транзакцията без записване на промените
                acTrans.Abort()
            Catch ex As Exception
                ' Обработка на грешка и прекратяване на транзакцията без записване на промените
                MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
                acTrans.Abort()
            End Try
        End Using
        ' Връщане на масива с резултатите
        Return arrBlock
    End Function
    ' Функцията InsertMText вмъква многострочен текст (MText) в AutoCAD чертеж на дадена точка на вмъкване, слой и с определена височина на текста. По избор може да се зададе и цвят на текста. Функцията връща ObjectId, който е идентификатор на вмъкнатия текстов обект.
    Public Function InsertMText(strMtext As String,                 ' Текст който се вмъква
                            InsertPoint As Point3d,                 ' Точка на вмъкване
                            layer As String,                        ' Слой на който се вмъква
                            dbTextHeight As Double,                 ' Височина на текста
                            Optional LineColor As Integer = 256,    ' Цвят на текста
                            Optional TextWidth As Integer = 0       ' Дължина на текста
                            ) As ObjectId

        ' Вземане на активния документ и база данни от AutoCAD.
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim docEdt As Editor = acDoc.Editor
        Dim blkRecIdNow As ObjectId = New ObjectId
        ' Започване на транзакция за работа с базата данни.
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction
            Try
                ' Вземане на таблицата с блоковете и запис на моделното пространство за запис.
                Dim bt As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim blTabRec As BlockTableRecord = acTrans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                ' Създаване на нов обект MText (многострочен текст).
                Using obMTXt As MText = New MText
                    ' Настройки на MText обекта.
                    With obMTXt
                        .TextHeight = dbTextHeight                      ' Задаване на височината на текста.
                        .Contents = strMtext                            ' Задаване на съдържанието на текста.
                        .Location = InsertPoint                         ' Задаване на точката на вмъкване.
                        .Layer = layer                                  ' Задаване на слоя.
                        If TextWidth > 0 Then
                            .Width = TextWidth
                        Else
                            ' Изчисляване и задаване на ширината на текста.
                            .Width = CalcArialMTextWidth(strMtext, dbTextHeight)
                        End If

                        .ColorIndex = LineColor                         ' Задаване на цвета на текста.

                        ' Други свойства на текста.
                        .LineWeight = LineWeight.ByLayer            ' Задаване на дебелината на линиите по слоя.
                        .Linetype = "ByLayer"                       ' Задаване на типа на линията по слоя.
                        .LinetypeScale = 1                          ' Задаване на мащаба на типа на линията.
                    End With
                    ' Добавяне на MText обекта в моделното пространство.
                    blTabRec.AppendEntity(obMTXt)
                    acTrans.AddNewlyCreatedDBObject(obMTXt, True)
                    ' Запазване на ObjectId на новия текстов обект.
                    blkRecIdNow = obMTXt.ObjectId
                End Using
                ' Комитване (запазване) на транзакцията.
                acTrans.Commit()
            Catch ex As Exception
                ' Обработка на грешка - показване на съобщение и отказ на транзакцията.
                MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
                acTrans.Abort()
            End Try
        End Using
        ' Връщане на ObjectId на нововмъкнатия текстов обект.
        Return blkRecIdNow
    End Function
    ' Функция за изчисляване на ширина на текст за Arial шрифт (приблизителна)
    Private Function CalcArialMTextWidth(strText As String, dbTextHeight As Double) As Double
        Dim maxWidth As Double = 0

        ' Разделяме текста на редове
        Dim lines() As String = strText.Split({vbCrLf, vbLf}, StringSplitOptions.None)

        For Each line As String In lines
            Dim width As Double = 0

            For Each ch As Char In line
                Select Case ch
                    Case "А"c To "Я"c, "а"c To "я"c
                        ' Кирилица
                        width += 0.6 * dbTextHeight
                    Case "A"c To "Z"c, "a"c To "z"c, "0"c To "9"c
                        ' Латиница и цифри
                        width += 0.5 * dbTextHeight
                    Case " "c
                        ' Интервал
                        width += 0.3 * dbTextHeight
                    Case Else
                        ' Специални символи и пунктуация
                        width += 0.5 * dbTextHeight
                End Select
            Next

            ' Малък марж за този ред
            width *= 1.35

            ' Проверяваме дали този ред е най-дългият
            If width > maxWidth Then
                maxWidth = width
            End If
        Next

        Return maxWidth
    End Function
    ' Функцията InsertMText вмъква многострочен текст (MText) в AutoCAD чертеж
    ' на дадена точка на вмъкване, слой и с определена височина на текста. По избор може да се зададе и цвят на текста. Функцията връща ObjectId, който е идентификатор на вмъкнатия текстов обект.
    Public Function InsertText(strText As String,               ' Текст който се вмъква
                           InsertPoint As Point3d,              ' Точка на вмъкване
                           layer As String,                     ' Слой на който се вмъква
                           Height As Double,                    ' Височина на текста
                           horAligm As TextHorizontalMode,      ' Хоризонтално подравняване
                           verAligm As TextVerticalMode,        ' Вертикално подравняване
                           Optional LineColor As Integer = 256  ' Цвят на текста, опционален параметър, по подразбиране е 256 (ByLayer)
                           ) As ObjectId
        ' Взимане на активния документ и базата данни на чертежа
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim blkRecIdNow As ObjectId = New ObjectId
        ' Започване на транзакция
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                ' Вземане на таблицата с блокове (BlockTable)
                Dim acBlkTbl As BlockTable
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                ' Отваряне на записа на блоковата таблица за Model space за писане
                Dim acBlkTblRec As BlockTableRecord
                acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                ' Създаване на нов обект DBText за текста
                Using acText As DBText = New DBText()
                    With acText
                        .SetDatabaseDefaults()                ' Настройки по подразбиране за базата данни
                        .Position = InsertPoint               ' Позиция на текста
                        .Layer = layer                        ' Слой на текста
                        .TextString = strText                 ' Текста
                        .ColorIndex = LineColor               ' Цвят на текста
                        ' Забележка: 256 е ByLayer, номера на цвета може да се вземе от таблицата на слоевете
                        .LineWeight = LineWeight.ByLayer      ' Дебелина на линията по слоя
                        .Linetype = "ByLayer"                 ' Тип на линията по слоя
                        .LinetypeScale = 1                    ' Скала на типа на линията
                        .HorizontalMode = horAligm            ' Хоризонтално подравняване
                        .VerticalMode = verAligm              ' Вертикално подравняване
                        .Height = Height                      ' Височина на текста
                    End With
                    ' Повторно задаване на позицията на текста
                    acText.Position = InsertPoint
                    ' Ако хоризонталното подравняване не е ляво, задаване на точката на подравняване
                    If acText.HorizontalMode <> TextHorizontalMode.TextLeft Then
                        acText.AlignmentPoint = InsertPoint
                    End If
                    ' Добавяне на текстовия обект в запис на блоковата таблица
                    acBlkTblRec.AppendEntity(acText)
                    acTrans.AddNewlyCreatedDBObject(acText, True)
                    blkRecIdNow = acText.ObjectId
                    ' Задаване на променлива за позицията на текста
                    Dim dddd As Point3d = acText.Position
                    dddd = dddd  ' Тази линия няма смисъл и може да бъде премахната
                End Using
                ' Потвърждаване на транзакцията
                acTrans.Commit()
            Catch ex As Exception
                ' Ако възникне грешка, показване на съобщение и анулиране на транзакцията
                MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
                acTrans.Abort()
            End Try
        End Using
        ' Връщане на идентификатора на новия текстов обект
        Return blkRecIdNow
    End Function
    'Този код създава линия в AutoCAD, като приема начална и крайна точка,
    'слой, дебелина, тип и цвят на линията като аргументи.
    Public Function DrowLine(pt1 As Point3d,                    '   Начална точка на линията
                         pt2 As Point3d,                        '   Крайна точка на линията
                         layerLine As String,                   '   Слой в който се чертае линията
                         WeightLine As LineWeight,              '   Дебелина на линия
                         LineTipe As String,                    '   Тип на линията 
                         Optional LineColor As Integer = 256    '   Цвят на линията (по подразбиране е 256, което означава "по слой")
                         ) As ObjectId
        ' Вземи текущия документ и базата данни на чертежа
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim blkRecIdNow As ObjectId = New ObjectId
        ' Започни транзакция
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            ' Отвори таблицата за типове линии за четене
            Dim acLineTypTbl As LinetypeTable
            acLineTypTbl = acTrans.GetObject(acCurDb.LinetypeTableId, OpenMode.ForRead)
            ' Провери дали съществува типът линия, ако не съществува, го зареди
            If acLineTypTbl.Has(LineTipe) = False Then
                acCurDb.LoadLineTypeFile(LineTipe, "acad.lin")
            End If
            ' Ако все още не съществува, използвай "ByLayer"
            If acLineTypTbl.Has(LineTipe) = False Then
                LineTipe = "ByLayer"
            End If
            ' Запази промените и завърши транзакцията
            acTrans.Commit()
        End Using
        ' Започни нова транзакция за създаване на линията
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                ' Отвори таблицата с блокове за четене
                Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                ' Отвори моделното пространство за писане
                Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                ' Създай нова линия
                Dim acLine As Line = New Line
                ' Настрой атрибутите на линията
                With acLine
                    .StartPoint = pt1                       ' Начална точка
                    .EndPoint = pt2                         ' Крайна точка
                    .Layer = layerLine                      ' Слой на линията
                    .ColorIndex = LineColor                 ' Цвят на линията (по подразбиране 256 = "по слой")
                    ' Забележка: 256 означава "по слой" (ByLayer),
                    ' номера на цвета може да се вземе от таблицата на слоевете
                    .LineWeight = WeightLine                ' Дебелина на линията
                    .Linetype = LineTipe                    ' Тип на линията
                    .LinetypeScale = 5                      ' Скала на тип на линията
                End With
                ' Добави линията към моделното пространство и я запази в базата данни
                acBlkTblRec.AppendEntity(acLine)
                acTrans.AddNewlyCreatedDBObject(acLine, True)
                blkRecIdNow = acLine.ObjectId               ' Запази ObjectId на новосъздадената линия
                acTrans.Commit()                            ' Запази транзакцията
            Catch ex As Exception
                ' В случай на грешка, покажи съобщение и прекрати транзакцията
                MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
                acTrans.Abort()
            End Try
        End Using
        ' Върни ObjectId на новосъздадената линия
        Return blkRecIdNow
    End Function

    <CommandMethod("TextAlignment")>
    Public Sub TextAlignment()
        '' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                               OpenMode.ForRead)

            '' Open the Block table record Model space for write
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace),
                                                  OpenMode.ForWrite)
            Dim textString(0 To 2) As String
            textString(0) = "Left"
            textString(1) = "Center"
            textString(2) = "Right"
            Dim textAlign(0 To 2) As Integer
            textAlign(0) = TextHorizontalMode.TextLeft
            textAlign(1) = TextHorizontalMode.TextCenter
            textAlign(2) = TextHorizontalMode.TextRight

            Dim acPtIns As Point3d = New Point3d(3, 3, 0)
            Dim acPtAlign As Point3d = New Point3d(3, 3, 0)
            Dim nCnt As Integer = 0
            For Each strVal As String In textString
                '' Create a single-line text object
                Dim acText As DBText = New DBText()
                acText.SetDatabaseDefaults()
                acText.Position = acPtIns
                acText.Height = 0.5
                acText.TextString = strVal
                '' Set the alignment for the text
                acText.HorizontalMode = textAlign(nCnt)
                If acText.HorizontalMode <> TextHorizontalMode.TextLeft Then
                    acText.AlignmentPoint = acPtAlign
                End If
                acBlkTblRec.AppendEntity(acText)
                acTrans.AddNewlyCreatedDBObject(acText, True)
                ' Create a point over the alignment point of the text
                'Dim acPoint As DBPoint = New DBPoint(acPtAlign)
                '' acPoint.SetDatabaseDefaults()
                'acPoint.ColorIndex = 1
                'acBlkTblRec.AppendEntity(acPoint)
                'acTrans.AddNewlyCreatedDBObject(acPoint, True)
                ' Adjust the insertion and alignment points
                acPtIns = New Point3d(acPtIns.X, acPtIns.Y + 3, 0)
                acPtAlign = acPtIns
                nCnt = nCnt + 1
            Next
            '' Set the point style to crosshair
            Application.SetSystemVariable("PDMODE", 2)
            '' Save the changes and dispose of the transaction
            acTrans.Commit()
        End Using
    End Sub
    Public Function GetAparati(SelectedSet As SelectionSet) As Array
        ' Получаване на текущия документ в AutoCAD
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim edt As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database
        ' Инициализиране на масив за съхранение на блоковете
        Dim arrBlock(500) As strТабло
        ' Променливи за съхранение на идентификатор на блока и индекс за масива
        Dim blkRecId As ObjectId = ObjectId.Null
        Dim index As Integer = 0
        ' Започване на транзакция за работа с база данни на AutoCAD
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            ' Обхождане на всеки избран обект от зададения SelectionSet
            For Each sObj As SelectedObject In SelectedSet
                ' Получаване на идентификатора на блока
                blkRecId = sObj.ObjectId
                ' Получаване на препратка към блока и неговите атрибути
                Dim acBlkRef As BlockReference = DirectCast(acTrans.GetObject(blkRecId, OpenMode.ForRead), BlockReference)
                Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                Dim blName As String = (CType(acBlkRef.DynamicBlockTableRecord.GetObject(OpenMode.ForRead), BlockTableRecord)).Name
                ' Променливи за съхранение на атрибутите на блока
                Dim iVisib As Integer = -1
                Dim strSHORTNAME As String = ""
                Dim strREFNB As String = ""
                Dim str_LONGNAME As String = ""
                Dim str_DESIGNATION As String = ""
                Dim str_1 As String = ""
                Dim str_2 As String = ""
                Dim str_3 As String = ""
                Dim str_4 As String = ""
                Dim str_5 As String = ""
                Dim str_6 As String = ""
                Dim str_7 As String = ""
                ' Обхождане на всеки атрибут на блока и съхранение на стойностите им
                For Each objID As ObjectId In attCol
                    Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                    Dim acAttRef As AttributeReference = dbObj
                    ' Съхранение на стойностите на атрибутите в съответните променливи
                    If acAttRef.Tag = "SHORTNAME" Then strSHORTNAME = acAttRef.TextString
                    If acAttRef.Tag = "REFNB" Then strREFNB = acAttRef.TextString
                    If acAttRef.Tag = "LONGNAME" Then str_LONGNAME = acAttRef.TextString
                    If acAttRef.Tag = "DESIGNATION" Then str_DESIGNATION = acAttRef.TextString
                    If acAttRef.Tag = "1" Then str_1 = acAttRef.TextString
                    If acAttRef.Tag = "2" Then str_2 = acAttRef.TextString
                    If acAttRef.Tag = "3" Then str_3 = acAttRef.TextString
                    If acAttRef.Tag = "4" Then str_4 = acAttRef.TextString
                    If acAttRef.Tag = "5" Then str_5 = acAttRef.TextString
                    If acAttRef.Tag = "6" Then str_6 = acAttRef.TextString
                    If acAttRef.Tag = "7" Then str_7 = acAttRef.TextString
                Next
                ' Проверка за специален случай в името на блока
                strSHORTNAME = IIf(Mid(strSHORTNAME, 1, 3) = "Тип", "_" + strSHORTNAME, strSHORTNAME)
                ' Проверка дали блокът вече съществува в масива
                iVisib = Array.FindIndex(arrBlock, Function(f) f.bl_ИмеБлок = blName And
                                                   f.bl_SHORTNAME = strSHORTNAME And
                                                   f.bl_DESIGNATION = str_DESIGNATION And
                                                   f.bl_Табло = strREFNB And
                                                   f.bl_1 = str_1 And
                                                   f.bl_2 = str_2 And
                                                   f.bl_3 = str_3 And
                                                   f.bl_4 = str_4 And
                                                   f.bl_5 = str_5 And
                                                   f.bl_6 = str_6 And
                                                   f.bl_7 = str_7)
                ' Ако блокът не съществува, добави го в масива
                If iVisib = -1 Then
                    arrBlock(index).bl_Брой = 1
                    arrBlock(index).bl_ИмеБлок = blName
                    arrBlock(index).bl_SHORTNAME = strSHORTNAME
                    arrBlock(index).bl_Табло = strREFNB
                    arrBlock(index).bl_REFNB = strREFNB
                    arrBlock(index).bl_LONGNAME = str_LONGNAME
                    arrBlock(index).bl_DESIGNATION = str_DESIGNATION
                    arrBlock(index).bl_1 = str_1
                    arrBlock(index).bl_2 = str_2
                    arrBlock(index).bl_3 = str_3
                    arrBlock(index).bl_4 = str_4
                    arrBlock(index).bl_5 = str_5
                    arrBlock(index).bl_6 = str_6
                    arrBlock(index).bl_7 = str_7
                    index += 1
                Else
                    ' Ако блокът съществува, увеличи броя му
                    arrBlock(iVisib).bl_Брой = arrBlock(iVisib).bl_Брой + 1
                End If
            Next
        End Using
        ' Връщане на масива с блокове
        Return arrBlock
    End Function
    ' Процедура за редактиране на динамичен блок, представящ кабел
    Public Sub EditDynamicBlockReferenceKabel(blkRecId As ObjectId)
        ' Получаване на текущия документ и база данни
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        ' Започване на транзакция за редактиране
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            ' Проверка дали ID на блока е валиден
            If blkRecId = ObjectId.Null Then
                Exit Sub
            End If
            ' Отваряне на блок референцията за писане
            Dim acBlkRef As BlockReference = DirectCast(acTrans.GetObject(blkRecId, OpenMode.ForWrite), BlockReference)
            Dim lenKabel As Double = 0
            ' Показване на таговете и стойностите на прикрепените атрибути
            Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
            For Each objID As ObjectId In attCol
                Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                Dim acAttRef As AttributeReference = dbObj
                lenKabel = Math.Max(lenKabel, Len(acAttRef.TextString)) ' Намиране на най-дългия текстов низ от атрибутите
            Next
            ' Декларация на променливи за позицията и ъглите
            Dim Position_X, Position_Y As Double
            Dim Angle1, Angle2 As Double
            Dim Distance1, Distance2, Distance3, Distance4 As Double
            ' Получаване на името на извикващия метод
            Dim stackTrace As New Diagnostics.StackFrame(1) ' Пропуска един стеков фрейм
            Dim callingMethodName As String = stackTrace.GetMethod.Name
            ' Проверка на извикващия метод
            If callingMethodName = "Insert_Block_Kabel" Then
                Dim br As Integer = 0
                For Each objID As ObjectId In attCol
                    Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForWrite)
                    Dim acAttRef As AttributeReference = dbObj
                    ' Проверка за атрибути, започващи с "NA4IN_"
                    If InStr(acAttRef.Tag, "NA4IN_") > 0 And acAttRef.TextString <> "" Then
                        Dim brLine As Integer = Val(Mid(acAttRef.TextString, 1, 2))
                        br += IIf(acAttRef.TextString <> "", brLine, 0)
                        ' Премахване на първите три знака от текста ако brLine е 1
                        If brLine = 1 Then acAttRef.TextString = Trim(Mid(acAttRef.TextString, 4, Len(acAttRef.TextString)))
                    End If
                Next
                br = IIf(br = 0, 1, br)
                Distance1 = 6.0 * acBlkRef.ScaleFactors.X * br ' Изчисляване на Distance1
            End If
            ' Дефиниране на допълнителни разстояния
            Distance3 = 6.0
            Distance4 = lenKabel * 9.4 * acBlkRef.ScaleFactors.X
            Dim BlockKabel As Boolean = False
            ' Получаване на динамичните свойства на блока
            Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection
            For Each prop As DynamicBlockReferenceProperty In props
                ' Проверка и настройка на състоянията на динамичните свойства
                If prop.PropertyName = "Position X" Then Position_X = prop.Value
                If prop.PropertyName = "Position Y" Then Position_Y = prop.Value
                If prop.Value.ToString = "Линии" Or prop.Value.ToString = "Точка" Then BlockKabel = True
            Next
            ' Проверка дали блока е кабел, ако не е - излизане от процедурата
            If Not BlockKabel Then Exit Sub
            Distance2 = Position_Y
            ' Логика за настройка на ъглите и разстоянията според позициите
            If Position_Y > 0 Then
                If Position_X > 0 Then
                    If Position_X < 87 * acBlkRef.ScaleFactors.X Then
                        If callingMethodName <> "Kabel_Aligment" Then Position_X = 87 * acBlkRef.ScaleFactors.X
                    End If
                    Angle1 = PI / 2
                    Angle2 = PI / 2
                    Distance3 = Position_X - 81.29468 * acBlkRef.ScaleFactors.X
                Else
                    If Math.Abs(Position_X) < Distance4 Then
                        If callingMethodName <> "Kabel_Aligment" Then Position_X = -1 * (Distance4 + 6 - 81.29468 * acBlkRef.ScaleFactors.X)
                    End If
                    Angle1 = PI / 2
                    Angle2 = PI / 2
                    Distance3 = 6
                    Distance4 = Math.Abs(Position_X) + 81.29468 * acBlkRef.ScaleFactors.X
                End If
            Else
                If Position_X < 0 Then
                    If Math.Abs(Position_X) < Distance4 Then
                        If callingMethodName <> "Kabel_Aligment" Then Position_X = -1 * (Distance4 + 6 - 81.29468 * acBlkRef.ScaleFactors.X)
                    End If
                    Angle1 = -1 * PI / 2
                    Angle2 = -1 * PI / 2
                    Distance3 = 6
                    Distance4 = Math.Abs(Position_X) + 81.29468 * acBlkRef.ScaleFactors.X
                Else
                    If Position_X < 87 And callingMethodName <> "Kabel_Aligment" Then Position_X = 87

                    Angle1 = -1 * PI / 2
                    Angle2 = -1 * PI / 2
                    Distance3 = Position_X - 81.29468 * acBlkRef.ScaleFactors.X
                End If
            End If
            ' Допълнителна логика за настройка на ъглите при малки стойности на Y
            If Math.Abs(Position_Y) < (25.0 * acBlkRef.ScaleFactors.X) Then
                If Position_X > 0 Then
                    Angle1 = 0
                    Angle2 = 0
                Else
                    Angle1 = PI
                    Angle2 = PI
                End If
                Position_Y = 0
                Distance2 = 0
            End If
            ' Логика за настройка на Distance3 и позицията по X
            If Distance3 < 0 Then
                Distance3 = 6 * acBlkRef.ScaleFactors.X
                If callingMethodName <> "Kabel_Aligment" Then Position_X = Position_X + 6 * acBlkRef.ScaleFactors.X
            End If

            If Position_Y = 0 Then
                Distance3 = 0
                Distance2 = Position_X
            End If
            ' Задаване на нови стойности за динамичните свойства на блока
            For Each prop As DynamicBlockReferenceProperty In props
                If prop.PropertyName = "Position X" Then prop.Value = Position_X
                If prop.PropertyName = "Position Y" Then prop.Value = Position_Y
                If prop.PropertyName = "Angle" Then prop.Value = Angle1
                If prop.PropertyName = "Angle2" Then prop.Value = Angle2
                If prop.PropertyName = "Visibility1" Then prop.Value = "Линии"
                If prop.PropertyName = "Distance1" Then
                    If Distance1 <> 0 Then
                        prop.Value = Distance1
                    End If
                End If
                If prop.PropertyName = "Distance2" Then prop.Value = Distance2
                If prop.PropertyName = "Distance3" Then prop.Value = Distance3
                If prop.PropertyName = "Distance4" Then prop.Value = Distance4
            Next
            ' Потвърждаване на транзакцията
            acTrans.Commit()
        End Using
    End Sub
    Public Sub Kabel_Aligment(blKabelАlign As strKabelАlign,
                          arrKabelАlign As List(Of strKabelАlign),
                          Stypka As Boolean,
                          Vertikal As Boolean,
                          Horizontal As Boolean
                          )
        ' Инициализира променливата Stypka_Raz, която представлява стъпка на която ще се разполагат блоковете.
        Dim Stypka_Raz As Double = 0
        ' Ако Stypka е зададена като True, тогава стъпката ще бъде 67.5.
        If Stypka Then Stypka_Raz = 67
        ' Вземи текущата база данни и започни транзакция.
        Dim acCurDb As Autodesk.AutoCAD.DatabaseServices.Database
        acCurDb = Application.DocumentManager.MdiActiveDocument.Database
        ' Запазва началната стъпка от първия блок.
        Dim br_NA4IN As Double = blKabelАlign.Stypka
        ' Използвай транзакция, за да направиш промени в базата данни.
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            ' Обработи всеки блок в списъка arrKabelАlign.
            For Each Obj As strKabelАlign In arrKabelАlign
                ' Вземи референцията на блока и неговата колекция от атрибути.
                Dim acBlkRef As BlockReference = DirectCast(acTrans.GetObject(Obj.Id, OpenMode.ForWrite), BlockReference)
                Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                ' Изчисли абсолютните стойности на позициите по X и Y за текущия и основния блок.
                Dim D_X_1 As Double = Math.Abs(blKabelАlign.Position_X)
                Dim D_Y_1 As Double = Math.Abs(blKabelАlign.Position_Y)
                Dim D_X_2 As Double = Math.Abs(Obj.Position_X)
                Dim D_Y_2 As Double = Math.Abs(Obj.Position_Y)
                ' Изчисли разликите в позициите по X и Y между текущия и основния блок.
                Dim D_X As Double = Obj.pInsert.X - blKabelАlign.pInsert.X
                Dim D_Y As Double = Obj.pInsert.Y - blKabelАlign.pInsert.Y
                Dim D_X_3 As Double = Math.Abs(D_X)
                Dim D_Y_3 As Double = Math.Abs(D_Y)
                ' Инициализира новите делти за X и Y.
                Dim New_Delta_X As Double = 0
                Dim New_Delta_Y As Double = 0
                ' Определи новата делта по X в зависимост от разликата D_X и позицията на основния блок.
                Select Case D_X
                    Case > 0
                        If blKabelАlign.Position_X > 0 Then
                            New_Delta_X = D_X_1 - D_X_3
                        Else
                            New_Delta_X = -1 * (D_X_1 + D_X_3)
                        End If
                    Case < 0
                        If blKabelАlign.Position_X > 0 Then
                            New_Delta_X = D_X_1 + D_X_3
                        Else
                            New_Delta_X = -1 * (D_X_1 - D_X_3)
                        End If
                    Case = 0
                        New_Delta_X = blKabelАlign.Position_X
                End Select
                ' Определи новата делта по Y в зависимост от разликата D_Y и позицията на основния блок.
                Select Case D_Y
                    Case > 0
                        If blKabelАlign.Position_Y > 0 Then
                            New_Delta_Y = D_Y_1 - D_Y_3
                        Else
                            New_Delta_Y = -1 * (D_Y_1 + D_Y_3)
                        End If
                    Case < 0
                        If blKabelАlign.Position_Y > 0 Then
                            New_Delta_Y = D_Y_1 + D_Y_3
                        Else
                            New_Delta_Y = -1 * (D_Y_1 - D_Y_3)
                        End If
                    Case = 0
                        New_Delta_Y = blKabelАlign.Position_Y
                End Select
                ' Коригира новата делта по Y със стъпката, ако е необходимо.
                New_Delta_Y = New_Delta_Y - Stypka_Raz * br_NA4IN
                ' Вземи колекцията от свойства на динамичния блок.
                Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection
                ' Промени стойностите на позицията по X и Y, ако съответните флагове са зададени.
                For Each prop As DynamicBlockReferenceProperty In props
                    If Vertikal Then
                        If prop.PropertyName = "Position Y" Then prop.Value = New_Delta_Y ' Позиция на вмъкване коорд. Y
                    End If
                    If Horizontal Then
                        If prop.PropertyName = "Position X" Then prop.Value = New_Delta_X ' Позиция на вмъкване коорд. Х
                    End If
                Next
                ' Извикай функцията за редактиране на динамичния блок.
                EditDynamicBlockReferenceKabel(Obj.Id)
                ' Увеличи стойността на стъпката за следващия блок.
                br_NA4IN += Obj.Stypka
            Next
            ' Потвърди транзакцията.
            acTrans.Commit()
        End Using
    End Sub
End Class
Public Class Commands_original
    <CommandMethod("GXD")>
    Public Shared Sub GetXData()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        Dim opt As PromptEntityOptions = New PromptEntityOptions(vbLf & "Select entity: ")
        Dim res As PromptEntityResult = ed.GetEntity(opt)

        If res.Status = PromptStatus.OK Then
            Dim tr As Transaction = doc.TransactionManager.StartTransaction()

            Using tr
                Dim obj As DBObject = tr.GetObject(res.ObjectId, OpenMode.ForRead)
                Dim rb As ResultBuffer = obj.XData

                If rb Is Nothing Then
                    ed.WriteMessage(vbLf & "Entity does not have XData attached.")
                Else
                    Dim n As Integer = 0

                    For Each tv As TypedValue In rb
                        ed.WriteMessage(vbLf & "TypedValue {0} - type: {1}, value: {2}",
                                        Math.Min(System.Threading.Interlocked.Increment(n), n - 1),
                                        tv.TypeCode,
                                        tv.Value)
                    Next
                    rb.Dispose()
                End If
            End Using
        End If
    End Sub
    <CommandMethod("SXD")>
    Public Shared Sub SetXData()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        Dim opt As PromptEntityOptions = New PromptEntityOptions(vbLf & "Select entity: ")
        Dim res As PromptEntityResult = ed.GetEntity(opt)

        If res.Status = PromptStatus.OK Then
            Dim tr As Transaction = doc.TransactionManager.StartTransaction()

            Using tr
                Dim obj As DBObject = tr.GetObject(res.ObjectId, OpenMode.ForWrite)
                AddRegAppTableRecord("EWG")
                Dim rb As ResultBuffer = New ResultBuffer(New TypedValue(1001, "KEAN"),
                                                          New TypedValue(1000, "This 1"),
                                                          New TypedValue(1000, "This 2"),
                                                          New TypedValue(1000, "This 3"))
                rb.Add(New TypedValue(1001, "This 4"))

                obj.XData = rb
                rb.Dispose()
                tr.Commit()
            End Using
        End If
    End Sub
    Private Shared Sub AddRegAppTableRecord(ByVal regAppName As String)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        Dim db As Database = doc.Database
        Dim tr As Transaction = doc.TransactionManager.StartTransaction()

        Using tr
            Dim rat As RegAppTable = CType(tr.GetObject(db.RegAppTableId, OpenMode.ForRead, False), RegAppTable)

            If Not rat.Has(regAppName) Then
                rat.UpgradeOpen()
                Dim ratr As RegAppTableRecord = New RegAppTableRecord()
                ratr.Name = regAppName
                rat.Add(ratr)
                tr.AddNewlyCreatedDBObject(ratr, True)
            End If
            tr.Commit()
        End Using
    End Sub

End Class
Public Class Commands_Modifi
    Public Shared Sub GetXData()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        Dim opt As PromptEntityOptions = New PromptEntityOptions(vbLf & "Select entity: ")
        Dim res As PromptEntityResult = ed.GetEntity(opt)

        If res.Status = PromptStatus.OK Then
            Dim tr As Transaction = doc.TransactionManager.StartTransaction()

            Using tr
                Dim obj As DBObject = tr.GetObject(res.ObjectId, OpenMode.ForRead)
                Dim rb As ResultBuffer = obj.XData

                If rb Is Nothing Then
                    ed.WriteMessage(vbLf & "Entity does not have XData attached.")
                Else
                    Dim n As Integer = 0

                    For Each tv As TypedValue In rb
                        ed.WriteMessage(vbLf & "TypedValue {0} - type: {1}, value: {2}",
                                        Math.Min(System.Threading.Interlocked.Increment(n), n - 1),
                                        tv.TypeCode,
                                        tv.Value)
                    Next
                    rb.Dispose()
                End If
            End Using
        End If
    End Sub
    Public Shared Sub SetXData()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        Dim opt As PromptEntityOptions = New PromptEntityOptions(vbLf & "Select entity: ")
        Dim res As PromptEntityResult = ed.GetEntity(opt)

        If res.Status = PromptStatus.OK Then
            Dim tr As Transaction = doc.TransactionManager.StartTransaction()

            Using tr
                Dim obj As DBObject = tr.GetObject(res.ObjectId, OpenMode.ForWrite)
                AddRegAppTableRecord("EWG_Tablo")
                Dim rb As ResultBuffer = New ResultBuffer(New TypedValue(1001, "EWG_Tablo"))

                obj.XData = rb
                rb.Dispose()
                tr.Commit()
            End Using
        End If
    End Sub
    Public Shared Sub AddRegAppTableRecord(ByVal regAppName As String)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        Dim db As Database = doc.Database
        Dim tr As Transaction = doc.TransactionManager.StartTransaction()

        Using tr
            Dim rat As RegAppTable = CType(tr.GetObject(db.RegAppTableId, OpenMode.ForRead, False), RegAppTable)
            Dim ratr As RegAppTableRecord = New RegAppTableRecord()

            If Not rat.Has(regAppName) Then
                rat.UpgradeOpen()
                ratr.Name = regAppName
                rat.Add(ratr)
                tr.AddNewlyCreatedDBObject(ratr, True)
            End If
            tr.Commit()
        End Using
    End Sub
End Class
