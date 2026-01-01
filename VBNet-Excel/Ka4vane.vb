Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Public Class Ka4vane
    Structure strLine
        Dim Layer As String
        Dim Linetype As String
        Dim count As Double
    End Structure

    <CommandMethod("Ka4vane")>
    Public Sub Ka4vane()
        ' Създаване на нов обект от класа CommonUtil
        Dim cu As CommonUtil = New CommonUtil()
        ' Извличане на селектирани обекти чрез метода GetObjects с параметри "INSERT" и съобщение "Изберете блок за качване: "
        Dim ss = cu.GetObjects("INSERT", "Изберете блок за качване: ")
        Dim nameBlock As String = "Качване"

        ' Получаване на текущия документ и база данни в AutoCAD
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim edt As Editor = acDoc.Editor

        ' Проверка дали има маркиран блок
        If ss Is Nothing Then
            MsgBox("Няма маркиран блок.")
            Exit Sub
        ElseIf ss.Count > 1 Then
            MsgBox("Маркиран е повече от един блок.")
            Exit Sub
        End If

        Try
            ' Създаване на нов обект Document за текущия документ
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            ' Деклариране на променливи за ключови думи и атрибути
            Dim pKeyOpts As PromptKeywordOptions = New PromptKeywordOptions("")
            Dim KOTA_1, KOTA_2 As String
            Dim ПОЛАГ_1, ПОЛАГ_2 As String
            Dim Visibility As String = ""
            Dim dist_1, dist_2, ang As String
            Dim blkRecId As ObjectId
            ' Стартиране на транзакция
            Using trans As Transaction = doc.TransactionManager.StartTransaction()
                ' Извличане на ObjectId на първия маркиран обект
                blkRecId = ss(0).ObjectId
                ' Получаване на BlockTable в режим на четене
                Dim acBlkTbl As BlockTable = trans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                ' Получаване на BlockReference в режим на писане
                Dim acBlkRef As BlockReference = DirectCast(trans.GetObject(blkRecId, OpenMode.ForWrite), BlockReference)

                ' Извличане на колекция от атрибути на блока
                Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                KOTA_1 = ""
                KOTA_2 = ""
                ПОЛАГ_1 = ""
                ПОЛАГ_2 = ""
                ' Извличане на стойностите на специфични атрибути
                For Each objID As ObjectId In attCol
                    Dim dbObj As DBObject = trans.GetObject(objID, OpenMode.ForRead)
                    Dim acAttRef As AttributeReference = dbObj
                    If acAttRef.Tag = "KOTA_1" Then KOTA_1 = acAttRef.TextString
                    If acAttRef.Tag = "KOTA_2" Then KOTA_2 = acAttRef.TextString
                    If acAttRef.Tag = "ТРЪБА_1" Then ПОЛАГ_1 = acAttRef.TextString
                    If acAttRef.Tag = "ТРЪБА_2" Then ПОЛАГ_2 = acAttRef.TextString
                Next

                ' Извличане на колекция от динамични свойства на блока
                Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection
                For Each prop As DynamicBlockReferenceProperty In props
                    ' Промяна на стойности на динамични свойства в зависимост от тяхното име
                    If prop.PropertyName = "Visibility" Then Visibility = prop.Value
                    If prop.PropertyName = "Distance1" Then dist_1 = prop.Value
                    If prop.PropertyName = "Distance2" Then dist_2 = prop.Value
                    If prop.PropertyName = "Angle1" Then ang = prop.Value
                Next

                ' Промяна на атрибути с тагове "Kabel_d_0" до "Kabel_d_10" и "Kabel_g_0" до "Kabel_g_10" на стойност "Кабел"
                For br = 0 To 10
                    Dim ddd As String = "Kabel_d_" + br.ToString
                    Dim sss As String = "Kabel_g_" + br.ToString
                    For Each objID As ObjectId In attCol
                        Dim dbObj As DBObject = trans.GetObject(objID, OpenMode.ForWrite)
                        Dim acAttRef As AttributeReference = dbObj
                        If acAttRef.Tag = ddd Then acAttRef.TextString = "Кабел"
                        If acAttRef.Tag = sss Then acAttRef.TextString = "Кабел"
                    Next
                Next
                ' Записване на транзакцията
                trans.Commit()
            End Using

            ' Изпълнение на специфични функции в зависимост от стойността на Visibility
            Select Case Visibility
                Case "от етаж"
                    Call Insert_Kabel_Ka4vane(blkRecId, KOTA_1, ПОЛАГ_1, True)
                Case "към етаж"
                    Call Insert_Kabel_Ka4vane(blkRecId, KOTA_2, ПОЛАГ_2, False)
                Case "преход"
                    Call Insert_Kabel_Ka4vane(blkRecId, KOTA_1, ПОЛАГ_1, True)
                    Call Insert_Kabel_Ka4vane(blkRecId, KOTA_2, ПОЛАГ_2, False)
                Case "текстове"
                    Call Insert_Kabel_Ka4vane(blkRecId, KOTA_1, ПОЛАГ_1, True)
                    Call Insert_Kabel_Ka4vane(blkRecId, KOTA_2, ПОЛАГ_2, False)
            End Select
        Catch ex As Exception
            ' Обработка на изключения (празно тяло)
        End Try
    End Sub

    Private Sub Insert_Kabel_Ka4vane(blkRecId As ObjectId,  ' ID на блока за редактиране
                                     KOTA As String,        ' Кота за която се отнася
                                     ПОЛАГ As String,       ' Начин на полагане
                                     OT_KYM As Boolean)     ' Определя ОТ или КЪМ кота
        ' True - ОТ етаж
        ' False - КЪМ етаж

        Dim cu As CommonUtil = New CommonUtil()
        Dim kab As Kabel = New Kabel
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim tgKOTA, tgTRYBA, tgKABEL, tgOT_KYM As String
        If OT_KYM Then
            tgKOTA = "KOTA_1"
            tgTRYBA = "ТРЪБА_1"
            tgKABEL = "Kabel_d_"
            tgOT_KYM = "ОТ"
        Else
            tgKOTA = "KOTA_2"
            tgTRYBA = "ТРЪБА_2"
            tgKABEL = "Kabel_g_"
            tgOT_KYM = "КЪМ"
        End If
        Using trans As Transaction = acDoc.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = trans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim acBlkRef As BlockReference = DirectCast(trans.GetObject(blkRecId, OpenMode.ForWrite), BlockReference)
            Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
            KOTA = Edit_Atribut_Double(KOTA, tgOT_KYM)
            For Each objID As ObjectId In attCol
                Dim dbObj As DBObject = trans.GetObject(objID, OpenMode.ForWrite)
                Dim acAttRef As AttributeReference = dbObj
                If acAttRef.Tag = tgKOTA Then acAttRef.TextString = KOTA
            Next
            ПОЛАГ = Edit_Atribut_Param(ПОЛАГ, tgOT_KYM)
            For Each objID As ObjectId In attCol
                Dim dbObj As DBObject = trans.GetObject(objID, OpenMode.ForWrite)
                Dim acAttRef As AttributeReference = dbObj
                If acAttRef.Tag = tgTRYBA Then acAttRef.TextString = ПОЛАГ
            Next
            Dim ss = cu.GetObjects("LINE", "Изберете линии преход " & tgOT_KYM & " етаж:")

            If ss Is Nothing Then
                MsgBox("Нама маркиран линия в слой 'EL'.")
                Exit Sub
            End If
            '
            ' Kabel (*,0) - Тип на линията
            ' Kabel (*,1) - Тип на тръбата
            ' Kabel (*,2) - брой маркирани линии от този тип
            '
            Dim Kabel(110) As strLine
            Dim Kabel_(10, 2) As String
            Dim Index As Integer
            For Each sObj As SelectedObject In ss
                Dim line As Line = TryCast(trans.GetObject(sObj.ObjectId, OpenMode.ForRead), Line)
                Dim iVisib As Integer = -1
                iVisib = Array.FindIndex(Kabel, Function(f) f.Layer = line.Layer)

                If iVisib = -1 Then
                    Kabel(Index).Layer = line.Layer
                    Kabel(Index).Linetype = line.Linetype
                    Kabel(Index).count = 1
                    Index += 1
                Else
                    Kabel(iVisib).count = Kabel(iVisib).count + 1
                End If
            Next
            '
            ' Запълва блока КАЧВАНЕ
            ' не запълва масива!!!
            '
            For br = 0 To UBound(Kabel)
                Dim ddd As String = tgKABEL + br.ToString
                If ПОЛАГ = "Скара" Or ПОЛАГ = "Канал" Then
                    Kabel(br).Linetype = ПОЛАГ
                Else
                    Kabel(br).Linetype = cu.line_Layer(Kabel(br).Layer)
                End If
                For Each objID As ObjectId In attCol
                    Dim dbObj As DBObject = trans.GetObject(objID, OpenMode.ForWrite)
                    Dim acAttRef As AttributeReference = dbObj
                    If acAttRef.Tag = ddd Then
                        acAttRef.TextString = Kabel(br).count & "л. " & Kabel(br).Layer
                        Kabel_(br, 0) = Kabel(br).Layer
                        Kabel_(br, 2) = Kabel(br).count
                        Exit For
                    End If
                Next
            Next
            Select Case ПОЛАГ
                Case "Скара"
                    For br = 0 To UBound(Kabel_)
                        cu.GET_LINE_TYPE_KABEL(Kabel_, ss, True)
                        Kabel_(br, 1) = "пол. по кабелна скара"
                    Next
                Case "Канал"
                    For br = 0 To UBound(Kabel_)
                        cu.GET_LINE_TYPE_KABEL(Kabel_, ss, True)
                        Kabel_(br, 1) = "изт. в каб.кан."
                    Next
                Case "Тръби"
                    For br = 0 To UBound(Kabel_)
                        '
                        ' Kabel (*,0) - Тип на линията
                        ' Kabel (*,1) - Тип на тръбата
                        ' Kabel (*,2) - брой маркирани линии от този тип
                        '
                        Kabel_(br, 1) = cu.SET_line_Type(Kabel_(br, 0))
                        Kabel_(br, 0) = cu.line_Layer(Kabel_(br, 0))
                        Kabel_(br, 1) = cu.GET_line_Type(Kabel_(br, 1), False)
                    Next
            End Select
            Dim posokaText As String = ""
            For br = 0 To UBound(Kabel_)
                If Kabel_(br, 2) = "0" Then
                    If OT_KYM = True Then
                        posokaText = "от кота "
                    Else
                        posokaText = "към кота "
                    End If
                    Kabel_(br, 0) = ""
                    Kabel_(br, 1) = posokaText
                    Kabel_(br, 2) = 100
                    Exit For
                End If
            Next
            kab.Insert_Block_Kabel(Kabel_)
            trans.Commit()
        End Using
    End Sub
    Private Function Edit_Atribut_Double(kota As String, OT As String) As String
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim pDouOpts As PromptDoubleOptions = New PromptDoubleOptions("")
        Dim atrib As String
        With pDouOpts
            .Keywords.Add(kota)
            .Keywords.Default = kota
            .Message = vbCrLf & "Въведете дължина на преход " & OT & " етаж: "
            .AllowZero = False
            .AllowNegative = False
        End With
        Dim pKeyRes As PromptDoubleResult = acDoc.Editor.GetDouble(pDouOpts)

        If pKeyRes.Status = PromptStatus.Keyword Then
            atrib = pKeyRes.StringResult
        Else
            atrib = pKeyRes.Value.ToString()
        End If
        Return atrib
    End Function
    Private Function Edit_Atribut_Param(ПОЛАГ As String, OT As String) As String
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim pKeyOpts As PromptKeywordOptions = New PromptKeywordOptions("")
        If ПОЛАГ <> "Тръби" Or
            ПОЛАГ <> "Скара" Or
            ПОЛАГ <> "Канал" Then ПОЛАГ = "Тръби"

        With pKeyOpts
            .Keywords.Add("Тръби")
            .Keywords.Add("Скара")
            .Keywords.Add("Канал")
            .Message = vbCrLf & "Въведете начина на полагане на преход " & OT & " етаж: "
            .Keywords.Default = ПОЛАГ
            .AllowNone = True
        End With
        Return acDoc.Editor.GetKeywords(pKeyOpts).StringResult
    End Function
    <CommandMethod("GetKeywordFromUser")>
    Public Sub GetKeywordFromUser()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim pKeyOpts As PromptKeywordOptions = New PromptKeywordOptions("")
        pKeyOpts.Message = vbLf & "Enter an option "
        pKeyOpts.Keywords.Add("Line")
        pKeyOpts.Keywords.Add("Circle")
        pKeyOpts.Keywords.Add("Arc")
        pKeyOpts.AllowNone = False
        Dim pKeyRes As PromptResult = acDoc.Editor.GetKeywords(pKeyOpts)
        Application.ShowAlertDialog("Entered keyword: " &
                                      pKeyRes.StringResult)

    End Sub
    <CommandMethod("GetKeywordFromUser2")>
    Public Sub GetKeywordFromUser2()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim pKeyOpts As PromptKeywordOptions = New PromptKeywordOptions("")
        pKeyOpts.Message = vbLf & "Enter an option "
        pKeyOpts.Keywords.Add("Тръба")
        pKeyOpts.Keywords.Add("Мазилка")
        pKeyOpts.Keywords.Add("Скара")
        pKeyOpts.Keywords.Default = "Тръба"
        pKeyOpts.AllowNone = True
        Dim pKeyRes As PromptResult = acDoc.Editor.GetKeywords(pKeyOpts)
        Application.ShowAlertDialog("Entered keyword: " &
                              pKeyRes.StringResult)
    End Sub
    <CommandMethod("GetIntegerOrKeywordFromUser")>
    Public Sub GetIntegerOrKeywordFromUser()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument

        Dim pIntOpts As PromptIntegerOptions = New PromptIntegerOptions("")
        pIntOpts.Message = vbCrLf & "Enter the size or "

        '' Restrict input to positive and non-negative values
        pIntOpts.AllowZero = False
        pIntOpts.AllowNegative = False

        '' Define the valid keywords and allow Enter
        pIntOpts.Keywords.Add("Big")
        pIntOpts.Keywords.Add("Small")
        pIntOpts.Keywords.Add("Regular")
        pIntOpts.Keywords.Default = "Regular"
        pIntOpts.AllowNone = True

        '' Get the value entered by the user
        Dim pIntRes As PromptIntegerResult = acDoc.Editor.GetInteger(pIntOpts)

        If pIntRes.Status = PromptStatus.Keyword Then
            Application.ShowAlertDialog("Entered keyword: " &
                                    pIntRes.StringResult)
        Else
            Application.ShowAlertDialog("Entered value: " &
                                    pIntRes.Value.ToString())
        End If
    End Sub
    <CommandMethod("GetStringFromUser")>
    Public Sub GetStringFromUser()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument

        Dim pStrOpts As PromptStringOptions = New PromptStringOptions(vbLf &
                                                                     "Enter your name: ")
        pStrOpts.AllowSpaces = True
        Dim pStrRes As PromptResult = acDoc.Editor.GetString(pStrOpts)

        Application.ShowAlertDialog("The name entered was: " &
                                    pStrRes.StringResult)
    End Sub
    <CommandMethod("GetIntegerOrKeywordFromUser2")>
    Public Sub GetIntegerOrKeywordFromUser2()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument

        Dim pIntOpts As PromptDoubleOptions = New PromptDoubleOptions("")
        pIntOpts.Message = vbCrLf & "Enter the size or "

        '' Restrict input to positive and non-negative values
        pIntOpts.AllowZero = True
        pIntOpts.AllowNegative = False

        '' Define the valid keywords and allow Enter
        pIntOpts.Keywords.Add("Big")
        pIntOpts.Keywords.Add("Small")
        pIntOpts.Keywords.Add("Regular")
        pIntOpts.Keywords.Default = "Regular"
        pIntOpts.AllowNone = True

        '' Get the value entered by the user
        Dim pIntRes As PromptDoubleResult = acDoc.Editor.GetDouble(pIntOpts)

        If pIntRes.Status = PromptStatus.Keyword Then
            Application.ShowAlertDialog("Entered keyword: " &
                                    pIntRes.StringResult)
        Else
            Application.ShowAlertDialog("Entered value: " &
                                    pIntRes.Value.ToString())
        End If
    End Sub
End Class