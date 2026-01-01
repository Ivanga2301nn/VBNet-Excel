Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Internal.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Imports Autodesk.AutoCAD.PlottingServices
Imports System.Collections.Generic
Public Class PIC
    Public Structure strDat4ik
        Dim strVisibility As String
        Dim strZN As String
        Dim strNOM As String
        Dim strAD As String
        Dim strPosition_X As Double
        Dim strPosition_Y As Double
        Dim ObjectId As ObjectId
    End Structure
    Public Structure strLine
        Dim strVisibility As String
        Dim strZN As String
        Dim strNOM As String
        Dim strAD As String
        Dim strPosition_X As Double
        Dim strPosition_Y As Double
    End Structure
    Dim cu As CommonUtil = New CommonUtil()
    <CommandMethod("PIC_Dat4ik_Zone")>
    Public Sub PIC_Dat4ik_Zone()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim edt As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database
        Dim SelectedSet = cu.GetObjects("INSERT", "Изберете блок")
        If SelectedSet Is Nothing Then
            MsgBox("Нама маркиран нито един блок.")
            Exit Sub
        End If
        Dim arrBlock(250) As strDat4ik
        Dim blkRecId As ObjectId = ObjectId.Null
        Dim pKeyOpts As PromptKeywordOptions = New PromptKeywordOptions("")
        With pKeyOpts
            .Keywords.Add("01")
            .Keywords.Add("02")
            .Keywords.Add("03")
            .Keywords.Add("04")
            .Keywords.Add("05")
            .Keywords.Add("06")
            .Keywords.Add("07")
            .Keywords.Add("08")
            .Message = vbCrLf & "Въведете номер на зона/контур:"
            .Keywords.Default = "01"
            .AllowNone = True
        End With
        Dim Zona_Dat4i As String = acDoc.Editor.GetKeywords(pKeyOpts).StringResult
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                Dim index As Integer = 0
                For Each sObj As SelectedObject In SelectedSet
                    blkRecId = sObj.ObjectId
                    Dim Visibility As String = ""
                    Dim acBlkRef As BlockReference = DirectCast(acTrans.GetObject(blkRecId, OpenMode.ForRead), BlockReference)
                    Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                    Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection
                    Dim blName As String = (CType(acBlkRef.DynamicBlockTableRecord.GetObject(OpenMode.ForRead), BlockTableRecord)).Name
                    If blName <> "Датчик_ПАБ" Then Continue For
                    For Each prop As DynamicBlockReferenceProperty In props
                        If prop.PropertyName = "Visibility" Then Visibility = prop.Value
                    Next
                    arrBlock(index).ObjectId = blkRecId
                    Dim ZONA As String = ""
                    Select Case Visibility
                        Case "ПАБ - Сирена конвенционална"
                            ZONA = "ЗС"
                        Case "ПАБ - Димооптичен конвенционален",
                             "ПАБ - Сирена адресируема",
                             "ПАБ - Термичен адресируем комбиниран",
                             "ПАБ - Термичен адресируем диференциален",
                             "ПАБ - Термичен адресируем с адаптер-7120",
                             "ПАБ - Термичен адресируем - 7101",
                             "ПАБ - Димооптичен адресируем",
                             "ПАБ - Пламъков конвенционален",
                             "ПАБ - Термичен конвенционален диференциален",
                             "ПАБ - Термичен конвенционален",
                             "ПАБ - Термичен конвенционален комбиниран",
                             "Линеен оптично димен приемник",
                             "Линеен оптично димен излъчвател"
                            ZONA = Zona_Dat4i
                        Case "ПАБ - Лампа"
                            ZONA = "ПС"
                        Case "ПАБ - Сирена и Звук"
                            ZONA = "СЗС"
                        Case "Ръчен пожароизвестител конвенционален", "Ръчен пожароизвестител адресируем"
                            ZONA = "РП"
                        Case "Изпълнително устройство"
                            ZONA = "ИУ"
                        Case "Изолатор"
                            ZONA = "ИЗ"
                        Case Else
                            ZONA = ZONA
                    End Select
                    For Each objID As ObjectId In attCol
                        Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForWrite)
                        Dim acAttRef As AttributeReference = dbObj
                        If acAttRef.Tag = "ZN" Then acAttRef.TextString = ZONA
                        If acAttRef.Tag = "NOM" Then acAttRef.TextString = "000"
                        If acAttRef.Tag = "AD" Then acAttRef.TextString = "000"
                    Next
                Next
                acTrans.Commit()
            Catch ex As Exception
                MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
                acTrans.Abort()
            End Try
        End Using
    End Sub
    <CommandMethod("PIC_Dat4ik_Nomber")>
    Public Sub PIC_Dat4ik_Nomber()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim edt As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database
        Dim arrBlock(250) As strDat4ik
        Dim blkRecId As ObjectId = ObjectId.Null
        Dim pStrOpts As PromptStringOptions = New PromptStringOptions(vbLf &
                                                               "Въведете начален номер: ")
        pStrOpts.AllowSpaces = True
        Dim pStrRes As PromptResult = acDoc.Editor.GetString(pStrOpts)
        Dim Nomer_dat4ik As Integer = Val(pStrRes.StringResult)
        Do
            Try
                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                    Dim SelectedSet = cu.GetObjects("INSERT", "Изберете блок за номериране:")
                    If SelectedSet Is Nothing Then
                        MsgBox("Нама маркиран нито един блок.")
                        Exit Sub
                    End If
                    Dim index As Integer = 0
                    For Each sObj As SelectedObject In SelectedSet
                        blkRecId = sObj.ObjectId
                        Dim Visibility As String = ""
                        Dim acBlkRef As BlockReference = DirectCast(acTrans.GetObject(blkRecId, OpenMode.ForRead), BlockReference)
                        Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                        Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection
                        Dim blName As String = (CType(acBlkRef.DynamicBlockTableRecord.GetObject(OpenMode.ForRead), BlockTableRecord)).Name
                        If Not (blName = "Датчик_ПАБ" Or blName = "Камери") Then
                            Continue For
                        End If
                        For Each prop As DynamicBlockReferenceProperty In props
                            If prop.PropertyName = "Visibility" Then Visibility = prop.Value
                        Next
                        arrBlock(index).ObjectId = blkRecId
                        Dim ZONA As String = ""
                        For Each objID As ObjectId In attCol
                            Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForWrite)
                            Dim acAttRef As AttributeReference = dbObj
                            If acAttRef.Tag = "NOM" Then acAttRef.TextString = Right("0000" & Nomer_dat4ik.ToString, 3)
                        Next
                    Next
                    Nomer_dat4ik += 1
                    acTrans.Commit()
                End Using
            Catch ex As Exception
                MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
            End Try
        Loop
    End Sub
    <CommandMethod("PIC_Dat4ik_Adres")>
    Public Sub PIC_Dat4ik_Adres()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim edt As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database
        Dim arrBlock(250) As strDat4ik
        Dim blkRecId As ObjectId = ObjectId.Null
        Dim pStrOpts As PromptStringOptions = New PromptStringOptions(vbLf &
                                                               "Въведете начален номер: ")
        pStrOpts.AllowSpaces = True
        Dim pStrRes As PromptResult = acDoc.Editor.GetString(pStrOpts)
        Dim Nomer_dat4ik As Integer = Val(pStrRes.StringResult)
        Do
            Try
                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                    Dim SelectedSet = cu.GetObjects("INSERT", "Изберете блок за номериране:")
                    If SelectedSet Is Nothing Then
                        MsgBox("Нама маркиран нито един блок.")
                        Exit Sub
                    End If
                    Dim index As Integer = 0
                    For Each sObj As SelectedObject In SelectedSet
                        blkRecId = sObj.ObjectId
                        Dim Visibility As String = ""
                        Dim acBlkRef As BlockReference = DirectCast(acTrans.GetObject(blkRecId, OpenMode.ForRead), BlockReference)
                        Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                        Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection
                        Dim blName As String = (CType(acBlkRef.DynamicBlockTableRecord.GetObject(OpenMode.ForRead), BlockTableRecord)).Name
                        If blName <> "Датчик_ПАБ" Then Continue For
                        For Each prop As DynamicBlockReferenceProperty In props
                            If prop.PropertyName = "Visibility" Then Visibility = prop.Value
                        Next
                        arrBlock(index).ObjectId = blkRecId
                        For Each objID As ObjectId In attCol
                            Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForWrite)
                            Dim acAttRef As AttributeReference = dbObj
                            If acAttRef.Tag = "AD" Then acAttRef.TextString = Right("0000" & Nomer_dat4ik.ToString, 3)
                        Next
                    Next
                    Nomer_dat4ik += 1
                    acTrans.Commit()
                End Using
            Catch ex As Exception
                MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
            End Try
        Loop
    End Sub
    <CommandMethod("PIC_Paralelen_Adres")>
    Public Sub PIC_Paralelen_Adres()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim edt As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database
        Dim arrBlock(250) As strDat4ik
        Dim SelectedSet = cu.GetObjects("INSERT", "Изберете паралелен сигнализатор за номериране:")
        Dim blkRecId As ObjectId = ObjectId.Null
        If SelectedSet Is Nothing Then
            MsgBox("Нама маркиран нито един блок.")
            Exit Sub
        End If
        If SelectedSet.Count > 1 Then
            MsgBox("Марлирани много блокове.")
            Exit Sub
        End If
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                Dim index As Integer = 0
                For Each sObj As SelectedObject In SelectedSet
                    arrBlock(index).ObjectId = sObj.ObjectId
                Next
                index += 1
                Do
                    SelectedSet = cu.GetObjects("INSERT", "Изберете блок за сигнализиране:")
                    If SelectedSet Is Nothing Then
                        MsgBox("Нама маркиран нито един блок.")
                        Exit Do
                    End If
                    For Each sObj As SelectedObject In SelectedSet
                        blkRecId = sObj.ObjectId
                        Dim acBlkRef As BlockReference = DirectCast(acTrans.GetObject(blkRecId, OpenMode.ForRead), BlockReference)
                        Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                        Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection
                        Dim blName As String = (CType(acBlkRef.DynamicBlockTableRecord.GetObject(OpenMode.ForRead), BlockTableRecord)).Name

                        If blName <> "Датчик_ПАБ" Then Continue For

                        Dim Visibility As String = ""

                        For Each prop As DynamicBlockReferenceProperty In props
                            If prop.PropertyName = "Visibility" Then Visibility = prop.Value
                        Next

                        If Visibility = "ПАБ - Димооптичен конвенционален" Or
                           Visibility = "ПАБ - Пламъков конвенционален" Or
                           Visibility = "ПАБ - Термичен конвенционален диференциален" Or
                           Visibility = "ПАБ - Термичен конвенционален" Or
                           Visibility = "ПАБ - Термичен конвенционален комбиниран" Or
                           Visibility = "Линеен оптично димен приемник" Or
                           Visibility = "Линеен оптично димен излъчвател" Then

                            arrBlock(index).ObjectId = blkRecId
                            arrBlock(index).strVisibility = Visibility
                            For Each objID As ObjectId In attCol
                                Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForWrite)
                                Dim acAttRef As AttributeReference = dbObj
                                If acAttRef.Tag = "AD" Then arrBlock(index).strAD = acAttRef.TextString
                                If acAttRef.Tag = "NOM" Then arrBlock(index).strNOM = acAttRef.TextString
                                If acAttRef.Tag = "ZN" Then arrBlock(index).strZN = acAttRef.TextString
                            Next
                            index += 1
                        End If
                    Next
                Loop
                Dim AdresMin As Integer = 999
                Dim AdresMax As Integer = 0
                For i As Integer = 1 To UBound(arrBlock)
                    If arrBlock(i).strVisibility = Nothing Then Exit For
                    AdresMin = Math.Min(AdresMin, Val(arrBlock(i).strNOM))
                    AdresMax = Math.Max(AdresMax, Val(arrBlock(i).strNOM))
                Next
                Dim adres As String = ""
                If AdresMin = AdresMax Then
                    adres = Right("0000" & AdresMin.ToString, 3)
                Else
                    adres = Right("0000" & AdresMin.ToString, 3) & "-" & Right("0000" & AdresMax.ToString, 3)
                End If

                Dim acBlkRef_1 As BlockReference = DirectCast(acTrans.GetObject(arrBlock(0).ObjectId, OpenMode.ForRead), BlockReference)
                Dim attCol_1 As AttributeCollection = acBlkRef_1.AttributeCollection
                arrBlock(index).ObjectId = blkRecId
                For Each objID As ObjectId In attCol_1
                    Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForWrite)
                    Dim acAttRef As AttributeReference = dbObj
                    If acAttRef.Tag = "AD" Then acAttRef.TextString = adres
                Next
                acTrans.Commit()
            Catch e As Exception
                MsgBox("Възникна грешка: " & e.Message & vbCrLf & vbCrLf & e.StackTrace.ToString)
                acTrans.Abort()
            End Try
        End Using
    End Sub
    <CommandMethod("Nomer_6ahta")>
    Public Sub Nomer_6ahta()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim edt As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database
        Dim arrBlock(250) As strDat4ik
        Dim blkRecId As ObjectId = ObjectId.Null
        Dim pStrOpts As PromptStringOptions = New PromptStringOptions(vbLf &
                                                               "Въведете начален номер: ")
        pStrOpts.AllowSpaces = True
        Dim pStrRes As PromptResult = acDoc.Editor.GetString(pStrOpts)
        Dim Nomer_dat4ik As Integer = Val(pStrRes.StringResult)
        Do
            Try
                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                    Dim SelectedSet = cu.GetObjects("INSERT", "Изберете блок за номериране:")
                    If SelectedSet Is Nothing Then
                        MsgBox("Нама маркиран нито един блок.")
                        Exit Sub
                    End If
                    Dim index As Integer = 0
                    For Each sObj As SelectedObject In SelectedSet
                        blkRecId = sObj.ObjectId
                        Dim Visibility As String = ""
                        Dim acBlkRef As BlockReference = DirectCast(acTrans.GetObject(blkRecId, OpenMode.ForRead), BlockReference)
                        Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                        Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection
                        Dim blName As String = (CType(acBlkRef.DynamicBlockTableRecord.GetObject(OpenMode.ForRead), BlockTableRecord)).Name
                        If Not (blName = "Единична шахта") Then
                            Continue For
                        End If
                        For Each prop As DynamicBlockReferenceProperty In props
                            If prop.PropertyName = "Visibility" Then Visibility = prop.Value
                        Next
                        arrBlock(index).ObjectId = blkRecId
                        Dim ZONA As String = ""
                        For Each objID As ObjectId In attCol
                            Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForWrite)
                            Dim acAttRef As AttributeReference = dbObj
                            If acAttRef.Tag = "ШАХТА" Then acAttRef.TextString = Right("0000" & Nomer_dat4ik.ToString, 3)
                        Next
                    Next
                    Nomer_dat4ik += 1
                    acTrans.Commit()
                End Using
            Catch ex As Exception
                MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
            End Try
        Loop
    End Sub

End Class
