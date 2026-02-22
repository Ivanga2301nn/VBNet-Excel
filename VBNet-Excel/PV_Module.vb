Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Public Class PV_Module
    Dim cu As CommonUtil = New CommonUtil()
    Dim form_PV As New PV()

    <CommandMethod("PV_Module")>
    Public Sub PV_Module()
        Application.ShowModalDialog(form_PV)
    End Sub
    <CommandMethod("PV_Panel_Izgled_Klemi")>
    Public Sub PV_Panel_Izgled_Klemi()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim edt As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database
        Dim SelectedSet = cu.GetObjects("INSERT", "Изберете блок")
        If SelectedSet Is Nothing Then
            MsgBox("НЕ Е маркиран нито един блок.")
            Exit Sub
        End If
        Dim blkRecId As ObjectId = ObjectId.Null
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                For Each sObj As SelectedObject In SelectedSet
                    blkRecId = sObj.ObjectId
                    Dim acBlkRef As BlockReference = DirectCast(acTrans.GetObject(blkRecId, OpenMode.ForWrite), BlockReference)
                    Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection
                    For Each prop As DynamicBlockReferenceProperty In props
                        If prop.PropertyName = "Visibility1" Then
                            If prop.Value = "Изглед ДВЕ" Then
                                prop.Value = "Със + -- -"
                            End If
                        End If
                    Next
                Next
                acTrans.Commit()
            Catch ex As Exception
                MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
                acTrans.Abort()
            End Try
        End Using
    End Sub
    <CommandMethod("PV_Panel_Izgled_Klemi_Optimized")>
    Public Sub PV_Panel_Izgled_Klemi_Optimized()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim edt As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database

        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim SelectedSet = cu.GetObjects("INSERT", "Изберете блок")
            If SelectedSet Is Nothing Then
                MsgBox("НЕ Е маркиран нито един блок.")
                Exit Sub
            End If
            Try
                For Each sObj As SelectedObject In SelectedSet
                    Dim blkRecId As ObjectId = sObj.ObjectId
                    Dim acBlkRef As BlockReference = DirectCast(acTrans.GetObject(blkRecId, OpenMode.ForWrite), BlockReference)
                    Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection

                    For Each prop As DynamicBlockReferenceProperty In props
                        If prop.PropertyName = "Visibility1" Then
                            prop.Value = "Със + -- -"
                        End If
                    Next
                Next
                acTrans.Commit()
            Catch ex As Exception
                MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
                acTrans.Abort()
            End Try
        End Using
    End Sub
    <CommandMethod("PV_Panel_Klemi_Izgled")>
    Public Sub PV_Panel_Klemi_Izgled()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim edt As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database
        Dim SelectedSet = cu.GetObjects("INSERT", "Изберете блок")
        If SelectedSet Is Nothing Then
            MsgBox("НЕ Е маркиран нито един блок.")
            Exit Sub
        End If

        Dim blkRecId As ObjectId = ObjectId.Null
        Dim ind As Integer = 0
        Dim index As Integer = 0

        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                For Each sObj As SelectedObject In SelectedSet
                    blkRecId = sObj.ObjectId
                    index += 1
                    Dim acBlkRef As BlockReference = DirectCast(acTrans.GetObject(blkRecId, OpenMode.ForWrite), BlockReference)
                    Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection
                    For Each prop As DynamicBlockReferenceProperty In props
                        If prop.PropertyName = "Visibility1" Then
                            If prop.Value = "Със + -- -" Then
                                prop.Value = "Изглед ДВЕ"
                            End If
                        End If
                    Next
                Next

                acTrans.Commit()
            Catch ex As Exception
                MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
                acTrans.Abort()
            End Try
        End Using
    End Sub
End Class

Module Example
    Function GetBlockReferences() As List(Of BlockReference)
        ' Реализирайте логиката за получаване на списък с обекти BlockReference тук
        ' ...
        Return New List(Of BlockReference)()
    End Function
End Module
