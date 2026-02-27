Imports System.Collections.Generic
Imports System.Drawing
Imports System.Security.Cryptography
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.DatabaseServices.Filters
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Internal.DatabaseServices
Imports Autodesk.AutoCAD.PlottingServices
Imports Autodesk.AutoCAD.Runtime
Imports Org.BouncyCastle.Bcpg
'Imports System.IO
'Imports System.Windows.Forms

Public Class Form_Tablo_new
    Dim PI As Double = 3.1415926535897931
    Dim cu As CommonUtil = New CommonUtil()
    Private ListKonsumator As New List(Of strKonsumator)
    Public Structure strKonsumator
        Dim Name As String              ' Име на блока
        Dim ID_Block As ObjectId        ' Връзка към AutoCAD
        Dim ТоковКръг As String         ' Токов кръг
        Dim strМОЩНОСТ As String        ' Мощност като текст от атрибут
        Dim doubМОЩНОСТ As Double       ' Мощност като число
        Dim ТАБЛО As String             ' Табло
        Dim Pewdn As String             ' Предназначение
        Dim PEWDN1 As String            ' Доп. предназначение
        Dim Dylvina_Led As Double       ' За LED ленти
        Dim Visibility As String        ' За динамични блокове
    End Structure
    Public Structure strTokow
        ' Идентификация
        Dim CountKonsumator As Integer
        Dim Tablo As String             ' Родителско табло
        Dim ТоковКръг As String         ' Име на кръга
        ' Броячи
        Dim brLamp As Integer
        Dim brKontakt As Integer
        ' Мощност и ток
        Dim Мощност As Double           ' kW
        Dim Tok As Double               ' A
        Dim faza As String              ' "L1", "L2", "3F"
        ' Кабел
        Dim Kabebel_Se4enie As String   ' "3x2.5"
        ' Защита (прекъсвач)
        Dim BlockName As String
        Dim Designation As String
        Dim ShortName As String
        Dim Type As String
        Dim NumberPoles As String
        Dim RatedCurrent As String
        Dim Curve As String
        Dim Current As String
        Dim Control As String
        Dim Sensitivity As String
        Dim Protection As String
        ' Изчислителни
        Dim BrojPol As String
        ' Консуматори - ПРОМЯНА: масив → List
        Dim Konsumator As List(Of strKonsumator)
    End Structure
    Public Structure strTablo
        Dim countTablo As Integer
        Dim Name As String              ' "АП-1"
        Dim prevTablo As String         ' "Гл.Р.Т."
        Dim countTokKryg As Integer
        ' За TreeView групиране - ДОБАВЕНО:
        Dim Floor As String             ' "Първи етаж", "Подземен"
        Dim Building As String          ' "Сграда А" (по желание)
        Dim Tokowkryg As List(Of strTokow)  ' ПРОМЯНА: масив → List
        Dim TabloType As String
        ' Изчислени (за показване в TreeView)
        Dim TotalPower As Double        ' Сума от кръговете
        Dim SupplyCable As String       ' "NYM 5x16"
        ' Допълнителни за таблото (по желание)
        Dim Width As Integer
        Dim Height As Integer
        Dim IP_Rating As String
    End Structure

    Private Sub Form_Tablo_new_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Height = 1000
        Me.Width = 1600
        DataGridView.Visible = False
        GetObjects()
        DataGridView.Visible = True
    End Sub
    <CommandMethod("NEW_Tablo")>
    Private Sub NEW_Tablo()

    End Sub
    Private Sub GetObjects()
        Me.Visible = False
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim edt As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database
        Dim SelectedSet = cu.GetObjects("INSERT", "Изберете блок")
        If SelectedSet Is Nothing Then
            MsgBox("НЕ Е маркиран нито един блок.")
            Exit Sub
        End If
        Me.Visible = True
        Dim blkRecId As ObjectId = ObjectId.Null
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                Dim acBlkTbl As BlockTable
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                ToolStripProgressBar1.Maximum = SelectedSet.Count
                ToolStripProgressBar1.Value = 0
                For Each sObj As SelectedObject In SelectedSet
                    blkRecId = sObj.ObjectId
                    Dim Kons As strKonsumator
                    Dim acBlkRef As BlockReference = DirectCast(acTrans.GetObject(blkRecId, OpenMode.ForWrite), BlockReference)
                    Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                    Kons.Name = (CType(acBlkRef.DynamicBlockTableRecord.GetObject(OpenMode.ForRead), BlockTableRecord)).Name
                    Kons.ID_Block = blkRecId
                    For Each attId As ObjectId In attCol
                        Dim dbObj As DBObject = acTrans.GetObject(attId, OpenMode.ForRead)
                        ' Преобразува обекта в AttributeReference, за да работи с атрибутите.
                        Dim acAttRef As AttributeReference = dbObj
                        If acAttRef.Tag = "ТАБЛО" Then Kons.ТАБЛО = acAttRef.TextString
                        If acAttRef.Tag = "КРЪГ" Then Kons.ТоковКръг = acAttRef.TextString
                        If acAttRef.Tag = "Pewdn" Then Kons.Pewdn = acAttRef.TextString
                        If acAttRef.Tag = "PEWDN1" Then Kons.PEWDN1 = acAttRef.TextString
                        If acAttRef.Tag = "LED" Then Kons.strМОЩНОСТ = acAttRef.TextString
                        If acAttRef.Tag = "МОЩНОСТ" Then Kons.strМОЩНОСТ = acAttRef.TextString

                    Next
                    ListKonsumator.Add(Kons)
                    ToolStripProgressBar1.Value += 1
                Next
                acTrans.Commit()
            Catch ex As Exception
                MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
                acTrans.Abort()
            End Try
        End Using
    End Sub
End Class