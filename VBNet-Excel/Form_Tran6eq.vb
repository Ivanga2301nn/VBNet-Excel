Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Windows.Forms
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime

Public Class Form_Tran6eq

    ' Това е AutoCAD команда, която стартира формата
    <CommandMethod("Tran6eq")>
    Public Sub Tran6eq()
        ' Създаваме нова инстанция на формата и я показваме
        Dim form As New Form_Tran6eq()
        form.Show()
    End Sub

    ' Събитие, което се изпълнява при зареждане на формата
    Private Sub Form_Tran6eq_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' При зареждане на формата извикваме метода за зареждане на preview изображение
        ' на блока с името "Траншея"
        LoadBlockPreviewImage("Траншея")
    End Sub

    ' Метод за зареждане на preview изображение на блок с дадено име
    Private Sub LoadBlockPreviewImage(blockName As String)
        Try
            ' Вземаме текущия активен документ в AutoCAD
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            'Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            ' Вземаме базата данни на документа
            Dim db As Database = doc.Database

            ' Стартираме транзакция за четене от базата данни
            Using trans As Transaction = db.TransactionManager.StartTransaction()
                ' Отваряме таблицата с блокове за четене
                Dim bt As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)

                ' Проверяваме дали блок с даденото име съществува
                If Not bt.Has(blockName) Then
                    ' Ако няма такъв блок, показваме съобщение и прекратяваме метода
                    MessageBox.Show("Блокът '" & blockName & "' не съществува в документа.")
                    Return
                End If

                ' Вземаме BlockTableRecord за блока с името blockName
                Dim btr As BlockTableRecord = trans.GetObject(bt(blockName), OpenMode.ForRead)

                '' Взимаме preview изображението на блока като Bitmap
                'Dim bmp As Bitmap = btr.GetPreviewBitmap()

                '' Проверяваме дали preview изображението не е Nothing
                'If bmp IsNot Nothing Then
                '    ' Задаваме изображението на PictureBox контрола pbCanvas
                '    pbCanvas.Image = bmp
                '    ' Задаваме режима на мащабиране, за да се вижда цялото изображение красиво
                '    pbCanvas.SizeMode = PictureBoxSizeMode.Zoom
                'Else
                '    ' Ако няма preview изображение, показваме съобщение
                '    MessageBox.Show("НЕ Е preview изображение за блока '" & blockName & "'.")
                'End If

                ' Завършваме транзакцията
                trans.Commit()
            End Using

        Catch ex As System.Exception
            ' Ако възникне грешка при изпълнение, показваме я на потребителя
            MessageBox.Show("Грешка при зареждане на изображение: " & ex.Message)
        End Try
    End Sub

    Private Sub LoadBlockThumbnailImage(blockName As String)
        Try
            Dim doc As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Autodesk.AutoCAD.DatabaseServices.Database = doc.Database

            Using trans As Autodesk.AutoCAD.DatabaseServices.Transaction = db.TransactionManager.StartTransaction()
                Dim bt As Autodesk.AutoCAD.DatabaseServices.BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)

                If Not bt.Has(blockName) Then
                    MessageBox.Show("Блокът '" & blockName & "' не съществува.")
                    Return
                End If

                Dim btr As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = trans.GetObject(bt(blockName), OpenMode.ForRead)



                trans.Commit()
            End Using
        Catch ex As Exception
            MessageBox.Show("Грешка: " & ex.Message)
        End Try
    End Sub

End Class
