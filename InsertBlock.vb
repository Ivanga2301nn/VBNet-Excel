Imports Microsoft.VisualBasic

Public Class Kabel

    <CommandMethod("InserKabel")>

        Public Sub InserKabel()
            Dim cu As CommonUtil = New CommonUtil()
            Dim blockRef As BlockReference
            Dim blockName As String
            Dim pp As Variable
            ' Съберете всички линии в набор за избор
            Dim cu As CommonUtil = New CommonUtil()
            Dim ss = cu.GetObjects("LINE")

            If ss Is Nothing Then
                Return "No Line found in the drawing."
            End If
            ggg

        End Sub

    End Class
