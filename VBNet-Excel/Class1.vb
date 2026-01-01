Imports acApp = Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Runtime
Imports System.Drawing.Imaging
Imports System.Drawing
Public Class Commands
    <CommandMethod("CSS")>
    Public Shared Sub CaptureScreenShot()
        ScreenShotToFile(acApp.Application.DocumentManager.MdiActiveDocument.Window, System.IO.Path.GetTempPath() + "doc-window.jpg", 30, 26, 10, 10)
    End Sub

    Private Shared Sub ScreenShotToFile(ByVal wd As Autodesk.AutoCAD.Windows.Window,
                                        ByVal filename As String,
                                        ByVal top As Integer,
                                        ByVal bottom As Integer,
                                        ByVal left As Integer,
                                        ByVal right As Integer)


        Dim pt As System.Windows.Point = wd.DeviceIndependentLocation
        Dim sz As System.Windows.Size = wd.DeviceIndependentSize

        pt.X += left
        pt.Y += top
        sz.Height -= top + bottom
        sz.Width -= left + right
        Dim bmp As Bitmap = New Bitmap(CInt(sz.Width), CInt(sz.Height), PixelFormat.Format32bppArgb)

        Using bmp
            Using gfx As Graphics = Graphics.FromImage(bmp)
                gfx.CopyFromScreen(CInt(pt.X), CInt(pt.Y), 0, 0, New System.Drawing.Size(CInt(sz.Width), CInt(sz.Height)), CopyPixelOperation.SourceCopy)
                bmp.Save(filename, ImageFormat.Jpeg)
            End Using
        End Using
    End Sub
End Class

