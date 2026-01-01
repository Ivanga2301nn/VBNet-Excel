' Standard .NET namespaces
Imports System.Runtime.InteropServices

' Main AutoCAD namespaces
Imports Autodesk.AutoCAD
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.PlottingServices
Imports Autodesk.AutoCAD.GraphicsInterface
Imports System.Drawing
Public Class CreateLayout
    Private Function CreateLayout(name As String)
        ' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        ' Get the layout and plot settings of the named pagesetup
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            ' Reference the Layout Manager
            Dim acLayoutMgr As LayoutManager = LayoutManager.Current

            ' Create the new layout with default settings
            Dim objID As ObjectId = acLayoutMgr.CreateLayout(name)

            ' Open the layout
            Dim acLayout As Layout = acTrans.GetObject(objID,
                                                       OpenMode.ForRead)

            ' Set the layout current if it is not already
            If acLayout.TabSelected = False Then
                acLayoutMgr.CurrentLayout = acLayout.LayoutName
            End If

            ' Output some information related to the layout object
            acDoc.Editor.WriteMessage(vbLf & "Tab Order: " & acLayout.TabOrder &
                                      vbLf & "Tab Selected: " & acLayout.TabSelected &
                                      vbLf & "Block Table Record ID: " &
                                      acLayout.BlockTableRecordId.ToString())

            Return acLayout.BlockTableRecordId
            ' Save the changes made
            acTrans.Commit()
        End Using
    End Function
    <CommandMethod("CreateLayoutAndViewport")>
    Public Sub CreateLayoutAndViewport(layoutName As String, viewportCenter As Point2d, viewportSize As Size)
        Dim acCurDb As Database = Application.DocumentManager.MdiActiveDocument.Database

        layoutName = "ooooooooo"
        viewportCenter = New Point2d(0, 0)
        viewportSize = New Size(200, 200)

        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            ' Получаване на LayoutManager
            Dim acLayoutMgr As LayoutManager = LayoutManager.Current

            ' Създаване на нов Layout, ако не съществува
            Dim acLayoutId As ObjectId = acLayoutMgr.GetLayoutId(layoutName)
            If acLayoutId.IsNull Then
                acLayoutId = acLayoutMgr.CreateLayout(layoutName)
            End If

            ' Установяване на новия Layout като текущ
            acLayoutMgr.CurrentLayout = layoutName

            ' Получаване на BlockTableRecord за новия Layout
            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acLayoutId, OpenMode.ForWrite)

            ' Създаване на нов Viewport
            Dim acVport As New Autodesk.AutoCAD.DatabaseServices.Viewport()
            acVport.CenterPoint = New Point3d(viewportCenter.X, viewportCenter.Y, 0)
            acVport.Width = viewportSize.Width
            acVport.Height = viewportSize.Height
            acVport.ViewCenter = New Point2d(viewportCenter.X, viewportCenter.Y)
            acVport.ViewHeight = viewportSize.Height
            acVport.Width = viewportSize.Width

            ' Добавяне на Viewport към BlockTableRecord
            acBlkTblRec.AppendEntity(acVport)
            acTrans.AddNewlyCreatedDBObject(acVport, True)

            ' Запазване на промените и приключване на транзакцията
            acTrans.Commit()
        End Using
    End Sub
    ' Used to create a rectangular and nonrectangular viewports - RapidRT example
    <CommandMethod("RRTCreatViewportsAndSetShadePlot")>
    Public Sub RRTCreatViewportsAndSetShadePlot()
        ' Get the current document and database, and start a transaction
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            ' Open the Block table for read
            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId,
                                                       OpenMode.ForRead)

            ' Open the Block table record Paper space for write
            Dim acBlkTblRec As BlockTableRecord =
            acTrans.GetObject(acBlkTbl(BlockTableRecord.PaperSpace),
                              OpenMode.ForWrite)

            ' Create a Viewport
            Using acVport1 As Autodesk.AutoCAD.DatabaseServices.Viewport =
                       New Autodesk.AutoCAD.DatabaseServices.Viewport()
                ' Set the center point and size of the viewport
                acVport1.CenterPoint = New Point3d(3.75, 4, 0)
                acVport1.Width = 7.5
                acVport1.Height = 7.5

                ' Lock the viewport
                acVport1.Locked = True

                ' Set the scale to 1" = 4'
                acVport1.CustomScale = 48

                ' Set visual style
                Dim vStyles As DBDictionary =
                acTrans.GetObject(acCurDb.VisualStyleDictionaryId,
                                  OpenMode.ForRead)

                acVport1.SetShadePlot(ShadePlotType.VisualStyle,
                                  vStyles.GetAt("Sketchy"))

                ' Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acVport1)
                acTrans.AddNewlyCreatedDBObject(acVport1, True)

                ' Change the view direction
                acVport1.ViewDirection = New Vector3d(-1, -1, 1)

                ' Create a rectangular viewport to change to a non-rectangular viewport
                Using acVport2 As Autodesk.AutoCAD.DatabaseServices.Viewport =
                       New Autodesk.AutoCAD.DatabaseServices.Viewport()

                    acVport2.CenterPoint = New Point3d(9, 6.5, 0)
                    acVport2.Width = 2.5
                    acVport2.Height = 2.5

                    ' Set the scale to 1" = 8'
                    acVport2.CustomScale = 96

                    ' Set render preset
                    Dim namedObjs As DBDictionary =
                    acTrans.GetObject(acCurDb.NamedObjectsDictionaryId,
                                      OpenMode.ForRead)

                    ' Check to see if the Render Settings dictionary already exists
                    Dim renderSettings As DBDictionary
                    If (namedObjs.Contains("ACAD_RENDER_RAPIDRT_SETTINGS") = True) Then
                        renderSettings = acTrans.GetObject(
                        namedObjs.GetAt("ACAD_RENDER_RAPIDRT_SETTINGS"),
                        OpenMode.ForWrite)
                    Else
                        ' If it does not exist, create it and add it to the drawing
                        acTrans.GetObject(acCurDb.NamedObjectsDictionaryId, OpenMode.ForWrite)
                        renderSettings = New DBDictionary()
                        namedObjs.SetAt("ACAD_RENDER_RAPIDRT_SETTINGS", renderSettings)
                        acTrans.AddNewlyCreatedDBObject(renderSettings, True)
                    End If

                    ' Create a new render preset and assign it to the new viewport
                    Dim renderSetting As RapidRTRenderSettings

                    If (renderSettings.Contains("MyPreset") = False) Then
                        renderSetting = New RapidRTRenderSettings()
                        renderSetting.Name = "MyPreset"
                        renderSetting.Description = "Custom new render preset"

                        renderSettings.SetAt("MyPreset", renderSetting)
                        acTrans.AddNewlyCreatedDBObject(renderSetting, True)
                    Else
                        renderSetting = acTrans.GetObject(
                            renderSettings.GetAt("MyPreset"), OpenMode.ForRead)
                    End If

                    acVport2.SetShadePlot(ShadePlotType.RenderPreset,
                                      renderSetting.ObjectId)
                    renderSetting.Dispose()

                    ' Create a circle
                    Using acCirc As Circle = New Circle()
                        acCirc.Center = acVport2.CenterPoint
                        acCirc.Radius = 1.25

                        ' Add the new object to the block table record and the transaction
                        acBlkTblRec.AppendEntity(acCirc)
                        acTrans.AddNewlyCreatedDBObject(acCirc, True)

                        ' Add the new object to the block table record and the transaction
                        acBlkTblRec.AppendEntity(acVport2)
                        acTrans.AddNewlyCreatedDBObject(acVport2, True)

                        ' Clip the viewport using the circle  
                        acVport2.NonRectClipEntityId = acCirc.ObjectId
                        acVport2.NonRectClipOn = True
                    End Using

                    ' Change the view direction
                    acVport2.ViewDirection = New Vector3d(0, 0, 1)

                    ' Enable the viewports
                    acVport1.On = True
                    acVport2.On = True
                End Using
            End Using

            ' Save the new objects to the database
            acTrans.Commit()
        End Using

        ' Switch to the last named layout
        acDoc.Database.TileMode = False
    End Sub
End Class