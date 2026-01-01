Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports System.IO
Imports System.Net
Imports System.Security.Cryptography
Imports System.Diagnostics.Eventing.Reader
Imports System.Drawing.Drawing2D
Imports Autodesk.AutoCAD.DatabaseServices.Filters
Imports Autodesk.AutoCAD.GraphicsInterface
Imports Autodesk.AutoCAD.Colors
Public Class Kabel
    Dim PI As Double = 3.1415926535897931
    Dim cu As CommonUtil = New CommonUtil()

    <CommandMethod("KabelKanal")>
    Public Sub KabelKanal()
        Dim cu As CommonUtil = New CommonUtil()
        Dim ss = cu.GetObjects("INSERT", "Изберете блок кабелен канал")

        If ss Is Nothing Then
            MsgBox("Нама маркиран блок кабелен канал в слой 'EL'.")
            Exit Sub
        End If

        Call Insert_Block_Kabel_Kanal(ss)
    End Sub
    Public Sub Insert_Block_Kabel_Kanal(ss As SelectionSet)
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        '
        ' Kabel (*,0) - Тип на линията
        ' Kabel (*,1) - Тип на тръбата
        ' Kabel (*,2) - брой маркирани линии от този тип
        '
        Dim Kabel(10, 2) As String
        Dim blkRecIdNow As ObjectId = ObjectId.Null

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim pPtRes As PromptPointResult
        Dim pPtRes1 As PromptPointResult
        Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")

        Dim Ширина As String = ""
        Dim Височина As String = ""
        Dim Текст As String = ""

        Try
            Using actrans As Transaction = doc.TransactionManager.StartTransaction()
                pPtOpts.Message = vbLf & "Изберете точка на вмъкване на блока: "
                pPtRes = acDoc.Editor.GetPoint(pPtOpts)
                pPtOpts.Message = vbLf & "Изберете точка на постаряне на надписа: "
                pPtRes1 = acDoc.Editor.GetPoint(pPtOpts)
                For Each sObj As SelectedObject In ss
                    Dim blkRecId As ObjectId = sObj.ObjectId
                    Dim acBlkRef As BlockReference = DirectCast(actrans.GetObject(blkRecId, OpenMode.ForWrite), BlockReference)
                    Dim attCol As AttributeCollection = acBlkRef.AttributeCollection

                    Dim blName As String = (CType(acBlkRef.DynamicBlockTableRecord.GetObject(OpenMode.ForRead), BlockTableRecord)).Name

                    For Each objID As ObjectId In attCol
                        Dim dbObj As DBObject = actrans.GetObject(objID, OpenMode.ForWrite)
                        Dim acAttRef As AttributeReference = dbObj
                        Dim ddd As String = ""
                        If acAttRef.Tag = "ШИРИНА" Then Ширина = acAttRef.TextString
                        If acAttRef.Tag = "ВИСОЧИНА" Then Височина = acAttRef.TextString
                    Next

                    Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection
                    For Each prop As DynamicBlockReferenceProperty In props
                        'This Is where you change states based on input
                        If prop.PropertyName = "Ширина" Then Ширина = prop.Value
                        If prop.PropertyName = "Височина" Then Височина = prop.Value
                    Next

                    If InStr(blName, "Скара") Then
                        Текст = "каб. скара "
                    Else
                        Текст = "каб. канал "
                    End If

                Next
            End Using

            Kabel(0, 0) = Текст
            Kabel(0, 1) = ""
            Kabel(0, 2) = "1"
            Kabel(1, 0) = (Val(Ширина) * 10).ToString & "х" & (Val(Височина) * 10).ToString & "mm"
            Kabel(1, 1) = ""
            Kabel(1, 2) = "1"

            If pPtRes Is Nothing Then
                MsgBox("Не е избрана точка на вмъкване на блока")
                Exit Sub
            End If

            If pPtRes1 Is Nothing Then
                MsgBox("Не е избрана точка на вмъкване на надписа")
                Exit Sub
            End If

            Dim InsertPoint As Point3d = pPtRes.Value
            Dim InsertPoint1 As Point3d = pPtRes1.Value

            Dim scale As Double = 1

            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                ' Вмъква блок
                If InStr(Application.DocumentManager.MdiActiveDocument.Name, "_PV_40") Then scale = 40
                If InStr(Application.DocumentManager.MdiActiveDocument.Name, "_PV_20") Then scale = 20
                If InStr(Application.DocumentManager.MdiActiveDocument.Name, "_PV_15") Then scale = 15

                blkRecIdNow = cu.InsertBlock("Кабел", InsertPoint, "EL__DIM", New Scale3d(scale, scale, scale))

                Dim Position_X, Position_Y As Double
                Position_X = InsertPoint1.X - InsertPoint.X
                Position_Y = InsertPoint1.Y - InsertPoint.Y
                Dim acBlkRef As BlockReference =
                    DirectCast(acTrans.GetObject(blkRecIdNow, OpenMode.ForWrite), BlockReference)
                Dim props As DynamicBlockReferencePropertyCollection =
                    acBlkRef.DynamicBlockReferencePropertyCollection
                For Each prop As DynamicBlockReferenceProperty In props
                    'This Is where you change states based on input
                    If prop.PropertyName = "Position X" Then prop.Value = Position_X    ' Позиция на вмъкване коорд. Х
                    If prop.PropertyName = "Position Y" Then prop.Value = Position_Y    ' Позиция на вмъкване коорд. Х
                    If prop.PropertyName = "Visibility" Then prop.Value = "Точка"
                Next

                EditAttributeCollectionKabel(blkRecIdNow, Kabel, "Канал")
                cu.EditDynamicBlockReferenceKabel(blkRecIdNow)

                acTrans.Commit()
            End Using
        Catch ex As Exception
            If blkRecIdNow.IsNull Then
                MsgBox("Възникна грешка: " & "Във файла липсва блок 'Кабел'!!!")
            Else
                MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
            End If
        End Try
    End Sub
    <CommandMethod("KabelInsert")>
    Public Sub KabelInsert()
        Dim cu As CommonUtil = New CommonUtil()

        Dim ss = cu.GetObjects("LINE", "Изберете Линия")

        Dim br As Integer = 0

        If ss Is Nothing Then
            MsgBox("Нама маркирана линия в слой 'EL'.")
            Exit Sub
        End If
        Dim Kabel(10, 2) As String
        Insert_Block_Kabel(cu.GET_LINE_TYPE_KABEL(Kabel, ss, True))
    End Sub
    <CommandMethod("KabelInsert_PV")>
    Public Sub KabelInsert_PV()
        Dim cu As CommonUtil = New CommonUtil()
        Dim br As Integer = 0

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim edt As Editor = doc.Editor
        Dim SelectedSet = cu.GetObjects("INSERT", "Изберете блок")
        Dim scale As Double = 1

        ' Get the current database and start a transaction
        Dim acCurDb As Autodesk.AutoCAD.DatabaseServices.Database
        acCurDb = Application.DocumentManager.MdiActiveDocument.Database
        Dim blkRecId As ObjectId = ObjectId.Null

        Dim Kabel(10, 2) As String

        If SelectedSet Is Nothing Then
            MsgBox("Нама маркирана БЛОК в слой 'EL'.")
            Exit Sub
        End If

        Dim obLayer As String = ""
        Dim obInver As String = ""

        For Each sObj As SelectedObject In SelectedSet
            blkRecId = sObj.ObjectId
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim acBlkRef As BlockReference = DirectCast(acTrans.GetObject(blkRecId, OpenMode.ForWrite), BlockReference)
                Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                obLayer = acBlkRef.Layer
                For Each objID As ObjectId In attCol
                    Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForWrite)
                    Dim acAttRef As AttributeReference = dbObj
                    If acAttRef.Tag = "ИНВЕРТОР" Then obInver = acAttRef.TextString
                Next
                acTrans.Commit()
            End Using
        Next

        Kabel(0, 0) = "Инв. " & obInver
        Kabel(0, 1) = ""
        Kabel(0, 2) = 1

        Kabel(1, 0) = "Стр. " & obInver & "." & Mid(obLayer, 10, Len(obLayer))
        Kabel(1, 1) = ""
        Kabel(1, 2) = 1

        Dim blkKabel As ObjectId = Insert_Block_Kabel(Kabel)

        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkRef As BlockReference = DirectCast(acTrans.GetObject(blkKabel, OpenMode.ForWrite), BlockReference)
            acBlkRef.Layer = obLayer
            acTrans.Commit()
        End Using
    End Sub
    <CommandMethod("KabelАlign")>
    Public Sub KabelАlign()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim edt As Editor = acDoc.Editor
        Dim SelectedSet = cu.GetObjects("INSERT", "Изберете блок които ще се подравняват")
        Dim scale As Double = 1
        ' Get the current database and start a transaction
        Dim acCurDb As Autodesk.AutoCAD.DatabaseServices.Database
        acCurDb = Application.DocumentManager.MdiActiveDocument.Database
        Dim blkRecId As ObjectId = ObjectId.Null
        If SelectedSet Is Nothing Then
            MsgBox("Няма маркиран линия в слой 'EL'.")
            Exit Sub
        End If
        '
        ' Потребителски избор подравняване по Вертикал
        '
        Dim Vertikal As Boolean = False
        Dim pDouOpts As PromptDoubleOptions = New PromptDoubleOptions("")
        With pDouOpts
            .Keywords.Add("Да")
            .Keywords.Add("Не")
            .Keywords.Default = "Да"
            .Message = vbCrLf & "Жалаете ли подравняване по ВЕРТИКАЛ: "
            .AllowZero = False
            .AllowNegative = False
        End With
        Dim pKeyRes As PromptDoubleResult = acDoc.Editor.GetDouble(pDouOpts)
        Vertikal = IIf(pKeyRes.StringResult = "Не", False, True)
        '
        ' Потребителски избор подравняване по Хоризонтал
        '
        Dim Horizontal As Boolean = False
        With pDouOpts
            .Keywords.Default = "Да"
            .Message = vbCrLf & "Жалаете ли подравняване по ХОРИЗОНТАЛ: "
        End With
        pKeyRes = acDoc.Editor.GetDouble(pDouOpts)
        Horizontal = IIf(pKeyRes.StringResult = "Не", False, True)
        '
        ' Потребителски избор подравняване със стъпка
        '
        Dim Stypka As Boolean = False
        With pDouOpts
            .Keywords.Default = "Не"
            .Message = vbCrLf & "Жалаете ли подравняване ВЕРТИКАЛНО със СТЪПКА: "
        End With
        pKeyRes = acDoc.Editor.GetDouble(pDouOpts)
        If pKeyRes.StringResult = "Не" Then
            Stypka = False
        Else
            Stypka = True
            Horizontal = True
            Vertikal = True
        End If

        Dim ssSet = cu.GetObjects("INSERT", "Изберете блок по който ще се подравнява", False)
        If IsNothing(ssSet) Then
            MsgBox("Е що се ебаваш? Помолих те да избереш БЛОК!")
            Exit Sub
        End If
        Dim arrKabelАlign(SelectedSet.Count - 1) As CommonUtil.strKabelАlign
        Dim blKabelАlign As CommonUtil.strKabelАlign
        Try
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                '
                '   Записва параметрите на блока по който ще се подравнява 
                '
                blKabelАlign.Id = ssSet(0).ObjectId
                Dim acSSBlkRef As BlockReference = DirectCast(acTrans.GetObject(blKabelАlign.Id, OpenMode.ForRead), BlockReference)
                Dim attSSCol As AttributeCollection = acSSBlkRef.AttributeCollection
                '
                '   Записва точка на вмъкване на блока
                '
                blKabelАlign.pInsert = New Point3d(acSSBlkRef.Position.X, acSSBlkRef.Position.Y, acSSBlkRef.Position.Z)
                '
                '   Записва точка позицията на линията
                '
                Dim ssprops As DynamicBlockReferencePropertyCollection = acSSBlkRef.DynamicBlockReferencePropertyCollection
                For Each prop As DynamicBlockReferenceProperty In ssprops
                    If prop.PropertyName = "Position X" Then blKabelАlign.Position_X = prop.Value ' Позиция на вмъкване коорд. Х
                    If prop.PropertyName = "Position Y" Then blKabelАlign.Position_Y = prop.Value ' Позиция на вмъкване коорд. Y
                Next
                Dim br_NA4IN As Integer = 0
                If Stypka Then
                    For Each objID As ObjectId In attSSCol
                        Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                        Dim acAttRef As AttributeReference = dbObj
                        If InStr(acAttRef.Tag, "NA4IN_") > 0 Then blKabelАlign.Stypka = blKabelАlign.Stypka + IIf(acAttRef.TextString <> "", 1, 0)
                    Next
                End If
                '
                '-------------------------------------------------------------------------------------------------------------------------------------
                '   Записва параметрите на блоковете които ще се подравняват
                '-------------------------------------------------------------------------------------------------------------------------------------
                '
                Dim Index As Integer = 0
                For Each sObj As SelectedObject In SelectedSet
                    arrKabelАlign(Index).Id = sObj.ObjectId
                    Dim acBlkRef As BlockReference = DirectCast(acTrans.GetObject(arrKabelАlign(Index).Id, OpenMode.ForWrite), BlockReference)
                    Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                    '
                    '   Записва точка на вмъкване на блока
                    '
                    arrKabelАlign(Index).pInsert = New Point3d(acBlkRef.Position.X, acBlkRef.Position.Y, acBlkRef.Position.Z)
                    '
                    '   Записва точка позицията на линията
                    '
                    Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection
                    For Each prop As DynamicBlockReferenceProperty In props
                        If prop.PropertyName = "Position X" Then arrKabelАlign(Index).Position_X = prop.Value ' Позиция на вмъкване коорд. Х
                        If prop.PropertyName = "Position Y" Then arrKabelАlign(Index).Position_Y = prop.Value ' Позиция на вмъкване коорд. Y
                    Next
                    br_NA4IN = 0
                    If Stypka Then
                        For Each objID As ObjectId In attCol
                            Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                            Dim acAttRef As AttributeReference = dbObj
                            If InStr(acAttRef.Tag, "NA4IN_") > 0 Then br_NA4IN = br_NA4IN + IIf(acAttRef.TextString <> "", 1, 0)
                        Next
                        arrKabelАlign(Index).Stypka = br_NA4IN
                    End If
                    Index += 1
                Next
                acTrans.Commit()
            End Using
            cu.Kabel_Aligment(blKabelАlign, New List(Of CommonUtil.strKabelАlign)(arrKabelАlign), Stypka, Vertikal, Horizontal)
        Catch ex As Exception
            MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
        End Try
    End Sub
    <CommandMethod("KabelEdit")>
    Public Sub KabelEdit()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim edt As Editor = acDoc.Editor
        Dim SelectedSet = cu.GetObjects("INSERT", "Изберете блок")
        Dim scale As Double = 1

        ' Get the current database and start a transaction
        Dim acCurDb As Autodesk.AutoCAD.DatabaseServices.Database
        acCurDb = Application.DocumentManager.MdiActiveDocument.Database
        Dim blkRecId As ObjectId = ObjectId.Null
        If SelectedSet Is Nothing Then
            MsgBox("Няма маркиран линия в слой 'EL'.")
            Exit Sub
        End If
        For Each sObj As SelectedObject In SelectedSet
            blkRecId = sObj.ObjectId
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                cu.EditDynamicBlockReferenceKabel(blkRecId)
                acTrans.Commit()
            End Using
        Next
    End Sub
    <CommandMethod("KabelLineType")>
    Public Sub KabelLineType()
        ' Публична процедура, която определя типа на линия на основата на определен слой
        ' Създаване на нов обект от клас CommonUtil
        Dim cu As CommonUtil = New CommonUtil()
        ' Извличане на обекти от типа "LINE" чрез метода GetObjects на обекта cu
        ' Потребителят е помолен да избере линия
        Dim ss = cu.GetObjects("LINE", "Изберете Линия")
        ' Получаване на текущия документ и база данни на AutoCAD
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        ' Проверка дали е избран обект (линия)
        If ss Is Nothing Then
            ' Ако няма избрани обекти, показване на съобщение и излизане от процедурата
            MsgBox("Няма маркирана линия в слой 'EL'.")
            Exit Sub
        End If
        ' Опит за изпълнение на следния блок код, с уловка на евентуални грешки
        Try
            ' Деклариране на променливи за слой и тип на линията
            Dim l_Layer As String = ""
            Dim l_Type_Get As String = ""
            Dim l_Type_Set As String = ""
            ' Започване на транзакция за промени в базата данни на AutoCAD
            Using trans As Transaction = acDoc.TransactionManager.StartTransaction()

                Dim idsToMove As New ObjectIdCollection()
                ' Вземане на текущата база данни на чертежа
                Dim db As Database = HostApplicationServices.WorkingDatabase
                ' Получаване на таблицата с типовете линии в чертежа
                Dim linetypeTable As LinetypeTable = trans.GetObject(db.LinetypeTableId, OpenMode.ForRead)

                ' Таблица за DrawOrder на текущото пространство
                Dim btr As BlockTableRecord = trans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite)
                Dim dot As DrawOrderTable = trans.GetObject(btr.DrawOrderTableId, OpenMode.ForWrite)

                ' Обхождане на всички избрани обекти
                For Each sObj As SelectedObject In ss
                    ' Опит за получаване на обекта като линия с режим за писане
                    Dim line As Line = TryCast(trans.GetObject(sObj.ObjectId, OpenMode.ForWrite), Line)
                    ' Проверка дали слоят на линията не е "ByLayer"
                    If line.Layer <> "ByLayer" Then
                        ' Извикване на метод от обекта cu за определяне на типа линия на базата на слоя
                        l_Type_Set = cu.SET_line_Type(line.Layer)
                        ' Проверка дали зададеният тип линия (l_Type_Set) съществува в таблицата с типовете линии
                        If linetypeTable.Has(l_Type_Set) Then
                            ' Ако типът линия съществува, той се задава на линията
                            line.Linetype = l_Type_Set
                        Else
                            ' Ако типът линия не съществува, се задава резервен вариант (ByLayer)
                            ' Можеш да изведеш съобщение за липсващ тип линия
                            MsgBox("Типът линия ''" & l_Type_Set & "'' не съществува в чертежа!", MsgBoxStyle.Exclamation)
                            Continue For
                        End If
                        line.LineWeight = LineWeight.ByLayer
                        line.ColorIndex = 256
                        line.LinetypeScale = 1.1
                        ' Промяна на Z координата на стартовата и крайната точка на линията на 0
                        line.StartPoint = New Point3d(line.StartPoint.X, line.StartPoint.Y, 0)
                        line.EndPoint = New Point3d(line.EndPoint.X, line.EndPoint.Y, 0)

                        'Добавяне към списъка за Draw Order
                        idsToMove.Add(line.ObjectId)
                    End If

                Next
                ' Преместване на филтрираните линии най-отгоре
                If idsToMove.Count > 0 Then
                    dot.MoveToTop(idsToMove)
                End If
                ' Потвърждаване на транзакцията, за да се запазят промените
                trans.Commit()
            End Using
            ' Улавяне на евентуална грешка по време на изпълнение на кода
        Catch ex As Exception
            ' Показване на съобщение за грешка с описание и стек на грешката
            MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
        End Try
    End Sub
    '==============================
    ' Функция: Insert_Block_Kabel
    ' Цел: Вмъкване на блок "Кабел" в текущ AutoCAD документ.
    ' Входни параметри:
    '   Kabel - масив с информация за кабела (за редактиране на атрибутите)
    ' Връща:
    '   ObjectId на вмъкнатия блок, или ObjectId.Null при грешка
    '==============================
    Function Insert_Block_Kabel(Kabel As Array) As ObjectId

        ' Вземане на текущия AutoCAD документ
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        ' Вземане на базата данни на документа
        Dim acCurDb As Database = acDoc.Database
        ' Инициализация на ObjectId за новия блок
        Dim blkRecIdNow As ObjectId = ObjectId.Null
        ' Определяне на начална видимост на блока в зависимост от името на документа
        Dim strVisib As String = IIf(InStr(Application.DocumentManager.MdiActiveDocument.Name, "_PV") > 0, "Точка", "Линии")
        ' Проверка на стека на извикванията (StackTrace), за да се промени видимостта, ако е извикано от конкретен метод
        Dim trace As New StackTrace()
        For i As Integer = 0 To trace.FrameCount - 1
            ' Ако методът "Insert_Kabel_Ka4vane" е в стека → задаваме видимост "Точка"
            If trace.GetFrame(i).GetMethod().Name = "Insert_Kabel_Ka4vane" Then strVisib = "Точка"
        Next

        Try
            ' Настройка на опциите за избор на точки от потребителя
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim pPtRes As PromptPointResult          ' Първата точка (вмъкване на блока)
            Dim pPtRes1 As PromptPointResult         ' Втората точка (позиция на надписа)
            Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")

            ' Извличане на точките от потребителя
            Using trans As Transaction = doc.TransactionManager.StartTransaction()
                ' Първата точка
                pPtOpts.Message = vbLf & "Изберете точка на вмъкване на блока: "
                pPtRes = acDoc.Editor.GetPoint(pPtOpts)

                ' Втората точка
                pPtOpts.Message = vbLf & "Изберете точка на поставяне на надписа: "
                pPtRes1 = acDoc.Editor.GetPoint(pPtOpts)
            End Using

            ' Проверка дали потребителят е избрал точките
            If pPtRes Is Nothing Then
                MsgBox("Не е избрана точка на вмъкване на блока")
                Exit Function
            End If
            If pPtRes1 Is Nothing Then
                MsgBox("Не е избрана точка на вмъкване на надписа")
                Exit Function
            End If

            ' Запис на избраните точки като Point3d обекти
            Dim InsertPoint As Point3d = pPtRes.Value
            Dim InsertPoint1 As Point3d = pPtRes1.Value

            ' Инициализация на мащаб
            Dim scale As Double = 1

            ' Отваряне на транзакция за редакция на блока
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                ' Определяне на мащаба на блока в зависимост от името на документа
                If InStr(Application.DocumentManager.MdiActiveDocument.Name, "_PV_40") > 0 Then scale = 40
                If InStr(Application.DocumentManager.MdiActiveDocument.Name, "_PV_20") > 0 Then scale = 20
                If InStr(Application.DocumentManager.MdiActiveDocument.Name, "_PV_30") > 0 Then scale = 30
                If InStr(Application.DocumentManager.MdiActiveDocument.Name, "_PV_15") > 0 Then scale = 15

                ' Вмъкване на блока "Кабел" с указаната точка и мащаб
                blkRecIdNow = cu.InsertBlock("Кабел", InsertPoint, "EL__DIM", New Scale3d(scale, scale, scale))

                ' Изчисляване на разликата между точката на блока и точката на надписа
                Dim Position_X, Position_Y As Double
                Position_X = InsertPoint1.X - InsertPoint.X
                Position_Y = InsertPoint1.Y - InsertPoint.Y

                ' Вземане на препратка към блока за редакция
                Dim acBlkRef As BlockReference = DirectCast(acTrans.GetObject(blkRecIdNow, OpenMode.ForWrite), BlockReference)

                ' Достъп до динамичните свойства на блока
                Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection

                ' Настройка на динамичните свойства
                For Each prop As DynamicBlockReferenceProperty In props
                    If prop.PropertyName = "Position X" Then prop.Value = Position_X
                    If prop.PropertyName = "Position Y" Then prop.Value = Position_Y
                    If prop.PropertyName = "Visibility" Then prop.Value = strVisib
                Next

                ' Редактиране на атрибутите на блока "Кабел" (например тръби и кабелни данни)
                EditAttributeCollectionKabel(blkRecIdNow, Kabel, "Тръби")
                ' Допълнителна обработка на динамичните свойства
                cu.EditDynamicBlockReferenceKabel(blkRecIdNow)

                ' Потвърждаване на всички промени в транзакцията
                acTrans.Commit()
            End Using

        Catch ex As Exception
            ' Обработка на грешки
            If blkRecIdNow.IsNull Then
                ' Ако блокът не е вмъкнат (липсва в библиотеката)
                MsgBox("Възникна грешка: " & "Във файла липсва блок 'Кабел'!!!")
            Else
                ' Показване на съобщение с грешката и стек на извикванията
                MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
            End If
        End Try
        ' Връщане на ObjectId на вмъкнатия блок
        Return blkRecIdNow
    End Function
    ' <summary>
    ' Функцията EditAttributeCollectionKabel има за цел да редактира атрибутите на блокови обекти в AutoCAD,
    ' като използва данни от подаден масив Kabel.
    ' Тя извършва сортиране на масива и след това обновява атрибутите на блоковия обект
    ' според тези данни и подадената стойност на Polagane.
    ' </summary>
    Private Sub EditAttributeCollectionKabel(blkRecId As ObjectId,
                                             Kabel As Array,
                                             Polagane As String)
        ' Взема активния документ на AutoCAD
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        ' Взема базата данни на документа
        Dim acCurDb As Database = acDoc.Database
        Dim br As Integer  ' Променлива за брояч/индекс на кабелите
        Dim posokaText As String = "" ' Променлива за посоката на кабела
        Dim KotaText As String = ""   ' Променлива за котата на кабела
        ' Търси кабел с дължина = 100 и запазва неговата кота и посока
        For i = 0 To Kabel.GetUpperBound(0)
            If Kabel(i, 2) = 100 Then
                KotaText = Kabel(i, 0)    ' Запазва котата
                posokaText = Kabel(i, 1)  ' Запазва посоката
                ' Обнулява стойностите в масива за този ред
                Kabel(i, 0) = Nothing
                Kabel(i, 1) = "ByLayer"
                Kabel(i, 2) = "0"
                KotaText = cu.GetObjects_TEXT("Изберете текст съдържаш котата :")
                Exit For
            End If
        Next
        '--------------------------------------------------------------------------------------------
        ' Сортиране на масива Kabel по първата колона (кота)
        '--------------------------------------------------------------------------------------------
        For i = 0 To Kabel.GetUpperBound(0)
            For j = i + 1 To Kabel.GetUpperBound(0)
                If Kabel(i, 0) < Kabel(j, 0) Then
                    ' Размяна на редовете
                    For k = 0 To Kabel.GetUpperBound(1)
                        Dim temp As String = Kabel(i, k)
                        Kabel(i, k) = Kabel(j, k)
                        Kabel(j, k) = temp
                    Next
                End If
            Next
        Next
        '--------------------------------------------------------------------------------------------
        ' Сортиране на масива Kabel по втората колона (посока/полагане)
        '--------------------------------------------------------------------------------------------
        For i = 0 To Kabel.GetUpperBound(0)
            For j = i + 1 To Kabel.GetUpperBound(0)
                If Kabel(i, 1) < Kabel(j, 1) Then
                    ' Размяна на редовете
                    For k = 0 To Kabel.GetUpperBound(1)
                        Dim temp As String = Kabel(i, k)
                        Kabel(i, k) = Kabel(j, k)
                        Kabel(j, k) = temp
                    Next
                End If
            Next
        Next
        '--------------------------------------------------------------------------------------------
        ' Инициализация на временен масив за обработка на кабели
        '--------------------------------------------------------------------------------------------
        Dim Kabel_Pom(10) As String
        For i = 0 To UBound(Kabel_Pom)
            Kabel_Pom(i) = "" ' Задава празен низ на всички позиции
        Next
        br = 1
        Dim polag As String = Kabel(0, 1) ' Запазва първоначалното място на полагане

        ' Добавя първия кабел в помощния масив, ако има дължина > 0
        Kabel_Pom(0) = IIf(Val(Kabel(0, 2)) > 0,
                           Kabel(0, 2) & "л. " & Kabel(0, 0), "")

        ' Обхожда останалите кабели и подготвя текстовете за атрибутите
        For i = 1 To Kabel.GetUpperBound(0)
            ' Ако дължината на кабела е 0, записва мястото на полагане и прекъсва
            If Val(Kabel(i, 2)) = 0 Then
                Kabel_Pom(br) = polag
                Exit For
            End If
            ' Ако мястото на полагане се променя, записва предишното и актуализира текущото
            If polag <> Kabel(i, 1) Then
                Kabel_Pom(br) = polag
                polag = Kabel(i, 1)
                br += 1
            End If
            ' Записва текста за атрибута, различно в зависимост от Polagane
            Kabel_Pom(br) = IIf(Polagane = "Канал", Kabel(i, 0), Kabel(i, 2) & "л. " & Kabel(i, 0))
            br += 1
        Next
        '--------------------------------------------------------------------------------------------
        ' Стартира транзакция за промяна на блоковите атрибути
        '--------------------------------------------------------------------------------------------
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            ' Взема референция към блока
            Dim acBlkRef As BlockReference = DirectCast(acTrans.GetObject(blkRecId, OpenMode.ForWrite), BlockReference)
            ' Взема колекцията от атрибути на блока
            Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
            ' Проверка дали имаме запазени стойности за посока и кота
            If Not String.IsNullOrWhiteSpace(posokaText) AndAlso Not String.IsNullOrWhiteSpace(KotaText) Then
                ' Намери първия незапълнен елемент в помощния масив и го запълва
                For i = 0 To Kabel_Pom.GetUpperBound(0)
                    If String.IsNullOrWhiteSpace(Kabel_Pom(i)) Then
                        Kabel_Pom(i) = posokaText & " " & KotaText
                        Exit For
                    End If
                Next
            End If
            For Each objID As ObjectId In attCol
                Dim acAttRef As AttributeReference = acTrans.GetObject(objID, OpenMode.ForWrite)
                ' Проверява дали атрибутът започва с "NA4IN_"
                If acAttRef.Tag.StartsWith("NA4IN_") Then
                    Dim idx As Integer
                    If Integer.TryParse(acAttRef.Tag.Substring(6), idx) Then
                        If idx <= UBound(Kabel_Pom) Then
                            acAttRef.TextString = Kabel_Pom(idx)
                        End If
                    End If
                End If
            Next
            acTrans.Commit() ' Потвърждава промените
        End Using
    End Sub


End Class