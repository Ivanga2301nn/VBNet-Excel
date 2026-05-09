Imports System.Windows.Forms
Imports Autodesk.AutoCAD
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime

Public Class Form_Tablo_new_AutoCadInserter
    Dim PI As Double = 3.1415926535897931
    Dim cu As CommonUtil = New CommonUtil()
    Dim twoBus As Boolean
    Dim hasDisconnector As Boolean
    '===============================================================================
    ' КОНСТРУКТОР – Инициализация на зависимостите
    '===============================================================================
    ' Приема списъка с моторни защити (GV_Database) при създаване на обекта.
    ' Предимства:
    ' • Данните са достъпни за всички методи в класа без да се препращат като аргументи
    ' • Класът не зависи от главната форма, а само от подадените му данни
    ' • По-лесна поддръжка, четимост и бъдещо тестване
    '===============================================================================
    Private ReadOnly gvDatabase As List(Of Form_Tablo_new.GV_Entry)
    Public Sub New(gvDb As List(Of Form_Tablo_new.GV_Entry))
        ' Защита срещу грешка при подаден празен/нищолен списък
        If gvDb Is Nothing Then
            Throw New ArgumentNullException(NameOf(gvDb), "GV_Database не може да бъде Nothing.")
        End If
        ' Запазваме референцията – ще остане непроменена през целия живот на обекта
        Me.gvDatabase = gvDb
    End Sub
    ''' <summary>
    ''' Глобални променливи за таблата
    ''' </summary>
    Public widthColom As Double = 120      ' Ширина на всяка колона в таблицата
    Public heightRow As Double = 25        ' Височина на редовете
    Public widthText As Double = 140       ' Ширина на колоната за текст (напр. "Токов кръг")
    Public widthTextDim As Double = 40     ' Допълнителна ширина за текстова колона (напр. за единици)
    Public lengthProw As Double = 90       ' Дължина на вертикалните линии между текст и блокове
    Public lengthProwBlock As Double = 0   ' Дължина на линията под блока за прекъсвач (ще се изчислява по-късно)
    Public padingText As Double = 3        ' Отстояние на текста от линиите
    Public widthTablo As Double = 410      ' Ширина на цялото табло (за блокове и линии)
    Public heightText As Double = 12       ' Височина на текста, използван в блоковете
    Public Y_Шина As Double = 620          ' Вертикална позиция на шината (Y координата)
    ''' <summary>
    ''' Структура за дефиниция на линия за чертане
    ''' </summary>
    Private Structure LineDefinition
        Public StartPoint As Point3d
        Public EndPoint As Point3d
        Public Layer As String
        Public LineWeightValue As Integer
        Public LineType As String
        Public ColorIndex As Integer
        Public Sub New(startPoint As Point3d, endPoint As Point3d, layer As String,
                   lineWeightValue As Integer,
                   lineType As String, Optional colorIndex As Integer = -1)
            Me.StartPoint = startPoint
            Me.EndPoint = endPoint
            Me.Layer = layer
            Me.LineWeightValue = lineWeightValue
            Me.LineType = lineType
            Me.ColorIndex = colorIndex
        End Sub
    End Structure
    ''' <summary>
    ''' Конфигурация за управляващо устройство
    ''' </summary>
    Private Structure ControlDeviceConfig
        Public Str_1 As String
        Public Str_2 As String
        Public Str_3 As String
        Public Str_4 As String
        Public Str_5 As String
        Public ShortName As String
        Public Sub New(str_1 As String, str_2 As String, str_3 As String,
                   str_4 As String, str_5 As String, shortName As String)
            Me.Str_1 = str_1
            Me.Str_2 = str_2
            Me.Str_3 = str_3
            Me.Str_4 = str_4
            Me.Str_5 = str_5
            Me.ShortName = shortName
        End Sub
    End Structure
    ' ============================================================
    ' РЕЧНИК ЗА УПРАВЛЕНИЕ -> БЛОК
    ' ============================================================
    Dim ControlBlockMap As New Dictionary(Of String, String) From {
            {"Импулсно реле", "s_tl"},
            {"Контактор", "s_ct_cont_no"},
            {"Моторна защита", "s_tesys_cont_no"},
            {"Моторен механизъм", "s_ns100_motor_fixed"},
            {"Честотен регулатор", "s_altivar"},
            {"Стълбищен автомат", "s_min"},
            {"Електромер", "s_Wh_meter"},
            {"Фото реле", "s_switch_light_sens"}
        }
    Public Sub ExecuteInsert(panelCircuits As List(Of Form_Tablo_new.strTokow), selectedTablo As String)
        ' Вземане на текущия AutoCAD документ, редактор и база
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim edt As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database
        ' ВЗЕМИ БАЗОВА ТОЧКА ОТ ПОТРЕБИТЕЛЯ
        Dim ptBasePointRes As PromptPointResult
        Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
        pPtOpts.Message = vbLf & "Изберете долен ляв ъгъл на таблото: "
        ptBasePointRes = acDoc.Editor.GetPoint(pPtOpts)
        If ptBasePointRes.Status = PromptStatus.Cancel Then Exit Sub
        Dim ptBasePoint As Point3d = ptBasePointRes.Value
        ' Проверяваме дали има кръгове на отделна шина
        twoBus = panelCircuits.Any(Function(c) c.Шина)
        hasDisconnector = panelCircuits.Any(Function(c) c.Device = "Разединител")
        If twoBus Then
            ' Проверяваме дали НЯМА нито един елемент с Device = "Разединител"
            If Not hasDisconnector Then
                ' Извеждаме съобщение и прекратяваме процедурата
                MessageBox.Show("Две шини – добре. Разединител – няма. Софтуерът изпада в депресия!")
                Return
            End If
        End If
        ' 6. СТАРТИРАЙ ЧЕРТАНЕТО В ТРАНЗАКЦИЯ
        Using trans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                ' ПРЕДИЗЧИСЛЯВАНЕ НА ПАРАМЕТРИТЕ
                ' Тук ще извикваме процедурите за чертане една по една
                DrawPanelFrame(acDoc, acCurDb, ptBasePoint, panelCircuits, selectedTablo)   ' Тук чертаем рамката на таблото
                DrawBusbars(acDoc, acCurDb, ptBasePoint, panelCircuits)                     ' Тук чертаем шините
                DrawCircuits(acDoc, acCurDb, ptBasePoint, panelCircuits)                    ' Тук чертаем всеки токов кръг (прекъсвачи, текстове, линии)
                DrawRCDBusbar(acDoc, acCurDb, ptBasePoint, panelCircuits)                   ' Тук чертаем ДЗТ за токовите кръгове (прекъсвачи, текстове, линии)


                DrawGrounding(acDoc, acCurDb, ptBasePoint.X, ptBasePoint, selectedTablo)   ' Чертaем заземление само за главно разпределително табло
                DrawAnnotations(ptBasePoint, panelCircuits)                                ' Процедурата създава текстови анотации
            Catch ex As Exception
                trans.Abort()
                MsgBox("Възникна грешка при чертане: " & vbCrLf & ex.Message & vbCrLf & vbCrLf & ex.StackTrace, MsgBoxStyle.Critical)
            Finally
                trans.Commit()
            End Try
        End Using
    End Sub
    ''' <summary>
    ''' Процедурата DrawBusbars отговаря за изчертаването на шините (busbars)
    ''' в електрическо табло в AutoCAD.
    ''' 
    ''' В зависимост от конфигурацията (една или две шини), процедурата:
    ''' - изчислява геометрията (позиции и дължини)
    ''' - определя фазите
    ''' - изчертава линии за шините
    ''' - добавя текстови надписи за фазите
    ''' - при две шини – визуализира връзката между тях
    ''' 
    ''' Използва данни от списък с токови кръгове (strTokow).
    ''' </summary>
    ''' <param name="acDoc">AutoCAD документ (не се използва директно, но е част от контекста).</param>
    ''' <param name="acCurDb">AutoCAD база данни (не се използва директно тук).</param>
    ''' <param name="basePoint">Базова точка за позициониране на всички елементи.</param>
    ''' <param name="circuits">Списък от токови кръгове, използван за изчисления и логика.</param>
    Private Sub DrawBusbars(acDoc As Document, acCurDb As Database, basePoint As Point3d, circuits As List(Of Form_Tablo_new.strTokow))
        Try
            ' 1️ ИЗЧИСЛЯВАНЕ НА ОСНОВНИТЕ РАЗМЕРИ
            Dim brColums As Integer = circuits.Count - 1
            Dim X_Start As Double = basePoint.X + widthText + widthTextDim
            Dim X_End As Double = basePoint.X + widthText + widthTextDim + brColums * widthColom + widthColom / 2
            Dim Y_Shina As Double = basePoint.Y + Y_Шина
            ' Брой токови кръгове, които принадлежат към първата шина.
            Dim brTokKrygoweNa6ina = circuits.Where(Function(c) c.Шина = True).Count()
            Dim Faza_Първа_шина = circuits.Any(Function(c) c.Фаза = "L1" Or c.Фаза = "L2" Or c.Фаза = "L3")
            Dim circuitOBSTO = circuits.FirstOrDefault(Function(c) c.ТоковКръг = "ОБЩО")
            Dim Faza_Втора_шина = circuitOBSTO.Фаза
            ' 3️ ТЕКСТ ЗА ФАЗИТЕ НА ШИНАТА
            Dim phaseText As String = IIf(Faza_Първа_шина, "L1,L2,L3,N,PE", "L,N,PE")
            ' 4️ ЧЕРТАЕНЕ НА ШИНИТЕ
            Dim X_Split As Double = 0
            Dim X_SecondStart As Double = 0
            Dim X_SecondEnd As Double = 0
            If Not twoBus Then
                ' ----- ЕДНА ШИНА -----
                cu.DrowLine(New Point3d(X_Start, Y_Shina, 0),
                        New Point3d(X_End, Y_Shina, 0),
                        "EL_ТАБЛА", Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight070, "ByLayer")
                cu.InsertText(Faza_Втора_шина & ",N,PE",
                          New Point3d(X_Start, Y_Shina + 2 * padingText, 0),
                          "EL__DIM", heightText,
                          TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
                X_SecondStart = X_Start
                X_SecondEnd = X_End
            Else
                ' ----- ДВЕ ШИНИ -----
                X_Split = X_Start + brTokKrygoweNa6ina * widthColom - widthColom / 2
                ' Чертае първата (лява) шина.
                cu.DrowLine(New Point3d(X_Start, Y_Shina, 0),
                        New Point3d(X_Split, Y_Shina, 0),
                        "EL_ТАБЛА", Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight070, "ByLayer")
                ' Начална позиция на втората (дясна) шина.
                X_SecondStart = X_Start + brTokKrygoweNa6ina * widthColom + widthColom / 2
                ' Край на втората шина (същият като X_End).
                X_SecondEnd = basePoint.X + widthText + widthTextDim + (brColums - 1) * widthColom + widthColom / 2
                ' Чертае втората шина.
                cu.DrowLine(New Point3d(X_SecondStart, Y_Shina, 0),
                        New Point3d(X_SecondEnd, Y_Shina, 0),
                        "EL_ТАБЛА", Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight070, "ByLayer")
                ' Надпис за първата шина (зависи от наличието на трифазни товари).
                cu.InsertText(phaseText,
                          New Point3d(X_Start, Y_Shina + 2 * padingText, 0),
                          "EL__DIM", heightText,
                          TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
                ' Надпис за втората шина (взет от "ОБЩО").
                cu.InsertText(Faza_Втора_шина & ",N,PE",
                          New Point3d(X_SecondStart, Y_Shina + 2 * padingText, 0),
                          "EL__DIM", heightText,
                          TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
                ' Чертае връзка между двете шини (между разединители).
                ' Позиционирана е над основните шини (+95 по Y).
                Dim X_6ina1 As Double = (X_SecondStart + X_End) / 2
                Dim X_6ina2 As Double = (X_Start + X_Split) / 2
                cu.DrowLine(New Point3d(X_6ina1, Y_Shina + 95, 0),
                        New Point3d(X_6ina2, Y_Shina + 95, 0),
                        "EL_ТАБЛА", Autodesk.AutoCAD.DatabaseServices.LineWeight.ByLayer, "ByLayer")
                ' Вмъква разединител в средата на шината и попълва атрибутите
                Dim circuit = circuits.FirstOrDefault(Function(c) c.Device = "Разединител")
                If circuit Is Nothing Then Return

                Dim X_disconn As Double = (X_Start + X_Split) / 2
                Dim Y_disconn As Double = Y_Shina + 95
                Dim blkRecId As ObjectId = cu.InsertBlock("s_i_ng_switch_disconn",
                                               New Point3d(X_disconn, Y_disconn, 0),
                                               "EL_ТАБЛА",
                                               New Scale3d(5, 5, 5))
                If Not blkRecId.IsNull Then
                    Using trans As Transaction = acCurDb.TransactionManager.StartTransaction()
                        Dim acBlkRef As BlockReference = DirectCast(trans.GetObject(blkRecId, OpenMode.ForWrite), BlockReference)
                        For Each objID As ObjectId In acBlkRef.AttributeCollection
                            Dim acAttRef As AttributeReference = DirectCast(trans.GetObject(objID, OpenMode.ForWrite), AttributeReference)
                            Select Case acAttRef.Tag
                                Case "1" : acAttRef.TextString = ""
                                Case "2" : acAttRef.TextString = circuit.Брой_Полюси & "p"
                                Case "3" : acAttRef.TextString = circuit.Breaker_Номинален_Ток & "A"
                                Case "4" : acAttRef.TextString = ""
                                Case "5" : acAttRef.TextString = ""
                                Case "SHORTNAME" : acAttRef.TextString = circuit.Breaker_Тип_Апарат
                                Case "REFNB" : acAttRef.TextString = circuit.Tablo
                                Case "DESIGNATION" : acAttRef.TextString = ""
                            End Select
                        Next
                        trans.Commit()
                    End Using
                End If
            End If
            Dim circuit_Общо = circuits.FirstOrDefault(Function(c) c.Device = "Табло")
            Dim X_Общо As Double = (X_SecondStart + X_SecondEnd) / 2
            Dim Y_Общо As Double = Y_Shina + 95
            Dim blkRecId_Общо As ObjectId = cu.InsertBlock("s_i_ng_switch_disconn",
                               New Point3d(X_Общо, Y_Общо, 0),
                               "EL_ТАБЛА",
                               New Scale3d(5, 5, 5))
            If Not blkRecId_Общо.IsNull Then
                Using trans As Transaction = acCurDb.TransactionManager.StartTransaction()
                    Dim acBlkRef As BlockReference = DirectCast(trans.GetObject(blkRecId_Общо, OpenMode.ForWrite), BlockReference)
                    For Each objID As ObjectId In acBlkRef.AttributeCollection
                        Dim acAttRef As AttributeReference = DirectCast(trans.GetObject(objID, OpenMode.ForWrite), AttributeReference)
                        Select Case acAttRef.Tag
                            Case "1" : acAttRef.TextString = ""
                            Case "2" : acAttRef.TextString = circuit_Общо.Брой_Полюси & "p"
                            Case "3" : acAttRef.TextString = circuit_Общо.Breaker_Номинален_Ток & "A"
                            Case "4" : acAttRef.TextString = ""
                            Case "5" : acAttRef.TextString = ""
                            Case "SHORTNAME" : acAttRef.TextString = circuit_Общо.Breaker_Тип_Апарат
                            Case "REFNB" : acAttRef.TextString = circuit_Общо.Tablo
                            Case "DESIGNATION" : acAttRef.TextString = ""
                        End Select
                    Next
                    trans.Commit()
                End Using
                ' Чертае вертикална линия над прекъсвача.
                cu.DrowLine(New Point3d(X_Общо, Y_Общо, 0),
                        New Point3d(X_Общо, Y_Общо + 125, 0),
                        "EL_ТАБЛА", Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight070, "ByLayer")

                Dim blkRecId_Текст As ObjectId = cu.InsertBlock("Кабел",
                                   New Point3d(X_Общо, Y_Общо + 90, 0),
                                   "EL__DIM",
                                   New Scale3d(1, 1, 1))
                Using trans As Transaction = acCurDb.TransactionManager.StartTransaction()
                    Dim acBlkRef As BlockReference = DirectCast(trans.GetObject(blkRecId_Текст, OpenMode.ForWrite), BlockReference)
                    For Each objID As ObjectId In acBlkRef.AttributeCollection
                        Dim acAttRef As AttributeReference = DirectCast(trans.GetObject(objID, OpenMode.ForWrite), AttributeReference)
                        Select Case acAttRef.Tag
                            Case "NA4IN_0" : acAttRef.TextString = circuit_Общо.Кабел_Сечение
                            Case "NA4IN_1" : acAttRef.TextString = "от табло " & circuit_Общо.Табло_Родител
                            Case "NA4IN_2" : acAttRef.TextString = ""
                            Case "NA4IN_3" : acAttRef.TextString = ""
                            Case "NA4IN_4" : acAttRef.TextString = ""
                            Case "NA4IN_5" : acAttRef.TextString = ""
                            Case "NA4IN_6" : acAttRef.TextString = ""
                            Case "NA4IN_7" : acAttRef.TextString = ""
                            Case "NA4IN_8" : acAttRef.TextString = ""
                            Case "NA4IN_9" : acAttRef.TextString = ""
                            Case "NA4IN_10" : acAttRef.TextString = ""
                        End Select
                    Next
                    Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection
                    For Each prop As DynamicBlockReferenceProperty In props
                        If prop.PropertyName = "Visibility" Then prop.Value = "Точка"
                    Next
                    trans.Commit()
                End Using
                cu.EditDynamicBlockReferenceKabel(blkRecId_Текст)
            End If
        Catch ex As Exception
            ' Показва съобщение при възникване на грешка,
            ' включително текста на грешката и stack trace.
            ' Полезно при дебъг, но не е подходящо за production среда.
            MsgBox("Възникна грешка:  " &
               ex.Message &
               vbCrLf & vbCrLf &
               ex.StackTrace.ToString)
        End Try
    End Sub
    ''' <summary>
    ''' Изчертава рамката на електрическо табло в AutoCAD.
    ''' Включва позициониране спрямо базова точка и използва данни от подадените токови кръгове.
    ''' </summary>
    ''' <param name="acDoc">Активният AutoCAD документ.</param>
    ''' <param name="acCurDb">Текущата база данни на чертежа.</param>
    ''' <param name="basePoint">Базова точка за позициониране на рамката.</param>
    ''' <param name="circuits">Списък с токови кръгове, използван за определяне на размерите и съдържанието.</param>
    ''' <param name="selectedTablo">Име на таблото, за което се чертае рамката.</param>
    ''' <remarks>
    ''' Процедурата изгражда графичната рамка на таблото, като използва
    ''' геометрични зависимости и данни от токовите кръгове.
    ''' Използва помощни функции за чертане на линии и текст.
    ''' </remarks>
    Private Sub DrawPanelFrame(acDoc As Document, acCurDb As Database, basePoint As Point3d,
                               circuits As List(Of Form_Tablo_new.strTokow), selectedTablo As String)
        Try
            ' =====================================================
            ' 1️ ИЗЧИСЛЯВАНЕ НА ОСНОВНИТЕ РАЗМЕРИ
            ' =====================================================
            Dim brColums As Integer = circuits.Count - If(twoBus, 1, 0)
            Dim tableWidth As Double = basePoint.X + widthText + widthTextDim + (brColums) * widthColom
            Dim tableHeight As Double = 10 * heightRow
            ' =====================================================
            ' 2️ СЪЗДАВАНЕ НА СПИСЪК С ЛИНИИТЕ
            ' =====================================================
            Dim lines As New List(Of LineDefinition)
            ' --- Хоризонтални линии на таблицата ---
            ' Долна линия (ред 0)
            AddLine(lines, New Point3d(basePoint.X, basePoint.Y, 0),
                       New Point3d(tableWidth, basePoint.Y, 0))
            ' Хоризонтални линии за редове 3-10
            For row As Integer = 3 To 10
                AddLine(lines, New Point3d(basePoint.X, basePoint.Y + row * heightRow, 0),
                           New Point3d(tableWidth, basePoint.Y + row * heightRow, 0))
            Next
            ' --- Вертикални линии на таблицата ---
            ' Ляв край
            AddLine(lines, New Point3d(basePoint.X, basePoint.Y, 0),
                       New Point3d(basePoint.X, basePoint.Y + tableHeight, 0))
            ' След "Токов кръг"
            AddLine(lines, New Point3d(basePoint.X + widthText, basePoint.Y, 0),
                       New Point3d(basePoint.X + widthText, basePoint.Y + tableHeight, 0))
            ' След "№"
            AddLine(lines, New Point3d(basePoint.X + widthText + widthTextDim, basePoint.Y, 0),
                       New Point3d(basePoint.X + widthText + widthTextDim, basePoint.Y + tableHeight, 0))
            ' Вертикални линии за всеки токов кръг
            For col As Integer = 1 To brColums
                Dim xLine As Double = basePoint.X + widthText + widthTextDim + col * widthColom
                AddLine(lines, New Point3d(xLine, basePoint.Y, 0),
                               New Point3d(xLine, basePoint.Y + tableHeight, 0))
            Next
            ' --- Рамка на блока с информация за шината ---
            Dim blockStartY As Double = basePoint.Y + tableHeight + lengthProw
            Dim blockEndY As Double = blockStartY + widthTablo
            ' Лява страна (CENTER тип)
            AddLine(lines, New Point3d(basePoint.X + widthText, blockStartY, 0),
                       New Point3d(basePoint.X + widthText, blockEndY, 0),
                       "EL_ТАБЛА", Autodesk.AutoCAD.DatabaseServices.LineWeight.ByLayer, "CENTER")
            ' Долна страна
            AddLine(lines, New Point3d(basePoint.X + widthText, blockStartY, 0),
                       New Point3d(tableWidth, blockStartY, 0),
                       "EL_ТАБЛА", Autodesk.AutoCAD.DatabaseServices.LineWeight.ByLayer, "CENTER")
            ' Дясна страна
            AddLine(lines, New Point3d(tableWidth, blockStartY, 0),
                       New Point3d(tableWidth, blockEndY, 0),
                       "EL_ТАБЛА", Autodesk.AutoCAD.DatabaseServices.LineWeight.ByLayer, "CENTER")
            ' Горна страна
            AddLine(lines, New Point3d(basePoint.X + widthText, blockEndY, 0),
                       New Point3d(tableWidth, blockEndY, 0),
                       "EL_ТАБЛА", Autodesk.AutoCAD.DatabaseServices.LineWeight.ByLayer, "CENTER")
            ' --- Червен кръст за маркировка (Defpoints) ---
            Dim crossCenterX As Double = basePoint.X + widthText + 18
            Dim crossCenterY As Double = blockEndY - 18
            ' Вертикална линия на кръста
            AddLine(lines, New Point3d(crossCenterX, blockEndY, 0),
                           New Point3d(crossCenterX, blockEndY - 36, 0),
                           "Defpoints", Autodesk.AutoCAD.DatabaseServices.LineWeight.ByLayer, "ByLayer", 1)
            ' Хоризонтална линия на кръста
            AddLine(lines, New Point3d(basePoint.X + widthText, crossCenterY, 0),
                       New Point3d(basePoint.X + widthText + 36, crossCenterY, 0),
                       "Defpoints", Autodesk.AutoCAD.DatabaseServices.LineWeight.ByLayer, "ByLayer", 1)
            ' =====================================================
            ' 3️ ЧЕРТАЕНЕ НА ВСИЧКИ ЛИНИИ ОТ СПИСЪКА
            ' =====================================================
            For Each line As LineDefinition In lines
                If line.ColorIndex = -1 Then
                    cu.DrowLine(line.StartPoint, line.EndPoint, line.Layer, line.LineWeightValue, line.LineType)
                Else
                    cu.DrowLine(line.StartPoint, line.EndPoint, line.Layer, line.LineWeightValue, line.LineType, line.ColorIndex)
                End If
            Next
            ' =====================================================
            ' 4️ ТЕКСТОВЕ - ПЪРВА КОЛОНА (ЗАГЛАВКИ)
            ' =====================================================
            Dim textX As Double = basePoint.X + padingText
            Dim textY As Double = basePoint.Y + (heightRow - heightText) / 2
            cu.InsertText("Токов кръг", New Point3d(textX, textY + 9 * heightRow, 0),
                      "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("Брой лампи", New Point3d(textX, textY + 8 * heightRow, 0),
                      "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("Брой контакти", New Point3d(textX, textY + 7 * heightRow, 0),
                      "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("Инстал. мощност", New Point3d(textX, textY + 6 * heightRow, 0),
                      "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("Тип кабел", New Point3d(textX, textY + 5 * heightRow, 0),
                      "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("Сечение кабел", New Point3d(textX, textY + 4 * heightRow, 0),
                      "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("Фаза", New Point3d(textX, textY + 3 * heightRow, 0),
                      "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("Консуматор", New Point3d(textX, textY + 2 * heightRow, 0),
                      "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            ' =====================================================
            ' 5️ ТЕКСТОВЕ - ВТОРА КОЛОНА (МЕРНИ ЕДИНИЦИ)
            ' =====================================================
            textX = textX + widthText
            cu.InsertText("№", New Point3d(textX, textY + 9 * heightRow, 0),
                      "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("бр.", New Point3d(textX, textY + 8 * heightRow, 0),
                      "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("бр.", New Point3d(textX, textY + 7 * heightRow, 0),
                      "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("kW", New Point3d(textX, textY + 6 * heightRow, 0),
                      "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("---", New Point3d(textX, textY + 5 * heightRow, 0),
                      "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("mm²", New Point3d(textX, textY + 4 * heightRow, 0),
                      "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("---", New Point3d(textX, textY + 3 * heightRow, 0),
                      "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
            cu.InsertText("---", New Point3d(textX, textY + 2 * heightRow, 0),
                      "EL__DIM", heightText, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)

            Dim X = basePoint.X + widthText + widthTextDim
            cu.InsertText(selectedTablo,
                          New Point3d(X + (brColums - 1) * widthColom,
                                      basePoint.Y + Y_Шина + 95,
                                      0),
                          "EL__DIM", heightText + 5, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
        Catch ex As Exception
            ' --------------------------------------------------------
            ' Обработка на грешки
            ' --------------------------------------------------------
            MsgBox("Възникна грешка: " &
                   ex.Message &
                   vbCrLf & vbCrLf &
                   ex.StackTrace.ToString)
        End Try
    End Sub
    ''' <summary>
    ''' Добавя нова линия към списъка от линии, които ще бъдат изчертани по-късно.
    ''' </summary>
    ''' <param name="lines">Списък, в който се съхраняват линиите за последващо чертане.</param>
    ''' <param name="startPoint">Начална точка на линията.</param>
    ''' <param name="endPoint">Крайна точка на линията.</param>
    ''' <param name="layer">Слой, на който ще бъде поставена линията. По подразбиране "EL_ТАБЛА".</param>
    ''' <param name="lineWeight">Дебелина на линията (AutoCAD LineWeight).</param>
    ''' <param name="lineType">Тип линия (ByLayer, CENTER и др.).</param>
    ''' <param name="colorIndex">Индекс на цвят. Ако е -1, се използва ByLayer.</param>
    ''' <remarks>
    ''' Помощна функция за централизирано създаване на LineDefinition обекти.
    ''' Улеснява управлението и стандартизацията на линиите преди реалното им чертане.
    ''' </remarks>
    Private Sub AddLine(lines As List(Of LineDefinition),
                        startPoint As Point3d,
                        endPoint As Point3d,
                        Optional layer As String = "EL_ТАБЛА",
                        Optional lineWeight As Autodesk.AutoCAD.DatabaseServices.LineWeight = Autodesk.AutoCAD.DatabaseServices.LineWeight.ByLayer,
                        Optional lineType As String = "ByLayer",
                        Optional colorIndex As Integer = -1)

        lines.Add(New LineDefinition(startPoint, endPoint, layer, lineWeight, lineType, colorIndex))

    End Sub
    ''' <summary>
    ''' Чертaе всички токови кръгове в таблото.
    ''' Логиката:
    ''' 1. Изчислява началните координати
    ''' 2. Обхожда всички токови кръгове
    ''' 3. Изчертава текстовете за всеки кръг
    ''' 4. Вмъква прекъсвач и управляващи елементи
    ''' 5. Чертaе свързващите линии
    ''' </summary>
    Private Sub DrawCircuits(acDoc As Document, acCurDb As Database, basePoint As Point3d, circuits As List(Of Form_Tablo_new.strTokow))
        ' Изчислява общия брой колони.
        ' Ако има двойна шина (twoBus=True), една колона се резервира и не участва.
        Dim brColums As Integer = circuits.Count - If(twoBus, 1, 0)
        ' Начална X координата след текстовата зона
        Dim X_Start As Double = basePoint.X + widthText + widthTextDim
        ' Y координата на шината
        Dim Y_Shina As Double = basePoint.Y + Y_Шина
        Try
            ' Индекс на текущата колона
            Dim colIndex As Integer = 0
            ' Обхождаме всички токови кръгове
            For Each circuit As Form_Tablo_new.strTokow In circuits
                ' Пропускаме специалните кръгове тип "Разединител"
                If circuit.Device = "Разединител" Then Continue For
                ' Изчисляване на X позицията за текущия кръг
                Dim X As Double =
                X_Start + colIndex * widthColom + widthColom / 2
                ' Чертaе текстовата информация за токовия кръг
                DrawCircuitTexts(acDoc, acCurDb, basePoint, circuit, X)
                ' Ако е запис "Табло" → пишем само текстовете
                ' Не се чертаят прекъсвачи и линии
                If circuit.Device = "Табло" Then Continue For
                ' Вмъква блок за прекъсвач
                DrawBreakerBlock(acDoc, acCurDb, basePoint, circuit, X, Y_Shina)
                ' Чертaе управляващо устройство (ако има)
                DrawControlDevice(acDoc, acCurDb, circuit, X, Y_Shina)
                ' Чертaе вертикалните линии за кръга
                DrawCircuitLines(X, circuit, Y_Shina)
                ' Преминава към следващата колона
                colIndex += 1
            Next
        Catch ex As Exception
            MsgBox("Възникна грешка: " &
               vbCrLf &
               ex.Message &
               vbCrLf &
               vbCrLf &
               ex.StackTrace,
               MsgBoxStyle.Critical)
        End Try
    End Sub
    ''' <summary>
    ''' Чертaе текстовата информация за един токов кръг в таблицата на таблото.
    ''' </summary>
    ''' <param name="acDoc">Текущият AutoCAD документ.</param>
    ''' <param name="acCurDb">Текущата база данни.</param>
    ''' <param name="basePoint">Начална точка на таблицата.</param>
    ''' <param name="circuit">Обект с данни за токовия кръг.</param>
    ''' <param name="X">X координата на колоната за съответния кръг.</param>
    ''' <remarks>
    ''' Процедурата позиционира и изчертава всички текстове за даден токов кръг
    ''' в съответната колона на таблицата.
    '''
    ''' Особености:
    ''' - Използва центрирано подравняване за числови и кратки стойности
    ''' - Използва ляво подравняване за текстови описания
    ''' - При нулеви стойности (лампи/контакти) показва "----"
    ''' - Всички координати се изчисляват спрямо basePoint
    '''
    ''' Възможни подобрения:
    ''' - Проверка за празни/null стойности (Nothing)
    ''' - Форматиране на текста спрямо дължината (truncate/auto-scale)
    ''' - Унифициране на височината на текста (в момента има 12 и heightText)
    ''' </remarks>
    Private Sub DrawCircuitTexts(acDoc As Document, acCurDb As Database, basePoint As Point3d,
                         circuit As Form_Tablo_new.strTokow, X As Double)
        Dim Y_Base As Double = basePoint.Y
        Dim textLayer As String = "EL__DIM"
        ' Токов кръг (ред 1)
        cu.InsertText(circuit.ТоковКръг,
                  New Point3d(X + padingText, Y_Base + 9 * heightRow + heightRow / 2, 0),
                  textLayer, heightText, TextHorizontalMode.TextMid, TextVerticalMode.TextBase)
        ' Брой лампи (ред 2)
        cu.InsertText(IIf(circuit.brLamp = 0, "----", circuit.brLamp.ToString()),
                  New Point3d(X + padingText, Y_Base + 8 * heightRow + heightRow / 2, 0),
                  textLayer, heightText, TextHorizontalMode.TextMid, TextVerticalMode.TextBase)
        ' Брой контакти (ред 3)
        cu.InsertText(IIf(circuit.brKontakt = 0, "----", circuit.brKontakt.ToString()),
                  New Point3d(X + padingText, Y_Base + 7 * heightRow + heightRow / 2, 0),
                  textLayer, heightText, TextHorizontalMode.TextMid, TextVerticalMode.TextBase)
        ' Мощност (ред 4)
        cu.InsertText(circuit.Мощност.ToString("0.000"),
                  New Point3d(X + padingText, Y_Base + 6 * heightRow + heightRow / 2, 0),
                  textLayer, heightText, TextHorizontalMode.TextMid, TextVerticalMode.TextBase)
        ' Тип кабел (ред 5)
        cu.InsertText(circuit.Кабел_Тип,
                  New Point3d(X + padingText, Y_Base + 5 * heightRow + heightRow / 2, 0),
                  textLayer, heightText, TextHorizontalMode.TextMid, TextVerticalMode.TextBase)
        ' Сечение кабел (ред 6)
        cu.InsertText(circuit.Кабел_Сечение,
                  New Point3d(X + padingText, Y_Base + 4 * heightRow + heightRow / 2, 0),
                  textLayer, heightText, TextHorizontalMode.TextMid, TextVerticalMode.TextBase)
        ' Фаза (ред 7)
        cu.InsertText(circuit.Фаза,
                  New Point3d(X + padingText, Y_Base + 3 * heightRow + heightRow / 2, 0),
                  textLayer, heightText, TextHorizontalMode.TextMid, TextVerticalMode.TextBase)
        ' Консуматор (ред 8) - ляво подравнен
        cu.InsertText(circuit.Консуматор,
                  New Point3d(X - widthColom / 2 + padingText, Y_Base + 2 * heightRow + (heightRow - heightText) / 2, 0),
                  textLayer, 12, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
        ' Предназначение (ред 9) - ляво подравнен
        cu.InsertText(circuit.предназначение,
                  New Point3d(X - widthColom / 2 + padingText, Y_Base + 1 * heightRow + (heightRow - heightText) / 2, 0),
                  textLayer, 12, TextHorizontalMode.TextLeft, TextVerticalMode.TextBase)
    End Sub
    ''' <summary>
    ''' Процедурата DrawBreakerBlock добавя визуално представяне (блок) на автоматичен прекъсвач или RCD в AutoCAD документ.
    ''' Използва се в контекста на проектиране на електрически табла и схеми на токови кръгове.
    ''' </summary>
    ''' <param name="acDoc">Документът на AutoCAD, в който се добавя блокът.</param>
    ''' <param name="acCurDb">Текущата база данни на AutoCAD, за извършване на транзакции и операции върху блокове.</param>
    ''' <param name="basePoint">Базова точка (не се използва директно в тази процедура, но може да е за разширения).</param>
    ''' <param name="circuit">Обект от тип strTokow, който съдържа данните за конкретния токов кръг, като тип апарат, RCD информация, брой полюси и др.</param>
    ''' <param name="X">X координата за позициониране на блока.</param>
    ''' <param name="Y_Shina">Y координата на шината, върху която се поставя блокът.</param>
    Private Sub DrawBreakerBlock(acDoc As Document, acCurDb As Database, basePoint As Point3d,
                                circuit As Form_Tablo_new.strTokow, X As Double, Y_Shina As Double)
        ' Име на блока по подразбиране – стандартен прекъсвач C60
        Dim blockName As String = "s_c60_circ_break"
        ' Масштаб на блока – зададен като 5х5х5.
        Dim blockScale As New Scale3d(5, 5, 5)
        ' Начална позиция за поставяне на блока
        Dim insertPoint As New Point3d(X, Y_Shina, 0)
        ' Флаг, който указва какво представлява блокът - за специално попълване на атрибути
        Dim rcd_Yes As String = ""
        ' Ако има RCD_Нула и тя не е "N", местим блока надолу по Y с 117.5 единици
        If Not String.IsNullOrEmpty(circuit.RCD_Нула) AndAlso
                    circuit.RCD_Нула.Trim().ToUpper() <> "N" Then
            insertPoint = New Point3d(X, Y_Shina - 117.5, 0)
        End If
        ' Избор на блок според типа апарат
        ' 1. Първоначални настройки
        rcd_Yes = "Прекъсвач"
        blockName = "s_c60_circ_break"
        ' 2. Логика за избор на тип апарат и блок
        Select Case True
    ' Първи приоритет: Резерви
            Case circuit.Device = "Резерва"
                rcd_Yes = "Резерва"
                blockName = "s_c60_circ_break"
    ' Първи приоритет: Резерви
            Case circuit.Device = "Съществуващ"
                rcd_Yes = "Съществуващ"
                blockName = "s_c60_circ_break"
    ' Втори приоритет: Моторна защита
            Case circuit.Управление = "Моторна защита"
                rcd_Yes = "Моторна защита"
                blockName = "s_GV2"
    ' Трети приоритет: Проверка за RCD (ако има попълнен тип)
            Case Not String.IsNullOrWhiteSpace(circuit.RCD_Тип)
                rcd_Yes = "RCD"
                blockName = "s_dpnn_vigi_circ_break"
                ' Всичко останало (Default)
            Case Else
                rcd_Yes = "Прекъсвач"
                blockName = "s_c60_circ_break"
        End Select
        ' Вмъкване на блока в AutoCAD с помощта на функция InsertBlock
        ' (предполага се, че cu е помощен модул/клас за CAD операции)
        Dim blkRecId As ObjectId = cu.InsertBlock(blockName, insertPoint, "EL_ТАБЛА", blockScale)
        ' Ако вмъкването е успешно (ObjectId не е Null)
        If Not blkRecId.IsNull Then
            ' Стартиране на транзакция за промяна на атрибутите на блока
            Try
                Using trans As Transaction = acCurDb.TransactionManager.StartTransaction()
                    ' Получаваме референция към блока за писане
                    Dim acBlkRef As BlockReference = DirectCast(trans.GetObject(blkRecId, OpenMode.ForWrite), BlockReference)
                    ' Обхождане на всички атрибути на блока
                    For Each objID As ObjectId In acBlkRef.AttributeCollection
                        Dim acAttRef As AttributeReference = DirectCast(trans.GetObject(objID, OpenMode.ForWrite), AttributeReference)
                        ' Попълване на атрибутите в зависимост дали е RCD или обикновен прекъсвач
                        If acAttRef.Tag = "DESIGNATION" Then acAttRef.TextString = ""
                        If acAttRef.Tag = "REFNB" Then acAttRef.TextString = circuit.Tablo
                        Select Case rcd_Yes
                            Case "Съществуващ"
                                Select Case acAttRef.Tag
                                    Case "1" : acAttRef.TextString = ""
                                    Case "2" : acAttRef.TextString = ""
                                    Case "3" : acAttRef.TextString = ""
                                    Case "4" : acAttRef.TextString = circuit.Breaker_Номинален_Ток
                                    Case "5" : acAttRef.TextString = ""
                                    Case "SHORTNAME" : acAttRef.TextString = ""
                                End Select
                            Case "Резерва"
                                Select Case acAttRef.Tag
                                    Case "1" : acAttRef.TextString = ""
                                    Case "2" : acAttRef.TextString = ""
                                    Case "3" : acAttRef.TextString = ""
                                    Case "4" : acAttRef.TextString = circuit.Breaker_Номинален_Ток
                                    Case "5" : acAttRef.TextString = ""
                                    Case "SHORTNAME" : acAttRef.TextString = ""
                                End Select
                            Case "Моторна защита"
                                Select Case acAttRef.Tag
                                    Case "1" : acAttRef.TextString = Calculate_GV2(circuit.Ток, 3)
                                    Case "2" : acAttRef.TextString = "3P"
                                    Case "3" : acAttRef.TextString = Calculate_GV2(circuit.Ток, 2)
                                    Case "4" : acAttRef.TextString = ""
                                    Case "5" : acAttRef.TextString = ""
                                    Case "SHORTNAME" : acAttRef.TextString = Calculate_GV2(circuit.Ток, 1)
                                End Select
                            Case "RCD"
                                ' Атрибути за RCD
                                Select Case acAttRef.Tag
                                    Case "SHORTNAME" : acAttRef.TextString = circuit.RCD_Тип
                                    Case "1" : acAttRef.TextString = circuit.RCD_Клас
                                    Case "2" : acAttRef.TextString = circuit.Брой_Полюси & "p"
                                    Case "3" : acAttRef.TextString = "C"
                                    Case "4" : acAttRef.TextString = circuit.RCD_Ток & "A"
                                    Case "5" : acAttRef.TextString = circuit.RCD_Чувствителност & "mA"
                                End Select
                            Case "Прекъсвач"
                                ' Атрибути за прекъсвач
                                Select Case acAttRef.Tag
                                    Case "SHORTNAME" : acAttRef.TextString = circuit.Breaker_Тип_Апарат
                                    Case "2" : acAttRef.TextString = circuit.Breaker_Крива
                                    Case "3" : acAttRef.TextString = circuit.Брой_Полюси & "p"
                                    Case "4" : acAttRef.TextString = circuit.Breaker_Номинален_Ток & "A"
                                End Select
                        End Select
                    Next
                    ' Потвърждаваме промяната на атрибутите
                    trans.Commit()
                End Using
            Catch ex As Exception

            End Try
        End If
    End Sub
    ''' <summary>
    ''' Функцията Calculate_GV2 избира подходящ моторен прекъсвач (тип GV2)
    ''' на база вече изчислен ток.
    ''' 
    ''' Логиката включва:
    ''' 1. Преобразуване на входния ток от текст към число
    ''' 2. Търсене на съвпадение в база данни (GV_Database)
    ''' 3. Връщане на конкретна информация според параметъра "Връща"
    ''' 
    ''' Функцията НЕ изчислява ток – очаква той да е подаден отвън.
    ''' Това я прави по-гъвкава и независима от начина на изчисление.
    ''' </summary>
    ''' <param name="Ток">
    ''' Ток като текст (например "10", "10.5", "10,5").
    ''' Допуска се използване на запетая или точка като десетичен разделител.
    ''' </param>
    ''' <param name="Връща">
    ''' Определя какъв резултат да бъде върнат:
    ''' 1 → Тип на защитата (например GV2-ME)
    ''' 2 → Мощност по каталог (при 400V)
    ''' 3 → Диапазон на настройка
    ''' </param>
    ''' <returns>
    ''' Връща String със съответния резултат или съобщение:
    ''' - "N/A" при невалиден ток
    ''' - "Out of range (...A)" ако няма подходящ апарат
    ''' - "Грешен параметър" при невалиден вход за "Връща"
    ''' </returns>
    Private Function Calculate_GV2(Ток As String, Връща As Integer) As String
        ' =====================================================
        ' 1️ ПРЕОБРАЗУВАНЕ НА ВХОДНИЯ ТОК
        ' =====================================================
        ' Замяна на запетая с точка, за да се осигури коректно
        ' преобразуване към числов тип.
        Dim I_val As String = Ток.Replace(",", ".")
        ' Преобразуване на текстовата стойност към Double.
        ' Val извлича числото от началото на низа.
        Dim I_double As Double = Val(I_val)
        ' Проверка за невалиден или нулев ток.
        If I_double <= 0 Then Return "N/A"
        ' =====================================================
        ' 2️ ТЪРСЕНЕ В БАЗАТА ДАННИ
        ' =====================================================
        ' Търсене на първия запис в GV_Database,
        ' при който токът попада в диапазона:
        ' MinCurrent ≤ I_double ≤ MaxCurrent
        Dim match = gvDatabase.FirstOrDefault(Function(x) I_double >= x.MinCurrent And I_double <= x.MaxCurrent)
        ' Ако няма намерен подходящ апарат,
        ' връщаме информация за тока.
        If match Is Nothing Then Return "Out of range (" & I_double.ToString("F2") & "A)"
        ' =====================================================
        ' 3️ ВРЪЩАНЕ НА РЕЗУЛТАТ
        ' =====================================================
        ' В зависимост от параметъра "Връща",
        Select Case Връща
            Case 1 : Return match.Type               ' Връща типа на апарата (например GV2-ME).
            Case 2 : Return match.MotorPower         ' Връща мощността по каталог (при 400V).
            Case 3 : Return match.SettingRange       ' Връща диапазона на настройка на тока.
            Case Else : Return "Грешен параметър"    ' Невалиден параметър "Връща".
        End Select
    End Function
    ''' <summary>
    ''' Връща конфигурацията за даден тип управление
    ''' </summary>
    Private Function GetControlDeviceConfig(circuit As Form_Tablo_new.strTokow) As ControlDeviceConfig
        ' New(str_1 , str_2 , str_3 , str_4 , str_5 , shortName)
        Select Case circuit.Управление
            Case "Импулсно реле"
                Return New ControlDeviceConfig("", "1p", If(circuit.Ток * 1.1 > 16, "32A", "16A"), "220VAC", "", "iTL")
            Case "Контактор"
                Return New ControlDeviceConfig(circuit.Брой_Полюси.ToString & "НО", "", If(circuit.Ток * 1.1 > 16, "25A", "16A"), "220VAC", "", "iCT")
                'Case "Моторна защита"
                '    Return New ControlDeviceConfig("3NO", "1NO+1NC", "9A", "230VAC", "LC1D")
                'Case "Моторен механизъм"
                '    Return New ControlDeviceConfig("", "", "", "", "NS100")
                'Case "Честотен регулатор"
                '    Return New ControlDeviceConfig("", "", "", "", "ATV")
            Case "Стълбищен автомат"
                Return New ControlDeviceConfig("", "", "0.5-20min",
                                               If(circuit.Ток * 1.1 > 16, "---", "16A"),
                                               "", "MINp")
                'Case "Електромер"
                '    Return New ControlDeviceConfig("", "", "", "", "kWh")
            Case "Фото реле"
                Return New ControlDeviceConfig("", "2-100 Lx",
                                               If(circuit.Ток * 1.1 > 10, "---", "10A"),
                                               "", "", "IC100")
            Case Else
                Return New ControlDeviceConfig("", "", "", "", "", "")
        End Select
    End Function
    ''' <summary>
    ''' Чертaе вертикална линия за токов кръг в таблото.
    ''' Позицията на линията зависи от:
    ''' - наличието на управление
    ''' - типа на устройството (напр. "Контакт")
    ''' </summary>
    ''' <param name="X">Х координата на линията</param>
    ''' <param name="circuit">Обект с данни за токовия кръг</param>
    ''' <param name="Y_Shina">Y координата на шината (референтна точка)</param>
    Public Sub DrawCircuitLines(X As Double, circuit As Form_Tablo_new.strTokow, Y_Shina As Double)
        ' Ако резерва не чертаем линия
        If circuit.Device = "Резерва" Then Return
        ' 1️ ДЕФИНИРАНЕ НА НАЧАЛНИ КООРДИНАТИ
        ' Начална позиция по Y (горна точка на линията)
        Dim startY As Double = Y_Shina - 135
        ' Крайна позиция по Y (долна точка на линията)
        Dim endY As Double = Y_Shina - 370
        ' 2️ ПРОВЕРКА ЗА УПРАВЛЕНИЕ
        ' Ако има управление (различно от Nothing, празно или "Няма"),
        ' линията се измества надолу (правим място за управление)
        If Not String.IsNullOrEmpty(circuit.Управление) AndAlso circuit.Управление <> "Няма" Then startY -= 95
        ' 3️⃣ ПРОВЕРКА ЗА ТИП УСТРОЙСТВО
        ' Ако устройството е "Контакт",
        ' използваме фиксирана позиция (override на startY)
        If circuit.Device = "Контакт" Then startY = Y_Shina - 253
        ' 4️ СЪЗДАВАНЕ НА ТОЧКИ
        ' Начална точка на линията
        Dim startPt As New Point3d(X, startY, 0)
        ' Крайна точка на линията
        Dim endPt As New Point3d(X, endY, 0)
        ' 5️ ЧЕРТАНЕ НА ЛИНИЯТА
        ' Чертaе линията в слой "EL_ТАБЛА"
        ' с настройки по слой (ByLayer)
        cu.DrowLine(startPt, endPt,
                "EL_ТАБЛА",
                DatabaseServices.LineWeight.ByLayer,
                "ByLayer")
    End Sub
    ''' <summary>
    ''' Чертaе RCD (ДТЗ) групи и разпределителни линии в таблото.
    ''' Логиката:
    ''' 1. Изчислява позициите на шините
    ''' 2. Обхожда всички токови кръгове
    ''' 3. Групира кръговете по RCD_Нула
    ''' 4. При смяна на групата затваря предишната
    ''' 5. Изчертава обща линия и RCD блок за всяка група
    ''' </summary>
    Private Sub DrawRCDBusbar(acDoc As Document, acCurDb As Database,
                          basePoint As Point3d,
                          circuits As List(Of Form_Tablo_new.strTokow))
        ' Изчисляване на началната X координата
        Dim X_Start As Double =
        basePoint.X + widthText + widthTextDim
        ' Y координата на главната шина
        Dim Y_Shina As Double = basePoint.Y + Y_Шина
        ' Y координата на RCD шината
        Dim Y_RCD As Double = Y_Shina - 118
        ' Начална колона на текущата RCD група
        Dim rcdGroupStart As Integer = 0
        ' RCD_Нула от предишния токов кръг
        Dim previousRCD_Null As String = ""
        ' Флаг дали в момента сме в активна RCD група
        Dim inRCDGroup As Boolean = False
        ' Текущ индекс на колоната
        Dim colIndex As Integer = 0
        ' Запазва последния токов кръг в активната група
        Dim currentGroupCircuit As Form_Tablo_new.strTokow = Nothing
        Try
            ' Обхождаме всички токови кръгове
            For Each circuit As Form_Tablo_new.strTokow In circuits
                ' Пропускаме специалните кръгове
                If circuit.ТоковКръг = "Разединител" OrElse
               circuit.ТоковКръг = "ОБЩО" Then Continue For
                ' Преминаваме към следващата колона
                colIndex += 1
                ' Проверка дали кръгът има ДТЗ
                Dim hasRCD As Boolean =
                Not String.IsNullOrEmpty(circuit.RCD_Нула) AndAlso
                circuit.RCD_Нула.Trim().ToUpper() <> "N"
                If hasRCD Then
                    ' Ако няма активна група → започваме нова
                    If Not inRCDGroup Then
                        rcdGroupStart = colIndex
                        previousRCD_Null = circuit.RCD_Нула.Trim().ToUpper()
                        currentGroupCircuit = circuit
                        inRCDGroup = True
                        ' Ако RCD_Нула е различно → затваряме старата група
                    ElseIf circuit.RCD_Нула.Trim().ToUpper() <> previousRCD_Null Then
                        DrawRCDGroupLine(
                                        acDoc,
                                        acCurDb,
                                        X_Start,
                                        Y_RCD,
                                        Y_Shina,
                                        rcdGroupStart,
                                        colIndex - 1,
                                        currentGroupCircuit
                                        )
                        ' Стартираме нова група
                        rcdGroupStart = colIndex
                        previousRCD_Null = circuit.RCD_Нула.Trim().ToUpper()
                        currentGroupCircuit = circuit
                    Else
                        ' Същото RCD_Нула → групата продължава
                        currentGroupCircuit = circuit
                    End If
                Else
                    ' Ако няма ДТЗ и има активна група → затваряме я
                    If inRCDGroup Then
                        DrawRCDGroupLine(
                                acDoc,
                                acCurDb,
                                X_Start,
                                Y_RCD,
                                Y_Shina,
                                rcdGroupStart,
                                colIndex - 1,
                                currentGroupCircuit)
                        inRCDGroup = False
                        previousRCD_Null = ""
                        currentGroupCircuit = Nothing
                    End If
                End If
            Next
            ' Ако последната група е останала отворена → затваряме я
            If inRCDGroup Then
                DrawRCDGroupLine(
                        acDoc,
                        acCurDb,
                        X_Start,
                        Y_RCD,
                        Y_Shina,
                        rcdGroupStart,
                        colIndex,
                        currentGroupCircuit)

            End If
        Catch ex As Exception
            MsgBox(
            "Възникна грешка: " &
            vbCrLf &
            ex.Message &
            vbCrLf &
            vbCrLf &
            ex.StackTrace,
            MsgBoxStyle.Critical
        )
        End Try
    End Sub
    ''' <summary>
    ''' Чертaе група с ДТЗ (RCD):
    ''' - хоризонтална шина
    ''' - блок на RCD в центъра
    ''' - текст с фази над шината
    ''' - попълва атрибутите на блока
    ''' </summary>
    ''' <param name="acDoc">Текущ документ</param>
    ''' <param name="acCurDb">Текуща база данни</param>
    ''' <param name="X_Start">Начална X позиция</param>
    ''' <param name="Y_RCD">Y позиция на шината (хоризонталната линия)</param>
    ''' <param name="Y_Shina">Y позиция за поставяне на RCD блока</param>
    ''' <param name="groupStart">Начална колона на групата</param>
    ''' <param name="groupEnd">Крайна колона на групата</param>
    ''' <param name="circuits">Данни за токовия кръг (RCD параметри)</param>
    Private Sub DrawRCDGroupLine(acDoc As Document, acCurDb As Database,
                         X_Start As Double, Y_RCD As Double, Y_Shina As Double,
                         groupStart As Integer, groupEnd As Integer,
                         circuits As Form_Tablo_new.strTokow
                         )
        Try
            ' 1 ИЗЧИСЛЯВАНЕ НА X ПОЗИЦИИТЕ
            ' Лява граница на групата
            Dim X_First As Double = X_Start + (groupStart - 1) * widthColom + widthColom / 4
            ' Дясна граница на групата
            Dim X_Last As Double = X_Start + (groupEnd) * widthColom - widthColom / 4
            ' Център на групата (за позициониране на RCD блока)
            Dim X_Center As Double = (X_First + X_Last) / 2
            ' 2️ ЧЕРТАНЕ НА ХОРИЗОНТАЛНА ЛИНИЯ (ШИНА)
            ' Чертaе шината между първата и последната позиция
            cu.DrowLine(New Point3d(X_First, Y_RCD, 0),
                New Point3d(X_Last, Y_RCD, 0),
                "EL_ТАБЛА",
                Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight070,
                "ByLayer")
            ' 3️ ВМЪКВАНЕ НА RCD БЛОК
            ' Поставя блока в центъра на групата
            Dim rcdBlockId As ObjectId = cu.InsertBlock("s_id_res_circ_break",
                                                 New Point3d(X_Center, Y_Shina, 0),
                                                 "EL_ТАБЛА",
                                                 New Scale3d(5, 5, 5))
            ' 4️⃣ ДОБАВЯНЕ НА ТЕКСТ НАД ШИНАТА
            ' Y позиция на текста (малко над линията)
            Dim textY As Double = Y_RCD + 15
            ' Текст с фази + нула + защитен проводник
            Dim phaseText As String = circuits.Фаза & "," & circuits.RCD_Нула & "," & "PE"
            ' Вмъкване на текста
            cu.InsertText(phaseText,
                  New Point3d(X_First, textY, 0),
                  "EL__DIM",
                  10,
                  TextHorizontalMode.TextLeft,
                  TextVerticalMode.TextBase)
            ' 5️ ПОПЪЛВАНЕ НА АТРИБУТИ НА БЛОКА
            ' Проверка дали блокът е създаден успешно
            If Not rcdBlockId.IsNull Then
                ' Стартираме транзакция за редакция
                Using trans As Transaction = acCurDb.TransactionManager.StartTransaction()
                    ' Вземаме референция към блока
                    Dim acBlkRef As BlockReference =
                                    DirectCast(trans.GetObject(rcdBlockId, OpenMode.ForWrite), BlockReference)
                    ' Обхождаме всички атрибути на блока
                    For Each objID As ObjectId In acBlkRef.AttributeCollection
                        ' Вземаме конкретен атрибут
                        Dim acAttRef As AttributeReference =
                                        DirectCast(trans.GetObject(objID, OpenMode.ForWrite), AttributeReference)
                        ' Попълваме според TAG-а
                        Select Case acAttRef.Tag
                            Case "1" : acAttRef.TextString = circuits.RCD_Клас
                            Case "2" : acAttRef.TextString = circuits.RCD_Полюси
                            Case "3" : acAttRef.TextString = circuits.RCD_Ток & "А"
                            Case "4" : acAttRef.TextString = "Мигновена"
                            Case "5" : acAttRef.TextString = circuits.RCD_Чувствителност & "mА"
                            Case "SHORTNAME" : acAttRef.TextString = circuits.RCD_Тип
                            Case "REFNB" : acAttRef.TextString = circuits.Tablo
                            Case "DESIGNATION" : acAttRef.TextString = ""
                        End Select
                    Next
                    ' Записваме промените
                    trans.Commit()
                End Using
            End If
        Catch ex As Exception
            ' Обработка на грешка – показва съобщение с детайли
            MsgBox("Възникна грешка: " & vbCrLf &
               ex.Message & vbCrLf & vbCrLf &
               ex.StackTrace,
               MsgBoxStyle.Critical)
        End Try
    End Sub
    ''' <summary>
    ''' Чертaе заземителната схема към главното табло.
    ''' Логиката:
    ''' 1. Проверява дали таблото е Гл.Р.Т.
    ''' 2. Чертaе връзката към заземителя
    ''' 3. Добавя текст за съпротивление на заземяване
    ''' 4. Вмъква динамичен блок "Заземление"
    ''' 5. Настройва параметрите и атрибутите на блока
    ''' 6. Добавя означение за PE проводник
    ''' </summary>
    Private Sub DrawGrounding(acDoc As Document, acCurDb As Database, X As Double, ptbasePoint As Point3d, panelName As String)
        ' Чертaем заземление само за главно разпределително табло
        If panelName <> "Гл.Р.Т." AndAlso panelName <> "ГлРТ" Then Return
        X = X +
            widthText +       ' Ширина на колоната за текст (напр. "Токов кръг")
            widthTextDim      ' Допълнителна ширина за текстова колона (напр. за единици)
        ' Хоризонтална линия към заземителя
        cu.DrowLine(
                    New Point3d(X, ptbasePoint.Y + Y_Шина, 0),
                    New Point3d(X - widthColom, ptbasePoint.Y + Y_Шина, 0),
                    "EL_ТАБЛА",
                    Autodesk.AutoCAD.DatabaseServices.LineWeight.ByLayer,
                    "ByLayer"
                    )
        ' Текст за съпротивление на заземяване
        cu.InsertText(
                "R<30Ω",
                New Point3d(X - widthColom,
                            ptbasePoint.Y + Y_Шина + 2 * padingText,
                            0),
                "EL__DIM",
                heightText,
                TextHorizontalMode.TextLeft,
                TextVerticalMode.TextBase
                )
        ' Вмъкване на блока "Заземление"
        Dim blkRecId =
        cu.InsertBlock(
            "Заземление",
            New Point3d(X - widthColom,
                         ptbasePoint.Y + Y_Шина,
                         0),
            "EL_ТАБЛА",
            New Scale3d(0.21, 0.21, 0.21)
        )
        ' Взимаме активния документ
        Dim doc As Document =
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Using trans As Transaction = doc.TransactionManager.StartTransaction()
            ' Взимаме BlockTable
            Dim acBlkTbl As BlockTable =
            trans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            ' Взимаме BlockReference на вмъкнатия блок
            Dim acBlkRef As BlockReference =
                            DirectCast(
                                trans.GetObject(blkRecId, OpenMode.ForWrite),
                                BlockReference
                            )
            ' Достъп до динамичните параметри на блока
            Dim props As DynamicBlockReferencePropertyCollection =
                         acBlkRef.DynamicBlockReferencePropertyCollection
            ' Настройване на параметрите на динамичния блок
            For Each prop As DynamicBlockReferenceProperty In props
                Select Case prop.PropertyName
                    Case "Visibility" : prop.Value = "Заземител-БЕЗ контролна клема"
                    Case "Distance2" : prop.Value = 15.0
                    Case "Position1 X" : prop.Value = 30.0
                    Case "Position1 Y" : prop.Value = -50.0
                    Case "Angle1" : prop.Value = 0.0
                End Select
            Next
            ' Достъп до атрибутите на блока
            Dim attCol As AttributeCollection =
            acBlkRef.AttributeCollection
            ' Попълване на атрибути
            For Each objID As ObjectId In attCol
                Dim dbObj As DBObject =
                trans.GetObject(objID, OpenMode.ForWrite)
                Dim acAttRef As AttributeReference = dbObj
                ' Попълване на атрибут "ТАБЛО"
                If acAttRef.Tag = "ТАБЛО" Then acAttRef.TextString = "2к"
            Next
            ' Запис на промените
            trans.Commit()
        End Using
        ' Добавяне на означение за защитен проводник
        cu.InsertText(
                "PE",
                New Point3d(
                    X - widthColom + 3 * padingText,
                    ptbasePoint.Y + Y_Шина - heightText - padingText,
                    0
                ),
                "EL__DIM",
                heightText,
                TextHorizontalMode.TextLeft,
                TextVerticalMode.TextBase
                )
    End Sub
    ''' <summary>
    ''' Процедурата DrawAnnotations създава текстови анотации (бележки) в AutoCAD чертеж,
    ''' свързани с електрическо табло, както и изчислява и визуализира общия брой полюси
    ''' на всички токови кръгове.
    ''' 
    ''' Използва се в контекста на автоматизирано чертане на табла, където освен графиката
    ''' е необходимо да се добавят и нормативни указания и обобщена информация.
    ''' </summary>
    ''' <param name="basePoint">
    ''' Базова точка за позициониране на текстовете. Всички анотации се разполагат
    ''' относително спрямо тази точка.
    ''' </param>
    ''' <param name="circuits">
    ''' Списък от токови кръгове (strTokow), използван за изчисляване на общия брой полюси.
    ''' </param>
    Private Sub DrawAnnotations(basePoint As Point3d, circuits As List(Of Form_Tablo_new.strTokow))
        Dim Zabelevka As String = "1. Таблото да се изпълни в съответствие с изискванията на БДС EN 61439-1."
        ' Добавяне на нови редове с допълнителни изисквания към таблото
        Zabelevka += vbCrLf & "2. Aпаратурата и тоководящите части да бъдат монтирани зад защитни капаци. "
        Zabelevka += vbCrLf & "3. Достъпа до палците и ръкохватките на комутационните апарати се осигурява посредством отвори в защитните капаци."
        Zabelevka += vbCrLf & "4. Апаратурата е избрана по каталога на SCHNEIDER ELECTRIC."
        Zabelevka += vbCrLf & "5. Изборът на автоматичните прекъсвачи е съобразен с токовете на к.с., спазени са изискванията за селективност."
        Zabelevka += vbCrLf & "6. При замяна типа на апаратурата да се преизчисли схемата."
        Zabelevka += vbCrLf & "7. При замяна номиналният ток на апаратурата да се преизчисли сечението на кабелите."
        cu.InsertMText("ЗАБЕЛЕЖКИ:",
                       New Point3d(basePoint.X,
                                   basePoint.Y - 20, 0),
                       "EL__DIM", 10)
        cu.InsertMText(Zabelevka,
                       New Point3d(basePoint.X + 30,
                                   basePoint.Y - 20 - heightRow, 0),
                       "EL__DIM", 10)
        Dim pol As Integer = 0
        For Each circuit As Form_Tablo_new.strTokow In circuits
            pol += circuit.Брой_Полюси
            Select Case circuit.Управление
                Case "Няма", "", "Електромер",
                 "Честотен регулатор",
                 "Моторен механизъм"
                    ' При тези типове управление не се добавят допълнителни полюси.
                Case "Стълбищен автомат", "Импулсно реле"
                    pol += 1
                Case "Фото реле"
                    pol += 3
                Case "Контактор", "Моторна защита"
                    pol += circuit.Брой_Полюси
            End Select
            Select Case circuit.RCD_Полюси
                Case "2p"
                    pol += 2
                Case "4p"
                    pol += 4
            End Select
        Next
        cu.InsertMText("Полюси -> " & pol.ToString(0),
                       New Point3d(basePoint.X + 160,
                                   basePoint.Y + 900, 0),
                       "Defpoints", 20, 1)
    End Sub
    ''' <summary>
    ''' Чертае управляващо устройство под прекъсвача
    ''' </summary>
    ''' <param name="acDoc">AutoCAD документ</param>
    ''' <param name="acCurDb">AutoCAD база данни</param>
    ''' <param name="circuit">Токов кръг</param>
    ''' <param name="X">X координата (център на колоната)</param>
    ''' <param name="breakerY">Y позиция на прекъсвача</param>
    Private Sub DrawControlDevice(acDoc As Document, acCurDb As Database,
                                  circuit As Form_Tablo_new.strTokow, X As Double, breakerY As Double)
        Try
            ' =====================================================
            ' 1️ ПРОВЕРКА ДАЛИ ИМА УПРАВЛЕНИЕ
            ' =====================================================
            If String.IsNullOrEmpty(circuit.Управление) Then Return
            If circuit.Управление = "Няма" Then Return
            ' =====================================================
            ' 2️ НАМИРАНЕ НА БЛОКА ОТ РЕЧНИКА
            ' =====================================================
            If Not ControlBlockMap.ContainsKey(circuit.Управление) Then Return
            Dim blockName As String = ControlBlockMap(circuit.Управление)
            If String.IsNullOrEmpty(blockName) Then Return
            ' =====================================================
            ' 3️ ИЗЧИСЛЯВАНЕ НА ПОЗИЦИЯТА (ВИНАГИ под прекъсвача)
            ' =====================================================
            Dim controlY As Double = breakerY - 135
            Dim insertPoint As New Point3d(X, controlY, 0)
            Dim blockScale As New Scale3d(5, 5, 5)
            ' =====================================================
            ' 4️ ПОЛУЧАВАНЕ НА ПАРАМЕТРИТЕ ЗА ТОЗИ ТИП
            ' =====================================================
            Dim config As ControlDeviceConfig = GetControlDeviceConfig(circuit)
            ' =====================================================
            ' 5️ ВМЪКВАНЕ НА БЛОКА
            ' =====================================================
            Dim blkRecId As ObjectId = cu.InsertBlock(blockName, insertPoint, "EL_ТАБЛА", blockScale)
            ' =====================================================
            ' 6️ ПОПЪЛВАНЕ НА АТРИБУТИТЕ
            ' =====================================================
            If Not blkRecId.IsNull Then
                Using trans As Transaction = acCurDb.TransactionManager.StartTransaction()
                    Dim acBlkRef As BlockReference = DirectCast(trans.GetObject(blkRecId, OpenMode.ForWrite), BlockReference)
                    For Each objID As ObjectId In acBlkRef.AttributeCollection
                        Dim acAttRef As AttributeReference = DirectCast(trans.GetObject(objID, OpenMode.ForWrite), AttributeReference)
                        Select Case acAttRef.Tag
                            Case "1" : acAttRef.TextString = config.Str_1
                            Case "2" : acAttRef.TextString = config.Str_2
                            Case "3" : acAttRef.TextString = config.Str_3
                            Case "4" : acAttRef.TextString = config.Str_4
                            Case "5" : acAttRef.TextString = ""
                            Case "SHORTNAME" : acAttRef.TextString = config.ShortName
                            Case "REFNB" : acAttRef.TextString = circuit.Tablo
                            Case "DESIGNATION" : acAttRef.TextString = ""
                        End Select
                    Next
                    Dim kvadrat As Boolean = True
                    If kvadrat Then
                        Dim Y_kvadrat As Double = controlY - 195
                        cu.InsertBlock("Ключ_квадрат",
                                       New Point3d(X - 32, Y_kvadrat, 0),
                                       "EL_ТАБЛА",
                                       New Scale3d(1, 1, 1))
                        cu.DrowLine(New Point3d(X - 32,
                                                Y_kvadrat + 25,
                                                0),
                                    New Point3d(X - 32,
                                                Y_kvadrat + 133,
                                                0),
                                    "EL_ТАБЛА",
                                    DatabaseServices.LineWeight.ByLayer,
                                    "ByLayer")
                    End If
                    trans.Commit()
                End Using
            End If
        Catch ex As Exception
            MsgBox("Възникна грешка: " & vbCrLf & ex.Message & vbCrLf & vbCrLf & ex.StackTrace, MsgBoxStyle.Critical)
        End Try
    End Sub
End Class
