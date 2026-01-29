Imports System.CodeDom
Imports System.IO
Imports System.Net
Imports System.Net.Security
Imports System.Security.Cryptography
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Public Class Legend
    Dim cu As CommonUtil = New CommonUtil()
    Public Structure strTablo
        Dim blName As String
        Dim blVisibility As String
        Dim blText As String
        Dim blInsert As Boolean
        Dim blScale As Double
        Dim blAngle As Double
        Dim ipY As Double
    End Structure
    Public Structure strLine
        Dim Layer As String
        Dim Linetype As String
        Dim count As Double
    End Structure
    Dim strLED_Lamp_Montav As String = ""
    Dim strLED_Lamp_SW_Potok As String = ""
    Dim strLED_Lamp_TIP As String = ""
    Dim strLamp_Power As String = ""
    Public Structure strТабло
        Dim bl_Табло As String
        Dim bl_Брой As Integer
        Dim bl_ИмеБлок As String
        Dim bl_1 As String
        Dim bl_2 As String
        Dim bl_3 As String
        Dim bl_4 As String
        Dim bl_5 As String
        Dim bl_6 As String
        Dim bl_7 As String
        Dim bl_8 As String
        Dim bl_9 As String
        Dim bl_10 As String
        Dim bl_DESIGNATION As String
        Dim bl_LONGNAME As String
        Dim bl_REFNB As String
        Dim bl_SHORTNAME As String
        Dim bl_RABATY As String
        Dim bl_RABATY2 As String
    End Structure
    <CommandMethod("LegendaKabeli")>
    Public Sub LegendaKabeli()
        ' Дефиниране на променлива ss_Kabeli, която събира обекти от тип "LINE" с помощта на метода cu.GetObjects
        ' Показва диалогов прозорец за избор на кабели в чертежа
        Dim ss_Kabeli = cu.GetObjects("LINE", "Изберете КАБЕЛИТЕ в чертежа:")
        ' Получаване на активния документ от DocumentManager
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        ' Създаване на обект от тип Editor за взаимодействие с потребителя
        Dim edt As Editor = acDoc.Editor
        ' Получаване на базата данни на текущия документ
        Dim acCurDb As Database = acDoc.Database
        ' Проверка дали няма маркирани кабели (обекти от тип LINE)
        If ss_Kabeli Is Nothing Then
            ' Показване на съобщение за грешка ако няма маркирани линии
            MsgBox("Няма маркиран нито едина линия.")
            ' Прекратяване изпълнението на текущата процедура
            Exit Sub
        End If
        ' Деклариране на двумерен масив Kabel за съхранение на до 50 кабела и информация за тях
        Dim Kabel(50, 1) As String
        Dim arrBlock(600) As strLine
        Dim index As Integer = 0
        ' Деклариране на променлива за резултата от избор на точка от потребителя
        Dim pPtRes As PromptPointResult
        ' Създаване на обект от тип PromptPointOptions за настройка на опции за избиране на точка
        Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
        ' Настройка на съобщението, което се показва на потребителя при избиране на точка
        pPtOpts.Message = vbLf & "Изберете точка на вмъкване на ЛЕГЕНДАТА: "
        ' Извикване на метода Editor.GetPoint за получаване на точка от потребителя въз основа на зададените опции
        pPtRes = acDoc.Editor.GetPoint(pPtOpts)
        ' Получаване на избраната от потребителя точка и запазване в променлива от тип Point3d
        Dim InsertPoint As Point3d = pPtRes.Value
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            ' Започва нова транзакция
            Try
                ' Опитва се да изпълни следния блок код
                For Each sObj As SelectedObject In ss_Kabeli
                    ' Цикъл през всички избрани обекти в ss_Kabeli
                    Dim line As Line = TryCast(acTrans.GetObject(sObj.ObjectId, OpenMode.ForWrite), Line)
                    ' Опитва се да преобразува текущия избран обект в тип Line
                    Dim iVisib As Integer = -1
                    If line.Linetype = "" Then Continue For
                    ' Пропуска обектите, които нямат зададен тип линия
                    Dim Line_Layer As String = line.Layer
                    line.LinetypeScale = 0.00000000001

                    iVisib = Array.FindIndex(arrBlock, Function(f) f.Layer = Line_Layer)
                    ' Проверява дали слоят вече съществува в arrBlock
                    If iVisib = -1 Then
                        ' Ако слоят не съществува, го добавя в масива arrBlock
                        arrBlock(index).Layer = Line_Layer
                        index += 1
                    End If
                Next
                ' Добавя текст "ЛЕГЕНДА НА ЦВЕТОВЕТЕ:" в чертежа
                cu.InsertMText("ЛЕГЕНДА НА ЦВЕТОВЕТЕ:", InsertPoint, "EL__DIM", 15, TextWidth:=290)
                ' Получава координатите на точката за поставяне
                Dim ipX As Double = pPtRes.Value.X
                Dim ipY As Double = pPtRes.Value.Y
                ' Задава начални и крайни точки за линия и ги рисува
                Dim insPointText1 As Point3d = New Point3d(ipX, ipY + 20, 0)
                Dim insPointLine_n As Point3d = New Point3d(ipX, ipY - 20, 0)
                Dim insPointLine_k As Point3d = New Point3d(ipX + 270, ipY - 20, 0)
                cu.DrowLine(insPointLine_n, insPointLine_k, "0", LineWeight.LineWeight030, "ByLayer", LineColor:=90)
                insPointLine_n = New Point3d(ipX, ipY - 22, 0)
                insPointLine_k = New Point3d(ipX + 270, ipY - 22, 0)
                cu.DrowLine(insPointLine_n, insPointLine_k, "0", LineWeight.LineWeight018, "ByLayer", LineColor:=130)
                ' Променя координатата Y за следващия елемент
                ipY -= 50
                ' Цикъл през всички елементи в arrBlock и рисува линии и текстове
                For i = 0 To UBound(arrBlock)
                    If arrBlock(i).Layer = "" Then Exit For
                    ' Прекъсва цикъла ако слоя е празен
                    If arrBlock(i).Layer = "EL__DIM" OrElse
                        arrBlock(i).Layer = "EL__ORAZ" OrElse
                        arrBlock(i).Layer = "ELEKTRO" OrElse
                        arrBlock(i).Layer = "EL_ТАБЛА" Then Continue For
                    ' Задава координати и рисува линии и текстове
                    insPointLine_n = New Point3d(ipX, ipY, 0)
                    insPointLine_k = New Point3d(ipX + 270, ipY, 0)
                    insPointText1 = New Point3d(ipX, ipY + 20, 0)
                    cu.DrowLine(insPointLine_n, insPointLine_k, arrBlock(i).Layer, LineWeight.LineWeight035, "ByLayer")
                    cu.InsertMText(cu.line_Layer(arrBlock(i).Layer),
                                                insPointText1,
                                                arrBlock(i).Layer,
                                                dbTextHeight:=12,
                                                TextWidth:=250)
                    ipY -= 30
                Next
                acTrans.Commit()
                ' Потвърждава транзакцията ако няма грешки
            Catch ex As Exception
                ' Обработва грешки при изпълнение на кода
                MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
                acTrans.Abort()
                ' Прекратява транзакцията при грешка
            End Try
        End Using
        ' Край на транзакцията
    End Sub
    <CommandMethod("Legenda")>
    Public Sub Legenda()
        ' Извличане на активния документ и свързаните обекти
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim edt As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database
        ' Избиране на обекти от типа "INSERT" в чертежа, които представляват блокови референции
        Dim SelectedSet = cu.GetObjects("INSERT", "Изберете БЛОКОВЕТЕ в чертежа")
        ' Инициализация на масив, който ще съхранява информация за блоковете, и дефиниране на PI
        Dim arrBlock(500) As strTablo
        Dim PI As Double = 3.1415926535897931
        ' Проверка дали са избрани блокове
        If SelectedSet Is Nothing Then
            MsgBox("Нама маркиран нито един блок.")
            Exit Sub
        End If
        ' Променливи за идентификатор на блок и индекс за масива
        Dim blkRecId As ObjectId = ObjectId.Null
        Dim index As Integer = 0
        ' Започване на транзакция за достъп до базата данни на чертежа
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                ' Цикъл през всички избрани блокови референции
                For Each sObj As SelectedObject In SelectedSet
                    blkRecId = sObj.ObjectId
                    ' Вземане на референция към текущия блок и неговите атрибути и динамични свойства
                    Dim acBlkRef As BlockReference = DirectCast(acTrans.GetObject(blkRecId, OpenMode.ForWrite), BlockReference)
                    Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                    Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection
                    Dim Visibility As String = ""
                    ' Обхождане на свойствата на блока, за да се извлече стойността на видимостта и други атрибути
                    For Each prop As DynamicBlockReferenceProperty In props
                        If prop.PropertyName = "Visibility1" Then Visibility = prop.Value
                        If prop.PropertyName = "Visibility" Then Visibility = prop.Value
                        If prop.PropertyName = "Тип" Then Visibility = prop.Value
                        If prop.PropertyName = "Тип" Then strLED_Lamp_TIP = prop.Value
                        If prop.PropertyName = "Св_поток" Then strLED_Lamp_SW_Potok = prop.Value
                        If prop.PropertyName = "Монтаж" Then strLED_Lamp_Montav = prop.Value
                    Next
                    ' Обхождане на атрибутите на блока за извличане на специфична информация
                    For Each objID As ObjectId In attCol
                        Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                        Dim acAttRef As AttributeReference = dbObj
                        If acAttRef.Tag = "LED" Then strLamp_Power = acAttRef.TextString
                    Next
                    ' Извличане на името на блока
                    Dim blName As String = (CType(acBlkRef.DynamicBlockTableRecord.GetObject(OpenMode.ForRead), BlockTableRecord)).Name
                    ' Проверка дали блокът вече е добавен в масива с блокове
                    Dim iVisib As Integer = Array.FindIndex(arrBlock, Function(f) f.blName = blName And f.blVisibility = Visibility)
                    If iVisib <> -1 Then Continue For
                    ' Инициализация на нов блок в масива с атрибути по подразбиране
                    arrBlock(index).blScale = 1
                    arrBlock(index).blAngle = 0
                    ' Обработка на блока според неговото име и видимост
                    Select Case blName
                        Case "Табло_Ново"
                            arrBlock(index).blName = blName
                            arrBlock(index).blVisibility = Visibility
                            arrBlock(index).blText = "- " + BlockText(blName, Visibility)
                            arrBlock(index).blInsert = vbFalse
                            index += 1
                        Case "Контакт", "LED_луна", "Линия МХЛ - 220V",
                             "Прожектор", "бойлерно табло", "Ключ_знак", "Датчик_ПАБ",
                             "Домофон", "Високоговорител", "Аудио система", "Камери",
                             "СОТ", "Аудио система", "Ключ_квадрат", "Розетка_1",
                             "Ключ_знак_WIFI", "Авария", "Авария_100", "LED_DENIMA",
                             "LED_ULTRALUX", "LED_ULTRALUX_100", "Луминисцентна лампа", "LED_lenta",
                             "LED_ULTRALUX_нов"
                            arrBlock(index).blName = blName
                            arrBlock(index).blVisibility = Visibility
                            arrBlock(index).blText = "- " + BlockText(blName, Visibility)
                            arrBlock(index).blInsert = vbFalse
                            index += 1
                        Case "Плафони"
                            arrBlock(index).blName = blName
                            arrBlock(index).blVisibility = Visibility
                            arrBlock(index).blText = "- " + BlockText(blName, Visibility)
                            arrBlock(index).blInsert = vbFalse
                            arrBlock(index).blAngle = 1.5 * PI
                            index += 1
                        Case "Полилей"
                            arrBlock(index).blName = blName
                            arrBlock(index).blVisibility = Visibility
                            arrBlock(index).blText = "- " + BlockText(blName, Visibility)
                            arrBlock(index).blInsert = vbFalse
                            Select Case Visibility
                                Case "1х60 - Рошава", "1х60 - Кръгла"
                                    arrBlock(index).blScale = 1
                                Case "4х60 - Рошава", "3х60 - Рошава",
                                     "2х60 - Кръгла", "4х60 - Кръгла", "3х60 - Кръгла"
                                    arrBlock(index).blScale = 0.5
                                Case "1х60 - Индийски"
                                    arrBlock(index).blScale = 0.75
                                Case "2х60 - Рошава"
                                    arrBlock(index).blScale = 0.5
                                    arrBlock(index).blAngle = 1.5 * PI
                            End Select
                            index += 1
                        Case "Металхаогенна лампа"
                            arrBlock(index).blName = blName
                            arrBlock(index).blVisibility = Visibility
                            arrBlock(index).blText = "- " + BlockText(blName, Visibility)
                            arrBlock(index).blInsert = vbFalse
                            Select Case Visibility
                                Case "3х35 - Кръг", "2х35 - Кръг", "1х35 - Кръг", "4х35 - Кръг"
                                    arrBlock(index).blAngle = 0
                                    arrBlock(index).blScale = 1
                                Case "2х35 - за картина", "1х35 - за картина", "1х35 - Право"
                                    arrBlock(index).blAngle = 1.5 * PI
                                    arrBlock(index).blScale = 1
                                Case "4х35 - Дъга", "3х35 - Дъга", "2х35 - Дъга",
                                     "3х35 - Право", "2х35 - Право", "4х35 - Право"
                                    arrBlock(index).blAngle = PI / 2
                                    arrBlock(index).blScale = 0.75
                                Case "1х35 - Дъга", "1х35 - 90°"
                                    arrBlock(index).blAngle = PI / 2
                                    arrBlock(index).blScale = 1
                            End Select
                            index += 1
                        Case "Бойлер"
                            Select Case Visibility
                                Case "Изход 1p", "Изход 3p", "ПВ"
                                    arrBlock(index).blName = blName
                                    arrBlock(index).blVisibility = Visibility
                                    arrBlock(index).blText = "- " + BlockText(blName, Visibility)
                                    arrBlock(index).blInsert = vbFalse
                                    index += 1
                            End Select
                        Case "Вентилации"
                            Select Case Visibility
                                Case "Вентилатор - кръг - баня",
                                             "Вентилатор - кръг",
                                             "Вентилатор - правоъг"
                                    arrBlock(index).blName = blName
                                    arrBlock(index).blVisibility = Visibility
                                    arrBlock(index).blText = "- " + BlockText(blName, Visibility)
                                    arrBlock(index).blInsert = vbFalse
                                    index += 1
                            End Select
                        Case Else
                            '"Кабел", "Вентилации", "uli4no", "Качване",
                            ' "Ramka_Samo_ramka_EWG", "Ramka_Samo_ramka",
                            ' "Атрибути_Таблица_Чужди", "Контактор", "Заземление",
                            ' "Конзола_розетка", "Конзола", "Конзола_точка",
                            '"Скара", "Канал"
                    End Select
                Next
                ' Извличане на точка за вмъкване на легендата от потребителя
                Dim pPtRes As PromptPointResult
                Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
                ' Настройка на съобщението, което ще се покаже в командния ред на AutoCAD
                pPtOpts.Message = vbLf & "Изберете точка на вмъкване на ЛЕГЕНДАТА: "
                ' Чакаме потребителят да избере точка в чертежа
                pPtRes = acDoc.Editor.GetPoint(pPtOpts)
                ' Запазваме избраната точка в променлива от тип Point3d
                Dim InsertPoint As Point3d = pPtRes.Value
                ' -------------------------------------------------------------
                ' Горен ляв ъгъл на легендата
                ' Започваме със стойността на точката на вмъкване, избрана от потребителя
                ' -------------------------------------------------------------
                Dim topLeft As Point3d = InsertPoint
                ' -------------------------------------------------------------
                ' Долен десен ъгъл на легендата
                ' Инициализираме с крайни стойности, за да може първият обект
                ' да зададе реалния долен десен ъгъл
                ' Тази променлива ще се актуализира, ако открием обект по-дясно или по-долу
                ' -------------------------------------------------------------
                'Dim bottomRight As Point3d = New Point3d(Double.MinValue, Double.MaxValue, 0)
                Dim bottomRight As Point3d = InsertPoint

                ' Вмъкване на текст "ЛЕГЕНДА" в избраната точка
                ' Параметри: текст, координати, стил (EL__DIM), височина на текста (15)
                cu.InsertMText("ЛЕГЕНДА", InsertPoint, "0", 15, 5)

                ' Извличане на X и Y координатите на избраната точка
                Dim ipX As Double = pPtRes.Value.X
                Dim ipY As Double = pPtRes.Value.Y

                ' Определяне на начална и крайна точка на линия под текста
                ' Линията е хоризонтална – започва от избраната X координата и свършва на +100 по X
                ' Y координатата е 20 единици под избраната точка
                Dim insPointLine_n As Point3d = New Point3d(ipX, ipY - 20, 0)
                Dim insPointLine_k As Point3d = New Point3d(ipX + 100, ipY - 20, 0)

                ' Определяне на позиции за бъдещи текстове под линията
                ' Текст1 ще е на 30 единици надясно и 30 надолу от избраната точка
                ' Текст2 ще е на 150 единици надясно и 30 надолу
                Dim insPointText1 As Point3d = New Point3d(ipX + 30, ipY - 30, 0)
                Dim insPointText2 As Point3d = New Point3d(ipX + 150, ipY - 30, 0)

                ' Начертаване на първата хоризонтална линия
                ' Параметри: начална и крайна точка, слой ("0"), дебелина (0.30), тип на линия ("ByLayer"), цвят (90)
                cu.DrowLine(insPointLine_n, insPointLine_k, "0", LineWeight.LineWeight030, "ByLayer", LineColor:=90)

                ' Начертаване на втора хоризонтална линия, която е на 2 единици под първата
                insPointLine_n = New Point3d(ipX, ipY - 22, 0)
                insPointLine_k = New Point3d(ipX + 100, ipY - 22, 0)
                cu.DrowLine(insPointLine_n, insPointLine_k, "0", LineWeight.LineWeight030, "ByLayer", LineColor:=90)

                ' Променяме координатите ipX и ipY за бъдещо използване
                ' X = избраната X координата + 50
                ' Y = избраната Y координата - 50
                ipX = pPtRes.Value.X + 50
                ipY = pPtRes.Value.Y - 50
                ' Обхождаме всички елементи (блокове) в масива arrBlock
                For Each iarrBlock In arrBlock
                    ' Ако блокът вече е вмъкнат (флаг blInsert = True), пропускаме го
                    If iarrBlock.blInsert Then Continue For
                    ' Ако блокът е - Силово разпределително табло, пропускаме го
                    ' Домофонна централа, Пожароизвестителна централа и Слаботоково табло ги слагаме
                    If iarrBlock.blVisibility = "Табло" Then Continue For
                    ' Определяме точка за вмъкване на блока – текущите координати ipX, ipY
                    Dim insPointBlock As Point3d = New Point3d(ipX, ipY, 0)
                    ' Вземаме името на блока от текущия елемент
                    Dim sss As String = iarrBlock.blName
                    ' Ако блокът няма име → прекратяваме цикъла (няма смисъл да продължаваме)
                    If iarrBlock.blName Is Nothing Then Exit For
                    ' Вмъкваме блока в чертежа на определената точка с мащаб по трите оси
                    Dim InsertBlock As ObjectId = cu.InsertBlock(iarrBlock.blName,
                                                 insPointBlock, "0",
                                                 New Scale3d(iarrBlock.blScale,
                                                             iarrBlock.blScale,
                                                             iarrBlock.blScale))

                    bottomRight = UpdateBottomRight(InsertBlock, acTrans, bottomRight)
                    ' Взимаме BlockReference, за да можем да променяме динамични свойства и атрибути
                    Dim acBlkRef As BlockReference = DirectCast(acTrans.GetObject(InsertBlock, OpenMode.ForWrite), BlockReference)
                    ' Събираме всички динамични свойства на блока (Dynamic Block Properties)
                    Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection
                    ' Определяме позиция за текста, който ще стои до блока
                    Dim insPointText As Point3d = New Point3d(ipX + 50, ipY + 10, 0)
                    ' Определяме начална и крайна точка на хоризонтална линия, която ще се чертае до блока
                    insPointLine_n = New Point3d(ipX - 100, ipY + 10, 0)
                    insPointLine_k = New Point3d(ipX + 200, ipY + 10, 0)
                    ' Начертаваме хоризонталната линия
                    Dim line As ObjectId = cu.DrowLine(insPointLine_n, insPointLine_k, "EL__DIM", LineWeight.ByLayer, "ByLayer")
                    ' Вмъкваме текст (описание към блока)
                    Dim objIdText As ObjectId = cu.InsertMText(iarrBlock.blText, insPointText, "0", 12)

                    bottomRight = UpdateBottomRight(objIdText, acTrans, bottomRight)

                    ' Подготвяме променливи за различни размери на лампите
                    Dim LumLamp_dylv, LumLamp_6ir As Double
                    ' Флаг (тук винаги се задава True, но явно е предвидено за условие)
                    Dim yesType As Boolean = vbFalse
                    yesType = vbTrue
                    ' Подготвяме стойности за някои параметри на динамичния блок
                    Dim Position_X As Double = -100
                    Dim Position_Y As Double = 0
                    Dim Distance As Double = 12
                    ' Обхождаме всички динамични свойства на блока
                    For Each prop As DynamicBlockReferenceProperty In props
                        ' Ако има свойство "Visibility" или "Visibility1" → задаваме видимостта от iarrBlock
                        If prop.PropertyName = "Visibility" Then prop.Value = iarrBlock.blVisibility
                        If prop.PropertyName = "Visibility1" Then prop.Value = iarrBlock.blVisibility
                        ' Задаваме позиция по X и Y, както и дистанция
                        If prop.PropertyName = "Position X" Then prop.Value = Position_X
                        If prop.PropertyName = "Position Y" Then prop.Value = Position_Y
                        If prop.PropertyName = "Distance" Then prop.Value = Distance
                        ' Задаваме ъгъл на завъртане от iarrBlock
                        If prop.PropertyName = "Angle" Then prop.Value = iarrBlock.blAngle
                        ' Ако има свойство "Тип" → задаваме специфични стойности за дължина и ширина на лампата
                        If prop.PropertyName = "Тип" Then
                            Select Case iarrBlock.blVisibility
                                Case "Кухня"
                                    LumLamp_dylv = 60
                                    LumLamp_6ir = 5
                                Case "IP-20"
                                    LumLamp_dylv = 40
                                    LumLamp_6ir = 40
                                Case "IP-54", "IP-65", "IP-66"
                                    LumLamp_dylv = 60
                                    LumLamp_6ir = 15
                            End Select
                            ' Задаваме тези стойности на други динамични свойства (Distance1…Distance4)
                            For Each prop1 As DynamicBlockReferenceProperty In props
                                If prop1.PropertyName = "Distance1" Then prop1.Value = LumLamp_dylv / 2
                                If prop1.PropertyName = "Distance2" Then prop1.Value = LumLamp_dylv / 2
                                If prop1.PropertyName = "Distance3" Then prop1.Value = LumLamp_6ir / 2
                                If prop1.PropertyName = "Distance4" Then prop1.Value = LumLamp_6ir / 2
                            Next
                            ' Преместваме надолу Y координатата (за да се остави място за следващия елемент)
                            ipY -= 65
                        End If
                    Next
                    ' Взимаме колекцията от атрибути на блока
                    Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
                    ' Обхождаме всеки атрибут
                    For Each objID As ObjectId In attCol
                        Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForRead)
                        Dim acAttRef As AttributeReference = dbObj
                        ' Изчистваме съдържанието на определени атрибути (оставяме ги празни)
                        If acAttRef.Tag = "МОЩНОСТ" Then acAttRef.TextString = ""
                        If acAttRef.Tag = "КРЪГ" Then acAttRef.TextString = ""
                        If acAttRef.Tag = "ТАБЛО" Then acAttRef.TextString = ""
                        If acAttRef.Tag = "Pewdn" Then acAttRef.TextString = ""
                        If acAttRef.Tag = "PEWDN1" Then acAttRef.TextString = ""
                        If acAttRef.Tag = "ВИС" Then acAttRef.TextString = ""
                        If acAttRef.Tag = "LED" Then acAttRef.TextString = ""
                    Next
                    ' След всеки блок местим Y координатата още надолу, за да има отстояние
                    ipY -= 50
                    ' Задаваме, че блокът вече е обработен (не е за ново вмъкване)
                    iarrBlock.blInsert = vbFalse
                Next
                acTrans.Commit()
            Catch ex As Exception
                MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
                acTrans.Abort()
            End Try
        End Using
    End Sub
    Public Function UpdateBottomRight(objId As ObjectId, acTrans As Transaction, bottomRight As Point3d) As Point3d
        ' Отваряме обекта за четене
        Dim acEnt As Entity = DirectCast(acTrans.GetObject(objId, OpenMode.ForRead), Entity)

        ' Проверка дали обектът е блок
        If TypeOf acEnt Is BlockReference Then
            Dim br As BlockReference = DirectCast(acEnt, BlockReference)
            Dim ext As Extents3d = br.GeometricExtents
            bottomRight = New Point3d(bottomRight.X, Math.Min(bottomRight.Y, ext.MinPoint.Y), bottomRight.Z)
            Return bottomRight
        End If

        ' Проверка дали обектът е MText
        If TypeOf acEnt Is MText Then
            Dim mt As MText = DirectCast(acEnt, MText)
            Dim ext As Extents3d = mt.GeometricExtents
            bottomRight = New Point3d(Math.Max(bottomRight.X, ext.MaxPoint.X), bottomRight.Y, bottomRight.Z)
            Return bottomRight
        End If

        ' Ако обектът не е блок или текст, връщаме без промяна
        Return bottomRight
    End Function
    <CommandMethod("LegendaTablo")>
    Public Sub LegendaTablo()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim edt As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database
        Dim SelectedSet = cu.GetObjects("INSERT", "Изберете блок")

        If SelectedSet Is Nothing Then
            MsgBox("Нама маркиран нито един блок.")
            Exit Sub
        End If
        Dim Брой_Апарати As Integer = 6
        ' Създаваме опции за въвеждане на броя колони
        Dim pDouOpts As PromptDoubleOptions = New PromptDoubleOptions("")
        With pDouOpts
            .Keywords.Add("1")  ' всички апарати в една колона
            .Keywords.Add("2")
            .Keywords.Add("3")
            .Keywords.Add("4")
            .Keywords.Add("5")
            .Keywords.Add("6")
            .Keywords.Default = "1"  ' по подразбиране - всички апарати в една колона
            .Message = vbCrLf & "Въведете броя на колоните (1 = всички апарати в една колона): "
            .AllowZero = True
            .AllowNegative = False
        End With
        Dim pKeyRes As PromptDoubleResult = acDoc.Editor.GetDouble(pDouOpts)
        Dim Брой_Колони As Integer
        If pKeyRes.Status = PromptStatus.Keyword Then
            Брой_Колони = CInt(pKeyRes.StringResult)
        Else
            Брой_Колони = CInt(pKeyRes.Value)
        End If
        Dim blkRecId As ObjectId = ObjectId.Null
        Dim index As Integer = 0
        Dim arrBlock = cu.GetAparati(SelectedSet)
        Dim pPtRes As PromptPointResult
        Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
        pPtOpts.Message = vbLf & "Изберете точка на вмъкване на ЛЕГЕНДАТА: "
        pPtRes = acDoc.Editor.GetPoint(pPtOpts)
        Dim InsertPoint As Point3d = pPtRes.Value
        Dim Tekst1 As String = ""
        Dim Tekst2 As String = ""
        ' Точка на вмъкване на текста
        Dim ipX As Double = pPtRes.Value.X
        Dim ipY As Double = pPtRes.Value.Y
        Dim extraRows As Double = 0     ' брой допълнителни редове
        ' Константи за положението на текста за таблата
        Dim ipY_Text As Double = 70
        Dim ipX_Text As Double = 50

        Dim Tablo_Layer As Boolean = False
        Dim Tablo_Name As Boolean = True
        Dim Tablo_Text As String = ""

        Dim offset_Колона As Double = 500
        Dim Br_kolona As Integer = 0

        Dim offset_X As Double = 0
        Dim offset_Y As Double = 0

        cu.InsertMText("ЛЕГЕНДА:", InsertPoint, "EL__DIM", 15)
        Dim insPointLine_n As Point3d = New Point3d(ipX, ipY - 20, 0)
        Dim insPointLine_k As Point3d = New Point3d(ipX + 100, ipY - 20, 0)

        Dim insPointText1 As Point3d = New Point3d(ipX + 30, ipY - 30, 0)

        cu.DrowLine(insPointLine_n, insPointLine_k, "0", LineWeight.LineWeight030, "ByLayer", 90)
        insPointLine_n = New Point3d(ipX, ipY - 22, 0)
        insPointLine_k = New Point3d(ipX + 100, ipY - 22, 0)
        cu.DrowLine(insPointLine_n, insPointLine_k, "0", LineWeight.LineWeight018, "ByLayer", 130)
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                Dim br_aparati As Integer = 0

                Dim realCount As Integer = 0
                ' realCount съдържа броя на реално използваните елементи
                For Each Apar In arrBlock
                    If Apar.bl_ИмеБлок Is Nothing Then
                        Exit For
                    End If
                    If Apar.bl_ИмеБлок = "Кабел" OrElse
                       Apar.bl_ИмеБлок = "Заземление" OrElse
                       Apar.bl_ИмеБлок = "Ключ_квадрат" Then Continue For
                    realCount += 1
                Next

                Dim rows As Integer
                If Брой_Колони = 0 Then
                    rows = realCount                              ' Ако не е зададено нищо, слагаме всички в една колона
                Else
                    rows = Math.Ceiling(realCount / Брой_Колони)  ' Разделяме равномерно апаратите по колоните
                End If

                Dim row As Integer = 0
                Dim col As Integer = 0

                For Each Apar In arrBlock
                    If Apar.bl_ИмеБлок Is Nothing Then Exit For
                    If Apar.bl_ИмеБлок = "Кабел" OrElse
                       Apar.bl_ИмеБлок = "Заземление" OrElse
                       Apar.bl_ИмеБлок = "Ключ_квадрат" Then Continue For

                    Tekst1 = Apar.bl_SHORTNAME
                    Tekst2 = ""

                    Select Case Apar.bl_SHORTNAME
                        Case ""
                            Continue For
                        Case "Метален шкаф стоящ", "Метален шкаф"
                            Tekst1 = Apar.bl_SHORTNAME
                            Tekst1 += "; В:"
                            Tekst1 += Apar.bl_1
                            Tekst1 += "; Ш:"
                            Tekst1 += Apar.bl_2
                            Tekst1 += "; Д:"
                            Tekst1 += Apar.bl_3 + ";" + vbCrLf
                            Tekst1 += "Врата: "
                            Tekst1 += Apar.bl_4 + vbCrLf
                            Tekst1 += "Степен на защита: IP66"
                            cu.InsertMText(Tekst1, New Point3d(InsertPoint.X + ipX_Text, InsertPoint.Y + ipY_Text, 0), "EL__DIM", 12)
                            Tablo_Name = False
                            Tablo_Text = Apar.bl_Табло
                            Continue For
                        Case "Kaedra", "Mini Kaedra"
                            Tekst1 = "Полиестерен шкаф; "
                            Tekst1 += Apar.bl_SHORTNAME + vbCrLf
                            Tekst1 += "Брой модули: "
                            Tekst1 += Apar.bl_2
                            Tekst1 += "; Брой редове: "
                            Select Case Apar.bl_3
                                Case ""
                                    Select Case Apar.bl_2
                                        Case "3", "4", "6", "8", "12", "18"
                                            Tekst1 += "1"
                                        Case "24"
                                            Tekst1 += "2"
                                        Case "54"
                                            Tekst1 += "3"
                                        Case "72"
                                            Tekst1 += "4"
                                    End Select
                                Case Else
                                    Tekst1 += Apar.bl_3
                            End Select
                            Tekst1 += vbCrLf + "Степен на защита: IP65"
                            cu.InsertMText(Tekst1, New Point3d(InsertPoint.X + ipX_Text, InsertPoint.Y + ipY_Text, 0), "EL__DIM", 12)
                            Tablo_Name = False
                            Tablo_Text = Apar.bl_Табло
                            Continue For
                        Case "Изпъкнал монтаж", "Вграден монтаж"
                            Tekst1 = "Полиестерен шкаф; "
                            Tekst1 += Apar.bl_SHORTNAME + vbCrLf
                            Tekst1 += "Брой модули: "
                            Tekst1 += Apar.bl_1
                            Tekst1 += "; Врата: "
                            Tekst1 += Apar.bl_2 + vbCrLf
                            Tekst1 += "Степен на защита: IP40"
                            cu.InsertMText(Tekst1, New Point3d(InsertPoint.X + ipX_Text, InsertPoint.Y + ipY_Text, 0), "EL__DIM", 12)
                            Tablo_Name = False
                            Tablo_Text = Apar.bl_Табло
                            Continue For
                        Case "E60", "Е60", "iC60", "C120",
                             "Е120", "iK60", "NG160",
                             "EZ9 MCB", "EZ9 MCB "
                            Tekst1 = Tekst1 + " " + Apar.bl_1
                            Tekst1 = Tekst1 + vbCrLf + Apar.bl_2 + ", " + Apar.bl_3
                            Tekst1 = Tekst1 + vbCrLf + Apar.bl_4
                            Tekst2 = "Автоматичен прекъсвач"
                            Tekst2 = Tekst2 + vbCrLf + "Крива: " + Apar.bl_2 + ", Брой полюси: " + Apar.bl_3 + ","
                            Tekst2 = Tekst2 + vbCrLf + "Номинален ток: " + Apar.bl_4
                        Case "NS1600", "NS1250", "NS1000", "NS800", "NS630b", "NSX250", "NSX160", "NSX100",
                             "EZCV250",
                             "NW63", "NW50", "NW40b", "NW40", "NW32", "NW25",
                             "NW20", "NW16", "NW12", "NW10", "NW08",
                             "NT16", "NT12", "NT10", "NT08", "NT06", "NB600",
                             "EZC400", "EZC250", "EZC100", "NSX630", "NSX400"
                            Tekst1 += " " + Apar.bl_1
                            Tekst1 += vbCrLf + Apar.bl_2
                            Tekst1 += vbCrLf + Apar.bl_3
                            Tekst1 += vbCrLf + Apar.bl_4
                            Tekst2 = "Автоматичен прекъсвач,"
                            Tekst2 = Tekst2 + vbCrLf + "Изкл. способност: " + Apar.bl_1 + ", " + "Брой полюси: " + Apar.bl_2 + ","
                            Tekst2 = Tekst2 + vbCrLf + "Защита: " + Apar.bl_3
                        Case "DPNа Vigi", "DPNa Vigi", "DPN N Vigi"
                            Tekst1 += vbCrLf + Apar.bl_1 + ", " + Apar.bl_5
                            Tekst1 += vbCrLf + Apar.bl_2 + ", " + Apar.bl_3 + ", " + Apar.bl_4
                            Tekst2 = "Автом. прекъсвач с дефектнотокова защита,"
                            Tekst2 += vbCrLf + "Тип: " + Apar.bl_1 + ", " + "Чуствителност: " + Apar.bl_5 + ","
                            Tekst2 += vbCrLf + "Бр. полюси: " + Apar.bl_2 + ", " + "Крива: " + Apar.bl_3 + ", " + "Ном. ток: " + Apar.bl_4
                        Case "EZ9 RCBO"
                            Tekst1 += vbCrLf + Apar.bl_1 + ", " + Apar.bl_5
                            Tekst1 += vbCrLf + Apar.bl_2 + ", " + Apar.bl_3 + ", " + Apar.bl_4
                            Tekst2 = "Автом. прекъсвач с дефектнотокова защита,"
                            Tekst2 += vbCrLf + "Тип: " + Apar.bl_1 + ", " + "Чуствителност: " + Apar.bl_5
                            Tekst2 += vbCrLf + "Бр. полюси: " + Apar.bl_2 + ", " + "Крива: " + Apar.bl_3 + ", " + "Ном. ток: " + Apar.bl_4
                            'Tekst2 += vbCrLf + "Чуствителност: " + Apar.bl_5
                        Case "ID Domae", "iID", "iID К", "EZ9 RCCB"
                            Tekst1 = Tekst1 + vbCrLf + Apar.bl_1 + ", " + Apar.bl_2
                            Tekst1 = Tekst1 + vbCrLf + Apar.bl_3
                            Tekst2 = "Дефектнотокова защита,"
                            Tekst2 = Tekst2 + vbCrLf + "Тип: " + Apar.bl_1 + ", " + "Брой полюси: " + Apar.bl_2 + ","
                            Tekst2 = Tekst2 + vbCrLf + "Номинален ток: " + Apar.bl_3
                        Case "Vigi NG125", "Vigi C120", "Vigi iC60", "EZ9 RCCB"
                            Tekst1 = Tekst1 + vbCrLf + Apar.bl_1 + ", " + Apar.bl_2
                            Tekst1 = Tekst1 + vbCrLf + Apar.bl_3
                            Tekst2 = "Дефектнотокова защита,"
                            Tekst2 = Tekst2 + vbCrLf + "Тип: " + Apar.bl_1 + ", " + "Брой полюси: " + Apar.bl_2 + ","
                            Tekst2 = Tekst2 + vbCrLf + "Номинален ток: " + Apar.bl_3
                        Case "Vigi NSX400/630", "Vigi NSX250", "Vigi NSX100/160"
                            Tekst1 = Tekst1 + vbCrLf + Apar.bl_1 + ", " + Apar.bl_2
                            Tekst1 = Tekst1 + vbCrLf + Apar.bl_3
                            Tekst2 = "Дефектнотокова защита,"
                            Tekst2 = Tekst2 + vbCrLf + "Тип: " + Apar.bl_1 + ", " + "Брой полюси: " + Apar.bl_2 + ","
                            Tekst2 = Tekst2 + vbCrLf + "Номинален ток: " + Apar.bl_3
                        Case "PRI", "PRC"
                            Tekst1 = Tekst1
                            Tekst2 = "Катоден отводител за "
                            Tekst2 = Tekst2 + vbCrLf + "телефонна/информационна линия"
                        Case "_Тип 2 iPRD", "_Тип 2 iPF", "_Тип 1+2 PRD1", "_Тип 1+2 PRF1",
                             "_Тип 1 PRF1 Master", "_Тип 1 PRD1 Master"
                            Tekst1 = Mid(Tekst1, 2, Len(Tekst1)) + vbCrLf + Apar.bl_1 + ", " + Apar.bl_2
                            Tekst1 = Tekst1 + vbCrLf + Apar.bl_3
                            Tekst1 = Tekst1 + vbCrLf + Apar.bl_4
                            Tekst2 = "Катоден отводител,"
                            Tekst2 = Tekst2 + vbCrLf + "Тип: " + Apar.bl_1 + ", " + "Брой полюси: " + Apar.bl_2 + ","
                            Tekst2 = Tekst2 + vbCrLf + "Максимален разряден ток: " + Apar.bl_3
                            Tekst2 = Tekst2 + vbCrLf + "Работно напрежение: " + Apar.bl_4
                        Case "LC7K", "LP4K", "LC1D", "LP1D", "LC1K", "LP1K"
                            Tekst1 = Tekst1 + vbCrLf + Apar.bl_1
                            Tekst1 = Tekst1 + vbCrLf + Apar.bl_2
                            Tekst1 = Tekst1 + vbCrLf + Apar.bl_3
                            Tekst1 = Tekst1 + vbCrLf + Apar.bl_4
                            Tekst2 = "Контактор,"
                            Tekst2 = Tekst2 + vbCrLf + "Тип на конактите: " + Apar.bl_1 + ", "
                            Tekst2 = Tekst2 + vbCrLf + "Помощни контакти: " + Apar.bl_2 + ","
                            Tekst2 = Tekst2 + vbCrLf + "Номинален ток: " + Apar.bl_3
                            Tekst2 = Tekst2 + vbCrLf + "Управ. напрежение: " + Apar.bl_4
                        Case "LP5K", "LC8K", "LP2K", "LC2D", "LC2K"
                            Tekst1 = Tekst1 + vbCrLf + Apar.bl_1
                            Tekst1 = Tekst1 + vbCrLf + Apar.bl_2
                            Tekst1 = Tekst1 + vbCrLf + Apar.bl_3
                            Tekst1 = Tekst1 + vbCrLf + Apar.bl_4
                            Tekst2 = "Реверсивен контактор,"
                            Tekst2 = Tekst2 + vbCrLf + "Тип на конактите: " + Apar.bl_1 + ", "
                            Tekst2 = Tekst2 + vbCrLf + "Помощни контакти: " + Apar.bl_2 + ","
                            Tekst2 = Tekst2 + vbCrLf + "Номинален ток: " + Apar.bl_3
                            Tekst2 = Tekst2 + vbCrLf + "Управ. напрежение: " + Apar.bl_4
                        Case "iCT"
                            Tekst1 = Tekst1 + vbCrLf + Apar.bl_1
                            Tekst1 = Tekst1 + vbCrLf + Apar.bl_3
                            Tekst1 = Tekst1 + vbCrLf + Apar.bl_4
                            Tekst2 = "Модулен контактор,"
                            Tekst2 = Tekst2 + vbCrLf + "Тип на конактите: " + Apar.bl_1 + ","
                            Tekst2 = Tekst2 + vbCrLf + "Номинален ток: " + Apar.bl_3
                            Tekst2 = Tekst2 + vbCrLf + "Управ. напрежение: " + Apar.bl_4
                        Case "iSW", "INS", "IN", "INV", "ING125", "NS", "NT", "NW"
                            Tekst1 = Tekst1 + vbCrLf + Apar.bl_2
                            Tekst1 = Tekst1 + vbCrLf + Apar.bl_3
                            Tekst2 = "Мощностен разединител,"
                            Tekst2 = Tekst2 + vbCrLf + "Брой полюси: " + Apar.bl_2 + ","
                            Tekst2 = Tekst2 + vbCrLf + "Номинален ток: " + Apar.bl_3
                        Case "NG125"
                            If Apar.bl_ИмеБлок = "s_ng125_circ_break" Then
                                Tekst1 = Tekst1 + " " + Apar.bl_1
                                Tekst1 = Tekst1 + vbCrLf + Apar.bl_2 + ", " + Apar.bl_3
                                Tekst1 = Tekst1 + vbCrLf + Apar.bl_4
                                Tekst2 = "Автоматичен прекъсвач, Крива: " + Apar.bl_2 + ","
                                Tekst2 = Tekst2 + vbCrLf + "Изкл. способност: " + Apar.bl_1 + ", " + "Брой полюси: " + Apar.bl_3 + ","
                                Tekst2 = Tekst2 + vbCrLf + "Номинален ток: " + Apar.bl_4
                            End If
                            If Apar.bl_ИмеБлок = "s_i_ng_switch_disconn" Then
                                Tekst1 = Tekst1 + vbCrLf + Apar.bl_2
                                Tekst1 = Tekst1 + vbCrLf + Apar.bl_3
                                Tekst2 = "Мощностен разединител,"
                                Tekst2 = Tekst2 + vbCrLf + "Брой полюси: " + Apar.bl_2 + ","
                                Tekst2 = Tekst2 + vbCrLf + "Номинален ток: " + Apar.bl_3
                            End If
                        Case "IC 100к", "IC Astro", "IC2000P+", "IC2000", "IC100"
                            Tekst1 = Tekst1 + vbCrLf + Apar.bl_2
                            Tekst1 = Tekst1 + vbCrLf + Apar.bl_3
                            Tekst2 = "Фотореле,"
                            Tekst2 = Tekst2 + vbCrLf + "Чувствителност, Lx: " + Apar.bl_2 + ","
                            Tekst2 = Tekst2 + vbCrLf + "Номинален ток: " + Apar.bl_3
                        Case "IHP"
                            Tekst1 = Tekst1 + vbCrLf + Apar.bl_1
                            Tekst1 = Tekst1 + vbCrLf + Apar.bl_4
                            Tekst2 = "Програмируемо реле за време,"
                            Tekst2 = Tekst2 + vbCrLf + "Тип: " + Apar.bl_1 + ", "
                            Tekst2 = Tekst2 + "Канали: " + Apar.bl_2 + ", "
                            Tekst2 = Tekst2 + vbCrLf + "Номинален ток: " + Apar.bl_4
                        Case "IH"
                            Tekst1 = Tekst1 + vbCrLf + Apar.bl_1
                            Tekst1 = Tekst1 + vbCrLf + Apar.bl_4
                            Tekst2 = "Програмируемо реле за време,"
                            Tekst2 = Tekst2 + vbCrLf + "Тип: " + Apar.bl_1 + ", "
                            Tekst2 = Tekst2 + "Канали: " + Apar.bl_2 + ", "
                            Tekst2 = Tekst2 + vbCrLf + "Номинален ток: " + Apar.bl_4
                        Case "iTL"
                            Tekst1 = Tekst1 + vbCrLf + Apar.bl_2
                            Tekst1 = Tekst1 + vbCrLf + Apar.bl_3
                            Tekst1 = Tekst1 + vbCrLf + Apar.bl_4
                            Tekst2 = "Импулсно реле,"
                            Tekst2 = Tekst2 + vbCrLf + "Брой полюси: " + Apar.bl_2 + ","
                            Tekst2 = Tekst2 + vbCrLf + "Номинален ток: " + Apar.bl_3 + ","
                            Tekst2 = Tekst2 + vbCrLf + "Управл. напрежение: " + Apar.bl_4
                        Case "MINp", "MIN"
                            Tekst1 = Tekst1 + vbCrLf + Apar.bl_3
                            Tekst1 = Tekst1 + vbCrLf + Apar.bl_4
                            Tekst2 = "Стълбищен автомат,"
                            Tekst2 += vbCrLf + "Времезакъснение: " + Apar.bl_3 + ","
                            Tekst2 += vbCrLf + "Номинален ток: " + Apar.bl_4
                        Case "C60H-DC"
                            Tekst1 = Tekst1 + vbCrLf + Apar.bl_2 + ", " + Apar.bl_3
                            Tekst1 = Tekst1 + vbCrLf + Apar.bl_4
                            Tekst2 = "Автоматичен прекъсвач за постоянен ток,"
                            Tekst2 += vbCrLf + "Крива: " + Apar.bl_2 + ", " + "Брой полюси: " + Apar.bl_3 + ","
                            Tekst2 += vbCrLf + "Номинален ток: " + Apar.bl_4
                        Case "GZ1", "GV2-ME", "GV3-P", "GV4P"
                            Tekst1 += vbCrLf + Apar.bl_1
                            Tekst1 += vbCrLf + Apar.bl_2
                            Tekst1 += vbCrLf + Apar.bl_3
                            Tekst2 = "Термомагнитен моторен прекъсвач "
                            Tekst2 += vbCrLf + "Pдвиг(400V): " + Apar.bl_1
                            Tekst2 += vbCrLf + "Брой полюси: " + Apar.bl_2
                            Tekst2 += vbCrLf + "Обхват защита: " + Apar.bl_3
                        Case "MTZ2 40", "MTZ2"
                            Tekst1 += vbCrLf + Apar.bl_1
                            Tekst1 += vbCrLf + Apar.bl_2
                            Tekst1 += vbCrLf + Apar.bl_3
                            Tekst2 = "Термомагнитен моторен прекъсвач "
                            Tekst2 += vbCrLf + "Pдвиг(400V): " + Apar.bl_1
                            Tekst2 += vbCrLf + "Брой полюси: " + Apar.bl_2
                            Tekst2 += vbCrLf + "Обхват защита: " + Apar.bl_3
                        Case "iEM2155", "iEM2155", "iEM3155", "iEM3250", "iEM3255"
                            Tekst1 = ""
                            Tekst1 += vbCrLf + Apar.bl_4
                            Tekst1 += vbCrLf + Apar.bl_2
                            Tekst2 = "Електромер "
                            Tekst2 += vbCrLf + "Напрежение: " + Apar.bl_4
                            Tekst2 += vbCrLf + "Брой полюси: " + Apar.bl_2
                        Case Else
                            Tekst1 = "Непознат елемент"
                    End Select

                    ' Изчисляваме броя на допълнителните редове над 3 за текущия елемент
                    ' Това гарантира, че дори последният елемент ще отчита редовете над 3
                    Dim additionalRows As Integer = CountLongLinesMax(Tekst1, Tekst2)

                    ' Изчисляваме текущата колона и ред
                    row = index Mod rows        ' редът в текущата колона
                    col = index \ rows          ' колоната, в която се намираме

                    ' Изместване по X и Y за позициониране на текста
                    offset_X = col * offset_Колона
                    offset_Y = row * (52.5 + 20) + extraRows * 20  ' extraRows се добавя за всички редове над 3

                    ' Текуща позиция за MText
                    Dim currentPoint As New Point3d(insPointText1.X + offset_X, insPointText1.Y - offset_Y, 0)

                    ' Вмъкваме Tekst1, тирето и Tekst2 в AutoCAD чрез твоята помощна функция
                    cu.InsertMText(Tekst1, currentPoint, "EL__DIM", 12, TextWidth:=115)
                    cu.InsertMText("-", New Point3d(currentPoint.X + 115, currentPoint.Y, 0), "EL__DIM", 12, TextWidth:=5)
                    cu.InsertMText(Tekst2, New Point3d(currentPoint.X + 125, currentPoint.Y, 0), "EL__DIM", 12, TextWidth:=370)

                    ' Добавяме допълнителните редове към общото за следващия елемент в същата колона
                    extraRows += additionalRows

                    ' Проверка дали следващият елемент ще бъде в нова колона
                    ' Ако е нова колона → нулираме extraRows, за да не се пренасят допълнителните редове
                    Dim nextCol As Integer = (index + 1) \ rows
                    If nextCol <> col Then extraRows = 0

                    ' Увеличаваме индекса за следващия елемент
                    index += 1
                Next
                acTrans.Commit()
            Catch ex As Exception
                MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
                acTrans.Abort()
            End Try
        End Using
        If Not Tablo_Name Then
            Dim текстове As String = "EZ9E108P2S, MIP12104, PRA20113, PRA16113, NSYCRN22150"
            For Each текст In текстове.Split(",")
                If Tablo_Text.StartsWith(текст.Substring(0, 3)) Then
                    MsgBox(String.Format("Името на табло е :'{0}'", текст), MsgBoxStyle.Information, "Името на табло не ми харесва!")
                    Exit For
                End If
            Next текст
        Else
            If Not Tablo_Layer Then
                MsgBox("Генерирам легенда за табло. Няма маркиран блок за табло!" + vbCrLf + "Провери дали има блок за табло и дали е в правилния слой!")
            End If
        End If
    End Sub
    ''' <summary>
    ''' Връща броя на редовете над 3 от текста с повече редове.
    ''' Ако и двата текста имат 3 или по-малко реда, връща 0.
    ''' </summary>
    ''' <param name="text1">Първи текст</param>
    ''' <param name="text2">Втори текст</param>
    ''' <returns>Брой редове над 3 от текста с повече редове</returns>
    Function CountLongLinesMax(text1 As String, text2 As String) As Integer
        ' Броим редовете за всеки текст
        Dim count1 As Integer = If(String.IsNullOrEmpty(text1), 0, text1.Split({vbCrLf, vbCr, vbLf}, StringSplitOptions.None).Length)
        Dim count2 As Integer = If(String.IsNullOrEmpty(text2), 0, text2.Split({vbCrLf, vbCr, vbLf}, StringSplitOptions.None).Length)

        ' Вземаме по-голямото
        Dim maxCount As Integer = Math.Max(count1, count2)

        ' Ако е <= 3, връщаме 0
        If maxCount <= 3 Then
            Return 0
        End If

        ' Връщаме броя на редовете над 3
        Return maxCount - 3
    End Function
    Public Function BlockText(blName As String, Visibility As String) As String
        Dim strBlockText As String = ""
        Select Case blName
            Case "Авария", "Авария_100"
                strBlockText = "Светодиодно осветително тяло за евакуационно осветление, с вградени акумулаторни батерии, 4W"
            Case "LED_DENIMA"
                strBlockText = "Светодиодно осветително тяло"
            Case "LED_ULTRALUX", "LED_ULTRALUX_100"
                strBlockText = "Светодиодно осветително тяло"
                Select Case Visibility
                    Case "Кухня"
                        strBlockText = strBlockText + " - за вграждане в кухненска мебел"
                        Exit Select
                    Case "IP-20"
                        'strBlockText = strBlockText + " - панел, IP 44"
                    Case "IP-54", "IP-65", "IP-66"
                        'strBlockText = strBlockText + " - промишлен тип"
                    Case Else
                        strBlockText = "################################"
                End Select
                strBlockText += vbLf + "- мощност - " + strLamp_Power + "W;"
                strBlockText += vbLf + "- монтаж - " + strLED_Lamp_Montav + ";"
                strBlockText += vbLf + "- светлинен поток - " + strLED_Lamp_SW_Potok + ";"
                strBlockText += vbLf + "- степен на защита - " + strLED_Lamp_TIP + ";"
            Case "Луминисцентна лампа"
                strBlockText = "Луминисцентна лампа"
            Case "LED_lenta"
                strBlockText = "Едноцветна светодиодна лента, 12V"
            Case "Бойлер"
                Select Case Visibility
                    Case "Изход 1p"
                        strBlockText = "Еднофазен свободен извод за присъединяване на съоръжение"
                    Case "Изход 3p"
                        strBlockText = "Трифазен свободен извод за присъединяване на съоръжение"
                    Case "ПВ"
                        strBlockText = "Пускател въздушен"
                End Select
            Case "Контакт"
                Select Case Visibility
                    Case "Обикновен"
                        strBlockText = "Контакт " + Chr(34) + "шуко" + Chr(34)
                    Case "Обикновен - противовлажен"
                        strBlockText = "Контакт " + Chr(34) + "шуко" + Chr(34) + " IP 44"
                    Case "Двугнездов"
                        strBlockText = "Контакт " + Chr(34) + "шуко" + Chr(34) + " - двугнездов"
                    Case "Двугнездов - противовлажен"
                        strBlockText = "Контакт " + Chr(34) + "шуко" + Chr(34) + " - двугнездов, IP44"
                    Case "Тригнездов"
                        strBlockText = "Контакт " + Chr(34) + "шуко" + Chr(34) + " - тригнездов"
                    Case "Тригнездов - противовлажен"
                        strBlockText = "Контакт " + Chr(34) + "шуко" + Chr(34) + " - тригнездов, IP44"
                    Case "Усилен"
                        strBlockText = "Контакт усилен - 25А"
                    Case "Твърда връзка"
                        strBlockText = "Контакт - твърда връзка, скрит монтаж"
                    Case "Монифазен - IP 54"
                        strBlockText = "Евроконтакт, монофазен, 1Р+N+PE, IP54, 16А"
                    Case "Евроамерикански стандарт"
                        strBlockText = "Контакт евроамерикански стандарт"
                    Case "С детска защита"
                        strBlockText = "Контакт " + Chr(34) + "шуко" + Chr(34) + " - детска защита"
                    Case "С детска защита - противовлажен"
                        strBlockText = "Контакт " + Chr(34) + "шуко" + Chr(34) + " - детска защита, IP44"
                    Case "За монтаж в канал"
                        strBlockText = "Контакт " + Chr(34) + "шуко" + Chr(34) + " - за монтаж в кабелен канал"
                    Case "Трифазен"
                        strBlockText = "Контакт трифазен"
                    Case "Трифазен - IP 54"
                        strBlockText = "Евроконтакт, трифазен, 3Р+N+PE, IP54, 25А"
                    Case "Трифазен - противовлажен"
                        strBlockText = "Контакт трифазен - IP44."
                    Case "1xU"
                        strBlockText = "Контакт " + Chr(34) + "шуко" + Chr(34) + " с един USB изход"
                    Case "2xU"
                        strBlockText = "Контакт " + Chr(34) + "шуко" + Chr(34) + " двугнездов с USB изход"
                    Case "ТР+2МФ"
                        strBlockText = "Контакт трифазен модул, трифазен + 2 монофазни контакта, открит монтаж"
                    Case Else
                        strBlockText = "################################"
                End Select
            Case "LED_луна"
                Select Case Visibility
                    Case "Лед луна"
                        strBlockText = "Светодиодно осветително тяло-луна, 12V, IP-21"
                    Case "Лед луна противовлажна"
                        strBlockText = "Светодиодно осветително тяло-луна, 12V, IP-54"
                    Case "Драйвер"
                        strBlockText = "Драйвер за светодиодна лента, 220/12V"
                    Case Else
                        strBlockText = "################################"
                End Select
            Case "Линия МХЛ - 220V"
                Select Case Visibility
                    Case "1x26-с решетка"
                        strBlockText = "Светодиодно осветително - 220V, модел по избор на архитекта"
                    Case "1x26 - без решетка"
                        strBlockText = "Светодиодно осветително тяло кръгло - 220V, IP-21"
                    Case "1х26 - квадрат"
                        strBlockText = "Светодиодно осветително тяло квадратно - 220V, IP-21"
                    Case "1х26 - квадрат датчик"
                        strBlockText = "Светодиодно осветително тяло квадратно - 220V, с датчик за движение; IP-21"
                    Case "1х26 - IP 54"
                        strBlockText = "Светодиодно осветително тяло кръгло - 220V, IP-54"
                    Case "1х26 - квадрат IP 54"
                        strBlockText = "Светодиодно осветително тяло квадратно - 220V, IP-54"
                    Case "1х26 - квадрат IP 54 датчик"
                        strBlockText = "Светодиодно осветително тяло квадратно - 220V, с датчик за движение; IP-54"
                    Case Else
                        strBlockText = "################################"
                End Select
            Case "Плафони"
                Select Case Visibility
                    Case "Общо означение"
                        strBlockText = "Осветително тяло"
                    Case "Датчик насочен"
                        strBlockText = "Датчик за чувствителен към движение - насочен"
                    Case "Датчик 360°"
                        strBlockText = "Датчик за чувствителен към движение - 360°"
                    Case "Фотодатчик"
                        strBlockText = "Датчик за чувствителен към светлина /фотодачик/"
                    Case "Аплик", "Аплик - Рошав"
                        strBlockText = "Светодиодно осветително тяло - аплик"
                    Case "Аплик с датчик", "Аплик - Рошав с датчик"
                        strBlockText = "Светодиодно осветително тяло - аплик с датчик за движение"
                    Case "Аплик - противовлажен", "Аплик - Рошав - противовлажен"
                        strBlockText = "Светодиодно осветително тяло - аплик, противовлажен"
                    Case "Аплик - противовлажен с датчик", "Аплик - Рошав - противовлажен с датчик"
                        strBlockText = "Светодиодно осветително тяло - аплик, противовлажен с датчик за движение"
                    Case "Плафон"
                        strBlockText = "Светодиодно осветително тяло - плафон"
                    Case "Плафон с датчик"
                        strBlockText = "Светодиодно осветително тяло - плафон с датчик за движение"
                    Case "Плафон - противовлажен"
                        strBlockText = "Светодиодно осветително тяло - плафон, противовлажен"
                    Case "Плафон - противовлажен с датчик"
                        strBlockText = "Светодиодно осветително тяло - плафон, противовлажен с датчик за движение"
                    Case "Пендел"
                        strBlockText = "Светодиодно осветително тяло - пендел"
                    Case "Пендел с датчик"
                        strBlockText = "Светодиодно осветително тяло - пендел с датчик за движени"
                    Case "Пендел - противовлажен"
                        strBlockText = "Светодиодно осветително тяло - пендел, противовлажен"
                    Case "Пендел - противовлажен с датчик"
                        strBlockText = "Светодиодно осветително тяло - пендел, противовлажен с датчик за движение"
                    Case "Бански аплик", "Бански аплик ЛЕД"
                        strBlockText = "Бански аплик LED - модел по избор на архитекта, IP54"
                    Case "Настолна лампа", "Настолна лампа - рошава"
                        strBlockText = "Настолна LED лампа - модел по избор на архитекта"
                    Case "Лампион", "Лампион - рошав"
                        strBlockText = "Лампион LED - модел по избор на архитекта."
                    Case "Фасадно"
                        strBlockText = "Фасадно осветително тяло"
                    Case "Фасадно с датчик"
                        strBlockText = "Фасадно осветително тяло с датчик за движение"
                    Case Else
                        strBlockText = "################################"
                End Select
            Case "Металхаогенна лампа"
                Select Case Visibility
                    Case "1х35 - Дъга", "1х35 - Кръг", "1х35 - Право"
                        strBlockText = "Светодиодно осветително тяло със спот-1 тяло модел по избор на архутекта"
                    Case "2х35 - Дъга", "2х35 - Кръг", "2х35 - Право"
                        strBlockText = "Светодиодно осветително тяло със спотове-2 тела модел по избор на архутекта"
                    Case "3х35 - Дъга", "3х35 - Кръг", "3х35 - Право"
                        strBlockText = "Светодиодно осветително тяло със спотове-3 тела модел по избор на архутекта"
                    Case "4х35 - Дъга", "4х35 - Кръг", "4х35 - Право"
                        strBlockText = "Светодиодно осветително тяло със спотове-4 тела модел по избор на архутекта"
                    Case "1х35 - 90°"
                        strBlockText = "Светодиодно осветително тяло със спот-1 тяло за монтаж на стена модел по избор на архутекта"
                    Case "1х35 - за картина", "2х35 - за картина"
                        strBlockText = "Светодиодно осветително тяло със спотове-за осветяване на картина"
                    Case Else
                        strBlockText = "################################"
                End Select
            Case "Прожектор"
                Select Case Visibility
                    Case "МХЛ", "МХЛ - кръгла"
                        strBlockText = "Светодиодно осветително тяло - за фасадно осветление, комплект с конзола, IP-65"
                    Case "МЛХ - с датчик"
                        strBlockText = "Светодиодно осветително тяло - за фасадно осветление, комплект с конзола, IP-65, с датчик за чувствителен към движение"
                    Case Else
                        strBlockText = "################################"
                End Select
            Case "бойлерно табло"
                Select Case Visibility
                    Case "Само ключ"
                        strBlockText = "Бойлерно табло"
                    Case "Ключ и контакт"
                        strBlockText = "Бойлерно табло с контакт " + Chr(34) + "шуко" + Chr(34)
                    Case "С два ключа и контакт"
                        strBlockText = "Бойлерно табло с два ключа и контакт " + Chr(34) + "шуко" + Chr(34)
                    Case "С два контакта и един ключ"
                        strBlockText = "Бойлерно табло с един ключ и два контакта " + Chr(34) + "шуко" + Chr(34)
                    Case Else
                        strBlockText = "################################"
                End Select
            Case "Ключ_знак"
                Select Case Visibility
                    Case "Еднопозиционен"
                        strBlockText = "Ключ обикновен"
                    Case "Двупозиционен"
                        strBlockText = "Ключ сериен"
                    Case "Трипозиционен"
                        strBlockText = "Ключ троен"
                    Case "Деветор"
                        strBlockText = "Ключ девиаторен"
                    Case "Кръстат"
                        strBlockText = "Ключ кръстат"
                    Case "Еднопозиционен - противовлажен"
                        strBlockText = "Ключ обикновен, противовлажен"
                    Case "Двупозиционен - противовлажен"
                        strBlockText = "Ключ сериен, противовлажен"
                    Case "Трипозиционен - противовлажен"
                        strBlockText = "Ключ троен, противовлажен"
                    Case "Девятор - противовлажен"
                        strBlockText = "Ключ девиаторен, противовлажен"
                    Case "Кръстат - противовлажен"
                        strBlockText = "Ключ кръстат, противовлажен"
                    Case "Еднопозиционен - светещ"
                        strBlockText = "Ключ обикновен, светещ"
                    Case "Двупозиционен - светещ"
                        strBlockText = "Ключ сериен, светещ"
                    Case "Трипозиционен - светещ"
                        strBlockText = "Ключ троен, светещ"
                    Case "Девятор светещ"
                        strBlockText = "Ключ девиаторен, светещ"
                    Case "Кръстат светещ"
                        strBlockText = "Ключ кръстат, светещ"
                    Case Else
                        strBlockText = "################################"
                End Select
            Case "Ключ_знак_WIFI"
                Select Case Visibility
                    Case "Еднопозиционен - Радио"
                        strBlockText = "Едноканален ключ с RF дистанционно и ръчно управление"
                    Case "Двупозиционен - Радио"
                        strBlockText = "Двуканален ключ с RF дистанционно и ръчно управление"
                    Case "Трипозиционен - Радио"
                        strBlockText = "Триканален ключ с RF дистанционно и ръчно управление"
                    Case "Четирипозиционен - Радио"
                        strBlockText = "Четириканален ключ с RF дистанционно и ръчно управление"
                    Case "Еднопозиционен - WiFi"
                        strBlockText = "WiFi Smart ключ с един канал с ръчно и дистанционно управление"
                    Case "Двупозиционен - WiFi"
                        strBlockText = "WiFi Smart ключ с два канала с ръчно и дистанционно управление"
                    Case "Трипозиционен - WiFi"
                        strBlockText = "WiFi Smart ключ с три канала с ръчно и дистанционно управление"
                    Case "Четирипозиционен - WiFi"
                        strBlockText = "WiFi Smart ключ с четири канала с ръчно и дистанционно управление"
                    Case "Еднопозиционен - Сенсор"
                        strBlockText = "Едноканален ключ със сензорно управление"
                    Case "Двупозиционен - Сенсор"
                        strBlockText = "Двуканален ключ със сензорно управление"
                    Case "Трипозиционен - Сенсор"
                        strBlockText = "Триканален ключ със сензорно управление"
                    Case "Четирипозиционен - Сенсор"
                        strBlockText = "Четириканален ключ със сензорно управление"
                    Case Else
                        strBlockText = "################################"
                End Select
            Case "Полилей"
                Select Case Visibility
                    Case "1х60 - Индийски"
                        strBlockText = "Плафониера с дистанционно управление"
                    Case "1х60 - Кръгла", "1х60 - Рошава"
                        strBlockText = "Полилей с едно тяло"
                    Case "2х60 - Кръгла", "2х60 - Рошава"
                        strBlockText = "Полилей с две тела"
                    Case "3х60 - Кръгла", "3х60 - Рошава"
                        strBlockText = "Полилей с три тела"
                    Case "4х60 - Кръгла", "4х60 - Рошава"
                        strBlockText = "Полилей с четири тела"
                    Case Else
                        strBlockText = "################################"
                End Select
            Case "Датчик_ПАБ"
                Select Case Visibility
                    Case "ПАБ - Димооптичен адресируем"
                        strBlockText = "Пожароизвестител адресируем; оптично-димен"
                    Case "ПАБ - Термичен адресируем - 7101"
                        strBlockText = "Пожароизвестител адресируем; топлинен; максимален"
                    Case "ПАБ - Термичен адресируем с адаптер-7120"
                        strBlockText = "Термо-диференциален пожароизвестител"
                    Case "ПАБ - Термичен адресируем диференциален"
                        strBlockText = "Термо-диференциален пожароизвестител-адресируем"
                    Case "ПАБ - Термичен адресируем комбиниран"
                        strBlockText = "Комбиниран оптично-димен и термо-диференциален пожароизвестител-адресируем"
                    Case "ПАБ - Сирена адресируема"
                        strBlockText = "Вътрешна адресируема сирена със звук и светлина"
                    Case "Ръчен пожароизвестител адресируем"
                        strBlockText = "Адресируем ръчен пожароизвестител с вграден изолатор на късо съединение"
                    Case "Изпълнително устройство"
                        strBlockText = "Адресируем входно изходен модул"
                    Case "ПАБ - Димооптичен конвенционален"
                        strBlockText = "Пожароизвестител конвенционален; оптично-димен"
                    Case "ПАБ - Термичен конвенционален"
                        strBlockText = "Пожароизвестител конвенционален; топлинен; максимален"
                    Case "ПАБ - Термичен конвенционален диференциален"
                        strBlockText = "Топлинен диференциален пожароизвестител-конвенционален"
                    Case "ПАБ - Пламъков конвенционален"
                        strBlockText = "Пожароизвестител оптичен пламъков-конвенционален"
                    Case "ПАБ - Термичен конвенционален комбиниран"
                        strBlockText = "Комбиниран оптично-димен и термо-диференциален пожароизвестител-конвенционален"
                    Case "ПАБ - Лампа"
                        strBlockText = "Паралелен сигнализатор"
                    Case "ПАБ - Сирена конвенционална"
                        strBlockText = "Конвенционална сирена за вътрешно използване"
                    Case "ПАБ - Сирена и Звук"
                        strBlockText = "Външна сирена – метален корпус, със светлина"
                    Case "Ръчен пожароизвестител конвенционален"
                        strBlockText = "Ръчен пожароизвестител - конвенционален"
                    Case "Изолатор"
                        strBlockText = "################################"
                    Case "Линеен оптично димен излъчвател"
                        strBlockText = "Линеен оптично димен пожароизвестител-излъчвател"
                    Case "Линеен оптично димен приемник"
                        strBlockText = "Линеен оптично димен пожароизвестител-приемник"
                    Case Else
                        strBlockText = "################################"
                End Select
            Case "Високоговорител"
                Select Case Visibility
                    Case "0,8W-таванен"
                        strBlockText = "Високоговорител за вграждане"
                    Case "0,8W-Насочен", "1,5W-Насочен",
                         "3W-Насочен", "6W-Насочен"
                        strBlockText = "Високоговорител за стенен монтаж"
                    Case "EOL"
                        strBlockText = "Микрофонен пулт за съобщения"
                    Case Else
                        strBlockText = "################################"
                End Select
            Case "Аудио система"
                Select Case Visibility
                    Case "Регулатор на звука"
                        strBlockText = "Микрофонен пулт за съобщения"
                    Case "Високоговорител"
                        strBlockText = "################################"
                    Case "Микрофон"
                        strBlockText = "Микрофонен пулт за съобщения"
                    Case Else
                        strBlockText = "################################"
                End Select
            Case "Камери"
                strBlockText = "Камера за видеонаблюдение"
                Select Case Visibility
                    Case "Камерa 360°"
                        strBlockText += vbLf + "куполна IP камера с моторизиран обектив"
                    Case "Насочен камера-20"
                        strBlockText += vbLf + "насочена IP камера; IP44"
                    Case "Насочен камера-IP66"
                        strBlockText += vbLf + "насочена IP камера;открит монтаж; IP67"
                    Case Else
                        strBlockText = "################################"
                End Select
            Case "СОТ"
                Select Case Visibility
                    Case "СОТ_Насочен"
                        strBlockText = "Детектор за движение"
                    Case "СОТ_Звуков"
                        strBlockText = "Акустичен датчик за стъкло"
                    Case "СОТ_МУК"
                        strBlockText = "Магнитен или механичен конткт"
                    Case "СОТ_360"
                        strBlockText = "360° таванен датчик за движение"
                    Case "СОТ_Сирена"
                        strBlockText = "Бронирана външна сирена"
                    Case "Приемник паник бутон"
                        strBlockText = "Приемник за паник бутон"
                    Case "Паник бутон"
                        strBlockText = "Паник бутон"
                    Case "Датчик каса"
                        strBlockText = "Датчик за защита на каси и трезори"
                    Case "Вибрационен"
                        strBlockText = "Вибрационен сеизмичен датчик"
                    Case "Микровълнов"
                        strBlockText = "Комбиниран инфрачервен и микровълнов датчик за движение"
                    Case "Клавиатура"
                        strBlockText = "Клавиатура за алармери системи"
                    Case Else
                        strBlockText = "################################"
                End Select
            Case "Домофон"
                Select Case Visibility
                    Case "Звънец"
                        strBlockText = "Звънец"
                    Case "Домофон"
                        strBlockText = "Силово разпределително табло"
                    Case "Брава"
                        strBlockText = "Заключващ механизъм"
                    Case "Бутон"
                        strBlockText = "Бутон за отключване"
                    Case "Табло"
                        strBlockText = "Входно табло"
                    Case "Клавиатура"
                        strBlockText = "Клавиатура"
                    Case "Карта четец"
                        strBlockText = "Карточетец за RF карти "
                    Case "Централа"
                        strBlockText = "Домофонна централа"
                    Case "Домофонна централа (контролер за достъп)"
                        strBlockText = "Контролер контрол на достъпа"
                    Case Else
                        strBlockText = "################################"
                End Select
            Case "Табло_Ново"
                Select Case Visibility
                    Case "Табло"
                        strBlockText = "Силово разпределително табло"
                    Case "Слаботоково табло"
                        strBlockText = "Слаботоково табло"
                    Case "ПИЦ"
                        strBlockText = "Пожароизвестителна централа"
                    Case "Домофонна"
                        strBlockText = "Домофонна централа"
                    Case "СОТ"
                        strBlockText = "Контролен панел на СОТ"
                    Case "Контрол"
                        strBlockText = "Централен контролер за контрол на достъпа"
                    Case Else
                        strBlockText = "################################"
                End Select
            Case "Розетка_1"
                Select Case Visibility
                    Case "TV"
                        strBlockText = "Телевизионна розетка"
                    Case "T"
                        strBlockText = "Телефонна розетка"
                    Case "@"
                        strBlockText = "Розетка RJ 45, категория 5е"
                    Case "HDMI"
                        strBlockText = "HDMI розетка"
                    Case "USB"
                        strBlockText = "USB розетка"
                    Case "WiFi"
                        strBlockText = "Рутер (с Wi-Fi)"
                    Case "Router"
                        strBlockText = "Рутер (без Wi-Fi)"
                    Case "Access point"
                        strBlockText = "Wi-Fi точка на достъп (AP)"
                    Case Else
                        strBlockText = "################################"
                End Select
            Case "Ключ_квадрат"
                Select Case Visibility
                    Case "Ключ управление"
                        strBlockText = "Ключ управление"
                    Case "Лихт бутон единичен", "Лихт бутон единичен светещ"
                        strBlockText = "Електрически ключ лихт бутон"
                    Case "Лихт бутон двоен", "Лихт бутон двоен светещ"
                        strBlockText = "Електрически ключ лихт бутон двоен"
                    Case "Лихт бутон троен", "Лихт бутон троен светещ"
                        strBlockText = "Електрически ключ лихт бутон троен"
                    Case "Чипкарта", "Чипкарта светещ"
                        strBlockText = "Електрически ключ за карта"
                    Case "Стълбищен бутон", "Стълбищен бутон светещ"
                        strBlockText = "Стълбищен електрически ключ (лихт бутон)"
                    Case "Звънец"
                        strBlockText = "Бутон за звънец"
                    Case "Звънец светещ"
                        strBlockText = "Бутон за звънец светещ"
                    Case "Сензор"
                        strBlockText = "Датчик за движение с инфрачервена детекция за вграждане в стенна конзола"
                    Case "Димер_обикновен", "Димер_сензорен"
                        strBlockText = "Ротационен димер, 10A, 230VAC"
                    Case "Щори"
                        strBlockText = "Електрически бутон за ролетни щори"
                    Case "С въженце"
                        strBlockText = "Електрически ключ бутон, начин на превключване чрез въже (корда)"
                    Case "Завеси"
                        strBlockText = "Електрически бутон за ролетни щори със стоп бутон"
                    Case "Регулатор температура"
                        strBlockText = "Терморегулатор"
                    Case "ДКУ"
                        strBlockText = "Двубутонна кнопка за управление /ДКУ/"
                    Case Else
                        strBlockText = "################################"
                End Select
            Case "Вентилации"
                Select Case Visibility
                    Case "Вентилатор - кръг - баня",
                         "Вентилатор - кръг",
                         "Вентилатор - правоъг"
                        strBlockText = "Вентилатор за баня с конрол на влажността"
                End Select
            Case Else
                strBlockText = "################################"
        End Select
        Return strBlockText
    End Function
    Public Sub SetAttribute(ByVal BlockID As ObjectId, ByVal blckname As String, ByVal AttTag As String, ByVal AttVal As String)
        Dim MyDb As Database = Application.DocumentManager.MdiActiveDocument.Database
        If BlockID.IsNull Then Exit Sub
        Try
            Using myTrans As Transaction = MyDb.TransactionManager.StartTransaction
                Dim myBlckRef As BlockReference
                Dim myAttColl As AttributeCollection
                Dim myBlckTable As BlockTableRecord
                myBlckRef = BlockID.GetObject(OpenMode.ForWrite)
                If myBlckRef.IsDynamicBlock Then
                    myBlckTable = myTrans.GetObject(myBlckRef.DynamicBlockTableRecord, OpenMode.ForRead)
                Else
                    myBlckTable = myTrans.GetObject(myBlckRef.BlockTableRecord, OpenMode.ForRead)
                End If

                If String.Compare(myBlckTable.Name, blckname, True) = 0 Then
                    myAttColl = myBlckRef.AttributeCollection
                    Dim myEnt As ObjectId
                    Dim myAttRef As AttributeReference
                    For Each myEnt In myAttColl
                        myAttRef = myEnt.GetObject(OpenMode.ForWrite)
                        If String.Compare(myAttRef.Tag, AttTag, True) = 0 Then
                            myAttRef.TextString = AttVal.ToString
                        End If
                    Next
                End If
                myTrans.Commit()
            End Using
        Catch ex As Exception
        End Try
    End Sub
    <CommandMethod("WriteToNOD")>
    Public Sub WriteToNOD()
        Dim db As Database = New Database()
        Try
            db.ReadDwgFile("C:\Temp\Test.dwg", FileShare.ReadWrite, False, Nothing)
            Using trans As Transaction = db.TransactionManager.StartTransaction()
                Dim nod As DBDictionary = CType(trans.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForWrite), DBDictionary)
                Dim myXrecord As Xrecord = New Xrecord()
                myXrecord.Data = New ResultBuffer(New TypedValue(CInt(DxfCode.Int16), 1234), New TypedValue(CInt(DxfCode.Text), "This drawing has been processed"))
                nod.SetAt("MyData", myXrecord)
                trans.AddNewlyCreatedDBObject(myXrecord, True)
                Dim myDataId As ObjectId = nod.GetAt("MyData")
                Dim readBack As Xrecord = CType(trans.GetObject(myDataId, OpenMode.ForRead), Xrecord)
                For Each value As TypedValue In readBack.Data
                    Debug.Print("===== OUR DATA: " & value.TypeCode.ToString() & ". " + value.Value.ToString())
                Next
                trans.Commit()
            End Using
            db.SaveAs("C:\Temp\Test.dwg", DwgVersion.Current)
        Catch e As Exception
            System.Diagnostics.Debug.Print(e.ToString())
        Finally
            db.Dispose()
        End Try
    End Sub
End Class