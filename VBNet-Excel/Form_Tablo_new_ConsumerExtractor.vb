Imports System.IO
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Public Class Form_Tablo_new_ConsumerExtractor
    ''' <summary>
    ''' Стъпка 2: Основен метод за улавяне на селекцията и стартиране на процеса.
    ''' </summary>
    Public Shared Function ExtractSelectedConsumers() As List(Of strKonsumator)
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        If acDoc Is Nothing Then Return Nothing
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        ' Улавяне на предварителната селекция
        Dim selRes As PromptSelectionResult = ed.SelectImplied()
        If selRes.Status <> PromptStatus.OK OrElse selRes.Value Is Nothing OrElse selRes.Value.Count = 0 Then
            Return Nothing
        End If
        ' Взимане на името на активния чертеж
        Dim sourceDwgName As String = Path.GetFileName(acCurDb.Filename)
        ' Извикваме процедурата за филтриране и обработка
        Return FilterAndProcessSelection(selRes.Value, acCurDb)
    End Function
    ''' <summary>
    ''' Процедура за филтриране по тип (INSERT), пространство (ModelSpace) и слой (EL*).
    ''' </summary>
    Private Shared Function FilterAndProcessSelection(ByVal selectedSet As SelectionSet, ByVal acCurDb As Database) As List(Of strKonsumator)
        Dim consumersList As New List(Of strKonsumator)()
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                Dim modelSpaceId As ObjectId = SymbolUtilityServices.GetBlockModelSpaceId(acCurDb)
                For Each sObj As SelectedObject In selectedSet
                    If sObj Is Nothing Then Continue For
                    Dim dbObj As DBObject = acTrans.GetObject(sObj.ObjectId, OpenMode.ForRead)
                    ' ФИЛТЪР 1: Дали е Блок
                    Dim acBlkRef As BlockReference = TryCast(dbObj, BlockReference)
                    If acBlkRef Is Nothing Then Continue For
                    ' ФИЛТЪР 2: Дали е в ModelSpace
                    If acBlkRef.BlockId <> modelSpaceId Then Continue For
                    ' ФИЛТЪР 3: Дали слоят започва с "EL"
                    Dim layerName As String = acBlkRef.Layer
                    If Not layerName.StartsWith("EL", StringComparison.OrdinalIgnoreCase) Then Continue For
                    ' ОБРАБОТКА: Ако премине филтрите, предаваме блока на отделната процедура за четене
                    Dim consumer As strKonsumator = ExtractConsumerFromBlock(acBlkRef, acTrans)
                    ' Ако блокът е валиден консуматор, го добавяме в списъка
                    If consumer IsNot Nothing Then consumersList.Add(consumer)
                Next
                acTrans.Commit()
            Catch ex As Exception
                MsgBox("Възникна грешка при филтрирането: " & ex.Message & vbCrLf & ex.StackTrace)
                acTrans.Abort()
                Return Nothing
            End Try
        End Using
        Return consumersList
    End Function
    ''' <summary>
    ''' Извлича данни от конкретен блок, чете динамични свойства, атрибути и филтрира системни обекти.
    ''' </summary>
    Private Shared Function ExtractConsumerFromBlock(ByVal acBlkRef As BlockReference, ByVal acTrans As Transaction) As strKonsumator
        ' 1) Вземане на динамичните свойства на блока
        Dim props As DynamicBlockReferencePropertyCollection = acBlkRef.DynamicBlockReferencePropertyCollection
        Dim currentVisibility As String = ""
        Dim Dylvina_Led As Double = 0
        For Each prop As DynamicBlockReferenceProperty In props
            Select Case prop.PropertyName
                Case "Visibility", "Visibility1"
                    currentVisibility = Convert.ToString(prop.Value)
                Case "Дължина" ' Използва се при LED линиите
                    If prop.Value IsNot Nothing AndAlso IsNumeric(prop.Value) Then
                        Dylvina_Led = Convert.ToDouble(prop.Value)
                    End If
            End Select
        Next
        ' 2) Филтриране според Visibility състоянието (прескачаме системни/управляващи блокове)
        Select Case currentVisibility
            Case "Само ключ", "текст",
                 "Лампион - рошав", "Лампион", "Настолна лампа - рошава",
                 "Настолна лампа", "Фотодатчик", "Датчик 360°", "Датчик насочен",
                 "Драйвер", "ПВ", "Линии", "Само текст", "Табло_Ново"
                ' Връщаме Nothing, за да укажем на главния цикъл, че този блок се прескача
                Return Nothing
        End Select
        ' 3) Инициализиране на обекта и попълване на основните AutoCAD идентификатори
        Dim Kons As New strKonsumator()
        Kons.Visibility = currentVisibility
        ' Забележка: В твоята структура е записано като Дължина_Led, напасваме го тук без да променяме нищо друго
        Kons.Дължина_Led = Dylvina_Led
        ' Вземане на истинското име на блока (поддържа динамични и анонимни блокове)
        Dim nameBlock As String = CType(acTrans.GetObject(acBlkRef.DynamicBlockTableRecord, OpenMode.ForRead), BlockTableRecord).Name
        Kons.Name = nameBlock
        ' Записване на ObjectId за текущата сесия и Handle за дългосрочно съхранение
        Kons.ID_Block = acBlkRef.ObjectId
        Kons.Handle_Block = acBlkRef.ObjectId.Handle.ToString()
        ' 4) Четене на всички атрибути, прикачени към конкретния блок
        Dim attCol As AttributeCollection = acBlkRef.AttributeCollection
        For Each attId As ObjectId In attCol
            Dim dbObj As DBObject = acTrans.GetObject(attId, OpenMode.ForRead)
            Dim acAttRef As AttributeReference = TryCast(dbObj, AttributeReference)
            If acAttRef IsNot Nothing Then
                ' Проверка на точното име на Tag на атрибута и записване (Case-Insensitive за сигурност)
                Select Case acAttRef.Tag
                    Case "ТАБЛО"
                        Kons.ТАБЛО = acAttRef.TextString
                    Case "КРЪГ"
                        Kons.ТоковКръг = acAttRef.TextString
                    Case "Pewdn"
                        Kons.PEWDN = acAttRef.TextString
                    Case "PEWDN1"
                        Kons.PEWDN1 = acAttRef.TextString
                    Case "LED", "МОЩНОСТ"
                        Kons.strМОЩНОСТ = acAttRef.TextString
                End Select
            End If
        Next
        ' 5) Валидация: Ако блокът изобщо няма атрибут за мощност или той е празен, го пропускаме
        If String.IsNullOrWhiteSpace(Kons.strМОЩНОСТ) Then Return Nothing
        ' 6) Изчисляване на мощността чрез твоята функция CalcPower
        Kons.doubМОЩНОСТ = CalcPower(Kons.strМОЩНОСТ, Dylvina_Led)
        ' 7) Допълнителна обработка според типа на блока чрез твоята функция ProcessBlockByType
        ProcessBlockByType(Kons, Kons.Name, Kons.Visibility)
        ' 8) Финална проверка за валидност: връщаме обекта само ако има реална мощност
        If Kons.doubМОЩНОСТ > 0 Then
            Return Kons
        Else
            Return Nothing
        End If
    End Function
    ''' <summary>
    ''' Обработва блока според неговото име и Visibility свойство
    ''' Определя типа и брой фази (1 или 3)
    ''' НЕ определя брой лампи/контакти и НЕ ползва мощност за фази
    ''' </summary>
    Private Shared Sub ProcessBlockByType(Kons As strKonsumator,
                                   blockName As String,
                                   visibility As String)
        ' ============================================================
        ' ПРОВЕРКА ЗА NOTHING VISIBILITY
        ' ============================================================
        Dim vis As String = ""
        If visibility IsNot Nothing Then
            vis = visibility.ToUpper()
        End If
        ' Нормализирай имената (главни букви)
        Dim name As String = blockName.ToUpper()
        ' По подразбиране - 1 фаза
        Kons.Phase = 1
        ' ============================================================
        ' SELECT CASE ПО ИМЕ НА БЛОКА
        ' ============================================================
        Select Case True
        ' ============================================================
        ' 1. БЛОКОВЕ КОИТО МОГАТ ДА СА 1P ИЛИ 3P (проверка visibility)
        ' ============================================================
        ' --- БОЙЛЕРИ ---
            Case name.Contains("БОЙЛЕР")
                Select Case True
                    Case vis.Contains("ПРОТОЧЕН - 380V"), vis.Contains("ХОРИЗОНТАЛЕН - 380V"),
                     vis.Contains("ВЕРТИКАЛЕН - 380V"), vis.Contains("Изход 3p")
                        Kons.Phase = 3
                    Case Else
                        Kons.Phase = 1
                End Select
        ' --- ВЕНТИЛАЦИИ / КЛИМАТИЦИ ---
            Case name.Contains("ВЕНТИЛАЦИИ"),
                 name.Contains("ВЕНТИЛАТОР"),
                 name.Contains("КЛИМАТИК"),
                 name.Contains("КОНВЕКТОР"),
                 name.Contains("ГОРЕЛКА"),
                 name.Contains("НАГРЕВАТЕЛ"),
                 name.Contains("ЕЛ. ЛИРА")
                Select Case True
                    Case vis.Contains("ПРОЗОРЧЕН 3P"), vis.Contains("КАНАЛЕН 3P")
                        Kons.Phase = 3
                    Case Else
                        Kons.Phase = 1
                End Select
        ' --- КОНТАКТИ ---
            Case name.Contains("КОНТАКТ")
                Select Case True
                    Case vis.Contains("ТРИФАЗЕН"), vis.Contains("ТР+2МФ"), vis.Contains("3P")
                        Kons.Phase = 3
                    Case Else
                        Kons.Phase = 1
                End Select
        ' ============================================================
        ' 2. ВСИЧКИ ОСТАНАЛИ БЛОКОВЕ - ВИНАГИ 1 ФАЗА
        ' ============================================================
            Case name.Contains("LED_DENIMA"), name.Contains("LED_LENTA"),
                 name.Contains("LED_ULTRALUX"), name.Contains("LED_ЛУНА"),
                 name.Contains("АВАРИЯ"), name.Contains("БОЙЛЕРНО ТАБЛО"),
                 name.Contains("ЛАМПИ_СПАЛНЯ"), name.Contains("ЛИНИЯ МХЛ"),
                 name.Contains("ЛУМИНЕСЦЕНТНА"), name.Contains("МЕТАЛХАЛОГЕННА"),
                 name.Contains("ПЛАФОНИ"), name.Contains("АПЛИК"),
                 name.Contains("ПЕНДЕЛ"), name.Contains("ЛАМПИОН"),
                 name.Contains("НАСТОЛНА ЛАМПА"), name.Contains("ФАСАДНО"),
                 name.Contains("БАНСКИ АПЛИК"), name.Contains("ДАТЧИК"),
                 name.Contains("ФОТОДАТЧИК"), name.Contains("ПОЛИЛЕЙ"),
                 name.Contains("ПРОЖЕКТОР")
                Kons.Phase = 1  ' Всички тези са винаги 1 фаза
        End Select
    End Sub
    ''' <summary>
    ''' Универсална функция за изчисляване на мощност.
    ''' Разпознава автоматично формата на входа:
    ''' - LED ленти: "60 led/m" + дължина
    ''' - Директна мощност: "3500" или "3.5"
    ''' - Контакти/Консуматори: "2х100", "3х100", "100"
    ''' </summary>
    ''' <param name="strМОЩНОСТ">Текст от атрибута "МОЩНОСТ"</param>
    ''' <param name="Dylvina_Led">Дължина на LED лента в метри (ако е приложимо)</param>
    ''' <returns>Обща мощност във Watt</returns>
    Private Shared Function CalcPower(strМОЩНОСТ As String,
                       Optional Dylvina_Led As Double = 0) As Double
        ' --- 1. Валидация ---
        If String.IsNullOrEmpty(strМОЩНОСТ) Then
            Return 0.0
        End If
        Dim input As String = strМОЩНОСТ.Trim().ToLower()
        ' --- 2. LED ЛЕНТИ (формат: "60 led/m", "120led/m") ---
        If input.Contains("led/m") Then
            ' Проверка дали текстът съдържа "led/m" (т.е. LED лента)
            ' Вземаме числото пред "led/m", което показва броя диоди на метър
            ' Превръщаме текста в малки букви, махаме "led/m" и изтриваме интервали
            Dim диоди As Double = Val(strМОЩНОСТ.ToLower().Replace("led/m", "").Trim())
            ' Декларираме променлива за мощността на метър (W/m)
            Dim мощностНаМетър As Double
            ' Определяме мощността на метър според таблица с известни стойности
            ' Ако броят диоди не е стандартен, използваме средна мощност на диод (0.24 W/диод)
            Select Case диоди
                Case 30
                    мощностНаМетър = 7.2       ' 30 диода/м → 7.2 W/м
                Case 60
                    мощностНаМетър = 14.4      ' 60 диода/м → 14.4 W/м
                Case 72
                    мощностНаМетър = 17.28     ' 72 диода/м → 17.28 W/м
                Case 120
                    мощностНаМетър = 28.8      ' 120 диода/м → 28.8 W/м
                Case Else
                    ' За непознат брой диоди използваме средна мощност на диод 0.24 W/диод
                    мощностНаМетър = диоди * 0.24
            End Select
            ' Изчисляваме мощността за реалната дължина на лентата (Dylvina_Led в см)
            Return (Dylvina_Led / 100) * мощностНаМетър
        End If
        ' --- 3. КОНТАКТИ/КОНСУМАТОРИ (формат: "2х100", "3х100", "100") ---
        ' Поддържа различни разделители: "х", "x", "*", "Х"
        Dim separators As String() = {"х", "x", "*", "Х", "X"}
        For Each sep As String In separators
            If input.Contains(sep) Then
                Dim parts As String() = input.Split(sep)
                If parts.Length = 2 Then
                    Dim count As Double = 0.0
                    Dim power As Double = 0.0
                    If Double.TryParse(parts(0).Trim(), count) AndAlso
                    Double.TryParse(parts(1).Trim(), power) Then
                        Return count * power  ' Брой × Мощност на бройка
                    End If
                End If
            End If
        Next
        ' --- 5. ОБИКНОВЕНО ЧИСЛО (формат: "3500", "3.5") ---
        Dim numericValue As Double = 0.0
        If Double.TryParse(input, numericValue) Then
            Return numericValue  ' Предполагаме W
        End If
        ' --- 6. НЕУСПЕШНО РАЗПОЗНАВАНЕ ---
        Return 0.0
    End Function
    Public Shared Function CreateTokowList(ListKonsumator As List(Of strKonsumator)) As List(Of strTokow)
        If ListKonsumator Is Nothing Then Return Nothing
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim dwgFileName As String = System.IO.Path.GetFileNameWithoutExtension(doc.Name)
        Dim fullBuildingName As String = "Сграда_" & dwgFileName
        Dim ListTokow = ListKonsumator.Where(Function(k) Not String.IsNullOrWhiteSpace(k.ТоковКръг)) _
                .GroupBy(Function(k) New With {Key k.ТАБЛО, Key k.ТоковКръг}) _
                .Select(Function(g) New strTokow With {
                            .BuildingName = fullBuildingName, ' Присвояваме името тук
                            .Tablo = g.Key.ТАБЛО,
                            .ТоковКръг = g.Key.ТоковКръг,
                            .Konsumator = g.ToList()
                        }).ToList()
        Return ListTokow
    End Function
End Class
