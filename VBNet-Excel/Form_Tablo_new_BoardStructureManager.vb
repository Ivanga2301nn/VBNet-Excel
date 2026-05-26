Public Class BoardStructureManager
    ''' <summary>
    ''' Извършва йерархично сортиране на списъка с токови кръгове (Natural Sort).
    ''' </summary>
    ''' <param name="tokowList">
    ''' Списъкът с токови кръгове (List(Of strTokow)), който се подава по референция (ByRef).
    ''' Тъй като се пренаписва оригиналната референция, промените се отразяват директно в извикващата форма.
    ''' </param>
    ''' <remarks>
    ''' Сортирането следва строга географско-електрическа йерархия в 3 стъпки:
    ''' 1. Сграда (BuildingName) - защитава от празни стойности в AutoCAD.
    ''' 2. Електрическо табло (Tablo) - групира кръговете към съответното табло.
    ''' 3. Токов кръг (ТоковКръг) - естествено сортиране по инженерни приоритети чрез GetCircuitSortKey.
    ''' -----------------------------------------------------------------------------------------
    ''' ПЪРВО НИВО: ГЕОГРАФСКО ГРУПИРАНЕ (ПО СГРАДА)
    ''' Използваме вградена проверка If(функция), за да се подсигурим срещу пропуски в AutoCAD.
    ''' Ако атрибутът за сграда е празен, обектът се насочва към служебна група "БЕЗ СГРАДА".
    ''' -----------------------------------------------------------------------------------------
    ''' ВТОРО НИВО: ЛОКАЛНО ГРУПИРАНЕ (ПО ЕЛЕКТРИЧЕСКО ТАБЛО)
    ''' Изпълнява се строго в границите на всяка отделна сграда, дефинирана на първото ниво.
    ''' Подрежда таблата азбучно по име (напр. ГРТ, Т1, ТО).
    ''' -----------------------------------------------------------------------------------------
    ''' ТРЕТО НИВО: ИНЖЕНЕРНО СОРТИРАНЕ (ПО ТОКОВ КРЪГ)
    ''' Извиква спомагателната функция GetCircuitSortKey, която анализира текстовия низ на кръга
    ''' и генерира изкуствен алфа-нумеричен ключ за "естествено подреждане" по приоритети
    ''' (напр. Съществуващи -> Аварийни -> Допълнителни -> Числа -> Резерви -> Общо).
    ''' -----------------------------------------------------------------------------------------
    ''' ФИНАЛИЗИРАНЕ НА СОРТИРАНЕТО
    ''' Материализира LINQ заявката и я превръща обратно в нов, чист и подреден List(Of strTokow).
    ''' -----------------------------------------------------------------------------------------
    ''' </remarks>
    Public Sub SortListTokow(ByRef tokowList As List(Of strTokow))
        ' Изпълняваме тристепенна LINQ щафета, за да подредим оригиналния списък.
        ' Тъй като променливата е предадена с ByRef, промените ще се отразят веднага.
        tokowList = tokowList.
                    OrderBy(Function(t) If(String.IsNullOrEmpty(t.BuildingName), "БЕЗ СГРАДА", t.BuildingName)). ' ПЪРВО НИВО: Групиране по Сграда.
                    ThenBy(Function(t) t.Tablo). ' ВТОРО НИВО: Подреждане по име на Електрическото Табло.
                    ThenBy(Function(t) GetCircuitSortKey(t.ТоковКръг)). 'ТРЕТО НИВО: Естествено сортиране на самите токови кръгове.
                    ToList() ' ФИНАЛИЗИРАНЕ: Превръщаме подредената колекция обратно в чист списък (List).
    End Sub
    ''' <summary>
    ''' Връща ключ за сортиране на токов кръг със специален приоритет
    ''' Порядок: 1.ав. → 2.до. → 3.други букви → 4.числа → 5.само букви
    ''' </summary>
    Private Function GetCircuitSortKey(circuitName As String) As String
        If String.IsNullOrEmpty(circuitName) Then Return "ZZZZZZZZZZ"
        Dim name As String = circuitName.Trim().ToUpper()
        Dim priority As String = "9"  ' По подразбиране най-нисък приоритет
        Dim numberPart As String = ""
        Dim letterPart As String = ""
        ' ============================================================
        ' 1. ОПРЕДЕЛИ КАТЕГОРИЯТА (ПРИОРИТЕТ)
        ' ============================================================
        Select Case True
            ' 1. СЪЩЕСТВУВАЩИ
            Case name = "СЪЩ."
                priority = "0"
                numberPart = ExtractNumber(name)
                letterPart = "СЪЩ"
            ' 2. АВАРИЙНИ
            Case name.Contains("АВ")
                priority = "1"
                numberPart = ExtractNumber(name)
                letterPart = "АВ"
            ' 3. ДОПЪЛНИТЕЛНИ
            Case name.Contains("ДО")
                priority = "2"
                numberPart = ExtractNumber(name)
                letterPart = "ДО"
            ' 4. ЧИСТИ ЧИСЛА
            Case IsNumeric(name)
                priority = "3"
                numberPart = name
                letterPart = ""
            ' 5. ЧИСЛО + БУКВА (напр. 1а, 2б)
            Case HasNumberAndLetters(name) AndAlso Char.IsDigit(name(0))
                priority = "4"
                numberPart = ExtractNumber(name)
                letterPart = ExtractLetters(name)
            ' 6. ОБЩО (Провери го ПРЕДИ общия случай за букви)
            Case name = "ОБЩО"
                priority = "9"
                numberPart = ""
                letterPart = "ZZZZZ"
            ' 7. РЕЗЕРВА
            Case name = "РЕЗ."
                priority = "8"
                numberPart = ""
                letterPart = "РЕЗ"
            ' 8. ВСИЧКО ЗАПОЧВАЩО С БУКВА (Основни кръгове като А1, Б1 и т.н.)
            Case Not String.IsNullOrEmpty(name) AndAlso Char.IsLetter(name(0))
                priority = "5"
                numberPart = ExtractNumber(name)
                letterPart = name
                ' 9. ВСИЧКО ОСТАНАЛО
            Case Else
                priority = "8"
                numberPart = ""
                letterPart = name
        End Select
        ' ============================================================
        ' 2. СЪЗДАЙ КЛЮЧ ЗА СОРТИРАНЕ
        ' ============================================================
        ' Формат: Приоритет + Номер (с водещи нули) + Букви
        ' Пример: "10000000001АВ" за "1ав."
        If numberPart.Length > 0 Then
            ' Подравняване на числото с водещи нули (до 10 цифри)
            numberPart = numberPart.PadLeft(10, "0"c)
            Return priority & numberPart & letterPart
        Else
            ' Само букви - сортирай азбучно
            Return priority & "0000000000" & letterPart
        End If
    End Function
    ''' <summary>
    ''' Извлича числото от низ (напр. "1АВ" → "1", "А2Б" → "2")
    ''' </summary>
    Private Function ExtractNumber(text As String) As String
        Dim result As String = ""
        For Each c As Char In text
            If Char.IsDigit(c) Then
                result &= c
            End If
        Next
        Return result
    End Function
    ''' <summary>
    ''' Извлича буквите от низ (напр. "1АВ" → "АВ", "А2Б" → "АБ")
    ''' </summary>
    Private Function ExtractLetters(text As String) As String
        Dim result As String = ""
        For Each c As Char In text
            If Char.IsLetter(c) Then
                result &= c
            End If
        Next
        Return result
    End Function
    ''' <summary>
    ''' Проверява дали низът съдържа и букви и числа
    ''' </summary>
    Private Function HasNumberAndLetters(text As String) As Boolean
        Dim hasNumber As Boolean = False
        Dim hasLetter As Boolean = False
        For Each c As Char In text
            If Char.IsDigit(c) Then hasNumber = True
            If Char.IsLetter(c) Then hasLetter = True
        Next
        Return hasNumber AndAlso hasLetter
    End Function
    ''' <summary>
    ''' Проверява дали низът е само число
    ''' </summary>
    Private Function IsNumeric(text As String) As Boolean
        For Each c As Char In text
            If Not Char.IsDigit(c) Then Return False
        Next
        Return text.Length > 0
    End Function
End Class