Public Class BoardStructureManager

End Class
tokowList = tokowList.OrderBy(Function(t) t.Tablo).                       ' 1) Първо групира/сортира по име на Електрическото Табло
                             ThenBy(Function(t) GetCircuitSortKey(t.ТоковКръг)) _ ' 2) След това сортира кръговете вътре в таблото по техния логически приоритет (напр. Гл. прекъсвач -> Осветление -> Контакти)
                             .ToList()                                            ' 3) Превръща подредената колекция обратно в чист List(Of strTokow)
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