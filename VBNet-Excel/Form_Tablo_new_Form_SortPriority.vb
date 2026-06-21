Imports System.Windows.Forms

Public Class Form_SortPriority
    ' Твоят стандартен инженерен ред (Гръбнакът на системата)
    ' Ключът е името на групата, а стойността е приоритетът (нивото)
    Public Shared StandardOrder As New Dictionary(Of String, Integer) From {
                                    {"СЪЩ", 0},
                                    {"АВ", 1},
                                    {"ДО", 2},
                                    {"ОС", 3},
                                    {"КО", 4},
                                    {"ОВ", 5},
                                    {"[ЧИСТИ ЧИСЛА]", 97},
                                    {"РЕЗ", 98},
                                    {"ОБЩО", 99}
}
#Region "Управление на формата"
    ''' <summary>
    ''' При зареждане на формата, тя сама събира и анализира данните
    ''' </summary>
    Private Sub Form_SortPriority_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Заковава формата точно в средата на екрана при стартиране
        Me.StartPosition = FormStartPosition.CenterScreen
        ' 1. Сканираме текущото табло за уникални групи
        Dim detectedPrefixes As New HashSet(Of String)()
        For Each tok In AppSettings.ListTokow
            Dim letterPart As String = ""
            Dim numberPart As String = ""
            AnalyzeCircuitName(tok.ТоковКръг, letterPart, numberPart)
            If Not String.IsNullOrEmpty(letterPart) Then
                detectedPrefixes.Add(letterPart)
            End If
        Next
        ' 2. Сглобяваме списъка за нашия ListBox на база речника в мениджъра
        Dim groupsForForm As New List(Of String)()
        ' Добавяме познатите групи, подредени по тежест
        Dim knownGroups = detectedPrefixes.
                          Where(Function(p) StandardOrder.ContainsKey(p)).
                          OrderBy(Function(p) StandardOrder(p)).
                          ToList()
        groupsForForm.AddRange(knownGroups)
        ' Добавяме непознатите (нови) групи азбучно в средата
        Dim unknownGroups = detectedPrefixes.
                            Where(Function(p) Not StandardOrder.ContainsKey(p) AndAlso p <> "РЕЗ" AndAlso p <> "ОБЩО").
                            OrderBy(Function(p) p).
                            ToList()
        Dim insertIndex As Integer = groupsForForm.FindIndex(Function(g) g = "РЕЗ" OrElse g = "ОБЩО")
        If insertIndex >= 0 Then
            groupsForForm.InsertRange(insertIndex, unknownGroups)
        Else
            groupsForForm.AddRange(unknownGroups)
        End If
        ' Слагаме РЕЗЕРВА и ОБЩО най-отзад
        If detectedPrefixes.Contains("РЕЗ") AndAlso Not groupsForForm.Contains("РЕЗ") Then groupsForForm.Add("РЕЗ")
        If detectedPrefixes.Contains("ОБЩО") AndAlso Not groupsForForm.Contains("ОБЩО") Then groupsForForm.Add("ОБЩО")
        ' 3. Пълним ListBox-а на екрана
        lstPrefixes.Items.Clear()
        For Each grp In groupsForForm
            lstPrefixes.Items.Add(grp)
        Next
        If lstPrefixes.Items.Count > 0 Then lstPrefixes.SelectedIndex = 0
    End Sub
    ' Бутон НАГОРЕ
    Private Sub btnUp_Click(sender As Object, e As EventArgs) Handles btnUp.Click
        Dim index As Integer = lstPrefixes.SelectedIndex
        If index > 0 Then
            Dim currentItem As String = lstPrefixes.SelectedItem.ToString()
            If currentItem = "ОБЩО" OrElse currentItem = "РЕЗ" Then Exit Sub
            Dim item As Object = lstPrefixes.SelectedItem
            lstPrefixes.Items.RemoveAt(index)
            lstPrefixes.Items.Insert(index - 1, item)
            lstPrefixes.SelectedIndex = index - 1
        End If
    End Sub
    ' Бутон НАДОЛУ
    Private Sub btnDown_Click(sender As Object, e As EventArgs) Handles btnDown.Click
        Dim index As Integer = lstPrefixes.SelectedIndex
        If index >= 0 AndAlso index < lstPrefixes.Items.Count - 1 Then
            Dim currentItem As String = lstPrefixes.SelectedItem.ToString()
            Dim nextItem As String = lstPrefixes.Items(index + 1).ToString()
            If nextItem = "РЕЗ" OrElse nextItem = "ОБЩО" Then Exit Sub
            If currentItem = "ОБЩО" OrElse currentItem = "РЕЗ" Then Exit Sub
            Dim item As Object = lstPrefixes.SelectedItem
            lstPrefixes.Items.RemoveAt(index)
            lstPrefixes.Items.Insert(index + 1, item)
            lstPrefixes.SelectedIndex = index + 1
        End If
    End Sub
    ' Бутон СОРТИРАЙ (ОК) - Този бутон сега управлява финализирането!
    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        StandardOrder.Clear()

        For i As Integer = 0 To lstPrefixes.Items.Count - 1
            Dim groupName As String = lstPrefixes.Items(i).ToString()

            If groupName = "ОБЩО" Then
                StandardOrder.Add(groupName, 99)
            ElseIf groupName = "РЕЗ" Then
                StandardOrder.Add(groupName, 98)
            Else
                StandardOrder.Add(groupName, i) ' Новата подредба от ListBox-а
            End If
        Next
        Me.DialogResult = Windows.Forms.DialogResult.OK
        ' Извикваме сортирането директно в мениджъра зад кулисите
        SortListTokow()
        Me.Close()
    End Sub
    ' Бутон ОТКАЗ
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        ' Казваме, че потребителят се е отказал и затваряме
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub
#End Region
#Region "Сортиране на списък с токови кръгове (List(Of strTokow))"
    ''' <summary>
    ''' Извършва йерархично сортиране на списъка с токови кръгове (Natural Sort).
    ''' </summary>
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
    ''' (ВИЖ StandardOrder).
    ''' -----------------------------------------------------------------------------------------
    ''' ФИНАЛИЗИРАНЕ НА СОРТИРАНЕТО
    ''' Материализира LINQ заявката и я превръща обратно в нов, чист и подреден List(Of strTokow).
    ''' -----------------------------------------------------------------------------------------
    ''' </remarks>
    Public Sub SortListTokow()
        ' Изпълняваме тристепенна LINQ щафета, за да подредим оригиналния списък.
        AppSettings.ListTokow = AppSettings.ListTokow.
            OrderBy(Function(t) If(String.IsNullOrEmpty(t.BuildingName),
                                  "БЕЗ СГРАДА", t.BuildingName)).   ' 1. По Сграда
            ThenBy(Function(t) t.Tablo).                            ' 2. По име на Електрическото Табло
            ThenBy(Function(t) GetUniversalSortKey(t.ТоковКръг)).   ' 3. НОВОТО универсално естествено сортиране
            ToList()
    End Sub
    ''' <summary>
    ''' Анализира името на кръга и отделя чистата буквена група от номера, прескачайки разделители
    ''' </summary>
    Private Sub AnalyzeCircuitName(circuitName As String, ByRef letterPart As String, ByRef numberPart As String)
        If String.IsNullOrEmpty(circuitName) Then
            letterPart = ""
            numberPart = ""
            Exit Sub
        End If
        ' 1. Нормализация (премахваме точки и интервали)
        Dim name As String = circuitName.Trim().ToUpper().Replace(".", "")
        ' Служебни твърди съвпадения за края на таблото
        If name = "ОБЩО" Then
            letterPart = "ОБЩО"
            numberPart = ""
            Exit Sub
        End If
        If name = "РЕЗ" OrElse name = "РЕЗЕРВА" Then
            letterPart = "РЕЗ"
            numberPart = ""
            Exit Sub
        End If
        ' 2. Извличане на чистите цифри и чистите букви (тирета и долни черти се прескачат)
        Dim digits As String = ""
        Dim letters As String = ""
        For Each c As Char In name
            If Char.IsDigit(c) Then
                digits &= c
            Else
                If Char.IsLetter(c) Then
                    letters &= c
                End If
            End If
        Next
        ' 3. Класификация на резултата
        If letters.Length > 0 AndAlso digits.Length > 0 Then
            ' Има и букви, и цифри (напр. "ОС-1", "2_КО", "1АВ")
            letterPart = letters
            numberPart = digits
        ElseIf digits.Length > 0 Then
            ' Само числа (напр. "15", "1")
            letterPart = "[ЧИСТИ ЧИСЛА]"
            numberPart = digits
        Else
            ' Само букви (напр. "ОС", "КО")
            letterPart = letters
            numberPart = ""
        End If
    End Sub
    ''' <summary>
    ''' Генерира мощен ключ за сортиране на база динамичното тегло на групите
    ''' </summary>
    Public Function GetUniversalSortKey(circuitName As String) As String
        ' Ако името е празно, връщаме максимално голям ключ, за да отиде най-отзад
        If String.IsNullOrEmpty(circuitName) Then Return "99_9999999999_ZZZZZ"
        Dim letterPart As String = ""
        Dim numberPart As String = ""
        ' Извикваме анализатора, който ни връща чиста група (без разделители) и числови номер
        AnalyzeCircuitName(circuitName, letterPart, numberPart)
        Dim priority As String = ""
        ' 1. Директна и светкавична проверка в Речника
        If StandardOrder.ContainsKey(letterPart) Then
            ' Взимаме нивото (напр. 0, 1, 3, или 99 за ОБЩО) 
            ' и го правим на текст с водеща нула (напр. "00", "01", "03", "99")
            Dim level As Integer = StandardOrder(letterPart)
            priority = level.ToString().PadLeft(2, "0"c)
        Else
            ' 2. За напълно нови букви (извън речника) - отиват в "златната среда" (приоритет 50),
            ' за да са след основните групи, но ПРЕД твърдо закованите РЕЗЕРВА (98) и ОБЩО (99).
            ' Добавяме и самите букви към приоритета, за да се подредят азбучно спрямо други нови букви.
            priority = "50_" & letterPart
        End If
        ' 3. Форматираме номера с водещи нули (до 10 цифри за правилно "естествено" сортиране)
        If numberPart.Length > 0 Then
            numberPart = numberPart.PadLeft(10, "0"c)
        Else
            numberPart = "0000000000"
        End If
        ' Крайният ключ комбинира: Приоритет на групата + Номер + Буква (за разграничение при еднакви номера)
        ' Пример за "ОС-1": "03_0000000001_ОС"
        ' Пример за "ОБЩО": "99_0000000000_ОБЩО"
        Return priority & "_" & numberPart & "_" & letterPart
    End Function
#End Region
End Class