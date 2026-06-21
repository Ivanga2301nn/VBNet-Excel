#Region "Изчисляване на натоварванията на токовите кръгове"
' За всеки токов кръг:
' - обработва всички консуматори
' - изчислява общата мощност
' - изчислява номиналния ток
' - избира подходящ прекъсвач
' - попълва параметрите на кръга
'
' Това е основната електротехническа част на алгоритъма.
#End Region
Public Class ElectricalCalculationEngine
    ' Пазим локални референции към каталозите, които формата вече е създала
    Private _breakerCatalog As BreakerCatalog
    Private _cableCatalog As CableCatalog
    Private _rcdCatalog As RCDCatalog
    ''' <summary>
    ''' КОНСТРУКТОР: Приема създадените каталози от формата, за да работи с техните актуални данни
    ''' </summary>
    Public Sub New()
        Me._breakerCatalog = AppSettings.BreakerCatalog
        Me._cableCatalog = AppSettings.CableCatalog
        Me._rcdCatalog = AppSettings.RcdCatalog
    End Sub
    Public Class BlockConfig
        Public BlockNames As List(Of String)        ' Възможни имена на блока
        Public Category As String                   ' "Lamp", "Contact", "Device", "Panel"
        Public DefaultPoles As Integer              ' "1p" или "3p"
        Public DefaultCable As String               ' "3x1.5", "3x2.5", "5x2.5"
        Public DefaultBreaker As String             ' "10", "16", "20"
        Public DefaultBreakerType As String         ' "10", "16", "20"
        Public DefaultPrednaz As String             ' Предназначение 
        Public DefaultPrednaz1 As String            ' Предназначение 
        Public VisibilityRules As List(Of VisRule)  ' Правила за visibility
    End Class
    ''' <summary>
    ''' Правило за конкретна visibility стойност
    ''' </summary>
    Public Class VisRule
        Public VisibilityPattern As String        ' "3P", "Двугнездов", "Проточен"
        Public Poles As Integer                    ' "1p" или "3p"
        Public Cable As String                    ' "3x2.5", "5x4"
        Public Breaker As String                  ' "16", "25", "32"
        Public Phase As String                    ' "L" или "L1,L2,L3"
        Public BreakerType As String              ' опционално за специфични правила
        Public ContactCount As Integer            ' Колко контакта добавя (1, 2, 3)
    End Class
    Private BlockConfigs As New List(Of BlockConfig)
    Private Const ZnakX As String = "х" ' Напиши го веднъж тук (на кирилица)
    Private Sub InitializeBlockConfigs()
        BlockConfigs = New List(Of BlockConfig) From {
                New BlockConfig With {        ' LED ОСВЕТЛЕНИЕ
                    .BlockNames = New List(Of String) From {"LED_DENIMA", "LED_LENTA", "LED_ULTRALUX", "LED_ULTRALUX_100", "LED_ULTRALUX_НОВ",
                                                            "LED_ЛУНА", "ПЛАФОНИ", "МЕТАЛХАОГЕННА ЛАМПА", "ЛИНИЯ МХЛ - 220V", "ПОЛИЛЕЙ", "ПРОЖЕКТОР"},
                    .Category = "Lamp",
                    .DefaultPoles = 1,
                    .DefaultCable = "3" & ZnakX & "1,5",
                    .DefaultBreaker = "10",
                    .DefaultBreakerType = "EZ9 MCB",
                    .DefaultPrednaz = "Общо",
                    .DefaultPrednaz1 = "осветление",
                    .VisibilityRules = New List(Of VisRule)()
                },
                New BlockConfig With {        ' УЛИЧНО ОСВЕТЛЕНИЕ
                    .BlockNames = New List(Of String) From {"ULI4NO"},
                    .Category = "Lamp",
                    .DefaultPoles = 1,
                    .DefaultCable = "3" & ZnakX & "1,5",
                    .DefaultBreaker = "10",
                    .DefaultBreakerType = "EZ9 MCB",
                    .DefaultPrednaz = "Улично",
                    .DefaultPrednaz1 = "осветление",
                    .VisibilityRules = New List(Of VisRule)()
                },
                New BlockConfig With {        ' АВАРИЙНО ОСВЕТЛЕНИЕ
                    .BlockNames = New List(Of String) From {"АВАРИЯ", "АВАРИЯ_100"},
                    .Category = "Lamp",
                    .DefaultPoles = 1,
                    .DefaultCable = "3" & ZnakX & "1,5",
                    .DefaultBreakerType = "EZ9 MCB",
                    .DefaultBreaker = "10",
                    .DefaultPrednaz = "Аварийно",
                    .DefaultPrednaz1 = "осветление",
                    .VisibilityRules = New List(Of VisRule)()
                },
                New BlockConfig With {        ' БОЙЛЕРНО ТАБЛО
                    .BlockNames = New List(Of String) From {"БОЙЛЕРНО ТАБЛО"},
                    .Category = "Contact",
                    .DefaultPoles = 1,
                    .DefaultBreakerType = "EZ9 MCB",
                    .DefaultCable = "3" & ZnakX & "2,5",
                    .DefaultBreaker = "10",
                    .VisibilityRules = New List(Of VisRule) From {
                        New VisRule With {.VisibilityPattern = "КЛЮЧ И КОНТАКТ", .ContactCount = 1},
                        New VisRule With {.VisibilityPattern = "С ДВА КОНТАКТА", .ContactCount = 2},
                        New VisRule With {.VisibilityPattern = "С ДВА КЛЮЧА", .ContactCount = 2}
                    }
                },
                New BlockConfig With {        ' КОНТАКТИ
                    .BlockNames = New List(Of String) From {"КОНТАКТ"},
                    .Category = "Contact",
                    .DefaultPoles = 1,
                    .DefaultCable = "3" & ZnakX & "2,5",
                    .DefaultBreakerType = "EZ9 MCB",
                    .DefaultBreaker = "20",
                    .DefaultPrednaz = "Контакти",
                    .DefaultPrednaz1 = "",
                    .VisibilityRules = New List(Of VisRule) From {
                        New VisRule With {.VisibilityPattern = "ДВУГНЕЗДОВ", .Poles = 1, .ContactCount = 2},
                        New VisRule With {.VisibilityPattern = "ТРИГНЕЗДОВ", .Poles = 1, .ContactCount = 3},
                        New VisRule With {.VisibilityPattern = "ТРИФАЗЕН", .Poles = 3, .Cable = "5" & ZnakX & "2,5", .Phase = "L1,L2,L3"},
                        New VisRule With {.VisibilityPattern = "ТР+2МФ", .Poles = 3, .Cable = "5" & ZnakX & "2,5", .Phase = "L1,L2,L3", .ContactCount = 2},
                        New VisRule With {.VisibilityPattern = "ТВЪРДА ВРЪЗКА", .Poles = 1, .Cable = "3" & ZnakX & "4,0"},
                        New VisRule With {.VisibilityPattern = "УСИЛЕН", .Poles = 1, .Cable = "3" & ZnakX & "4,0"},
                        New VisRule With {.VisibilityPattern = "IP 54", .Poles = 1, .Cable = "3" & ZnakX & "2,5"},
                        New VisRule With {.VisibilityPattern = "МОНТАЖ В КАНАЛ", .Poles = 1, .Cable = "3" & ZnakX & "2,5"}
                    }
                },
                New BlockConfig With {        ' ВЕНТИЛАЦИИ, КЛИМАТИЦИ, КОНВЕКТОРИ
                    .BlockNames = New List(Of String) From {"ВЕНТИЛАЦИИ"},
                    .Category = "Мotor",
                    .DefaultPoles = 1,
                    .DefaultCable = "3" & ZnakX & "1,5",
                    .DefaultBreakerType = "EZ9 MCB",
                    .DefaultBreaker = "10",
                    .VisibilityRules = New List(Of VisRule) From {
                        New VisRule With {.VisibilityPattern = "3P", .Poles = 3, .Cable = "5" & ZnakX & "2,5", .Phase = "L1,L2,L3"},
                        New VisRule With {.VisibilityPattern = "КАНАЛЕН 3P", .Poles = 3, .Cable = "5" & ZnakX & "2,5", .Phase = "L1,L2,L3"},
                        New VisRule With {.VisibilityPattern = "ПРОЗОРЧЕН 3P", .Poles = 3, .Cable = "5" & ZnakX & "2,5", .Phase = "L1,L2,L3"}
                    }
                },
                New BlockConfig With {        ' БОЙЛЕРИ
                    .BlockNames = New List(Of String) From {"БОЙЛЕР"},
                    .Category = "Device",
                    .DefaultPoles = 1,
                    .DefaultCable = "3" & ZnakX & "2,5",
                    .DefaultBreakerType = "EZ9 MCB",
                    .DefaultBreaker = "20",
                    .VisibilityRules = New List(Of VisRule) From {
                        New VisRule With {.VisibilityPattern = "ИЗХОД 3P", .Poles = 3, .Cable = "5" & ZnakX & "2,5", .Phase = "L1,L2,L3"},
                        New VisRule With {.VisibilityPattern = "380V", .Poles = 3, .Cable = "5" & ZnakX & "2,5", .Phase = "L1,L2,L3"},
                        New VisRule With {.VisibilityPattern = "ПРОТОЧЕН", .Poles = 1, .Breaker = "20"},
                        New VisRule With {.VisibilityPattern = "СЕШОАР", .Poles = 1, .Breaker = "16"},
                        New VisRule With {.VisibilityPattern = "СЕШОАР С КОНТАКТ", .Poles = 1, .Breaker = "16"},
                        New VisRule With {.VisibilityPattern = "ИЗХОД ГАЗ", .Cable = "3" & ZnakX & "2,5", .Breaker = "6"}
                    }
                }
            }
    End Sub
    ''' <summary>
    ''' Главният метод, който задвижва изчисленията за всички токови кръгове.
    ''' Изчислява електрическите параметри на всички токови кръгове в ListTokow.
    '''
    ''' Логика на работа:
    ''' 1) Уверява се, че конфигурацията на блоковете (BlockConfigs) е инициализирана.
    ''' 2) За всеки токов кръг:
    '''    - Нулира броячите и натрупаната мощност.
    '''    - Обработва всички консуматори в кръга чрез ProcessConsumerByConfig().
    '''    - Изчислява номиналния ток на кръга.
    '''    - Проверява дали конфигурираният прекъсвач е достатъчен.
    '''    - При нужда избира нов прекъсвач според тока.
    '''
    ''' Цел:
    ''' Да осигури коректно оразмеряване на защита (прекъсвач)
    ''' спрямо реално изчисленото натоварване на всеки токов кръг.
    ''' </summary>
    Public Sub ExecuteCalculations()
        ' Проверка за празен списък (защита от грешки)
        If AppSettings.ListTokow Is Nothing OrElse AppSettings.ListTokow.Count = 0 Then Exit Sub
        ' ------------------------------------------------------------
        ' 1) Проверка дали конфигурацията на блоковете е инициализирана
        ' ------------------------------------------------------------
        If BlockConfigs Is Nothing OrElse BlockConfigs.Count = 0 Then InitializeBlockConfigs()
        ' Завъртаме цикъл през всеки токов кръг, извлечен от AutoCAD
        For Each tokow As clsTokow In AppSettings.ListTokow
            ' Ако апаратът е Главен разединител на таблото, прескачаме стандартните изчисления за товар
            If tokow.Device = "Разединител" Then Continue For
            ' Нулиране на броячи и стойности преди ново (преизчисляване)
            tokow.brLamp = 0
            tokow.brKontakt = 0
            tokow.Мощност = 0
            tokow.Брой_Полюси = 1
            tokow.Device = ""
            ' --------------------------------------------------------
            ' 2) Обработка на всички консуматори, закачени към този кръг
            ' --------------------------------------------------------
            For Each kons As strKonsumator In tokow.Konsumator
                ProcessConsumerByConfig(kons, tokow)
            Next
            ' Пазим първоначално зададения ток от потребителя (ако има такъв в Grid-а)
            Dim I_Def As Double = 0
            Double.TryParse(tokow.Breaker_Номинален_Ток, I_Def)
            ' --------------------------------------------------------
            ' 3) Изчисляване на номиналния ток на кръга (Inom)
            '    Подаваме мощността, полюсите (вече като Integer) и дали е мотор
            ' --------------------------------------------------------
            tokow.Ток = calc_Inom(tokow.Мощност, tokow.Брой_Полюси, tokow.Device = "Motor")
            ' --------------------------------------------------------
            ' 4) Автоматичен избор на прекъсвач от базата данни (Excel)
            ' --------------------------------------------------------
            _breakerCatalog.CalculateBreaker(tokow)
            ' Проверка дали автоматично изчисленият ток не е по-малък от ръчно зададения
            Dim I_Get As Double = 0
            Double.TryParse(tokow.Breaker_Номинален_Ток, I_Get)
            If I_Def > I_Get Then
                tokow.Breaker_Номинален_Ток = I_Def.ToString()
            Else
                tokow.Breaker_Номинален_Ток = I_Get.ToString()
            End If
            ' ----------------------------------------------------
            ' 5) Избор на сечение на кабела според тока и полюсите
            ' ----------------------------------------------------
            _cableCatalog.CalculateCable(tokow)
            If tokow.ДТЗ_RCD Then _rcdCatalog.SetRCD(tokow)
            If tokow.RCD_Автомат Then _breakerCatalog.ClearBreaker(tokow)
        Next
    End Sub
    ''' <summary>
    ''' Обработва един консуматор спрямо конфигурацията му (BlockConfigs)
    ''' и прехвърля необходимата информация към съответния токов кръг.
    '''
    ''' Логика:
    ''' 1) Намира конфигурация по име на блок.
    ''' 2) Проверява дали има специфично правило според Visibility.
    ''' 3) Попълва кабел, прекъсвач, полюси, фаза и предназначение.
    ''' 4) Натрупва мощност и броячи (лампи/контакти).
    ''' </summary>
    Private Sub ProcessConsumerByConfig(kons As strKonsumator, ByRef tokow As clsTokow)
        ' ------------------------------------------------------------
        ' 0) Подготвяме данните (унифицираме текста в UpperCase)
        ' ------------------------------------------------------------
        Dim blockName As String = kons.Name.ToUpper()
        Dim visibility As String = If(kons.Visibility IsNot Nothing,
                                  kons.Visibility.ToUpper(),
                                  "")
        ' ------------------------------------------------------------
        ' 1) Търсим основната конфигурация по име на блок
        '    Проверява дали blockName съдържа някое от имената
        '    в BlockNames списъка.
        ' ------------------------------------------------------------
        Dim config = BlockConfigs.FirstOrDefault(
                                 Function(c) c.BlockNames.Any(
                                 Function(n) blockName.Contains(n))
                                 )
        ' Ако няма намерена конфигурация → прекратяваме
        If config Is Nothing Then
            MsgBox("Блок '" & blockName & "' не е намерен в InitializeBlockConfigs!",
                   MsgBoxStyle.Critical)
            Return
        End If
        ' ------------------------------------------------------------
        ' 2) Проверяваме дали има специфично правило според Visibility
        ' ------------------------------------------------------------
        Dim visRule = config.VisibilityRules.FirstOrDefault(Function(r) visibility.Contains(r.VisibilityPattern))
        ' ------------------------------------------------------------
        ' 3) ПРЕХВЪРЛЯНЕ НА ДАННИ ОТ КОНФИГУРАЦИЯТА
        ' ------------------------------------------------------------
        '
        ' Кабел – ако има правило по Visibility → вземаме от него,
        ' иначе използваме Default стойност от конфигурацията
        tokow.Кабел_Сечение = If(visRule IsNot Nothing AndAlso
                                Not String.IsNullOrEmpty(visRule.Cable),
                                visRule.Cable,
                                config.DefaultCable)
        ' Тип кабел – фиксирана стойност
        tokow.Кабел_Тип = "СВТ"
        ' Номинален ток на прекъсвача
        Dim breakerVal As String = If(visRule IsNot Nothing AndAlso
                                    Not String.IsNullOrEmpty(visRule.Breaker),
                                    visRule.Breaker,
                                    config.DefaultBreaker)
        tokow.Breaker_Номинален_Ток = breakerVal
        ' Полюси – от правило или default
        tokow.Брой_Полюси = If(visRule IsNot Nothing AndAlso visRule.Poles <> 0,
                                     visRule.Poles,
                                     config.DefaultPoles)
        ' Числова стойност на полюсите (1 или 3)
        ' Тип апарат – от правило или default
        tokow.Breaker_Тип_Апарат = If(visRule IsNot Nothing AndAlso
                            Not String.IsNullOrEmpty(visRule.BreakerType),
                            visRule.BreakerType,
                            config.DefaultBreakerType)
        ' ------------------------------------------------------------
        ' ФАЗА
        ' ------------------------------------------------------------
        ' Ако е триполюсен → автоматично задаваме трите фази
        If tokow.Брой_Полюси = 3 Then
            tokow.Фаза = "L1,L2,L3"
        Else
            ' Ако не е 3P – запазваме съществуващата фаза
            ' или задаваме по подразбиране
            If String.IsNullOrEmpty(tokow.Фаза) Then tokow.Фаза = "L"
        End If
        ' ------------------------------------------------------------
        ' ПРЕДНАЗНАЧЕНИЕ (Default от глобалната Config)
        ' ------------------------------------------------------------
        tokow.Консуматор = config.DefaultPrednaz
        tokow.предназначение = config.DefaultPrednaz1
        ' ------------------------------------------------------------
        ' 4) МОЩНОСТ И БРОЯЧИ
        ' ------------------------------------------------------------
        ' Добавяме мощността (превръщаме W → kW)
        tokow.Мощност += kons.doubМОЩНОСТ / 1000.0
        ' Извличаме брой от текстовата мощност (ако има множител)
        Dim count As Integer = ExtractCountFromPower(kons.strМОЩНОСТ)
        ' Логика според категорията на конфигурацията
        Select Case config.Category
            Case "Lamp"
                ' Увеличаваме броя лампи
                tokow.brLamp += 1
                tokow.Device = "Лампа"
            Case "Contact"
                ' Ако има специфично правило за брой контакти
                If visRule IsNot Nothing AndAlso
                    visRule.ContactCount > 0 Then
                    tokow.brKontakt += visRule.ContactCount
                Else
                    tokow.brKontakt += count
                End If
                tokow.Device = "Контакт"
            Case "Device"
                ' За устройства – предназначението идва от консуматора
                tokow.Консуматор = kons.PEWDN
                tokow.предназначение = kons.PEWDN1
                ' ============================================================
                ' ПРОВЕРКА ЗА БОЙЛЕР - ТРЯБВА ЛИ ДЗТ ЗАЩИТА
                ' ============================================================
                Dim boilerTypes As String() = {
                                   "Хоризонтален",
                                   "Хоризонтален - 380V",
                                   "Вертикален",
                                   "Вертикален - 380V",
                                   "Проточен",
                                   "Проточен - 380V",
                                   "Бойлер кухня"
                }
                ' Проверяваме дали консуматорът е бойлер
                If boilerTypes.Contains(kons.Visibility) Then
                    tokow.ДТЗ_RCD = True
                    tokow.RCD_Автомат = True
                    tokow.Device = "Бойлер"
                Else
                    tokow.Device = "Консуматор"
                End If
        End Select
    End Sub
    ''' <summary>
    ''' Извлича брой от стойност като "3x100" → 3, "4х18" → 4, "100" → 1
    ''' Поддържа както латиница (x), така и кирилица (х)
    ''' </summary>
    Private Function ExtractCountFromPower(powerStr As String) As Integer
        If String.IsNullOrEmpty(powerStr) Then Return 1
        ' Нормализирай - превърни в малки букви за по-лесно сравнение
        Dim normalized As String = powerStr.ToLower()
        ' Проверка за "x" на латиница ИЛИ "х" на кирилица
        If normalized.Contains("x") OrElse normalized.Contains("х") Then
            ' Разделяй и по двата вида "x"
            Dim separators() As Char = {"x"c, "X"c, "х"c, "Х"c}
            Dim parts() As String = powerStr.Split(separators)
            If parts.Length >= 1 Then
                Dim count As Integer
                ' Опитай да парснеш първата част като число
                If Integer.TryParse(parts(0).Trim(), count) AndAlso count > 0 Then
                    Return count  ' Напр. "3x100" → 3, "4х18" → 4
                End If
            End If
        End If
        Return 1
    End Function
    ''' <summary>
    ''' Изчислява номиналния ток за токов кръг
    ''' </summary>
    ''' <param name="Pkryg">Мощност в kW</param>
    ''' <param name="NumberPoles">Брой фази: "1P" или "3P"</param>
    ''' <param name="Motor">True за двигатели (cos φ = 0.85, КПД = 0.9)</param>
    ''' <returns>Номинален ток в Ampere</returns>
    Public Function calc_Inom(Pkryg As Double,                      ' мощност
                       NumberPoles As Integer,                       ' брой фази
                       Optional Motor As Boolean = False            ' Ако е двигател True - КПД и cos FI да са по 0,83
                       ) As Double                                  ' Изчислява номинален ток за товар
        Dim CosFI As Double                                         ' Декларира променлива за cos φ (фактор на мощността)
        Dim KPD As Double                                           ' Декларира променлива за КПД (коефициент на полезно действие)
        Const U380 As Double = 0.4                                  ' Дефинира константа за напрежение при 380V, преобразувано в kV (киловолти)
        Const U220 As Double = 0.23                                 ' Дефинира константа за напрежение при 220V, преобразувано в kV (киловолти)
        Dim Inom As Double = 0                                      ' Инициализира променлива за номиналния ток с начална стойност 0
        If Motor Then                                               ' Проверява дали токовият кръг е двигател
            CosFI = 0.85                                            ' Ако е двигател, задава фактор на мощността 0.83
            KPD = 0.9                                               ' Ако е двигател, задава КПД 0.83
        Else                                                        ' Ако токовият кръг не е двигател
            CosFI = 0.9                                             ' Задава фактор на мощността 0.9
            KPD = 1                                                 ' Задава КПД 1
        End If
        If NumberPoles = 3 Then                                     ' Проверява дали токовият кръг е трифазен (3 полюса)
            Inom = Pkryg / (U380 * Math.Sqrt(3) * CosFI * KPD)      ' Изчислява номиналния ток за трифазен кръг по формулата
        Else                                                        ' Ако токовият кръг е монофазен (2 полюса)
            Inom = Pkryg / (U220 * CosFI * KPD)                     ' Изчислява номиналния ток за монофазен кръг по формулата
        End If
        Return Inom                                                 ' Връща изчисления номинален ток
    End Function
End Class