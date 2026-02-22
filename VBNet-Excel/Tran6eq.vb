Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports System.IO
Imports System.Drawing.Imaging

Public Class Tran6eq
    Dim PI As Double = 3.1415926535897931
    Dim cu As CommonUtil = New CommonUtil()
    Dim kab As Kabel = New Kabel()

    ' Вземете препратка към обекта Sheet Set Manager
    Public Structure strTryba
        Dim Type As String      ' Тип на тръбата
        ' PVC
        ' HDPE
        ' В земята
        ' В бетон
        Dim Layer As String         ' Слой в който да се постави надписа
        Dim D_OD As Double          ' Тръба Външен диаметър
        Dim D_IN As Double          ' Тръба Вътрешен диаметър
        Dim Kabel As String         ' Вид на кабела                 ' рез. - резервна тръба
        Dim Kabel_Diam As Double    ' Диаметър на кабела
        Dim Kabel_Type As String    ' Какво има в тръбата
        Dim Poz_X As Double         ' Позиция X на която положана тръбата
        Dim Poz_Y As Double         ' Позиция Y на която е положана тръбата
        Dim Red As Double           ' Ред на полагане броено от дъното
        Dim Count_Cab As Integer    ' Брой на кабелите в тръбата
        Dim max_Dia_Red As Double   ' Максимален диаметър на реда
    End Structure

    ' Създаваме структура за данни, които ще съхраняваме за всеки ред
    Public Structure RedInfo
        Public broiTrubi As Integer
        Public sumaOD As Double
        Public obshtoShirina As Double
        Public Razst_Trybi As Double
        Public Max_Diam As Double
        Public Poz_Y As Double
    End Structure

    Dim Tryba() As strTryba
    Dim Razst_Trybi As Double = 0
    Dim Razst_Kabeli As Double = 100
    Dim Vertikal_Kabeli As Double = 100

    Dim Dyno As Double = 0
    Dim Dulbochina As Double = 0
    Dim Dulbochina_Poslden_Red As Double = 600  ' Разсточние от последната тръба до пясъка 

    <CommandMethod("KabelIzkop")>
    Public Sub KabelIzkop()
        Dim cu As CommonUtil = New CommonUtil()

        Dim ss = cu.GetObjects("LINE", "Изберете Линия")
        Dim br As Integer = 0
        If ss Is Nothing Then
            MsgBox("НЕ Е маркирана линия в слой 'EL'.")
            Exit Sub
        End If
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim pPtRes As PromptPointResult
        Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
        Dim i As Integer
        '
        ' Ако ИМА кабели от този тип - True
        ' Ако НЯМА кабели от този тип - False
        '
        Dim Yes_Солар As Boolean = False
        Dim Yes_НН As Boolean = False
        Dim Yes_СрН As Boolean = False
        Dim Yes_Слабо As Boolean = False
        Dim Yes_шина As Boolean = False
        Dim Yes_бетон As Boolean = False
        '
        ' Запълва масива Tryba с празни символи
        '

        ReDim Tryba(ss.Count - 1)

        For i = 0 To UBound(Tryba)
            Tryba(i).Type = ""
            Tryba(i).Kabel = ""
            Tryba(i).Kabel_Type = ""
            Tryba(i).Count_Cab = 0
        Next
        Dim Tryba_max_kabeli As Integer = 6
        '
        ' Потребителски избор брой СКАБОТОКОВИ кабели в една търба
        '
        Dim pDouOpts As PromptDoubleOptions = New PromptDoubleOptions("")
        With pDouOpts
            .Keywords.Add("1")
            .Keywords.Add("5")
            .Keywords.Add("6")
            .Keywords.Add("7")
            .Keywords.Add("8")
            .Keywords.Add("9")
            .Keywords.Add("10")
            .Keywords.Default = "1"
            .Message = vbCrLf & "Въведете Брой СКАБОТОКОВИ кабели в една търба: "
            .AllowZero = False
            .AllowNegative = False
        End With
        Dim pKeyRes As PromptDoubleResult = acDoc.Editor.GetDouble(pDouOpts)
        If pKeyRes.Status = PromptStatus.Keyword Then
            Tryba_max_kabeli = pKeyRes.StringResult
        Else
            Tryba_max_kabeli = pKeyRes.Value
        End If
        '
        ' Потребителски избор разстояние между тръбите
        '
        pDouOpts.Keywords.Clear()
        With pDouOpts
            .Keywords.Add("1")
            .Keywords.Add("50")
            .Keywords.Add("100")
            .Keywords.Add("150")
            .Keywords.Default = "100"
            .Message = vbCrLf & "Въведете разсроянието между ТРЪБИТЕ: "
            .AllowZero = False
            .AllowNegative = False
        End With
        pKeyRes = acDoc.Editor.GetDouble(pDouOpts)
        If pKeyRes.Status = PromptStatus.Keyword Then
            Razst_Trybi = pKeyRes.StringResult
        Else
            Razst_Trybi = pKeyRes.Value
        End If
        '
        ' Потребителски избор разстояние между тръбите
        '
        pDouOpts.Keywords.Clear()
        With pDouOpts
            .Keywords.Add("1")
            .Keywords.Add("50")
            .Keywords.Add("100")
            .Keywords.Add("150")
            .Keywords.Default = "100"
            .Message = vbCrLf & "Въведете разсроянието между КАБЕЛИТЕ: "
            .AllowZero = False
            .AllowNegative = False
        End With
        pKeyRes = acDoc.Editor.GetDouble(pDouOpts)
        If pKeyRes.Status = PromptStatus.Keyword Then
            Razst_Kabeli = pKeyRes.StringResult
        Else
            Razst_Kabeli = pKeyRes.Value
        End If
        '
        ' Запълва масива Tryba с данни за кабелите
        '
        Using trans As Transaction = acDoc.TransactionManager.StartTransaction()

            For Each sObj As SelectedObject In ss
                Dim line As Line = TryCast(trans.GetObject(sObj.ObjectId, OpenMode.ForRead), Line)
                If line.Linetype = "" Then Continue For

                Dim line_Type As String = cu.GET_line_Type(line.Linetype, vbTrue)
                Dim Kabel_Type = cu.line_Layer(line.Layer)
                '
                '   Групира слаботоковите и соларните кабели
                '
                If InStr(Kabel_Type, "FTP") Or
                    InStr(Kabel_Type, "RG6/64") Or
                    InStr(Kabel_Type, "H1Z2Z2") Or
                    InStr(Kabel_Type, "RG59CU") Or
                    InStr(Kabel_Type, "Оптичен") Or
                    InStr(Kabel_Type, "FS ") Then
                    Dim iVisib As Integer = -1
                    iVisib = Array.FindIndex(Tryba, Function(f) f.Kabel = Kabel_Type And
                                                                f.Count_Cab < Tryba_max_kabeli)
                    If iVisib <> -1 Then
                        Tryba(iVisib).Count_Cab += 1
                        Continue For
                    End If
                End If
                Tryba(br).Kabel = Kabel_Type
                Tryba(br).Layer = line.Layer
                '
                '   Записва какво има в тръбата
                '
                If InStr(Tryba(br).Kabel, "H1Z2Z2") Then Tryba(br).Kabel_Type = "Солар"
                If InStr(Tryba(br).Kabel, "САВТ") Then Tryba(br).Kabel_Type = "НН"
                If InStr(Tryba(br).Kabel, "СВТ") Then Tryba(br).Kabel_Type = "НН"
                If InStr(Tryba(br).Kabel, "САХЕк(вн)П") Then Tryba(br).Kabel_Type = "СрН"
                If InStr(Tryba(br).Kabel, "FTP") Then Tryba(br).Kabel_Type = "Слабо"
                If InStr(Tryba(br).Kabel, "RG6/64") Then Tryba(br).Kabel_Type = "Слабо"
                If InStr(Tryba(br).Kabel, "FS ") Then Tryba(br).Kabel_Type = "Слабо"
                If InStr(Tryba(br).Kabel, "Оптичен") Then Tryba(br).Kabel_Type = "Слабо"

                If InStr(Tryba(br).Kabel, "RG59CU") Then Tryba(br).Kabel_Type = "Слабо"
                If InStr(Tryba(br).Kabel, "ПВ-A2") Then Tryba(br).Kabel_Type = "ПВ-A2"
                If InStr(Tryba(br).Kabel, "CAB/6") Then Tryba(br).Kabel_Type = "Слабо"
                If InStr(Tryba(br).Kabel, "Резервна") Then Tryba(br).Kabel_Type = "Резервна"

                If InStr(Tryba(br).Kabel, "поц.шина") Then Tryba(br).Kabel_Type = "шина"
                If InStr(Tryba(br).Kabel, "EL_РЕЗЕРВА") Then Tryba(br).Kabel_Type = "РЕЗЕРВА"
                '
                '   Записва какъв е типа на тръбата
                '
                If InStr(line_Type, "HDPE") Then Tryba(br).Type = "HDPE"
                If InStr(line_Type, "PVC") Then Tryba(br).Type = "PVC"
                If InStr(line_Type, "изкоп") Then Tryba(br).Type = "изкоп"
                If InStr(line_Type, "бетон") Then Tryba(br).Type = "бетон"
                '
                '   Записва диаметъра на кабела
                '
                Tryba(br).Kabel_Diam = cu.GET_line_Diamet(line.Layer)
                '
                ' Запизва диаметрите на тръбата или кабела
                '
                Dim sss As String = ""
                Dim we As Integer = InStr(line_Type, "р.ф")
                Dim ew As Integer = InStr(line_Type, "/")
                If (we * ew) > 0 Then
                    sss = Mid(line_Type, we + 3, ew - we - 3)
                    Tryba(br).D_OD = Val(sss)
                    we = InStr(line_Type, "/")
                    ew = InStr(line_Type, "mm")
                    sss = Mid(line_Type, we + 1, ew - we - 1)
                    Tryba(br).D_IN = Val(sss)
                Else
                    If Tryba(br).Kabel_Type = "СрН" Then
                        Tryba(br).D_OD = 2 * Tryba(br).Kabel_Diam
                    Else
                        Tryba(br).D_OD = Tryba(br).Kabel_Diam
                    End If
                    Tryba(br).D_IN = 0
                End If
                Tryba(br).Count_Cab = 1
                Select Case Tryba(br).Kabel_Type
                    Case "Солар"
                        Yes_Солар = True
                    Case "НН"
                        Yes_НН = True
                    Case "СрН"
                        Yes_СрН = True
                    Case "Слабо"
                        Yes_Слабо = True
                End Select
                br += 1
            Next
        End Using

        Dim br_tryba As Integer = -1 ' Номер на тръба с по-малко кабели
        For i = 0 To UBound(Tryba)
            If Tryba(i).Kabel_Type <> "Слабо" Then Continue For
            If Tryba(i).Count_Cab < Tryba_max_kabeli Then
                If br_tryba = -1 Then
                    br_tryba = i
                Else
                    If (Tryba(i).Count_Cab + Tryba(br_tryba).Count_Cab) <= Tryba_max_kabeli Then
                        Tryba(br_tryba).Count_Cab += Tryba(i).Count_Cab
                        Tryba(i).Type = ""
                        Tryba(i).D_OD = 0
                        Tryba(i).D_IN = 0
                        Tryba(i).Kabel = ""
                        Tryba(i).Kabel_Diam = 0
                        Tryba(i).Kabel_Type = ""
                        Tryba(i).Poz_X = 0
                        Tryba(i).Poz_Y = 0
                        Tryba(i).Count_Cab = 0
                        Tryba(i).max_Dia_Red = 0
                    End If
                End If
            End If
        Next
        '
        '   Събира соларните кабели в една тръба
        '
        br_tryba = -1 ' Номер на тръба с по-малко кабели
        For i = 0 To UBound(Tryba)
            If Tryba(i).Kabel_Type <> "Солар" Then Continue For
            If Tryba(i).Count_Cab < Tryba_max_kabeli Then
                If br_tryba = -1 Then
                    br_tryba = i
                Else
                    If (Tryba(i).Count_Cab + Tryba(br_tryba).Count_Cab) <= Tryba_max_kabeli Then
                        Tryba(br_tryba).Count_Cab += Tryba(i).Count_Cab
                        Tryba(i).Type = ""
                        Tryba(i).D_OD = 0
                        Tryba(i).D_IN = 0
                        Tryba(i).Kabel = ""
                        Tryba(i).Kabel_Diam = 0
                        Tryba(i).Kabel_Type = ""
                        Tryba(i).Poz_X = 0
                        Tryba(i).Poz_Y = 0
                        Tryba(i).Count_Cab = 0
                        Tryba(i).max_Dia_Red = 0
                    End If
                End If
            End If
        Next

        '
        ' Изчисля дъното на траншеята
        '
        Dim sum_trybi_SN As Double = 100
        Dim sum_trybi_NN As Double = 100
        Dim sum_trybi_Sl As Double = 100
        Dim sum_trybi_So As Double = 100
        Dim sum_trybi As Double = 0

        Dim br_trybi_SN As Double = 0
        Dim br_trybi_NN As Double = 0
        Dim br_trybi_Sl As Double = 0
        Dim br_trybi_So As Double = 0

        Dim br_kab_SN As Double = 0
        Dim br_kab_NN As Double = 0
        Dim br_kab_Sl As Double = 0
        Dim br_kab_So As Double = 0

        For i = 0 To UBound(Tryba)
            If Tryba(i).D_OD = 0 Then Continue For
            Select Case Tryba(i).Kabel_Type
                Case "Солар"
                    sum_trybi_So += Tryba(i).D_OD + IIf(Tryba(i).Type = "изкоп", 100, 0)
                    br_trybi_So += IIf(Tryba(i).Type = "изкоп", 1, 0)
                    br_kab_So += IIf(Tryba(i).Type = "изкоп", 0, 1)
                Case "НН"
                    sum_trybi_NN += Tryba(i).D_OD + IIf(Tryba(i).Type = "изкоп", 100, 0)
                    br_trybi_NN += IIf(Tryba(i).Type = "изкоп", 1, 0)
                    br_kab_NN += IIf(Tryba(i).Type = "изкоп", 0, 1)
                Case "СрН"
                    br_trybi_SN += IIf(Tryba(i).Type = "изкоп", 1, 0)
                    br_kab_SN += IIf(Tryba(i).Type = "изкоп", 0, 1)
                    sum_trybi_SN += Tryba(i).D_OD + IIf(Tryba(i).Type = "изкоп", 100, 0)
                Case "Слабо"
                    sum_trybi_Sl += Tryba(i).D_OD + IIf(Tryba(i).Type = "изкоп", 100, 0)
                    br_trybi_Sl += IIf(Tryba(i).Type = "изкоп", 1, 0)
                    br_kab_Sl += IIf(Tryba(i).Type = "изкоп", 0, 1)
            End Select
        Next
        sum_trybi = sum_trybi_SN + sum_trybi_NN + sum_trybi_Sl + sum_trybi_So
        Select Case sum_trybi
            Case > 1200
                Dyno = 1400
            Case 1000 To 1200
                Dyno = 1200
            Case 800 To 1000
                Dyno = 1000
            Case 500 To 800
                Dyno = 800
            Case < 500
                Dyno = 500
        End Select
        Dim blkRecId_Tran6 As ObjectId = ObjectId.Null
        Try
            Dim pDouOpts_Dylbo As PromptDoubleOptions = New PromptDoubleOptions("")

            '
            ' Потребителски избор ДЪЛБОЧИНАТА на траншеята
            '
            pDouOpts_Dylbo.Keywords.Clear()
            With pDouOpts_Dylbo
                .Keywords.Add("800")
                .Keywords.Add("1000")
                .Keywords.Add("1100")
                .Keywords.Add("1300")
                .Keywords.Default = "800"
                .Message = vbCrLf & "Въведете ДЪЛБОЧИНАТА на полагане: "
                .AllowZero = False
                .AllowNegative = False
            End With
            pKeyRes = acDoc.Editor.GetDouble(pDouOpts_Dylbo)
            If pKeyRes.Status = PromptStatus.Keyword Then
                Dulbochina = Val(pKeyRes.StringResult)
            Else
                Dulbochina = Val(pKeyRes.Value.ToString())
            End If
            '
            ' Потребителски избор ШИРИНАТА на траншеята
            '
            Dim pDouOpts_Dyno As PromptDoubleOptions = New PromptDoubleOptions("")
            With pDouOpts_Dyno
                .Keywords.Add("500")
                .Keywords.Add("800")
                .Keywords.Add("1000")
                .Keywords.Add("1200")
                .Keywords.Add("1400")
                .Keywords.Add("1600")
                .Keywords.Add("2000")
                .Keywords.Default = "500"
                .Message = vbCrLf & "Въведете ШИРИНАТА на дъното ОБЩО{" & sum_trybi.ToString & "} НИСКО НАПРЕЖЕНИЕ {" & sum_trybi_NN.ToString & "}: "
                .AllowZero = False
                .AllowNegative = False
            End With
            pKeyRes = acDoc.Editor.GetDouble(pDouOpts_Dyno)
            If pKeyRes.Status = PromptStatus.Keyword Then
                Dyno = pKeyRes.StringResult
            Else
                Dyno = pKeyRes.Value.ToString()
            End If
            'Dim poz_X As Double = 0
            'Dim poz_Y As Double = 100
            'Dim Diam_Old As Double = 0
            'Dim Аrranges As Boolean = False
            'pDouOpts = New PromptDoubleOptions("")
            'With pDouOpts
            '    .Keywords.Add("Да")
            '    .Keywords.Add("Не")
            '    .Keywords.Default = "Не"
            '    .Message = vbCrLf & "Жалаете ли подрежане по тръби: "
            '    .AllowZero = False
            '    .AllowNegative = False
            'End With
            ''
            '' Разпределя трабите в траншеята
            ''
            'pKeyRes = acDoc.Editor.GetDouble(pDouOpts)
            'Аrranges = IIf(pKeyRes.StringResult = "Не", False, True)
            'If Аrranges Then
            '    '
            '    poz_X = Razpred_Trybi_X(poz_X, "НН", "тръба", 1, Аrranges)        ' Разпределя кабелите които СА в ТРЪБИ
            '    poz_X = Razpred_Trybi_X(poz_X, "НН", "изкоп", 2, Аrranges)        ' Разпределя кабелите които СА в ИЗКОП    
            '    '
            '    poz_X = Razpred_Trybi_X(poz_X, "СрН", "тръба", 1, Аrranges)       ' Разпределя кабелите които СА в ТРЪБИ
            '    poz_X = Razpred_Trybi_X(poz_X, "СрН", "изкоп", 1, Аrranges)       ' Разпределя кабелите които СА в ИЗКОП   
            '    '
            '    poz_X = Razpred_Trybi_X(poz_X, "Слабо", "тръба", 1, Аrranges)     ' Разпределя кабелите които СА в ТРЪБИ
            '    poz_X = Razpred_Trybi_X(poz_X, "Слабо", "изкоп", 1, Аrranges)     ' Разпределя кабелите които СА в ИЗКОП   
            '    '
            '    poz_X = Razpred_Trybi_X(poz_X, "Солар", "тръба", -1, Аrranges)    ' Разпределя кабелите които СА в ТРЪБИ
            '    poz_X = Razpred_Trybi_X(poz_X, "Солар", "изкоп", -1, Аrranges)    ' Разпределя кабелите които СА в ИЗКОП  
            'Else
            '    poz_X = Razpred_Trybi_X(0, "", "", 1, Аrranges)    ' Разпределя по начина на маркиране
            'End If
            '
            ' Разпределя трабите в траншеята
            '
            Call ArrangePipesAndCables()
            '
            ' Потребителски избор точка на въкване на блока
            '
            Using trans As Transaction = acDoc.TransactionManager.StartTransaction()
                pPtOpts.Message = vbLf & "Изберете точка на вмъкване на блока: "
                pPtRes = acDoc.Editor.GetPoint(pPtOpts)
            End Using
            Dim InsertPoint As Point3d = pPtRes.Value
            Dim poz_Tryba As Point3d
            Dim acBlkRef As BlockReference
            Dim props As DynamicBlockReferencePropertyCollection
            '
            ' Вмъква блок Траншея
            '
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                blkRecId_Tran6 = cu.InsertBlock("Траншея", InsertPoint, "EL_РАМКА", New Scale3d(1, 1, 1))
                acBlkRef = DirectCast(acTrans.GetObject(blkRecId_Tran6, OpenMode.ForWrite), BlockReference)
                props = acBlkRef.DynamicBlockReferencePropertyCollection
                For Each prop As DynamicBlockReferenceProperty In props
                    If prop.PropertyName = "Сигнална лента" Then prop.Value = 350.0
                    If prop.PropertyName = "Дълбочина" Then prop.Value = Dulbochina
                    If prop.PropertyName = "D_1" Then prop.Value = 100.0
                    If prop.PropertyName = "D_2" Then prop.Value = 100.0
                    If prop.PropertyName = "D_3" Then prop.Value = 100.0
                    If prop.PropertyName = "D_4" Then prop.Value = 400.0
                    If prop.PropertyName = "Пясък" Then prop.Value = 600.0
                    If prop.PropertyName = "D_Пясък" Then prop.Value = Dyno
                    If prop.PropertyName = "Дъно" Then prop.Value = Dyno
                    If prop.PropertyName = "Лента" Then prop.Value = Dyno / 2
                Next
                acTrans.Commit()
            End Using
            '
            ' Вмъква блок ТРЪБИТЕ и КАБЕЛИТЕ
            '
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                For i = 0 To Tryba.Length - 1
                    If Tryba(i).Count_Cab = 0 Then Exit For

                    poz_Tryba = New Point3d(InsertPoint.X + Tryba(i).Poz_X, InsertPoint.Y - Dulbochina, 0.0)

                    If Tryba(i).Kabel_Type <> "Резервна" Then
                        Dim blkRecId_Kabel = cu.InsertBlock("Траншея_кабел", poz_Tryba, "EL_РАМКА", New Scale3d(1, 1, 1))
                        acBlkRef = DirectCast(acTrans.GetObject(blkRecId_Kabel, OpenMode.ForWrite), BlockReference)
                        props = acBlkRef.DynamicBlockReferencePropertyCollection

                        For Each prop As DynamicBlockReferenceProperty In props
                            If prop.PropertyName = "Visibility" Then prop.Value = IIf(Tryba(i).Kabel_Type = "СрН", "СрН", "НН")
                            If prop.PropertyName = "Diam" Then prop.Value = Tryba(i).Kabel_Diam
                            If prop.PropertyName = "Височина" Then prop.Value = Tryba(i).Poz_Y + IIf(Tryba(i).Type <> "изкоп", Tryba(i).D_OD - Tryba(i).D_IN, 0)
                            If prop.PropertyName = "Начало" Then prop.Value = Tryba(i).Poz_X
                        Next
                    End If

                    If Tryba(i).Type <> "изкоп" Then
                        Dim blkRecId_tryba = cu.InsertBlock("Траншея_тръба", poz_Tryba, "EL_РАМКА", New Scale3d(1, 1, 1))
                        acBlkRef = DirectCast(acTrans.GetObject(blkRecId_tryba, OpenMode.ForWrite), BlockReference)
                        props = acBlkRef.DynamicBlockReferencePropertyCollection
                        For Each prop As DynamicBlockReferenceProperty In props
                            If prop.PropertyName = "D_OD" Then prop.Value = Tryba(i).D_OD
                            If prop.PropertyName = "D_IN" Then prop.Value = Tryba(i).D_IN
                            If prop.PropertyName = "Височина" Then prop.Value = Tryba(i).Poz_Y
                            If prop.PropertyName = "Visibility" Then prop.Value = "Без размери"
                        Next
                    End If
                Next
                acTrans.Commit()
            End Using
            '
            ' Вмъква блок Надписите
            '
            Dim arrKabelАlign As New List(Of CommonUtil.strKabelАlign)
            Dim blKabelАlign As CommonUtil.strKabelАlign
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim Poz_nadpis As Double = 67.5
                Dim Position_X As Double = 0
                Dim Position_Y As Double
                Dim scale As Double = 4
                Dim strVisib As String = "Точка"

                Dim Point_Kabel As Point3d
                Dim blkRecIdNow As ObjectId
                Dim blkRecId_OB_Nasip As ObjectId
                Dim attCol As AttributeCollection
                '
                ' Вмъква блок Надпис "Обратен насип"
                '
                Point_Kabel = New Point3d(InsertPoint.X + Dyno, InsertPoint.Y - 160, 0.0)
                blkRecId_OB_Nasip = cu.InsertBlock("Кабел", Point_Kabel, "EL__DIM", New Scale3d(scale, scale, scale))
                acBlkRef = DirectCast(acTrans.GetObject(blkRecId_OB_Nasip, OpenMode.ForWrite), BlockReference)
                attCol = acBlkRef.AttributeCollection
                props = acBlkRef.DynamicBlockReferencePropertyCollection
                Position_Y = 0
                Position_X = 560
                For Each prop As DynamicBlockReferenceProperty In props
                    'This Is where you change states based on input
                    If prop.PropertyName = "Position X" Then prop.Value = Position_X    ' Позиция на вмъкване коорд. Х
                    If prop.PropertyName = "Position Y" Then prop.Value = Position_Y    ' Позиция на вмъкване коорд. Х
                    If prop.PropertyName = "Visibility" Then prop.Value = strVisib
                Next
                For Each objID As ObjectId In attCol
                    Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForWrite)
                    Dim acAttRef As AttributeReference = dbObj
                    If acAttRef.Tag = "NA4IN_0" Then
                        acAttRef.TextString = "Обратен насип"
                    Else
                        acAttRef.TextString = ""
                    End If
                Next
                cu.EditDynamicBlockReferenceKabel(blkRecId_OB_Nasip)
                '
                ' Вмъква блок Надпис "Сигнална лента"
                '
                Point_Kabel = New Point3d(InsertPoint.X + Dyno / 2, InsertPoint.Y - 350, 0.0)
                blkRecIdNow = cu.InsertBlock("Кабел", Point_Kabel, "EL__DIM", New Scale3d(scale, scale, scale))
                acBlkRef = DirectCast(acTrans.GetObject(blkRecIdNow, OpenMode.ForWrite), BlockReference)
                attCol = acBlkRef.AttributeCollection
                props = acBlkRef.DynamicBlockReferencePropertyCollection
                Position_Y = 121.7471
                Position_X = Dyno / 2 + 560
                For Each prop As DynamicBlockReferenceProperty In props
                    If prop.PropertyName = "Position X" Then prop.Value = Position_X    ' Позиция на вмъкване коорд. Х
                    If prop.PropertyName = "Position Y" Then prop.Value = Position_Y    ' Позиция на вмъкване коорд. Х
                    If prop.PropertyName = "Visibility" Then prop.Value = strVisib
                Next
                For Each objID As ObjectId In attCol
                    Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForWrite)
                    Dim acAttRef As AttributeReference = dbObj
                    If acAttRef.Tag = "NA4IN_0" Then
                        acAttRef.TextString = "Сигнална лента"
                    Else
                        acAttRef.TextString = ""
                    End If
                Next
                cu.EditDynamicBlockReferenceKabel(blkRecIdNow)
                '
                ' Вмъква блок Надпис "Пясък"
                '
                Point_Kabel = New Point3d(InsertPoint.X + Dyno, InsertPoint.Y - 660, 0.0)
                blkRecIdNow = cu.InsertBlock("Кабел", Point_Kabel, "EL__DIM", New Scale3d(scale, scale, scale))
                acBlkRef = DirectCast(acTrans.GetObject(blkRecIdNow, OpenMode.ForWrite), BlockReference)
                attCol = acBlkRef.AttributeCollection
                props = acBlkRef.DynamicBlockReferencePropertyCollection
                Position_Y = 121.7471
                Position_X = 560
                For Each prop As DynamicBlockReferenceProperty In props
                    If prop.PropertyName = "Position X" Then prop.Value = Position_X    ' Позиция на вмъкване коорд. Х
                    If prop.PropertyName = "Position Y" Then prop.Value = Position_Y    ' Позиция на вмъкване коорд. Х
                    If prop.PropertyName = "Visibility" Then prop.Value = strVisib
                Next
                For Each objID As ObjectId In attCol
                    Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForWrite)
                    Dim acAttRef As AttributeReference = dbObj
                    If acAttRef.Tag = "NA4IN_0" Then
                        acAttRef.TextString = "Пресята пръст/Пясък"
                    Else
                        acAttRef.TextString = ""
                    End If
                Next

                cu.EditDynamicBlockReferenceKabel(blkRecIdNow)


                For i = 0 To UBound(Tryba)
                    Point_Kabel = New Point3d(InsertPoint.X + Tryba(i).Poz_X, InsertPoint.Y - Dulbochina + Tryba(i).Poz_Y, 0.0)
                    If Tryba(i).Count_Cab = 0 Then Exit For


                    blkRecIdNow = cu.InsertBlock("Кабел", Point_Kabel, Tryba(i).Layer, New Scale3d(scale, scale, scale))

                    acBlkRef = DirectCast(acTrans.GetObject(blkRecIdNow, OpenMode.ForWrite), BlockReference)
                    props = acBlkRef.DynamicBlockReferencePropertyCollection
                    If i = 0 Then
                        Position_X = 500 + Dyno
                    Else
                        Position_X = 500 + Dyno - Tryba(i - 1).Poz_X
                    End If
                    Position_Y -= Poz_nadpis
                    For Each prop As DynamicBlockReferenceProperty In props
                        'This Is where you change states based on input
                        If prop.PropertyName = "Position X" Then prop.Value = Position_X    ' Позиция на вмъкване коорд. Х
                        If prop.PropertyName = "Position Y" Then prop.Value = Position_Y    ' Позиция на вмъкване коорд. Х
                        If prop.PropertyName = "Visibility" Then prop.Value = strVisib
                    Next
                    '
                    ' Поставя надписите
                    '
                    attCol = acBlkRef.AttributeCollection
                    ' Обхождаме всеки ObjectId в колекцията с атрибути attCol (най-вероятно от блок)
                    For Each objID As ObjectId In attCol
                        ' Взимаме DBObject от транзакцията (в режим за запис, за да можем да променяме стойности)
                        Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForWrite)
                        ' Преобразуваме DBObject към AttributeReference, за да работим с текстовите атрибути в блока
                        Dim acAttRef As AttributeReference = dbObj
                        ' Проверяваме по етикета (Tag) на атрибута, за да определим какво да записваме в него
                        Select Case acAttRef.Tag
                             'Първи атрибут: "NA4IN_0"
                            Case "NA4IN_0"
                                ' Записваме брой + мерна единица + тип кабел:
                                ' - Ако кабелът е "Резервна тръба", мерната единица е "бр."
                                ' - В противен случай е "л." (линейни метри)
                                acAttRef.TextString = Tryba(i).Count_Cab.ToString +
                                  If(Tryba(i).Kabel = "Резервна тръба", "бр.", "л. ") +
                                  Tryba(i).Kabel
                            ' Втори атрибут: "NA4IN_1"
                            Case "NA4IN_1"
                                ' Определяме типа на тръбата и съответно подготвяме надписа
                                Select Case Tryba(i).Type
                                    Case "HDPE"
                                        ' Ако кабелът НЕ е "Резервна тръба", добавяме префикс "изт. в"
                                        ' След това добавяме текст с типа на тръбата и диаметъра ѝ
                                        acAttRef.TextString = If(Tryba(i).Kabel = "Резервна тръба", "", "изт. в ") +
                                          "HDPE тр.ф" + Tryba(i).D_OD.ToString + "mm"
                                        ' Намаляваме Y-позицията на надписа, ако се налага да има отстояние между редовете
                                        Position_Y -= Poz_nadpis
                                    Case "PVC"
                                        ' Аналогично, но за PVC тръба
                                        acAttRef.TextString = If(Tryba(i).Kabel = "Резервна тръба", "", "изт. в ") +
                                          "PVC тр.ф" + Tryba(i).D_OD.ToString + "mm"
                                        Position_Y -= Poz_nadpis
                                    Case Else
                                        ' За всички други типове (или неразпознати), използваме общ текст
                                        acAttRef.TextString = "положен в изкоп"
                                End Select
                                ' За всички други атрибути, които не са "NA4IN_0" или "NA4IN_1"
                            Case Else
                                ' Изчистваме текста (можеш да сложиш друго поведение при нужда)
                                acAttRef.TextString = ""
                        End Select
                    Next
                    blKabelАlign.pInsert = Point_Kabel
                    blKabelАlign.Id = blkRecIdNow
                    blKabelАlign.Position_X = Position_X
                    blKabelАlign.Position_Y = Position_Y
                    blKabelАlign.Stypka = 2
                    arrKabelАlign.Add(blKabelАlign)
                Next

                arrKabelАlign = arrKabelАlign.OrderByDescending(Function(k) k.pInsert.X).ToList()

                blKabelАlign = arrKabelАlign(0)
                blKabelАlign.Position_X = 550
                blKabelАlign.Position_Y = -75

                cu.Kabel_Aligment(blKabelАlign, arrKabelАlign, True, True, True)
                acTrans.Commit()
            End Using
            '
            ' Добавя заземителна шина
            '
        Catch ex As Exception
            If blkRecId_Tran6.IsNull Then
                MsgBox("Възникна грешка: " & "Във файла липсва блок 'ТРАНШЕЯ'!!!")
            Else
                MsgBox("Възникна грешка: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace.ToString)
            End If
        End Try
    End Sub
    Public Function Razpred_Trybi_X(' Обработва масив Tryba който се обработва
                              poziciq As Double,        ' Начална позиция от която започват да се редят тръбите/кабелите
                              Kabel_Type As String,     ' Напрeжние на кабела        -> "НН", "Слабо", "СрН", "Солар"
                              Type As String,           ' Начин на полагане         -> Ако е "изкоп" 
                              Kabel_Red As Integer,
                              Аrranges As Boolean       ' Да се продрежат ли по траби или не ! тъпо в и май трябва да го махна
                              ) As Double
        ' Паранетри за Kabel_Red - Подреждане на редовете:
        ' -> -1 - Ако трябва да е на един ред по-високо от основния
        ' -> 1 - на един ред в основния ред
        ' -> 2 - на подрежда на два реда в основния ред
        ' -> 3 - Ако трябва да е на един ред по-ниско от основния

        ' Групиране по редове
        '--------------------------------------------------------------------------------
        Dim grupiPoRed = Tryba.GroupBy(Function(t) t.Red)
        ' Списък за съхранение на информацията за всеки ред
        Dim redInfoList As New List(Of RedInfo)
        For Each grupa In grupiPoRed
            Dim trubiNaRed = grupa.ToList()
            Dim broiTrubi = trubiNaRed.Count
            ' Сума от външните диаметри
            Dim sumaOD = trubiNaRed.Sum(Function(t) t.D_OD)
            ' Обща ширина: сума от диаметри + разстояния между тръбите
            Dim obshtoShirina = sumaOD + Razst_Trybi * (broiTrubi - 1)
            Dim RazsTrybi = (Dyno - obshtoShirina) / broiTrubi

            ' Записваме информацията за текущия ред
            redInfoList.Add(New RedInfo With {
                                    .broiTrubi = broiTrubi,
                                    .sumaOD = sumaOD,
                                    .obshtoShirina = obshtoShirina,
                                    .Razst_Trybi = RazsTrybi
                                    })
        Next

        Dim Razst As Double = IIf(Type = "изкоп", Razst_Kabeli, Razst_Trybi)
        Dim red_kabel As Double = 0
        Dim first As Boolean = True
        Dim Vert_kabeli As Double

        Select Case Kabel_Red
            Case -1
                Vert_kabeli = 250
                red_kabel = 1
            Case 1
                Vert_kabeli = 0
                red_kabel = 1
            Case 2
                red_kabel = 0
            Case 3
                red_kabel = 1
                Vert_kabeli = 250
        End Select

        For i As Integer = Tryba.Length - 1 To 0 Step -1
            If Аrranges Then
                If Tryba(i).Kabel_Type <> Kabel_Type Then Continue For
                If Type = "изкоп" And Tryba(i).Type <> "изкоп" Then Continue For
                If Type <> "изкоп" And Tryba(i).Type = "изкоп" Then Continue For
            End If

            If Kabel_Red = 2 Then
                Vert_kabeli = red_kabel * (200 - Vertikal_Kabeli - Tryba(i).D_OD)
                red_kabel = IIf(red_kabel = 1, 0, 1)
            End If

            Tryba(i).Poz_X = poziciq + red_kabel * (Tryba(i).D_OD / 2 + redInfoList(Tryba(i).Red).Razst_Trybi / 2)
            poziciq += red_kabel * (Tryba(i).D_OD + redInfoList(Tryba(i).Red).Razst_Trybi)
            Tryba(i).Poz_Y = Vertikal_Kabeli + Vert_kabeli
        Next
        Return poziciq
    End Function
    Public Sub ArrangePipesAndCables()
        ' 1. Сортиране на масива Tryba()
        ' Извикваме функцията SortTrybaArray(), която вероятно сортира елементите в масива Tryba.
        ' Това може да бъде сортиране по някаква характеристика на обектите в Tryba (например по диаметър, ID и т.н.).
        Call SortTrybaArray()
        ' 2. Разпределяне на елементите по редове
        ' Инициализираме брояча на редовете (RedCounter) на 1.
        ' Инициализираме текущата ширина на реда (CurrentWidth) на 0.
        Dim RedCounter As Integer = 1
        Dim CurrentWidth As Double = 0

        Dim Razst_Redove As Double = 50             ' Разстоянието между редовете

        ' Започваме цикъл през всички елементи в масива Tryba.
        For i As Integer = 0 To Tryba.Length - 1
            ' Изчисляваме пълната ширина на елемента (включително разстоянието между елементите) като
            ' събираме диаметъра на елемента (D_OD) и разстоянието между елементите (Razst_Trybi).
            Dim fullWidth As Double = Tryba(i).D_OD + Razst_Trybi
            ' Проверяваме дали елементът може да се побере в текущия ред:
            ' - Ако текущата ширина на реда плюс ширината на елемента не надвишава максималната ширина на реда (Dyno),
            ' - Или ако редът е първоначален (CurrentWidth = 0), добавяме елемента в текущия ред.
            ' Ако не може, започваме нов ред.
            If CurrentWidth + fullWidth <= Dyno Or CurrentWidth = 0 Then
                Tryba(i).Red = RedCounter ' Присвояваме текущия ред на елемента.
                CurrentWidth += fullWidth ' Актуализираме текущата ширина на реда.
            Else
                RedCounter += 1 ' Увеличаваме броя на редовете (начин за започване на нов ред).
                Tryba(i).Red = RedCounter ' Присвояваме новия ред на елемента.
                CurrentWidth = fullWidth ' Рестартираме текущата ширина на реда с ширината на текущия елемент.
            End If
        Next
        '
        ' Групиране по редове
        '
        Dim grupiPoRed = Tryba.GroupBy(Function(t) t.Red)
        ' Списък за съхранение на информацията за всеки ред
        Dim redInfoList As New List(Of RedInfo)
        Dim PozY As Double = Razst_Redove
        For Each grupa In grupiPoRed
            Dim trubiNaRed = grupa.ToList()
            Dim broiTrubi = trubiNaRed.Count
            ' Сума от външните диаметри
            Dim sumaOD = trubiNaRed.Sum(Function(t) t.D_OD)
            ' Обща ширина: сума от диаметри + разстояния между тръбите
            Dim obshtoShirina = sumaOD + Razst_Trybi * (broiTrubi - 1)
            ' Разстояние между тръбите
            Dim RazsTrybi = (Dyno - sumaOD) / broiTrubi
            ' Максимален диаметър в реда
            Dim DiamMax = trubiNaRed.Max(Function(t) t.D_OD)
            ' Записваме информацията за текущия ред
            redInfoList.Add(New RedInfo With {
                                    .broiTrubi = broiTrubi,
                                    .sumaOD = sumaOD,
                                    .obshtoShirina = obshtoShirina,
                                    .Razst_Trybi = RazsTrybi,
                                    .Max_Diam = DiamMax,
                                    .Poz_Y = PozY
                                    })
            PozY += Razst_Redove + DiamMax
        Next
        ' Променлива за текущото натрупано разстояние между тръбите в даден ред
        Dim current_Raz_Trybi As Double = 0
        ' Обхождаме всяка тръба от масива Tryba
        For i As Integer = 0 To Tryba.Length - 1
            ' Задаваме Y-позицията на тръбата според реда ѝ (редовете започват от 1, затова -1)
            Tryba(i).Poz_Y = redInfoList(Tryba(i).Red - 1).Poz_Y
            ' Ако това е първата тръба, или ако текущата тръба е в различен ред от предишната —
            ' започваме нов ред и нулираме натрупаното разстояние
            If i = 0 OrElse Tryba(i).Red <> Tryba(i - 1).Red Then
                current_Raz_Trybi = 0
            End If
            ' Изчисляваме X-позицията на текущата тръба:
            ' - текущото натрупано разстояние
            ' - плюс половината от разстоянието между тръбите (за да има отстояние отляво)
            ' - плюс половината от външния диаметър на тръбата (за да центрираме тръбата)
            Tryba(i).Poz_X = current_Raz_Trybi + redInfoList(Tryba(i).Red - 1).Razst_Trybi / 2 + Tryba(i).D_OD / 2
            ' Обновяваме натрупаното разстояние:
            ' - добавяме разстоянието между тръбите
            ' - добавяме външния диаметър на тръбата (за да "освободим място" за следващата)
            current_Raz_Trybi += redInfoList(Tryba(i).Red - 1).Razst_Trybi + Tryba(i).D_OD
        Next
        Dim new_Dylb = 600 + Razst_Redove + redInfoList(redInfoList.Count - 1).Poz_Y + redInfoList(redInfoList.Count - 1).Max_Diam
        new_Dylb = Math.Ceiling(new_Dylb / 50) * 50
        Dulbochina = Math.Max(Dulbochina, new_Dylb)
    End Sub
    Private Function GetTypePriority(kabelType As String) As Integer
        Select Case True
            Case kabelType.ToLower().Contains("срн")
                Return 1
            Case kabelType.ToLower().Contains("нн")
                Return 2
            Case kabelType.ToLower().Contains("солар")
                Return 3
            Case kabelType.ToLower().Contains("слабо")
                Return 4
            Case kabelType.ToLower().Contains("рез")
                Return 5
            Case Else
                Return 6 ' ако не разпознаем типа – най-отзад
        End Select
    End Function
    Public Sub SortTrybaArray()
        Array.Sort(Tryba, Function(a As strTryba, b As strTryba) As Integer
                              ' 1. По приоритет на типа
                              Dim typeComp As Integer = GetTypePriority(a.Kabel_Type).CompareTo(GetTypePriority(b.Kabel_Type))
                              If typeComp <> 0 Then Return typeComp
                              ' 2. По външен диаметър (низходящо)
                              Dim diamComp As Integer = b.D_OD.CompareTo(a.D_OD)
                              If diamComp <> 0 Then Return diamComp
                              ' 3. По азбучен ред на името на кабела
                              Return String.Compare(a.Kabel, b.Kabel, StringComparison.CurrentCultureIgnoreCase)
                          End Function)
    End Sub
End Class
