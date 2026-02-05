Imports System.IO
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports ACSMCOMPONENTS24Lib ' Използваме същата библиотека като твоя стар код
Imports AcApp = Autodesk.AutoCAD.ApplicationServices.Application

Public Class SimpleTest

    <CommandMethod("SimpleTest")>
    Public Sub RunUpdate()
        Dim acDoc As Document = AcApp.DocumentManager.MdiActiveDocument ' Активен документ
        Dim acDb As Database = acDoc.Database
        Dim dbMod As Integer = Convert.ToInt32(AcApp.GetSystemVariable("DBMOD"))
        If dbMod <> 0 Then
            ' Ако има промени, но не искаш да записваш автоматично:
            MsgBox("Внимание: Файлът има незаписани промени!", MsgBoxStyle.Information, "Инфо")
            Exit Sub
        End If
        ' 1. Взимаме пълния път на текущия DWG файл
        Dim dwgPath As String = acDb.Filename
        Dim name_file As String = acDoc.Name                                    ' Име на DWG файла
        Dim File_Path As String = Path.GetDirectoryName(name_file)              ' Път до папката
        Dim Path_Name As String = Path.GetFileName(File_Path)                   ' Име на папката (име на проекта)
        Dim File_DST As String = Path.Combine(File_Path, Path_Name & ".dst")    ' Пълен път до DST файла
        Dim Set_Desc As String = "Създадено от Бат Генчо"                        ' Описание на Sheet Set-а

        Dim sheetSetManager As IAcSmSheetSetMgr = New AcSmSheetSetMgr            ' Sheet Set Manager
        Dim sheetSetDatabase As AcSmDatabase
        ' Проверка дали DST файлът съществува
        If System.IO.File.Exists(File_DST) Then
            sheetSetDatabase = sheetSetManager.OpenDatabase(File_DST, False)    ' Отваряме съществуващ DST
        Else
            sheetSetDatabase = sheetSetManager.CreateDatabase(File_DST, "", True) ' Създаваме нов DST
        End If
        Try
            Dim sheetSet As AcSmSheetSet = sheetSetDatabase.GetSheetSet()
            If LockDatabase(sheetSetDatabase, True) = False Then                 ' Заключване за запис
                MsgBox("Sheet set не може да бъде отворен за четене.")
                Exit Sub
            End If
            ' Връща списък с всички Sheet-и от дадена Sheet Set база данни (DST).
            sheetSet.SetName(Path_Name)                                                         ' Име на Sheet Set-а
            sheetSet.SetDesc(Set_Desc)                                                          ' Описание на Sheet Set-а
        Catch ex As Exception
            MsgBox("Грешка: " & ex.Message)
        Finally
            If sheetSetDatabase IsNot Nothing Then LockDatabase(sheetSetDatabase, False) ' Отключване на DST
        End Try
        MsgBox("Sheet Set Name: " & sheetSetDatabase.GetSheetSet().GetName() & vbCrLf &
           "Sheet Set Description: " & sheetSetDatabase.GetSheetSet().GetDesc())
    End Sub
    Public Function LockDatabase(database As AcSmDatabase,
                             lockFlag As Boolean) As Boolean
        ' Променлива за състоянието на заключване
        Dim dbLock As Boolean = False
        ' Ако искаме да заключим и базата в момента е отключена
        If lockFlag = True And
            database.GetLockStatus() = AcSmLockStatus.AcSmLockStatus_UnLocked Then
            ' Заключваме базата
            database.LockDb(database)
            dbLock = True
            ' Ако искаме да отключим и базата е локално заключена
        ElseIf lockFlag = False And
            database.GetLockStatus() = AcSmLockStatus.AcSmLockStatus_Locked_Local Then
            ' Отключваме базата
            database.UnlockDb(database)
            dbLock = True
            ' Във всички останали случаи операцията не е приложима
        Else
            dbLock = False
        End If
        ' Връщаме резултата от операцията
        LockDatabase = dbLock
    End Function
End Class