Imports System.Drawing
Imports System.Windows.Forms
Imports Button = System.Windows.Forms.Button
Imports Font = System.Drawing.Font
Public Class Form_BatchAddCircuits
    Inherits Form
    ' Данни, подадени от основната форма
    Private _targetList As List(Of Form_Tablo_new.strTokow)
    Private _tabloName As String
    ' Контроли, до които ще имаме достъп по-късно
    Private numExist As NumericUpDown
    Private numReserve As NumericUpDown
    Private btnOk As Button
    Private btnCancel As Button
    Private lblInfo As Label
    ''' <summary>
    ''' Инициализира формата, приема входните данни и извиква процедурите за изграждане.
    ''' </summary>
    Public Sub New(targetList As List(Of Form_Tablo_new.strTokow), tabloName As String)
        ' 1. Записваме подадените данни в локални полета
        _targetList = targetList
        _tabloName = tabloName
        ' 2. Извикваме процедурите в строго определен ред
        ConfigureFormSettings()   ' Настройки на самата форма
        BuildUserInterface()      ' Създаване и подреждане на визуалните елементи
        SetupEventHandlers()      ' Свързване на събития и клавишни преки пътища
    End Sub
    ''' <summary>
    ''' Задава базовите свойства на прозореца (размер, стил, позиция, цветове).
    ''' </summary>
    Private Sub ConfigureFormSettings()
        Me.Text = "Добавяне на кръгове за Същеструващи/Резерви"
        Me.Size = New Size(400, 220)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.StartPosition = FormStartPosition.CenterParent
        Me.BackColor = Color.White
        Me.MaximizeBox = False
        Me.MinimizeBox = False
    End Sub
    ''' <summary>
    ''' Създава, конфигурира и добавя всички контроли към формата.
    ''' Отговаря само за визуалната структура.
    ''' </summary>
    Private Sub BuildUserInterface()
        ' --- Хедър ---
        lblInfo = New Label With {
            .Text = "Табло: " & _tabloName,
            .Dock = DockStyle.Top,
            .Height = 35,
            .TextAlign = ContentAlignment.MiddleCenter,
            .BackColor = Color.FromArgb(0, 102, 204),
            .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 14, FontStyle.Bold)
        }
        Me.Controls.Add(lblInfo)
        ' --- GroupBox контейнер ---
        Dim grp As New GroupBox With {
            .Text = " Брой кръгове за добавяне ",
            .Location = New Point(15, 50),
            .Size = New Size(355, 70),
            .Font = New Font("Segoe UI", 12, FontStyle.Bold)
        }
        Me.Controls.Add(grp)
        ' --- TableLayoutPanel за подредба (Етикет - Поле - Етикет - Поле) ---
        Dim tbl As New TableLayoutPanel With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 4,
            .RowCount = 1,
            .Padding = New Padding(5)
        }
        tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 45))
        tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 20))
        tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 45))
        tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 20))
        grp.Controls.Add(tbl)
        ' --- Лейбъл за Съществуващи ---
        Dim lblE As New Label With {
                .Text = "Същеструващи:",
                .AutoSize = False,
                .TextAlign = ContentAlignment.MiddleRight,
                .Dock = DockStyle.Fill,
                .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }

        ' --- Лейбъл за Резерви ---
        Dim lblR As New Label With {
                .Text = "Резерви:",
                .AutoSize = False,
                .TextAlign = ContentAlignment.MiddleRight,
                .Dock = DockStyle.Fill,
                .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }
        numExist = New NumericUpDown With {
                .Minimum = 0,
                .Maximum = 100,
                .Value = 1,
                .Width = 60,
                .TextAlign = HorizontalAlignment.Center,
                .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }
        numReserve = New NumericUpDown With {
            .Minimum = 0,
            .Maximum = 100,
            .Value = 1,
            .Width = 60,
            .TextAlign = HorizontalAlignment.Center,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }
        tbl.Controls.Add(lblE, 0, 0)
        tbl.Controls.Add(numExist, 1, 0)
        tbl.Controls.Add(lblR, 2, 0)
        tbl.Controls.Add(numReserve, 3, 0)
        ' --- Панел за бутоните (долу вдясно) ---
        Dim pnlBtns As New FlowLayoutPanel With {
            .FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft,
            .Location = New Point(15, 130),
            .Size = New Size(355, 45),
            .WrapContents = False,
            .Padding = New Padding(15, 5, 0, 0)
        }
        Me.Controls.Add(pnlBtns)
        ' --- Бутон ГЕНЕРИРАЙ ---
        btnOk = New Button With {
            .Text = "ГЕНЕРИРАЙ",
            .Size = New Size(95, 32),
            .BackColor = Color.FromArgb(0, 102, 204),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.System,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .DialogResult = DialogResult.OK,
            .Margin = New Padding(0, 0, 10, 0)
        }
        btnOk.FlatAppearance.BorderSize = 0
        btnOk.FlatAppearance.MouseOverBackColor = Color.FromArgb(0, 120, 240)
        ' --- Бутон ОТКАЗ ---
        btnCancel = New Button With {
            .Text = "ОТКАЗ",
            .Size = New Size(95, 32),
            .BackColor = Color.FromArgb(240, 240, 240),
            .ForeColor = Color.FromArgb(60, 60, 60),
            .FlatStyle = FlatStyle.System,
            .Font = New Font("Segoe UI", 10, FontStyle.Regular),
            .DialogResult = DialogResult.Cancel, .Margin = New Padding(0)
        }
        btnCancel.FlatAppearance.BorderSize = 0
        btnCancel.FlatAppearance.MouseOverBackColor = Color.FromArgb(220, 220, 220)

        pnlBtns.Controls.Add(btnOk)
        pnlBtns.Controls.Add(btnCancel)
    End Sub
    ''' <summary>
    ''' Дефинира кои бутони затварят формата и какви действия изпълняват.
    ''' </summary>
    Private Sub SetupEventHandlers()
        Me.AcceptButton = btnOk
        Me.CancelButton = btnCancel

        ' Свързваме логиката с бутона. 
        ' (DialogResult затваря формата автоматично след изпълнение на този handler)
        AddHandler btnOk.Click, AddressOf ProcessGeneration
    End Sub
    ''' <summary>
    ''' Прочита въведените стойности, създава обектите и ги добавя към списъка.
    ''' Отговаря САМО за логиката и данните, не за визуалната част.
    ''' </summary>
    Private Sub ProcessGeneration(sender As Object, e As EventArgs)
        ' Бърза валидация: да не се добавя нищо при нули
        If numExist.Value = 0 AndAlso numReserve.Value = 0 Then
            MessageBox.Show("Моля, въведете поне 1 за добавяне.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.DialogResult = DialogResult.None ' Спира автоматичното затваряне
            Return
        End If
        ' 1. Генериране на "съществуващи" кръгове
        For i As Integer = 1 To CInt(numExist.Value)
            _targetList.Add(New Form_Tablo_new.strTokow With {
                .Tablo = _tabloName, .Device = "Съществуващ", .ТоковКръг = "същ.",
                .Консуматор = "Съществуващ", .предназначение = "не се променя",
                .Breaker_Номинален_Ток = "Същ.", .Мощност = 0, .Ток = 0,
                .Брой_Полюси = 0, .Фаза = "---"
            })
        Next
        ' 2. Генериране на резервни кръгове
        For i As Integer = 1 To CInt(numReserve.Value)
            _targetList.Add(New Form_Tablo_new.strTokow With {
                .Tablo = _tabloName, .Device = "Резерва", .ТоковКръг = "рез.",
                .Консуматор = "Резерв", .предназначение = "",
                .Breaker_Номинален_Ток = "Същ.", .Breaker_Тип_Апарат = "EZ9 MCB",
                .Брой_Полюси = 1, .Фаза = "---", .Мощност = 0, .Ток = 0
            })
        Next
    End Sub
End Class