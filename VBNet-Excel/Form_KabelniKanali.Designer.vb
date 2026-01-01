<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form_KabelniKanali
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form_KabelniKanali))
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.NewToolStripButton = New System.Windows.Forms.ToolStripButton()
        Me.OpenToolStripButton = New System.Windows.Forms.ToolStripButton()
        Me.SaveToolStripButton = New System.Windows.Forms.ToolStripButton()
        Me.PrintToolStripButton = New System.Windows.Forms.ToolStripButton()
        Me.toolStripSeparator = New System.Windows.Forms.ToolStripSeparator()
        Me.CutToolStripButton = New System.Windows.Forms.ToolStripButton()
        Me.CopyToolStripButton = New System.Windows.Forms.ToolStripButton()
        Me.PasteToolStripButton = New System.Windows.Forms.ToolStripButton()
        Me.toolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.HelpToolStripButton = New System.Windows.Forms.ToolStripButton()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.DataGridView_Кабели = New System.Windows.Forms.DataGridView()
        Me.Вид = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.Жила = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.Сечение = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.Кабели = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Диаметър = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Площ = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GroupBox_Избор = New System.Windows.Forms.GroupBox()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.ProgressBar_Procent = New System.Windows.Forms.ProgressBar()
        Me.TextBox_Кабелна_Скара = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label_Сечение_кабели = New System.Windows.Forms.Label()
        Me.Label_Процент_Запълване = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ComboBox_Процент_Запълване = New System.Windows.Forms.ComboBox()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.Button_В_4 = New System.Windows.Forms.Button()
        Me.Button_В_3 = New System.Windows.Forms.Button()
        Me.Button_В_2 = New System.Windows.Forms.Button()
        Me.Button_В_1 = New System.Windows.Forms.Button()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.Button_Ш_10 = New System.Windows.Forms.Button()
        Me.Button_Ш_9 = New System.Windows.Forms.Button()
        Me.Button_Ш_8 = New System.Windows.Forms.Button()
        Me.Button_Ш_5 = New System.Windows.Forms.Button()
        Me.Button_Ш_7 = New System.Windows.Forms.Button()
        Me.Button_Ш_6 = New System.Windows.Forms.Button()
        Me.Button_Ш_4 = New System.Windows.Forms.Button()
        Me.Button_Ш_3 = New System.Windows.Forms.Button()
        Me.Button_Ш_2 = New System.Windows.Forms.Button()
        Me.Button_Ш_1 = New System.Windows.Forms.Button()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.RadioButton_Тръба = New System.Windows.Forms.RadioButton()
        Me.RadioButton_Канал = New System.Windows.Forms.RadioButton()
        Me.RadioButton_Скара = New System.Windows.Forms.RadioButton()
        Me.ToolStrip1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.DataGridView_Кабели, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox_Избор.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'ToolStrip1
        '
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.NewToolStripButton, Me.OpenToolStripButton, Me.SaveToolStripButton, Me.PrintToolStripButton, Me.toolStripSeparator, Me.CutToolStripButton, Me.CopyToolStripButton, Me.PasteToolStripButton, Me.toolStripSeparator1, Me.HelpToolStripButton})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Padding = New System.Windows.Forms.Padding(0, 0, 2, 0)
        Me.ToolStrip1.Size = New System.Drawing.Size(995, 25)
        Me.ToolStrip1.TabIndex = 0
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'NewToolStripButton
        '
        Me.NewToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.NewToolStripButton.Image = CType(resources.GetObject("NewToolStripButton.Image"), System.Drawing.Image)
        Me.NewToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.NewToolStripButton.Name = "NewToolStripButton"
        Me.NewToolStripButton.Size = New System.Drawing.Size(23, 22)
        Me.NewToolStripButton.Text = "&New"
        '
        'OpenToolStripButton
        '
        Me.OpenToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.OpenToolStripButton.Image = CType(resources.GetObject("OpenToolStripButton.Image"), System.Drawing.Image)
        Me.OpenToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.OpenToolStripButton.Name = "OpenToolStripButton"
        Me.OpenToolStripButton.Size = New System.Drawing.Size(23, 22)
        Me.OpenToolStripButton.Text = "&Open"
        '
        'SaveToolStripButton
        '
        Me.SaveToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.SaveToolStripButton.Image = CType(resources.GetObject("SaveToolStripButton.Image"), System.Drawing.Image)
        Me.SaveToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.SaveToolStripButton.Name = "SaveToolStripButton"
        Me.SaveToolStripButton.Size = New System.Drawing.Size(23, 22)
        Me.SaveToolStripButton.Text = "&Save"
        '
        'PrintToolStripButton
        '
        Me.PrintToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.PrintToolStripButton.Image = CType(resources.GetObject("PrintToolStripButton.Image"), System.Drawing.Image)
        Me.PrintToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.PrintToolStripButton.Name = "PrintToolStripButton"
        Me.PrintToolStripButton.Size = New System.Drawing.Size(23, 22)
        Me.PrintToolStripButton.Text = "&Print"
        '
        'toolStripSeparator
        '
        Me.toolStripSeparator.Name = "toolStripSeparator"
        Me.toolStripSeparator.Size = New System.Drawing.Size(6, 25)
        '
        'CutToolStripButton
        '
        Me.CutToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.CutToolStripButton.Image = CType(resources.GetObject("CutToolStripButton.Image"), System.Drawing.Image)
        Me.CutToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.CutToolStripButton.Name = "CutToolStripButton"
        Me.CutToolStripButton.Size = New System.Drawing.Size(23, 22)
        Me.CutToolStripButton.Text = "C&ut"
        '
        'CopyToolStripButton
        '
        Me.CopyToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.CopyToolStripButton.Image = CType(resources.GetObject("CopyToolStripButton.Image"), System.Drawing.Image)
        Me.CopyToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.CopyToolStripButton.Name = "CopyToolStripButton"
        Me.CopyToolStripButton.Size = New System.Drawing.Size(23, 22)
        Me.CopyToolStripButton.Text = "&Copy"
        '
        'PasteToolStripButton
        '
        Me.PasteToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.PasteToolStripButton.Image = CType(resources.GetObject("PasteToolStripButton.Image"), System.Drawing.Image)
        Me.PasteToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.PasteToolStripButton.Name = "PasteToolStripButton"
        Me.PasteToolStripButton.Size = New System.Drawing.Size(23, 22)
        Me.PasteToolStripButton.Text = "&Paste"
        '
        'toolStripSeparator1
        '
        Me.toolStripSeparator1.Name = "toolStripSeparator1"
        Me.toolStripSeparator1.Size = New System.Drawing.Size(6, 25)
        '
        'HelpToolStripButton
        '
        Me.HelpToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.HelpToolStripButton.Image = CType(resources.GetObject("HelpToolStripButton.Image"), System.Drawing.Image)
        Me.HelpToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.HelpToolStripButton.Name = "HelpToolStripButton"
        Me.HelpToolStripButton.Size = New System.Drawing.Size(23, 22)
        Me.HelpToolStripButton.Text = "He&lp"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.DataGridView_Кабели)
        Me.GroupBox1.Controls.Add(Me.GroupBox_Избор)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(0, 25)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(995, 420)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        '
        'DataGridView_Кабели
        '
        Me.DataGridView_Кабели.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView_Кабели.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Вид, Me.Жила, Me.Сечение, Me.Кабели, Me.Диаметър, Me.Площ})
        Me.DataGridView_Кабели.Location = New System.Drawing.Point(6, 24)
        Me.DataGridView_Кабели.Name = "DataGridView_Кабели"
        Me.DataGridView_Кабели.Size = New System.Drawing.Size(523, 389)
        Me.DataGridView_Кабели.TabIndex = 4
        '
        'Вид
        '
        Me.Вид.Frozen = True
        Me.Вид.HeaderText = "Вид"
        Me.Вид.Items.AddRange(New Object() {"СВТ/САВТ", "Слаботоков", "Соларен"})
        Me.Вид.Name = "Вид"
        '
        'Жила
        '
        Me.Жила.Frozen = True
        Me.Жила.HeaderText = "Брой жила"
        Me.Жила.Items.AddRange(New Object() {"1", "2", "3", "3+", "4", "5"})
        Me.Жила.Name = "Жила"
        Me.Жила.Width = 50
        '
        'Сечение
        '
        Me.Сечение.Frozen = True
        Me.Сечение.HeaderText = "Сечение"
        Me.Сечение.Items.AddRange(New Object() {"1,5", "2,5", "4", "6", "10", "16", "25", "35", "50", "70", "95", "120", "150", "185", "240", "300", "400", "500"})
        Me.Сечение.Name = "Сечение"
        Me.Сечение.Width = 75
        '
        'Кабели
        '
        Me.Кабели.HeaderText = "Брой Кабели"
        Me.Кабели.Name = "Кабели"
        Me.Кабели.Width = 65
        '
        'Диаметър
        '
        Me.Диаметър.HeaderText = "Диаметър"
        Me.Диаметър.Name = "Диаметър"
        Me.Диаметър.ReadOnly = True
        Me.Диаметър.Width = 90
        '
        'Площ
        '
        Me.Площ.HeaderText = "Площ"
        Me.Площ.Name = "Площ"
        Me.Площ.ReadOnly = True
        '
        'GroupBox_Избор
        '
        Me.GroupBox_Избор.BackColor = System.Drawing.SystemColors.Control
        Me.GroupBox_Избор.Controls.Add(Me.GroupBox5)
        Me.GroupBox_Избор.Controls.Add(Me.GroupBox4)
        Me.GroupBox_Избор.Controls.Add(Me.GroupBox3)
        Me.GroupBox_Избор.Location = New System.Drawing.Point(536, 90)
        Me.GroupBox_Избор.Name = "GroupBox_Избор"
        Me.GroupBox_Избор.Size = New System.Drawing.Size(451, 323)
        Me.GroupBox_Избор.TabIndex = 3
        Me.GroupBox_Избор.TabStop = False
        Me.GroupBox_Избор.Text = "Избор на начин на полагане"
        '
        'GroupBox5
        '
        Me.GroupBox5.BackColor = System.Drawing.SystemColors.Control
        Me.GroupBox5.Controls.Add(Me.ProgressBar_Procent)
        Me.GroupBox5.Controls.Add(Me.TextBox_Кабелна_Скара)
        Me.GroupBox5.Controls.Add(Me.Label4)
        Me.GroupBox5.Controls.Add(Me.Label_Сечение_кабели)
        Me.GroupBox5.Controls.Add(Me.Label_Процент_Запълване)
        Me.GroupBox5.Controls.Add(Me.Label1)
        Me.GroupBox5.Controls.Add(Me.ComboBox_Процент_Запълване)
        Me.GroupBox5.Location = New System.Drawing.Point(7, 25)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(434, 132)
        Me.GroupBox5.TabIndex = 4
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Предложение"
        '
        'ProgressBar_Procent
        '
        Me.ProgressBar_Procent.BackColor = System.Drawing.Color.Red
        Me.ProgressBar_Procent.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ProgressBar_Procent.ForeColor = System.Drawing.Color.Orange
        Me.ProgressBar_Procent.Location = New System.Drawing.Point(3, 113)
        Me.ProgressBar_Procent.Name = "ProgressBar_Procent"
        Me.ProgressBar_Procent.Size = New System.Drawing.Size(428, 16)
        Me.ProgressBar_Procent.Step = 1
        Me.ProgressBar_Procent.TabIndex = 9
        Me.ProgressBar_Procent.Value = 50
        '
        'TextBox_Кабелна_Скара
        '
        Me.TextBox_Кабелна_Скара.Location = New System.Drawing.Point(176, 67)
        Me.TextBox_Кабелна_Скара.Name = "TextBox_Кабелна_Скара"
        Me.TextBox_Кабелна_Скара.Size = New System.Drawing.Size(85, 26)
        Me.TextBox_Кабелна_Скара.TabIndex = 8
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 73)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(128, 20)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Скара [ШхВ],mm"
        '
        'Label_Сечение_кабели
        '
        Me.Label_Сечение_кабели.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label_Сечение_кабели.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label_Сечение_кабели.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label_Сечение_кабели.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label_Сечение_кабели.Location = New System.Drawing.Point(176, 25)
        Me.Label_Сечение_кабели.Name = "Label_Сечение_кабели"
        Me.Label_Сечение_кабели.Size = New System.Drawing.Size(85, 28)
        Me.Label_Сечение_кабели.TabIndex = 4
        Me.Label_Сечение_кабели.Text = "0"
        Me.Label_Сечение_кабели.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Label_Процент_Запълване
        '
        Me.Label_Процент_Запълване.BackColor = System.Drawing.Color.Red
        Me.Label_Процент_Запълване.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label_Процент_Запълване.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label_Процент_Запълване.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label_Процент_Запълване.Location = New System.Drawing.Point(267, 67)
        Me.Label_Процент_Запълване.Name = "Label_Процент_Запълване"
        Me.Label_Процент_Запълване.Size = New System.Drawing.Size(32, 28)
        Me.Label_Процент_Запълване.TabIndex = 3
        Me.Label_Процент_Запълване.Text = "0"
        Me.Label_Процент_Запълване.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 28)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(117, 20)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Запълване, %"
        '
        'ComboBox_Процент_Запълване
        '
        Me.ComboBox_Процент_Запълване.FormattingEnabled = True
        Me.ComboBox_Процент_Запълване.Items.AddRange(New Object() {"5", "10", "15", "20", "25", "30", "35", "40", "45", "50", "55", "60", "65", "70", "75", "80", "85", "90", "95"})
        Me.ComboBox_Процент_Запълване.Location = New System.Drawing.Point(126, 25)
        Me.ComboBox_Процент_Запълване.Name = "ComboBox_Процент_Запълване"
        Me.ComboBox_Процент_Запълване.Size = New System.Drawing.Size(44, 28)
        Me.ComboBox_Процент_Запълване.TabIndex = 0
        Me.ComboBox_Процент_Запълване.Text = "40"
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.Button_В_4)
        Me.GroupBox4.Controls.Add(Me.Button_В_3)
        Me.GroupBox4.Controls.Add(Me.Button_В_2)
        Me.GroupBox4.Controls.Add(Me.Button_В_1)
        Me.GroupBox4.Location = New System.Drawing.Point(6, 254)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(435, 60)
        Me.GroupBox4.TabIndex = 3
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Височина, мм"
        '
        'Button_В_4
        '
        Me.Button_В_4.Location = New System.Drawing.Point(263, 17)
        Me.Button_В_4.Name = "Button_В_4"
        Me.Button_В_4.Size = New System.Drawing.Size(80, 35)
        Me.Button_В_4.TabIndex = 7
        Me.Button_В_4.Text = "100"
        Me.Button_В_4.UseVisualStyleBackColor = True
        '
        'Button_В_3
        '
        Me.Button_В_3.Location = New System.Drawing.Point(177, 17)
        Me.Button_В_3.Name = "Button_В_3"
        Me.Button_В_3.Size = New System.Drawing.Size(80, 35)
        Me.Button_В_3.TabIndex = 6
        Me.Button_В_3.Text = "85"
        Me.Button_В_3.UseVisualStyleBackColor = True
        '
        'Button_В_2
        '
        Me.Button_В_2.Location = New System.Drawing.Point(92, 17)
        Me.Button_В_2.Name = "Button_В_2"
        Me.Button_В_2.Size = New System.Drawing.Size(80, 35)
        Me.Button_В_2.TabIndex = 5
        Me.Button_В_2.Text = "60"
        Me.Button_В_2.UseVisualStyleBackColor = True
        '
        'Button_В_1
        '
        Me.Button_В_1.Location = New System.Drawing.Point(6, 17)
        Me.Button_В_1.Name = "Button_В_1"
        Me.Button_В_1.Size = New System.Drawing.Size(80, 35)
        Me.Button_В_1.TabIndex = 4
        Me.Button_В_1.Text = "50"
        Me.Button_В_1.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.BackColor = System.Drawing.SystemColors.Control
        Me.GroupBox3.Controls.Add(Me.Button_Ш_10)
        Me.GroupBox3.Controls.Add(Me.Button_Ш_9)
        Me.GroupBox3.Controls.Add(Me.Button_Ш_8)
        Me.GroupBox3.Controls.Add(Me.Button_Ш_5)
        Me.GroupBox3.Controls.Add(Me.Button_Ш_7)
        Me.GroupBox3.Controls.Add(Me.Button_Ш_6)
        Me.GroupBox3.Controls.Add(Me.Button_Ш_4)
        Me.GroupBox3.Controls.Add(Me.Button_Ш_3)
        Me.GroupBox3.Controls.Add(Me.Button_Ш_2)
        Me.GroupBox3.Controls.Add(Me.Button_Ш_1)
        Me.GroupBox3.Location = New System.Drawing.Point(6, 138)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(435, 110)
        Me.GroupBox3.TabIndex = 1
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Ширина, мм"
        '
        'Button_Ш_10
        '
        Me.Button_Ш_10.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Button_Ш_10.Location = New System.Drawing.Point(349, 66)
        Me.Button_Ш_10.Name = "Button_Ш_10"
        Me.Button_Ш_10.Size = New System.Drawing.Size(80, 35)
        Me.Button_Ш_10.TabIndex = 9
        Me.Button_Ш_10.Text = "300"
        Me.Button_Ш_10.UseVisualStyleBackColor = False
        '
        'Button_Ш_9
        '
        Me.Button_Ш_9.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Button_Ш_9.Location = New System.Drawing.Point(263, 66)
        Me.Button_Ш_9.Name = "Button_Ш_9"
        Me.Button_Ш_9.Size = New System.Drawing.Size(80, 35)
        Me.Button_Ш_9.TabIndex = 8
        Me.Button_Ш_9.Text = "200"
        Me.Button_Ш_9.UseVisualStyleBackColor = False
        '
        'Button_Ш_8
        '
        Me.Button_Ш_8.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Button_Ш_8.Location = New System.Drawing.Point(177, 66)
        Me.Button_Ш_8.Name = "Button_Ш_8"
        Me.Button_Ш_8.Size = New System.Drawing.Size(80, 35)
        Me.Button_Ш_8.TabIndex = 7
        Me.Button_Ш_8.Text = "600"
        Me.Button_Ш_8.UseVisualStyleBackColor = False
        '
        'Button_Ш_5
        '
        Me.Button_Ш_5.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Button_Ш_5.Location = New System.Drawing.Point(349, 25)
        Me.Button_Ш_5.Name = "Button_Ш_5"
        Me.Button_Ш_5.Size = New System.Drawing.Size(80, 35)
        Me.Button_Ш_5.TabIndex = 6
        Me.Button_Ш_5.Text = "300"
        Me.Button_Ш_5.UseVisualStyleBackColor = False
        '
        'Button_Ш_7
        '
        Me.Button_Ш_7.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Button_Ш_7.Location = New System.Drawing.Point(92, 66)
        Me.Button_Ш_7.Name = "Button_Ш_7"
        Me.Button_Ш_7.Size = New System.Drawing.Size(80, 35)
        Me.Button_Ш_7.TabIndex = 5
        Me.Button_Ш_7.Text = "500"
        Me.Button_Ш_7.UseVisualStyleBackColor = False
        '
        'Button_Ш_6
        '
        Me.Button_Ш_6.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Button_Ш_6.Location = New System.Drawing.Point(6, 66)
        Me.Button_Ш_6.Name = "Button_Ш_6"
        Me.Button_Ш_6.Size = New System.Drawing.Size(80, 35)
        Me.Button_Ш_6.TabIndex = 4
        Me.Button_Ш_6.Text = "400"
        Me.Button_Ш_6.UseVisualStyleBackColor = False
        '
        'Button_Ш_4
        '
        Me.Button_Ш_4.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Button_Ш_4.Location = New System.Drawing.Point(263, 25)
        Me.Button_Ш_4.Name = "Button_Ш_4"
        Me.Button_Ш_4.Size = New System.Drawing.Size(80, 35)
        Me.Button_Ш_4.TabIndex = 3
        Me.Button_Ш_4.Text = "200"
        Me.Button_Ш_4.UseVisualStyleBackColor = False
        '
        'Button_Ш_3
        '
        Me.Button_Ш_3.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Button_Ш_3.Location = New System.Drawing.Point(177, 25)
        Me.Button_Ш_3.Name = "Button_Ш_3"
        Me.Button_Ш_3.Size = New System.Drawing.Size(80, 35)
        Me.Button_Ш_3.TabIndex = 2
        Me.Button_Ш_3.Text = "150"
        Me.Button_Ш_3.UseVisualStyleBackColor = False
        '
        'Button_Ш_2
        '
        Me.Button_Ш_2.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Button_Ш_2.Location = New System.Drawing.Point(92, 25)
        Me.Button_Ш_2.Name = "Button_Ш_2"
        Me.Button_Ш_2.Size = New System.Drawing.Size(80, 35)
        Me.Button_Ш_2.TabIndex = 1
        Me.Button_Ш_2.Text = "100"
        Me.Button_Ш_2.UseVisualStyleBackColor = False
        '
        'Button_Ш_1
        '
        Me.Button_Ш_1.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Button_Ш_1.Location = New System.Drawing.Point(6, 25)
        Me.Button_Ш_1.Name = "Button_Ш_1"
        Me.Button_Ш_1.Size = New System.Drawing.Size(80, 35)
        Me.Button_Ш_1.TabIndex = 0
        Me.Button_Ш_1.Text = "50"
        Me.Button_Ш_1.UseVisualStyleBackColor = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.RadioButton_Тръба)
        Me.GroupBox2.Controls.Add(Me.RadioButton_Канал)
        Me.GroupBox2.Controls.Add(Me.RadioButton_Скара)
        Me.GroupBox2.Location = New System.Drawing.Point(535, 25)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(452, 59)
        Me.GroupBox2.TabIndex = 2
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Избери начин на полагане"
        '
        'RadioButton_Тръба
        '
        Me.RadioButton_Тръба.AutoSize = True
        Me.RadioButton_Тръба.Location = New System.Drawing.Point(299, 26)
        Me.RadioButton_Тръба.Name = "RadioButton_Тръба"
        Me.RadioButton_Тръба.Size = New System.Drawing.Size(74, 24)
        Me.RadioButton_Тръба.TabIndex = 2
        Me.RadioButton_Тръба.Text = "Тръба"
        Me.RadioButton_Тръба.UseVisualStyleBackColor = True
        '
        'RadioButton_Канал
        '
        Me.RadioButton_Канал.AutoSize = True
        Me.RadioButton_Канал.Location = New System.Drawing.Point(152, 25)
        Me.RadioButton_Канал.Name = "RadioButton_Канал"
        Me.RadioButton_Канал.Size = New System.Drawing.Size(141, 24)
        Me.RadioButton_Канал.TabIndex = 1
        Me.RadioButton_Канал.Text = "Кабелен канал"
        Me.RadioButton_Канал.UseVisualStyleBackColor = True
        '
        'RadioButton_Скара
        '
        Me.RadioButton_Скара.AutoSize = True
        Me.RadioButton_Скара.Location = New System.Drawing.Point(7, 26)
        Me.RadioButton_Скара.Name = "RadioButton_Скара"
        Me.RadioButton_Скара.Size = New System.Drawing.Size(139, 24)
        Me.RadioButton_Скара.TabIndex = 0
        Me.RadioButton_Скара.Text = "Кабелна скара"
        Me.RadioButton_Скара.UseVisualStyleBackColor = True
        '
        'Form_KabelniKanali
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(995, 445)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "Form_KabelniKanali"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "KabelniKanali"
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.DataGridView_Кабели, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox_Избор.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ToolStrip1 As Windows.Forms.ToolStrip
    Friend WithEvents NewToolStripButton As Windows.Forms.ToolStripButton
    Friend WithEvents OpenToolStripButton As Windows.Forms.ToolStripButton
    Friend WithEvents SaveToolStripButton As Windows.Forms.ToolStripButton
    Friend WithEvents PrintToolStripButton As Windows.Forms.ToolStripButton
    Friend WithEvents toolStripSeparator As Windows.Forms.ToolStripSeparator
    Friend WithEvents CutToolStripButton As Windows.Forms.ToolStripButton
    Friend WithEvents CopyToolStripButton As Windows.Forms.ToolStripButton
    Friend WithEvents PasteToolStripButton As Windows.Forms.ToolStripButton
    Friend WithEvents toolStripSeparator1 As Windows.Forms.ToolStripSeparator
    Friend WithEvents HelpToolStripButton As Windows.Forms.ToolStripButton
    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As Windows.Forms.GroupBox
    Friend WithEvents RadioButton_Канал As Windows.Forms.RadioButton
    Friend WithEvents RadioButton_Скара As Windows.Forms.RadioButton
    Friend WithEvents GroupBox_Избор As Windows.Forms.GroupBox
    Friend WithEvents RadioButton_Тръба As Windows.Forms.RadioButton
    Friend WithEvents GroupBox3 As Windows.Forms.GroupBox
    Friend WithEvents Button_Ш_1 As Windows.Forms.Button
    Friend WithEvents GroupBox4 As Windows.Forms.GroupBox
    Friend WithEvents Button_Ш_6 As Windows.Forms.Button
    Friend WithEvents Button_Ш_4 As Windows.Forms.Button
    Friend WithEvents Button_Ш_3 As Windows.Forms.Button
    Friend WithEvents Button_Ш_2 As Windows.Forms.Button
    Friend WithEvents Button_Ш_8 As Windows.Forms.Button
    Friend WithEvents Button_Ш_5 As Windows.Forms.Button
    Friend WithEvents Button_Ш_7 As Windows.Forms.Button
    Friend WithEvents Button_В_4 As Windows.Forms.Button
    Friend WithEvents Button_В_3 As Windows.Forms.Button
    Friend WithEvents Button_В_2 As Windows.Forms.Button
    Friend WithEvents Button_В_1 As Windows.Forms.Button
    Friend WithEvents GroupBox5 As Windows.Forms.GroupBox
    Friend WithEvents ComboBox_Процент_Запълване As Windows.Forms.ComboBox
    Friend WithEvents Label_Процент_Запълване As Windows.Forms.Label
    Friend WithEvents Label_Сечение_кабели As Windows.Forms.Label
    Friend WithEvents Label4 As Windows.Forms.Label
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents TextBox_Кабелна_Скара As Windows.Forms.TextBox
    Friend WithEvents DataGridView_Кабели As Windows.Forms.DataGridView
    Friend WithEvents Вид As Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents Жила As Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents Сечение As Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents Кабели As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Диаметър As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Площ As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Button_Ш_10 As Windows.Forms.Button
    Friend WithEvents Button_Ш_9 As Windows.Forms.Button
    Friend WithEvents ProgressBar_Procent As Windows.Forms.ProgressBar
End Class
