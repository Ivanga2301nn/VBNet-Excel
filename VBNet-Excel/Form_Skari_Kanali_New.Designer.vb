<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form_Skari_Kanali_New
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form_Skari_Kanali_New))
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
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.DataGridView_Кабели = New System.Windows.Forms.DataGridView()
        Me.Вид = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.Жила = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.Сечение = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.Кабели = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Диаметър = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Площ = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.ComboBox_Процент_Запълване = New System.Windows.Forms.ComboBox()
        Me.Label_Площ = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.RadioButton_Тръба = New System.Windows.Forms.RadioButton()
        Me.RadioButton_Канал = New System.Windows.Forms.RadioButton()
        Me.RadioButton_Скара = New System.Windows.Forms.RadioButton()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.NumericUpDown_Razdelitel = New System.Windows.Forms.NumericUpDown()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label_Skara = New System.Windows.Forms.Label()
        Me.TextBox_Кабелна_Скара = New System.Windows.Forms.TextBox()
        Me.GroupBox_Размери_Скари = New System.Windows.Forms.GroupBox()
        Me.TableLayoutPanel = New System.Windows.Forms.TableLayoutPanel()
        Me.Label_0_9 = New System.Windows.Forms.Label()
        Me.Label_0_8 = New System.Windows.Forms.Label()
        Me.Label_0_7 = New System.Windows.Forms.Label()
        Me.Label_0_6 = New System.Windows.Forms.Label()
        Me.Label_0_5 = New System.Windows.Forms.Label()
        Me.Label_0_4 = New System.Windows.Forms.Label()
        Me.Label_0_3 = New System.Windows.Forms.Label()
        Me.Label_0_2 = New System.Windows.Forms.Label()
        Me.Label_0_1 = New System.Windows.Forms.Label()
        Me.Label_1_1 = New System.Windows.Forms.Label()
        Me.Label_1_2 = New System.Windows.Forms.Label()
        Me.Label_1_3 = New System.Windows.Forms.Label()
        Me.Label_1_4 = New System.Windows.Forms.Label()
        Me.Label_1_5 = New System.Windows.Forms.Label()
        Me.Label_1_6 = New System.Windows.Forms.Label()
        Me.Label_1_7 = New System.Windows.Forms.Label()
        Me.Label_1_8 = New System.Windows.Forms.Label()
        Me.Label_1_9 = New System.Windows.Forms.Label()
        Me.ToolStrip1.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.DataGridView_Кабели, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.TableLayoutPanel2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        CType(Me.NumericUpDown_Razdelitel, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox_Размери_Скари.SuspendLayout()
        Me.TableLayoutPanel.SuspendLayout()
        Me.SuspendLayout()
        '
        'ToolStrip1
        '
        Me.ToolStrip1.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.NewToolStripButton, Me.OpenToolStripButton, Me.SaveToolStripButton, Me.PrintToolStripButton, Me.toolStripSeparator, Me.CutToolStripButton, Me.CopyToolStripButton, Me.PasteToolStripButton, Me.toolStripSeparator1, Me.HelpToolStripButton})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Padding = New System.Windows.Forms.Padding(0, 0, 2, 0)
        Me.ToolStrip1.Size = New System.Drawing.Size(1450, 25)
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
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 25)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.DataGridView_Кабели)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer2)
        Me.SplitContainer1.Size = New System.Drawing.Size(1450, 595)
        Me.SplitContainer1.SplitterDistance = 520
        Me.SplitContainer1.SplitterWidth = 6
        Me.SplitContainer1.TabIndex = 1
        '
        'DataGridView_Кабели
        '
        Me.DataGridView_Кабели.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView_Кабели.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Вид, Me.Жила, Me.Сечение, Me.Кабели, Me.Диаметър, Me.Площ})
        Me.DataGridView_Кабели.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView_Кабели.Location = New System.Drawing.Point(0, 0)
        Me.DataGridView_Кабели.Name = "DataGridView_Кабели"
        Me.DataGridView_Кабели.Size = New System.Drawing.Size(520, 595)
        Me.DataGridView_Кабели.TabIndex = 5
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
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.TableLayoutPanel2)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.GroupBox_Размери_Скари)
        Me.SplitContainer2.Size = New System.Drawing.Size(924, 595)
        Me.SplitContainer2.SplitterDistance = 216
        Me.SplitContainer2.SplitterWidth = 6
        Me.SplitContainer2.TabIndex = 0
        '
        'TableLayoutPanel2
        '
        Me.TableLayoutPanel2.ColumnCount = 2
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel2.Controls.Add(Me.GroupBox1, 0, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.GroupBox2, 0, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.GroupBox3, 1, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.GroupBox5, 0, 2)
        Me.TableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.RowCount = 3
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 60.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 60.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 60.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(924, 216)
        Me.TableLayoutPanel2.TabIndex = 0
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.SystemColors.Control
        Me.GroupBox1.Controls.Add(Me.ComboBox_Процент_Запълване)
        Me.GroupBox1.Controls.Add(Me.Label_Площ)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(3, 63)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(456, 54)
        Me.GroupBox1.TabIndex = 7
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Процент запълване"
        '
        'ComboBox_Процент_Запълване
        '
        Me.ComboBox_Процент_Запълване.Dock = System.Windows.Forms.DockStyle.Right
        Me.ComboBox_Процент_Запълване.FormattingEnabled = True
        Me.ComboBox_Процент_Запълване.Items.AddRange(New Object() {"5", "10", "15", "20", "25", "30", "35", "40", "45", "50", "55", "60", "65", "70", "75", "80", "85", "90", "95"})
        Me.ComboBox_Процент_Запълване.Location = New System.Drawing.Point(409, 22)
        Me.ComboBox_Процент_Запълване.Name = "ComboBox_Процент_Запълване"
        Me.ComboBox_Процент_Запълване.Size = New System.Drawing.Size(44, 28)
        Me.ComboBox_Процент_Запълване.TabIndex = 8
        Me.ComboBox_Процент_Запълване.Text = "40"
        '
        'Label_Площ
        '
        Me.Label_Площ.AutoSize = True
        Me.Label_Площ.Location = New System.Drawing.Point(18, 25)
        Me.Label_Площ.Name = "Label_Площ"
        Me.Label_Площ.Size = New System.Drawing.Size(39, 20)
        Me.Label_Площ.TabIndex = 7
        Me.Label_Площ.Text = ",mm"
        Me.Label_Площ.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'GroupBox2
        '
        Me.TableLayoutPanel2.SetColumnSpan(Me.GroupBox2, 2)
        Me.GroupBox2.Controls.Add(Me.RadioButton_Тръба)
        Me.GroupBox2.Controls.Add(Me.RadioButton_Канал)
        Me.GroupBox2.Controls.Add(Me.RadioButton_Скара)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox2.Location = New System.Drawing.Point(3, 3)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(918, 54)
        Me.GroupBox2.TabIndex = 3
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
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.NumericUpDown_Razdelitel)
        Me.GroupBox3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox3.Location = New System.Drawing.Point(465, 63)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(456, 54)
        Me.GroupBox3.TabIndex = 6
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Брой разделители"
        '
        'NumericUpDown_Razdelitel
        '
        Me.NumericUpDown_Razdelitel.Location = New System.Drawing.Point(3, 22)
        Me.NumericUpDown_Razdelitel.Name = "NumericUpDown_Razdelitel"
        Me.NumericUpDown_Razdelitel.Size = New System.Drawing.Size(285, 26)
        Me.NumericUpDown_Razdelitel.TabIndex = 0
        '
        'GroupBox5
        '
        Me.GroupBox5.BackColor = System.Drawing.SystemColors.Control
        Me.GroupBox5.Controls.Add(Me.Label5)
        Me.GroupBox5.Controls.Add(Me.Label_Skara)
        Me.GroupBox5.Controls.Add(Me.TextBox_Кабелна_Скара)
        Me.GroupBox5.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox5.Location = New System.Drawing.Point(3, 123)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(456, 90)
        Me.GroupBox5.TabIndex = 5
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Предложение"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(190, 28)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(35, 20)
        Me.Label5.TabIndex = 10
        Me.Label5.Text = "mm"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label_Skara
        '
        Me.Label_Skara.AutoSize = True
        Me.Label_Skara.Location = New System.Drawing.Point(18, 25)
        Me.Label_Skara.Name = "Label_Skara"
        Me.Label_Skara.Size = New System.Drawing.Size(98, 20)
        Me.Label_Skara.TabIndex = 7
        Me.Label_Skara.Text = "Скара [ШхВ]"
        Me.Label_Skara.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TextBox_Кабелна_Скара
        '
        Me.TextBox_Кабелна_Скара.Location = New System.Drawing.Point(122, 22)
        Me.TextBox_Кабелна_Скара.Name = "TextBox_Кабелна_Скара"
        Me.TextBox_Кабелна_Скара.Size = New System.Drawing.Size(69, 26)
        Me.TextBox_Кабелна_Скара.TabIndex = 9
        Me.TextBox_Кабелна_Скара.Text = "0x0"
        Me.TextBox_Кабелна_Скара.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'GroupBox_Размери_Скари
        '
        Me.GroupBox_Размери_Скари.Controls.Add(Me.TableLayoutPanel)
        Me.GroupBox_Размери_Скари.Location = New System.Drawing.Point(196, 45)
        Me.GroupBox_Размери_Скари.Name = "GroupBox_Размери_Скари"
        Me.GroupBox_Размери_Скари.Size = New System.Drawing.Size(557, 275)
        Me.GroupBox_Размери_Скари.TabIndex = 1
        Me.GroupBox_Размери_Скари.TabStop = False
        Me.GroupBox_Размери_Скари.Text = "Размер на кабелната скара, mm"
        '
        'TableLayoutPanel
        '
        Me.TableLayoutPanel.ColumnCount = 10
        Me.TableLayoutPanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel.Controls.Add(Me.Label_0_9, 0, 9)
        Me.TableLayoutPanel.Controls.Add(Me.Label_0_8, 0, 8)
        Me.TableLayoutPanel.Controls.Add(Me.Label_0_7, 0, 7)
        Me.TableLayoutPanel.Controls.Add(Me.Label_0_6, 0, 6)
        Me.TableLayoutPanel.Controls.Add(Me.Label_0_5, 0, 5)
        Me.TableLayoutPanel.Controls.Add(Me.Label_0_4, 0, 4)
        Me.TableLayoutPanel.Controls.Add(Me.Label_0_3, 0, 3)
        Me.TableLayoutPanel.Controls.Add(Me.Label_0_2, 0, 2)
        Me.TableLayoutPanel.Controls.Add(Me.Label_0_1, 0, 1)
        Me.TableLayoutPanel.Controls.Add(Me.Label_1_1, 1, 0)
        Me.TableLayoutPanel.Controls.Add(Me.Label_1_2, 2, 0)
        Me.TableLayoutPanel.Controls.Add(Me.Label_1_3, 3, 0)
        Me.TableLayoutPanel.Controls.Add(Me.Label_1_4, 4, 0)
        Me.TableLayoutPanel.Controls.Add(Me.Label_1_5, 5, 0)
        Me.TableLayoutPanel.Controls.Add(Me.Label_1_6, 6, 0)
        Me.TableLayoutPanel.Controls.Add(Me.Label_1_7, 7, 0)
        Me.TableLayoutPanel.Controls.Add(Me.Label_1_8, 8, 0)
        Me.TableLayoutPanel.Controls.Add(Me.Label_1_9, 9, 0)
        Me.TableLayoutPanel.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.TableLayoutPanel.Location = New System.Drawing.Point(3, 22)
        Me.TableLayoutPanel.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TableLayoutPanel.Name = "TableLayoutPanel"
        Me.TableLayoutPanel.RowCount = 10
        Me.TableLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel.Size = New System.Drawing.Size(551, 250)
        Me.TableLayoutPanel.TabIndex = 0
        '
        'Label_0_9
        '
        Me.Label_0_9.AutoSize = True
        Me.Label_0_9.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label_0_9.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label_0_9.Location = New System.Drawing.Point(3, 225)
        Me.Label_0_9.Name = "Label_0_9"
        Me.Label_0_9.Size = New System.Drawing.Size(49, 25)
        Me.Label_0_9.TabIndex = 59
        Me.Label_0_9.Text = "35"
        Me.Label_0_9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label_0_8
        '
        Me.Label_0_8.AutoSize = True
        Me.Label_0_8.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label_0_8.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label_0_8.Location = New System.Drawing.Point(3, 200)
        Me.Label_0_8.Name = "Label_0_8"
        Me.Label_0_8.Size = New System.Drawing.Size(49, 25)
        Me.Label_0_8.TabIndex = 58
        Me.Label_0_8.Text = "600"
        Me.Label_0_8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label_0_7
        '
        Me.Label_0_7.AutoSize = True
        Me.Label_0_7.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label_0_7.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label_0_7.Location = New System.Drawing.Point(3, 175)
        Me.Label_0_7.Name = "Label_0_7"
        Me.Label_0_7.Size = New System.Drawing.Size(49, 25)
        Me.Label_0_7.TabIndex = 57
        Me.Label_0_7.Text = "500"
        Me.Label_0_7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label_0_6
        '
        Me.Label_0_6.AutoSize = True
        Me.Label_0_6.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label_0_6.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label_0_6.Location = New System.Drawing.Point(3, 150)
        Me.Label_0_6.Name = "Label_0_6"
        Me.Label_0_6.Size = New System.Drawing.Size(49, 25)
        Me.Label_0_6.TabIndex = 56
        Me.Label_0_6.Text = "400"
        Me.Label_0_6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label_0_5
        '
        Me.Label_0_5.AutoSize = True
        Me.Label_0_5.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label_0_5.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label_0_5.Location = New System.Drawing.Point(3, 125)
        Me.Label_0_5.Name = "Label_0_5"
        Me.Label_0_5.Size = New System.Drawing.Size(49, 25)
        Me.Label_0_5.TabIndex = 55
        Me.Label_0_5.Text = "300"
        Me.Label_0_5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label_0_4
        '
        Me.Label_0_4.AutoSize = True
        Me.Label_0_4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label_0_4.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label_0_4.Location = New System.Drawing.Point(3, 100)
        Me.Label_0_4.Name = "Label_0_4"
        Me.Label_0_4.Size = New System.Drawing.Size(49, 25)
        Me.Label_0_4.TabIndex = 54
        Me.Label_0_4.Text = "200"
        Me.Label_0_4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label_0_3
        '
        Me.Label_0_3.AutoSize = True
        Me.Label_0_3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label_0_3.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label_0_3.Location = New System.Drawing.Point(3, 75)
        Me.Label_0_3.Name = "Label_0_3"
        Me.Label_0_3.Size = New System.Drawing.Size(49, 25)
        Me.Label_0_3.TabIndex = 53
        Me.Label_0_3.Text = "150"
        Me.Label_0_3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label_0_2
        '
        Me.Label_0_2.AutoSize = True
        Me.Label_0_2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label_0_2.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label_0_2.Location = New System.Drawing.Point(3, 50)
        Me.Label_0_2.Name = "Label_0_2"
        Me.Label_0_2.Size = New System.Drawing.Size(49, 25)
        Me.Label_0_2.TabIndex = 52
        Me.Label_0_2.Text = "100"
        Me.Label_0_2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label_0_1
        '
        Me.Label_0_1.AutoSize = True
        Me.Label_0_1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label_0_1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label_0_1.Location = New System.Drawing.Point(3, 25)
        Me.Label_0_1.Name = "Label_0_1"
        Me.Label_0_1.Size = New System.Drawing.Size(49, 25)
        Me.Label_0_1.TabIndex = 51
        Me.Label_0_1.Text = "50"
        Me.Label_0_1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label_1_1
        '
        Me.Label_1_1.AutoSize = True
        Me.Label_1_1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label_1_1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label_1_1.Location = New System.Drawing.Point(58, 0)
        Me.Label_1_1.Name = "Label_1_1"
        Me.Label_1_1.Size = New System.Drawing.Size(49, 25)
        Me.Label_1_1.TabIndex = 45
        Me.Label_1_1.Text = "35"
        Me.Label_1_1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label_1_2
        '
        Me.Label_1_2.AutoSize = True
        Me.Label_1_2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label_1_2.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label_1_2.Location = New System.Drawing.Point(113, 0)
        Me.Label_1_2.Name = "Label_1_2"
        Me.Label_1_2.Size = New System.Drawing.Size(49, 25)
        Me.Label_1_2.TabIndex = 46
        Me.Label_1_2.Text = "60"
        Me.Label_1_2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label_1_3
        '
        Me.Label_1_3.AutoSize = True
        Me.Label_1_3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label_1_3.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label_1_3.Location = New System.Drawing.Point(168, 0)
        Me.Label_1_3.Name = "Label_1_3"
        Me.Label_1_3.Size = New System.Drawing.Size(49, 25)
        Me.Label_1_3.TabIndex = 47
        Me.Label_1_3.Text = "85"
        Me.Label_1_3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label_1_4
        '
        Me.Label_1_4.AutoSize = True
        Me.Label_1_4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label_1_4.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label_1_4.Location = New System.Drawing.Point(223, 0)
        Me.Label_1_4.Name = "Label_1_4"
        Me.Label_1_4.Size = New System.Drawing.Size(49, 25)
        Me.Label_1_4.TabIndex = 48
        Me.Label_1_4.Text = "110"
        Me.Label_1_4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label_1_5
        '
        Me.Label_1_5.AutoSize = True
        Me.Label_1_5.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label_1_5.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label_1_5.Location = New System.Drawing.Point(278, 0)
        Me.Label_1_5.Name = "Label_1_5"
        Me.Label_1_5.Size = New System.Drawing.Size(49, 25)
        Me.Label_1_5.TabIndex = 60
        Me.Label_1_5.Text = "110"
        Me.Label_1_5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label_1_6
        '
        Me.Label_1_6.AutoSize = True
        Me.Label_1_6.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label_1_6.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label_1_6.Location = New System.Drawing.Point(333, 0)
        Me.Label_1_6.Name = "Label_1_6"
        Me.Label_1_6.Size = New System.Drawing.Size(49, 25)
        Me.Label_1_6.TabIndex = 61
        Me.Label_1_6.Text = "110"
        Me.Label_1_6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label_1_7
        '
        Me.Label_1_7.AutoSize = True
        Me.Label_1_7.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label_1_7.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label_1_7.Location = New System.Drawing.Point(388, 0)
        Me.Label_1_7.Name = "Label_1_7"
        Me.Label_1_7.Size = New System.Drawing.Size(49, 25)
        Me.Label_1_7.TabIndex = 62
        Me.Label_1_7.Text = "110"
        Me.Label_1_7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label_1_8
        '
        Me.Label_1_8.AutoSize = True
        Me.Label_1_8.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label_1_8.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label_1_8.Location = New System.Drawing.Point(443, 0)
        Me.Label_1_8.Name = "Label_1_8"
        Me.Label_1_8.Size = New System.Drawing.Size(49, 25)
        Me.Label_1_8.TabIndex = 63
        Me.Label_1_8.Text = "110"
        Me.Label_1_8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label_1_9
        '
        Me.Label_1_9.AutoSize = True
        Me.Label_1_9.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label_1_9.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label_1_9.Location = New System.Drawing.Point(498, 0)
        Me.Label_1_9.Name = "Label_1_9"
        Me.Label_1_9.Size = New System.Drawing.Size(50, 25)
        Me.Label_1_9.TabIndex = 64
        Me.Label_1_9.Text = "110"
        Me.Label_1_9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Form_Skari_Kanali_New
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1450, 620)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "Form_Skari_Kanali_New"
        Me.Text = "Skari_Kanali"
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.DataGridView_Кабели, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        CType(Me.NumericUpDown_Razdelitel, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.GroupBox_Размери_Скари.ResumeLayout(False)
        Me.TableLayoutPanel.ResumeLayout(False)
        Me.TableLayoutPanel.PerformLayout()
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
    Friend WithEvents SplitContainer1 As Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer2 As Windows.Forms.SplitContainer
    Friend WithEvents TableLayoutPanel As Windows.Forms.TableLayoutPanel
    Friend WithEvents GroupBox_Размери_Скари As Windows.Forms.GroupBox
    Friend WithEvents Label_0_9 As Windows.Forms.Label
    Friend WithEvents Label_0_8 As Windows.Forms.Label
    Friend WithEvents Label_0_7 As Windows.Forms.Label
    Friend WithEvents Label_0_6 As Windows.Forms.Label
    Friend WithEvents Label_0_5 As Windows.Forms.Label
    Friend WithEvents Label_0_4 As Windows.Forms.Label
    Friend WithEvents Label_0_3 As Windows.Forms.Label
    Friend WithEvents Label_0_2 As Windows.Forms.Label
    Friend WithEvents Label_0_1 As Windows.Forms.Label
    Friend WithEvents Label_1_1 As Windows.Forms.Label
    Friend WithEvents Label_1_2 As Windows.Forms.Label
    Friend WithEvents Label_1_3 As Windows.Forms.Label
    Friend WithEvents Label_1_4 As Windows.Forms.Label
    Friend WithEvents DataGridView_Кабели As Windows.Forms.DataGridView
    Friend WithEvents Вид As Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents Жила As Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents Сечение As Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents Кабели As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Диаметър As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Площ As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TableLayoutPanel2 As Windows.Forms.TableLayoutPanel
    Friend WithEvents GroupBox2 As Windows.Forms.GroupBox
    Friend WithEvents RadioButton_Тръба As Windows.Forms.RadioButton
    Friend WithEvents RadioButton_Канал As Windows.Forms.RadioButton
    Friend WithEvents RadioButton_Скара As Windows.Forms.RadioButton
    Friend WithEvents GroupBox5 As Windows.Forms.GroupBox
    Friend WithEvents Label_Skara As Windows.Forms.Label
    Friend WithEvents GroupBox3 As Windows.Forms.GroupBox
    Friend WithEvents NumericUpDown_Razdelitel As Windows.Forms.NumericUpDown
    Friend WithEvents TextBox_Кабелна_Скара As Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents Label_Площ As Windows.Forms.Label
    Friend WithEvents ComboBox_Процент_Запълване As Windows.Forms.ComboBox
    Friend WithEvents Label5 As Windows.Forms.Label
    Friend WithEvents Label_1_5 As Windows.Forms.Label
    Friend WithEvents Label_1_6 As Windows.Forms.Label
    Friend WithEvents Label_1_7 As Windows.Forms.Label
    Friend WithEvents Label_1_8 As Windows.Forms.Label
    Friend WithEvents Label_1_9 As Windows.Forms.Label
End Class
