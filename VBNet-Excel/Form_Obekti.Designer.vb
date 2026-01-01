<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Obekti
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.DataGridView = New System.Windows.Forms.DataGridView()
        Me.Занчение = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Обект = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.butGet = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.butExit = New System.Windows.Forms.Button()
        CType(Me.DataGridView, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView
        '
        Me.DataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Занчение, Me.Обект})
        Me.DataGridView.Dock = System.Windows.Forms.DockStyle.Top
        Me.DataGridView.Location = New System.Drawing.Point(0, 0)
        Me.DataGridView.Name = "DataGridView"
        Me.DataGridView.Size = New System.Drawing.Size(1233, 493)
        Me.DataGridView.TabIndex = 0
        '
        'Занчение
        '
        Me.Занчение.HeaderText = ""
        Me.Занчение.Name = "Занчение"
        '
        'Обект
        '
        Me.Обект.HeaderText = ""
        Me.Обект.Name = "Обект"
        '
        'butGet
        '
        Me.butGet.Location = New System.Drawing.Point(0, 499)
        Me.butGet.Name = "butGet"
        Me.butGet.Size = New System.Drawing.Size(75, 23)
        Me.butGet.TabIndex = 1
        Me.butGet.Text = "Вземи данни"
        Me.butGet.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(82, 499)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "Button2"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'butExit
        '
        Me.butExit.Location = New System.Drawing.Point(163, 499)
        Me.butExit.Name = "butExit"
        Me.butExit.Size = New System.Drawing.Size(75, 23)
        Me.butExit.TabIndex = 3
        Me.butExit.Text = "Изход"
        Me.butExit.UseVisualStyleBackColor = True
        '
        'Obekti
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1233, 534)
        Me.Controls.Add(Me.butExit)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.butGet)
        Me.Controls.Add(Me.DataGridView)
        Me.Name = "Obekti"
        Me.Text = "Obekti"
        CType(Me.DataGridView, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents DataGridView As Windows.Forms.DataGridView
    Friend WithEvents Занчение As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Обект As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents butGet As Windows.Forms.Button
    Friend WithEvents Button2 As Windows.Forms.Button
    Friend WithEvents butExit As Windows.Forms.Button
End Class
