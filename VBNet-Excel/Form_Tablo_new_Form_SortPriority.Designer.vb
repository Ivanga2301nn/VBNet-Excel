<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form_SortPriority
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
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.btnUp = New System.Windows.Forms.Button()
        Me.btnDown = New System.Windows.Forms.Button()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.lstPrefixes = New System.Windows.Forms.ListBox()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 74.47552!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25.52448!))
        Me.TableLayoutPanel1.Controls.Add(Me.btnUp, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.btnDown, 1, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.btnOK, 1, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.btnCancel, 1, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.lstPrefixes, 0, 0)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 5
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(284, 361)
        Me.TableLayoutPanel1.TabIndex = 1
        '
        'btnUp
        '
        Me.btnUp.Dock = System.Windows.Forms.DockStyle.Fill
        Me.btnUp.Location = New System.Drawing.Point(214, 3)
        Me.btnUp.Name = "btnUp"
        Me.btnUp.Size = New System.Drawing.Size(67, 24)
        Me.btnUp.TabIndex = 0
        Me.btnUp.Text = "▲ Нагоре"
        Me.btnUp.UseVisualStyleBackColor = True
        '
        'btnDown
        '
        Me.btnDown.Dock = System.Windows.Forms.DockStyle.Fill
        Me.btnDown.Location = New System.Drawing.Point(214, 33)
        Me.btnDown.Name = "btnDown"
        Me.btnDown.Size = New System.Drawing.Size(67, 24)
        Me.btnDown.TabIndex = 1
        Me.btnDown.Text = "▼ Надолу"
        Me.btnDown.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Dock = System.Windows.Forms.DockStyle.Fill
        Me.btnOK.Location = New System.Drawing.Point(214, 63)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(67, 24)
        Me.btnOK.TabIndex = 2
        Me.btnOK.Text = "Сортирай"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Dock = System.Windows.Forms.DockStyle.Fill
        Me.btnCancel.Location = New System.Drawing.Point(214, 93)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(67, 24)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "Отказ"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'lstPrefixes
        '
        Me.lstPrefixes.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstPrefixes.FormattingEnabled = True
        Me.lstPrefixes.Location = New System.Drawing.Point(3, 3)
        Me.lstPrefixes.Name = "lstPrefixes"
        Me.TableLayoutPanel1.SetRowSpan(Me.lstPrefixes, 5)
        Me.lstPrefixes.Size = New System.Drawing.Size(205, 355)
        Me.lstPrefixes.TabIndex = 5
        '
        'Form_SortPriority
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 361)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Name = "Form_SortPriority"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "СОРТИРАНЕ ТОКОВИ КРЪГОВЕ"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TableLayoutPanel1 As Windows.Forms.TableLayoutPanel
    Friend WithEvents btnUp As Windows.Forms.Button
    Friend WithEvents btnDown As Windows.Forms.Button
    Friend WithEvents btnOK As Windows.Forms.Button
    Friend WithEvents btnCancel As Windows.Forms.Button
    Friend WithEvents lstPrefixes As Windows.Forms.ListBox
End Class
