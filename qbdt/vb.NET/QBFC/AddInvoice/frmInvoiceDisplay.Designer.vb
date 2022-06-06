<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmInvoiceDisplay
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents txtInvoiceTotal As System.Windows.Forms.TextBox
	Public WithEvents txtInvoiceLines As System.Windows.Forms.TextBox
	Public WithEvents txtSalesTax As System.Windows.Forms.TextBox
	Public WithEvents txtSubtotal As System.Windows.Forms.TextBox
	Public WithEvents txtDueDate As System.Windows.Forms.TextBox
	Public WithEvents txtInvoiceDate As System.Windows.Forms.TextBox
	Public WithEvents txtInvoiceNumber As System.Windows.Forms.TextBox
	Public WithEvents cmdCloseWindow As System.Windows.Forms.Button
	Public WithEvents Label8 As System.Windows.Forms.Label
	Public WithEvents Label7 As System.Windows.Forms.Label
	Public WithEvents Label6 As System.Windows.Forms.Label
	Public WithEvents Label5 As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtInvoiceTotal = New System.Windows.Forms.TextBox()
        Me.txtInvoiceLines = New System.Windows.Forms.TextBox()
        Me.txtSalesTax = New System.Windows.Forms.TextBox()
        Me.txtSubtotal = New System.Windows.Forms.TextBox()
        Me.txtDueDate = New System.Windows.Forms.TextBox()
        Me.txtInvoiceDate = New System.Windows.Forms.TextBox()
        Me.txtInvoiceNumber = New System.Windows.Forms.TextBox()
        Me.cmdCloseWindow = New System.Windows.Forms.Button()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'txtInvoiceTotal
        '
        Me.txtInvoiceTotal.AcceptsReturn = True
        Me.txtInvoiceTotal.BackColor = System.Drawing.SystemColors.Window
        Me.txtInvoiceTotal.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtInvoiceTotal.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtInvoiceTotal.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtInvoiceTotal.Location = New System.Drawing.Point(96, 160)
        Me.txtInvoiceTotal.MaxLength = 0
        Me.txtInvoiceTotal.Name = "txtInvoiceTotal"
        Me.txtInvoiceTotal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtInvoiceTotal.Size = New System.Drawing.Size(113, 19)
        Me.txtInvoiceTotal.TabIndex = 15
        '
        'txtInvoiceLines
        '
        Me.txtInvoiceLines.AcceptsReturn = True
        Me.txtInvoiceLines.BackColor = System.Drawing.SystemColors.Window
        Me.txtInvoiceLines.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtInvoiceLines.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtInvoiceLines.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtInvoiceLines.Location = New System.Drawing.Point(8, 208)
        Me.txtInvoiceLines.MaxLength = 0
        Me.txtInvoiceLines.Multiline = True
        Me.txtInvoiceLines.Name = "txtInvoiceLines"
        Me.txtInvoiceLines.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtInvoiceLines.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtInvoiceLines.Size = New System.Drawing.Size(369, 121)
        Me.txtInvoiceLines.TabIndex = 12
        '
        'txtSalesTax
        '
        Me.txtSalesTax.AcceptsReturn = True
        Me.txtSalesTax.BackColor = System.Drawing.SystemColors.Window
        Me.txtSalesTax.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSalesTax.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSalesTax.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSalesTax.Location = New System.Drawing.Point(96, 136)
        Me.txtSalesTax.MaxLength = 0
        Me.txtSalesTax.Name = "txtSalesTax"
        Me.txtSalesTax.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSalesTax.Size = New System.Drawing.Size(113, 19)
        Me.txtSalesTax.TabIndex = 11
        '
        'txtSubtotal
        '
        Me.txtSubtotal.AcceptsReturn = True
        Me.txtSubtotal.BackColor = System.Drawing.SystemColors.Window
        Me.txtSubtotal.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSubtotal.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSubtotal.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSubtotal.Location = New System.Drawing.Point(96, 112)
        Me.txtSubtotal.MaxLength = 0
        Me.txtSubtotal.Name = "txtSubtotal"
        Me.txtSubtotal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSubtotal.Size = New System.Drawing.Size(113, 19)
        Me.txtSubtotal.TabIndex = 10
        '
        'txtDueDate
        '
        Me.txtDueDate.AcceptsReturn = True
        Me.txtDueDate.BackColor = System.Drawing.SystemColors.Window
        Me.txtDueDate.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDueDate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDueDate.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDueDate.Location = New System.Drawing.Point(96, 88)
        Me.txtDueDate.MaxLength = 0
        Me.txtDueDate.Name = "txtDueDate"
        Me.txtDueDate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDueDate.Size = New System.Drawing.Size(113, 19)
        Me.txtDueDate.TabIndex = 9
        '
        'txtInvoiceDate
        '
        Me.txtInvoiceDate.AcceptsReturn = True
        Me.txtInvoiceDate.BackColor = System.Drawing.SystemColors.Window
        Me.txtInvoiceDate.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtInvoiceDate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtInvoiceDate.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtInvoiceDate.Location = New System.Drawing.Point(96, 64)
        Me.txtInvoiceDate.MaxLength = 0
        Me.txtInvoiceDate.Name = "txtInvoiceDate"
        Me.txtInvoiceDate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtInvoiceDate.Size = New System.Drawing.Size(113, 19)
        Me.txtInvoiceDate.TabIndex = 8
        '
        'txtInvoiceNumber
        '
        Me.txtInvoiceNumber.AcceptsReturn = True
        Me.txtInvoiceNumber.BackColor = System.Drawing.SystemColors.Window
        Me.txtInvoiceNumber.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtInvoiceNumber.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtInvoiceNumber.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtInvoiceNumber.Location = New System.Drawing.Point(96, 40)
        Me.txtInvoiceNumber.MaxLength = 0
        Me.txtInvoiceNumber.Name = "txtInvoiceNumber"
        Me.txtInvoiceNumber.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtInvoiceNumber.Size = New System.Drawing.Size(113, 19)
        Me.txtInvoiceNumber.TabIndex = 7
        '
        'cmdCloseWindow
        '
        Me.cmdCloseWindow.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCloseWindow.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCloseWindow.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCloseWindow.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCloseWindow.Location = New System.Drawing.Point(136, 344)
        Me.cmdCloseWindow.Name = "cmdCloseWindow"
        Me.cmdCloseWindow.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCloseWindow.Size = New System.Drawing.Size(113, 41)
        Me.cmdCloseWindow.TabIndex = 6
        Me.cmdCloseWindow.Text = "Close Window"
        Me.cmdCloseWindow.UseVisualStyleBackColor = False
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.SystemColors.Control
        Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label8.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label8.Location = New System.Drawing.Point(8, 160)
        Me.Label8.Name = "Label8"
        Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label8.Size = New System.Drawing.Size(70, 17)
        Me.Label8.TabIndex = 14
        Me.Label8.Text = "Invoice Total"
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(16, 192)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(85, 17)
        Me.Label7.TabIndex = 13
        Me.Label7.Text = "Invoice Lines"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(8, 8)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(377, 25)
        Me.Label6.TabIndex = 5
        Me.Label6.Text = "Invoice Successfully Added"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(8, 136)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(70, 17)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "Sales Tax"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(8, 112)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(70, 17)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "Subtotal"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(8, 88)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(70, 17)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Due Date"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(8, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(75, 17)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Invoice Date"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(8, 40)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(81, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Invoice Number"
        '
        'frmInvoiceDisplay
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(388, 387)
        Me.Controls.Add(Me.txtInvoiceTotal)
        Me.Controls.Add(Me.txtInvoiceLines)
        Me.Controls.Add(Me.txtSalesTax)
        Me.Controls.Add(Me.txtSubtotal)
        Me.Controls.Add(Me.txtDueDate)
        Me.Controls.Add(Me.txtInvoiceDate)
        Me.Controls.Add(Me.txtInvoiceNumber)
        Me.Controls.Add(Me.cmdCloseWindow)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(292, 220)
        Me.Name = "frmInvoiceDisplay"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Invoice Display"
        Me.ResumeLayout(False)

    End Sub
#End Region
End Class