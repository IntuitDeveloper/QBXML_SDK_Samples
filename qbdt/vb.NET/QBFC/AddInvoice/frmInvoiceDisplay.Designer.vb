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
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmInvoiceDisplay))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.txtInvoiceTotal = New System.Windows.Forms.TextBox
		Me.txtInvoiceLines = New System.Windows.Forms.TextBox
		Me.txtSalesTax = New System.Windows.Forms.TextBox
		Me.txtSubtotal = New System.Windows.Forms.TextBox
		Me.txtDueDate = New System.Windows.Forms.TextBox
		Me.txtInvoiceDate = New System.Windows.Forms.TextBox
		Me.txtInvoiceNumber = New System.Windows.Forms.TextBox
		Me.cmdCloseWindow = New System.Windows.Forms.Button
		Me.Label8 = New System.Windows.Forms.Label
		Me.Label7 = New System.Windows.Forms.Label
		Me.Label6 = New System.Windows.Forms.Label
		Me.Label5 = New System.Windows.Forms.Label
		Me.Label4 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.Text = "Invoice Display"
		Me.ClientSize = New System.Drawing.Size(388, 387)
		Me.Location = New System.Drawing.Point(292, 220)
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MaximizeBox = True
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmInvoiceDisplay"
		Me.txtInvoiceTotal.AutoSize = False
		Me.txtInvoiceTotal.Size = New System.Drawing.Size(113, 19)
		Me.txtInvoiceTotal.Location = New System.Drawing.Point(96, 160)
		Me.txtInvoiceTotal.TabIndex = 15
		Me.txtInvoiceTotal.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtInvoiceTotal.AcceptsReturn = True
		Me.txtInvoiceTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtInvoiceTotal.BackColor = System.Drawing.SystemColors.Window
		Me.txtInvoiceTotal.CausesValidation = True
		Me.txtInvoiceTotal.Enabled = True
		Me.txtInvoiceTotal.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtInvoiceTotal.HideSelection = True
		Me.txtInvoiceTotal.ReadOnly = False
		Me.txtInvoiceTotal.Maxlength = 0
		Me.txtInvoiceTotal.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtInvoiceTotal.MultiLine = False
		Me.txtInvoiceTotal.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtInvoiceTotal.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtInvoiceTotal.TabStop = True
		Me.txtInvoiceTotal.Visible = True
		Me.txtInvoiceTotal.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtInvoiceTotal.Name = "txtInvoiceTotal"
		Me.txtInvoiceLines.AutoSize = False
		Me.txtInvoiceLines.Size = New System.Drawing.Size(369, 121)
		Me.txtInvoiceLines.Location = New System.Drawing.Point(8, 208)
		Me.txtInvoiceLines.MultiLine = True
		Me.txtInvoiceLines.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
		Me.txtInvoiceLines.TabIndex = 12
		Me.txtInvoiceLines.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtInvoiceLines.AcceptsReturn = True
		Me.txtInvoiceLines.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtInvoiceLines.BackColor = System.Drawing.SystemColors.Window
		Me.txtInvoiceLines.CausesValidation = True
		Me.txtInvoiceLines.Enabled = True
		Me.txtInvoiceLines.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtInvoiceLines.HideSelection = True
		Me.txtInvoiceLines.ReadOnly = False
		Me.txtInvoiceLines.Maxlength = 0
		Me.txtInvoiceLines.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtInvoiceLines.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtInvoiceLines.TabStop = True
		Me.txtInvoiceLines.Visible = True
		Me.txtInvoiceLines.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtInvoiceLines.Name = "txtInvoiceLines"
		Me.txtSalesTax.AutoSize = False
		Me.txtSalesTax.Size = New System.Drawing.Size(113, 19)
		Me.txtSalesTax.Location = New System.Drawing.Point(96, 136)
		Me.txtSalesTax.TabIndex = 11
		Me.txtSalesTax.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtSalesTax.AcceptsReturn = True
		Me.txtSalesTax.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtSalesTax.BackColor = System.Drawing.SystemColors.Window
		Me.txtSalesTax.CausesValidation = True
		Me.txtSalesTax.Enabled = True
		Me.txtSalesTax.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtSalesTax.HideSelection = True
		Me.txtSalesTax.ReadOnly = False
		Me.txtSalesTax.Maxlength = 0
		Me.txtSalesTax.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtSalesTax.MultiLine = False
		Me.txtSalesTax.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtSalesTax.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtSalesTax.TabStop = True
		Me.txtSalesTax.Visible = True
		Me.txtSalesTax.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtSalesTax.Name = "txtSalesTax"
		Me.txtSubtotal.AutoSize = False
		Me.txtSubtotal.Size = New System.Drawing.Size(113, 19)
		Me.txtSubtotal.Location = New System.Drawing.Point(96, 112)
		Me.txtSubtotal.TabIndex = 10
		Me.txtSubtotal.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtSubtotal.AcceptsReturn = True
		Me.txtSubtotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtSubtotal.BackColor = System.Drawing.SystemColors.Window
		Me.txtSubtotal.CausesValidation = True
		Me.txtSubtotal.Enabled = True
		Me.txtSubtotal.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtSubtotal.HideSelection = True
		Me.txtSubtotal.ReadOnly = False
		Me.txtSubtotal.Maxlength = 0
		Me.txtSubtotal.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtSubtotal.MultiLine = False
		Me.txtSubtotal.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtSubtotal.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtSubtotal.TabStop = True
		Me.txtSubtotal.Visible = True
		Me.txtSubtotal.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtSubtotal.Name = "txtSubtotal"
		Me.txtDueDate.AutoSize = False
		Me.txtDueDate.Size = New System.Drawing.Size(113, 19)
		Me.txtDueDate.Location = New System.Drawing.Point(96, 88)
		Me.txtDueDate.TabIndex = 9
		Me.txtDueDate.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtDueDate.AcceptsReturn = True
		Me.txtDueDate.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtDueDate.BackColor = System.Drawing.SystemColors.Window
		Me.txtDueDate.CausesValidation = True
		Me.txtDueDate.Enabled = True
		Me.txtDueDate.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtDueDate.HideSelection = True
		Me.txtDueDate.ReadOnly = False
		Me.txtDueDate.Maxlength = 0
		Me.txtDueDate.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtDueDate.MultiLine = False
		Me.txtDueDate.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtDueDate.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtDueDate.TabStop = True
		Me.txtDueDate.Visible = True
		Me.txtDueDate.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtDueDate.Name = "txtDueDate"
		Me.txtInvoiceDate.AutoSize = False
		Me.txtInvoiceDate.Size = New System.Drawing.Size(113, 19)
		Me.txtInvoiceDate.Location = New System.Drawing.Point(96, 64)
		Me.txtInvoiceDate.TabIndex = 8
		Me.txtInvoiceDate.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtInvoiceDate.AcceptsReturn = True
		Me.txtInvoiceDate.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtInvoiceDate.BackColor = System.Drawing.SystemColors.Window
		Me.txtInvoiceDate.CausesValidation = True
		Me.txtInvoiceDate.Enabled = True
		Me.txtInvoiceDate.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtInvoiceDate.HideSelection = True
		Me.txtInvoiceDate.ReadOnly = False
		Me.txtInvoiceDate.Maxlength = 0
		Me.txtInvoiceDate.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtInvoiceDate.MultiLine = False
		Me.txtInvoiceDate.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtInvoiceDate.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtInvoiceDate.TabStop = True
		Me.txtInvoiceDate.Visible = True
		Me.txtInvoiceDate.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtInvoiceDate.Name = "txtInvoiceDate"
		Me.txtInvoiceNumber.AutoSize = False
		Me.txtInvoiceNumber.Size = New System.Drawing.Size(113, 19)
		Me.txtInvoiceNumber.Location = New System.Drawing.Point(96, 40)
		Me.txtInvoiceNumber.TabIndex = 7
		Me.txtInvoiceNumber.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtInvoiceNumber.AcceptsReturn = True
		Me.txtInvoiceNumber.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtInvoiceNumber.BackColor = System.Drawing.SystemColors.Window
		Me.txtInvoiceNumber.CausesValidation = True
		Me.txtInvoiceNumber.Enabled = True
		Me.txtInvoiceNumber.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtInvoiceNumber.HideSelection = True
		Me.txtInvoiceNumber.ReadOnly = False
		Me.txtInvoiceNumber.Maxlength = 0
		Me.txtInvoiceNumber.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtInvoiceNumber.MultiLine = False
		Me.txtInvoiceNumber.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtInvoiceNumber.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtInvoiceNumber.TabStop = True
		Me.txtInvoiceNumber.Visible = True
		Me.txtInvoiceNumber.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtInvoiceNumber.Name = "txtInvoiceNumber"
		Me.cmdCloseWindow.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdCloseWindow.Text = "Close Window"
		Me.cmdCloseWindow.Size = New System.Drawing.Size(113, 41)
		Me.cmdCloseWindow.Location = New System.Drawing.Point(136, 344)
		Me.cmdCloseWindow.TabIndex = 6
		Me.cmdCloseWindow.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCloseWindow.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCloseWindow.CausesValidation = True
		Me.cmdCloseWindow.Enabled = True
		Me.cmdCloseWindow.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCloseWindow.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCloseWindow.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCloseWindow.TabStop = True
		Me.cmdCloseWindow.Name = "cmdCloseWindow"
		Me.Label8.Text = "Invoice Total"
		Me.Label8.Size = New System.Drawing.Size(65, 17)
		Me.Label8.Location = New System.Drawing.Point(8, 160)
		Me.Label8.TabIndex = 14
		Me.Label8.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label8.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label8.BackColor = System.Drawing.SystemColors.Control
		Me.Label8.Enabled = True
		Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label8.UseMnemonic = True
		Me.Label8.Visible = True
		Me.Label8.AutoSize = False
		Me.Label8.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label8.Name = "Label8"
		Me.Label7.Text = "Invoice Lines"
		Me.Label7.Size = New System.Drawing.Size(65, 17)
		Me.Label7.Location = New System.Drawing.Point(16, 192)
		Me.Label7.TabIndex = 13
		Me.Label7.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label7.BackColor = System.Drawing.SystemColors.Control
		Me.Label7.Enabled = True
		Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label7.UseMnemonic = True
		Me.Label7.Visible = True
		Me.Label7.AutoSize = False
		Me.Label7.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label7.Name = "Label7"
		Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.Label6.Text = "Invoice Successfully Added"
		Me.Label6.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label6.Size = New System.Drawing.Size(377, 25)
		Me.Label6.Location = New System.Drawing.Point(8, 8)
		Me.Label6.TabIndex = 5
		Me.Label6.BackColor = System.Drawing.SystemColors.Control
		Me.Label6.Enabled = True
		Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label6.UseMnemonic = True
		Me.Label6.Visible = True
		Me.Label6.AutoSize = False
		Me.Label6.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label6.Name = "Label6"
		Me.Label5.Text = "Sales Tax"
		Me.Label5.Size = New System.Drawing.Size(49, 17)
		Me.Label5.Location = New System.Drawing.Point(8, 136)
		Me.Label5.TabIndex = 4
		Me.Label5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label5.BackColor = System.Drawing.SystemColors.Control
		Me.Label5.Enabled = True
		Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label5.UseMnemonic = True
		Me.Label5.Visible = True
		Me.Label5.AutoSize = False
		Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label5.Name = "Label5"
		Me.Label4.Text = "Subtotal"
		Me.Label4.Size = New System.Drawing.Size(41, 17)
		Me.Label4.Location = New System.Drawing.Point(8, 112)
		Me.Label4.TabIndex = 3
		Me.Label4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label4.BackColor = System.Drawing.SystemColors.Control
		Me.Label4.Enabled = True
		Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label4.UseMnemonic = True
		Me.Label4.Visible = True
		Me.Label4.AutoSize = False
		Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label4.Name = "Label4"
		Me.Label3.Text = "Due Date"
		Me.Label3.Size = New System.Drawing.Size(49, 17)
		Me.Label3.Location = New System.Drawing.Point(8, 88)
		Me.Label3.TabIndex = 2
		Me.Label3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label3.BackColor = System.Drawing.SystemColors.Control
		Me.Label3.Enabled = True
		Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label3.UseMnemonic = True
		Me.Label3.Visible = True
		Me.Label3.AutoSize = False
		Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label3.Name = "Label3"
		Me.Label2.Text = "Invoice Date"
		Me.Label2.Size = New System.Drawing.Size(65, 17)
		Me.Label2.Location = New System.Drawing.Point(8, 64)
		Me.Label2.TabIndex = 1
		Me.Label2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label2.BackColor = System.Drawing.SystemColors.Control
		Me.Label2.Enabled = True
		Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label2.UseMnemonic = True
		Me.Label2.Visible = True
		Me.Label2.AutoSize = False
		Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label2.Name = "Label2"
		Me.Label1.Text = "Invoice Number"
		Me.Label1.Size = New System.Drawing.Size(81, 17)
		Me.Label1.Location = New System.Drawing.Point(8, 40)
		Me.Label1.TabIndex = 0
		Me.Label1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.BackColor = System.Drawing.SystemColors.Control
		Me.Label1.Enabled = True
		Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Controls.Add(txtInvoiceTotal)
		Me.Controls.Add(txtInvoiceLines)
		Me.Controls.Add(txtSalesTax)
		Me.Controls.Add(txtSubtotal)
		Me.Controls.Add(txtDueDate)
		Me.Controls.Add(txtInvoiceDate)
		Me.Controls.Add(txtInvoiceNumber)
		Me.Controls.Add(cmdCloseWindow)
		Me.Controls.Add(Label8)
		Me.Controls.Add(Label7)
		Me.Controls.Add(Label6)
		Me.Controls.Add(Label5)
		Me.Controls.Add(Label4)
		Me.Controls.Add(Label3)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label1)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class