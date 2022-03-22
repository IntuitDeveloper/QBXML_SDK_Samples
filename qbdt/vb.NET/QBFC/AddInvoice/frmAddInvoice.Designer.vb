<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmAddInvoice
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
	Public WithEvents cmdExit As System.Windows.Forms.Button
	Public WithEvents cmdAddInvoice As System.Windows.Forms.Button
	Public WithEvents optQBFC As System.Windows.Forms.RadioButton
	Public WithEvents optqbXMLRP As System.Windows.Forms.RadioButton
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmAddInvoice))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdExit = New System.Windows.Forms.Button
		Me.cmdAddInvoice = New System.Windows.Forms.Button
		Me.optQBFC = New System.Windows.Forms.RadioButton
		Me.optqbXMLRP = New System.Windows.Forms.RadioButton
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.Text = "Add Invoice Sample"
		Me.ClientSize = New System.Drawing.Size(406, 294)
		Me.Location = New System.Drawing.Point(282, 230)
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
		Me.Name = "frmAddInvoice"
		Me.cmdExit.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdExit.Text = "Exit"
		Me.cmdExit.Size = New System.Drawing.Size(113, 57)
		Me.cmdExit.Location = New System.Drawing.Point(272, 216)
		Me.cmdExit.TabIndex = 5
		Me.cmdExit.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdExit.BackColor = System.Drawing.SystemColors.Control
		Me.cmdExit.CausesValidation = True
		Me.cmdExit.Enabled = True
		Me.cmdExit.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdExit.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdExit.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdExit.TabStop = True
		Me.cmdExit.Name = "cmdExit"
		Me.cmdAddInvoice.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdAddInvoice.Text = "Add Invoice"
		Me.cmdAddInvoice.Size = New System.Drawing.Size(113, 57)
		Me.cmdAddInvoice.Location = New System.Drawing.Point(128, 216)
		Me.cmdAddInvoice.TabIndex = 4
		Me.cmdAddInvoice.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdAddInvoice.BackColor = System.Drawing.SystemColors.Control
		Me.cmdAddInvoice.CausesValidation = True
		Me.cmdAddInvoice.Enabled = True
		Me.cmdAddInvoice.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdAddInvoice.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdAddInvoice.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdAddInvoice.TabStop = True
		Me.cmdAddInvoice.Name = "cmdAddInvoice"
		Me.optQBFC.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optQBFC.Text = "Use QBFC"
		Me.optQBFC.Size = New System.Drawing.Size(97, 17)
		Me.optQBFC.Location = New System.Drawing.Point(16, 248)
		Me.optQBFC.TabIndex = 1
		Me.optQBFC.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optQBFC.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optQBFC.BackColor = System.Drawing.SystemColors.Control
		Me.optQBFC.CausesValidation = True
		Me.optQBFC.Enabled = True
		Me.optQBFC.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optQBFC.Cursor = System.Windows.Forms.Cursors.Default
		Me.optQBFC.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optQBFC.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optQBFC.TabStop = True
		Me.optQBFC.Checked = False
		Me.optQBFC.Visible = True
		Me.optQBFC.Name = "optQBFC"
		Me.optqbXMLRP.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optqbXMLRP.Text = "Use qbXMLRP"
		Me.optqbXMLRP.Size = New System.Drawing.Size(97, 17)
		Me.optqbXMLRP.Location = New System.Drawing.Point(16, 216)
		Me.optqbXMLRP.TabIndex = 0
		Me.optqbXMLRP.Checked = True
		Me.optqbXMLRP.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optqbXMLRP.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optqbXMLRP.BackColor = System.Drawing.SystemColors.Control
		Me.optqbXMLRP.CausesValidation = True
		Me.optqbXMLRP.Enabled = True
		Me.optqbXMLRP.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optqbXMLRP.Cursor = System.Windows.Forms.Cursors.Default
		Me.optqbXMLRP.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optqbXMLRP.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optqbXMLRP.TabStop = True
		Me.optqbXMLRP.Visible = True
		Me.optqbXMLRP.Name = "optqbXMLRP"
        Me.Label2.Text = "You may choose to have the program use either qbXML and the MSXML6 DOM parser to build the message or QBFC to build the message."
        Me.Label2.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Size = New System.Drawing.Size(393, 49)
		Me.Label2.Location = New System.Drawing.Point(8, 136)
		Me.Label2.TabIndex = 3
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
		Me.Label1.Text = "This sample program adds an invoice to the sample product-based company file which comes with QuickBooks.  The program adds the invoice for customer Abercrombie, Kristy with the shipping address set to a different value than that in the customer record.  The invoice has three lines including one service item, one non-inventory item and one item group."
		Me.Label1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.Size = New System.Drawing.Size(385, 97)
		Me.Label1.Location = New System.Drawing.Point(8, 24)
		Me.Label1.TabIndex = 2
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
		Me.Controls.Add(cmdExit)
		Me.Controls.Add(cmdAddInvoice)
		Me.Controls.Add(optQBFC)
		Me.Controls.Add(optqbXMLRP)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label1)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class