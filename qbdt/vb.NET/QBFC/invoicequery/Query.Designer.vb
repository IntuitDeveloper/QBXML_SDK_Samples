<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class Query
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
	Public WithEvents UseQBOE As System.Windows.Forms.CheckBox
	Public WithEvents IncludeLineItems As System.Windows.Forms.CheckBox
	Public WithEvents Exit_Renamed As System.Windows.Forms.Button
	Public WithEvents Submit As System.Windows.Forms.Button
	Public WithEvents ToTxnDate As System.Windows.Forms.TextBox
	Public WithEvents FromTxnDate As System.Windows.Forms.TextBox
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Query))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.UseQBOE = New System.Windows.Forms.CheckBox
		Me.IncludeLineItems = New System.Windows.Forms.CheckBox
		Me.Exit_Renamed = New System.Windows.Forms.Button
		Me.Submit = New System.Windows.Forms.Button
		Me.ToTxnDate = New System.Windows.Forms.TextBox
		Me.FromTxnDate = New System.Windows.Forms.TextBox
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "Invoice Query"
		Me.ClientSize = New System.Drawing.Size(318, 213)
		Me.Location = New System.Drawing.Point(4, 23)
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
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
		Me.Name = "Query"
		Me.UseQBOE.Text = "Communicate with QuickBooks Online Edition"
		Me.UseQBOE.Size = New System.Drawing.Size(265, 25)
		Me.UseQBOE.Location = New System.Drawing.Point(16, 128)
		Me.UseQBOE.TabIndex = 7
		Me.UseQBOE.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.UseQBOE.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.UseQBOE.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.UseQBOE.BackColor = System.Drawing.SystemColors.Control
		Me.UseQBOE.CausesValidation = True
		Me.UseQBOE.Enabled = True
		Me.UseQBOE.ForeColor = System.Drawing.SystemColors.ControlText
		Me.UseQBOE.Cursor = System.Windows.Forms.Cursors.Default
		Me.UseQBOE.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.UseQBOE.Appearance = System.Windows.Forms.Appearance.Normal
		Me.UseQBOE.TabStop = True
		Me.UseQBOE.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.UseQBOE.Visible = True
		Me.UseQBOE.Name = "UseQBOE"
		Me.IncludeLineItems.Text = "Include Line Items"
		Me.IncludeLineItems.Size = New System.Drawing.Size(177, 17)
		Me.IncludeLineItems.Location = New System.Drawing.Point(16, 96)
		Me.IncludeLineItems.TabIndex = 6
		Me.IncludeLineItems.CheckState = System.Windows.Forms.CheckState.Checked
		Me.IncludeLineItems.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.IncludeLineItems.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.IncludeLineItems.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.IncludeLineItems.BackColor = System.Drawing.SystemColors.Control
		Me.IncludeLineItems.CausesValidation = True
		Me.IncludeLineItems.Enabled = True
		Me.IncludeLineItems.ForeColor = System.Drawing.SystemColors.ControlText
		Me.IncludeLineItems.Cursor = System.Windows.Forms.Cursors.Default
		Me.IncludeLineItems.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.IncludeLineItems.Appearance = System.Windows.Forms.Appearance.Normal
		Me.IncludeLineItems.TabStop = True
		Me.IncludeLineItems.Visible = True
		Me.IncludeLineItems.Name = "IncludeLineItems"
		Me.Exit_Renamed.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Exit_Renamed.Text = "Exit"
		Me.Exit_Renamed.Size = New System.Drawing.Size(105, 33)
		Me.Exit_Renamed.Location = New System.Drawing.Point(176, 168)
		Me.Exit_Renamed.TabIndex = 5
		Me.Exit_Renamed.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Exit_Renamed.BackColor = System.Drawing.SystemColors.Control
		Me.Exit_Renamed.CausesValidation = True
		Me.Exit_Renamed.Enabled = True
		Me.Exit_Renamed.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Exit_Renamed.Cursor = System.Windows.Forms.Cursors.Default
		Me.Exit_Renamed.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Exit_Renamed.TabStop = True
		Me.Exit_Renamed.Name = "Exit_Renamed"
		Me.Submit.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Submit.Text = "Submit"
		Me.AcceptButton = Me.Submit
		Me.Submit.Size = New System.Drawing.Size(105, 33)
		Me.Submit.Location = New System.Drawing.Point(32, 168)
		Me.Submit.TabIndex = 4
		Me.Submit.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Submit.BackColor = System.Drawing.SystemColors.Control
		Me.Submit.CausesValidation = True
		Me.Submit.Enabled = True
		Me.Submit.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Submit.Cursor = System.Windows.Forms.Cursors.Default
		Me.Submit.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Submit.TabStop = True
		Me.Submit.Name = "Submit"
		Me.ToTxnDate.AutoSize = False
		Me.ToTxnDate.Size = New System.Drawing.Size(113, 19)
		Me.ToTxnDate.Location = New System.Drawing.Point(176, 64)
		Me.ToTxnDate.TabIndex = 3
		Me.ToTxnDate.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ToTxnDate.AcceptsReturn = True
		Me.ToTxnDate.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.ToTxnDate.BackColor = System.Drawing.SystemColors.Window
		Me.ToTxnDate.CausesValidation = True
		Me.ToTxnDate.Enabled = True
		Me.ToTxnDate.ForeColor = System.Drawing.SystemColors.WindowText
		Me.ToTxnDate.HideSelection = True
		Me.ToTxnDate.ReadOnly = False
		Me.ToTxnDate.Maxlength = 0
		Me.ToTxnDate.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.ToTxnDate.MultiLine = False
		Me.ToTxnDate.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ToTxnDate.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.ToTxnDate.TabStop = True
		Me.ToTxnDate.Visible = True
		Me.ToTxnDate.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.ToTxnDate.Name = "ToTxnDate"
		Me.FromTxnDate.AutoSize = False
		Me.FromTxnDate.Size = New System.Drawing.Size(113, 19)
		Me.FromTxnDate.Location = New System.Drawing.Point(176, 32)
		Me.FromTxnDate.TabIndex = 2
		Me.FromTxnDate.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.FromTxnDate.AcceptsReturn = True
		Me.FromTxnDate.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.FromTxnDate.BackColor = System.Drawing.SystemColors.Window
		Me.FromTxnDate.CausesValidation = True
		Me.FromTxnDate.Enabled = True
		Me.FromTxnDate.ForeColor = System.Drawing.SystemColors.WindowText
		Me.FromTxnDate.HideSelection = True
		Me.FromTxnDate.ReadOnly = False
		Me.FromTxnDate.Maxlength = 0
		Me.FromTxnDate.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.FromTxnDate.MultiLine = False
		Me.FromTxnDate.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.FromTxnDate.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.FromTxnDate.TabStop = True
		Me.FromTxnDate.Visible = True
		Me.FromTxnDate.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.FromTxnDate.Name = "FromTxnDate"
		Me.Label2.Text = "To Txn Date (mm/dd/yyyy)"
		Me.Label2.Size = New System.Drawing.Size(145, 17)
		Me.Label2.Location = New System.Drawing.Point(16, 64)
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
		Me.Label1.Text = "From Txn Date (mm/dd/yyyy)"
		Me.Label1.Size = New System.Drawing.Size(137, 17)
		Me.Label1.Location = New System.Drawing.Point(16, 33)
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
		Me.Label1.AutoSize = True
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Controls.Add(UseQBOE)
		Me.Controls.Add(IncludeLineItems)
		Me.Controls.Add(Exit_Renamed)
		Me.Controls.Add(Submit)
		Me.Controls.Add(ToTxnDate)
		Me.Controls.Add(FromTxnDate)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label1)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class