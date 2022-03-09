<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class UIForm
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			Static fTerminateCalled As Boolean
			If Not fTerminateCalled Then
				Form_Terminate_renamed()
				fTerminateCalled = True
			End If
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
	Public WithEvents Command1 As System.Windows.Forms.Button
	Public WithEvents inputFile As System.Windows.Forms.TextBox
	Public WithEvents browseButton As System.Windows.Forms.Button
	Public browseDialogOpen As System.Windows.Forms.OpenFileDialog
	Public WithEvents logBox As System.Windows.Forms.TextBox
	Public WithEvents goBut As System.Windows.Forms.Button
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(UIForm))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.UseQBOE = New System.Windows.Forms.CheckBox
		Me.Command1 = New System.Windows.Forms.Button
		Me.inputFile = New System.Windows.Forms.TextBox
		Me.browseButton = New System.Windows.Forms.Button
		Me.browseDialogOpen = New System.Windows.Forms.OpenFileDialog
		Me.logBox = New System.Windows.Forms.TextBox
		Me.goBut = New System.Windows.Forms.Button
		Me.Label1 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "SDKTest"
		Me.ClientSize = New System.Drawing.Size(445, 327)
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
		Me.Name = "UIForm"
		Me.UseQBOE.Text = "Send to QuickBooks Online Edition (QBOE)"
		Me.UseQBOE.Size = New System.Drawing.Size(241, 17)
		Me.UseQBOE.Location = New System.Drawing.Point(192, 8)
		Me.UseQBOE.TabIndex = 6
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
		Me.Command1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command1.Text = "Show Request"
		Me.Command1.Size = New System.Drawing.Size(73, 33)
		Me.Command1.Location = New System.Drawing.Point(80, 72)
		Me.Command1.TabIndex = 5
		Me.Command1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Command1.BackColor = System.Drawing.SystemColors.Control
		Me.Command1.CausesValidation = True
		Me.Command1.Enabled = True
		Me.Command1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command1.TabStop = True
		Me.Command1.Name = "Command1"
		Me.inputFile.AutoSize = False
		Me.inputFile.Size = New System.Drawing.Size(329, 25)
		Me.inputFile.Location = New System.Drawing.Point(16, 32)
		Me.inputFile.TabIndex = 3
		Me.inputFile.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.inputFile.AcceptsReturn = True
		Me.inputFile.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.inputFile.BackColor = System.Drawing.SystemColors.Window
		Me.inputFile.CausesValidation = True
		Me.inputFile.Enabled = True
		Me.inputFile.ForeColor = System.Drawing.SystemColors.WindowText
		Me.inputFile.HideSelection = True
		Me.inputFile.ReadOnly = False
		Me.inputFile.Maxlength = 0
		Me.inputFile.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.inputFile.MultiLine = False
		Me.inputFile.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.inputFile.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.inputFile.TabStop = True
		Me.inputFile.Visible = True
		Me.inputFile.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.inputFile.Name = "inputFile"
		Me.browseButton.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.browseButton.Text = "Browse"
		Me.browseButton.Size = New System.Drawing.Size(81, 25)
		Me.browseButton.Location = New System.Drawing.Point(352, 32)
		Me.browseButton.TabIndex = 2
		Me.browseButton.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.browseButton.BackColor = System.Drawing.SystemColors.Control
		Me.browseButton.CausesValidation = True
		Me.browseButton.Enabled = True
		Me.browseButton.ForeColor = System.Drawing.SystemColors.ControlText
		Me.browseButton.Cursor = System.Windows.Forms.Cursors.Default
		Me.browseButton.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.browseButton.TabStop = True
		Me.browseButton.Name = "browseButton"
		Me.logBox.AutoSize = False
		Me.logBox.Size = New System.Drawing.Size(425, 193)
		Me.logBox.Location = New System.Drawing.Point(8, 120)
		Me.logBox.MultiLine = True
		Me.logBox.ScrollBars = System.Windows.Forms.ScrollBars.Both
		Me.logBox.WordWrap = False
		Me.logBox.TabIndex = 1
		Me.logBox.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.logBox.AcceptsReturn = True
		Me.logBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.logBox.BackColor = System.Drawing.SystemColors.Window
		Me.logBox.CausesValidation = True
		Me.logBox.Enabled = True
		Me.logBox.ForeColor = System.Drawing.SystemColors.WindowText
		Me.logBox.HideSelection = True
		Me.logBox.ReadOnly = False
		Me.logBox.Maxlength = 0
		Me.logBox.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.logBox.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.logBox.TabStop = True
		Me.logBox.Visible = True
		Me.logBox.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.logBox.Name = "logBox"
		Me.goBut.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.goBut.Text = "Process Request"
		Me.goBut.Size = New System.Drawing.Size(81, 33)
		Me.goBut.Location = New System.Drawing.Point(176, 72)
		Me.goBut.TabIndex = 0
		Me.goBut.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.goBut.BackColor = System.Drawing.SystemColors.Control
		Me.goBut.CausesValidation = True
		Me.goBut.Enabled = True
		Me.goBut.ForeColor = System.Drawing.SystemColors.ControlText
		Me.goBut.Cursor = System.Windows.Forms.Cursors.Default
		Me.goBut.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.goBut.TabStop = True
		Me.goBut.Name = "goBut"
		Me.Label1.Text = "qbXML Request File"
		Me.Label1.Size = New System.Drawing.Size(137, 17)
		Me.Label1.Location = New System.Drawing.Point(16, 8)
		Me.Label1.TabIndex = 4
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
		Me.Controls.Add(UseQBOE)
		Me.Controls.Add(Command1)
		Me.Controls.Add(inputFile)
		Me.Controls.Add(browseButton)
		Me.Controls.Add(logBox)
		Me.Controls.Add(goBut)
		Me.Controls.Add(Label1)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class