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
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents showResButton As System.Windows.Forms.Button
	Public WithEvents showReqButton As System.Windows.Forms.Button
	Public WithEvents exitButton As System.Windows.Forms.Button
	Public WithEvents goButton As System.Windows.Forms.Button
	Public WithEvents qbFile As System.Windows.Forms.TextBox
	Public WithEvents browseButton As System.Windows.Forms.Button
	Public WithEvents htmlFile As System.Windows.Forms.TextBox
	Public WithEvents embedBrowser As System.Windows.Forms.WebBrowser
	Public browseDialogOpen As System.Windows.Forms.OpenFileDialog
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(UIForm))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.showResButton = New System.Windows.Forms.Button
		Me.showReqButton = New System.Windows.Forms.Button
		Me.exitButton = New System.Windows.Forms.Button
		Me.goButton = New System.Windows.Forms.Button
		Me.qbFile = New System.Windows.Forms.TextBox
		Me.browseButton = New System.Windows.Forms.Button
		Me.htmlFile = New System.Windows.Forms.TextBox
		Me.embedBrowser = New System.Windows.Forms.WebBrowser
		Me.browseDialogOpen = New System.Windows.Forms.OpenFileDialog
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "QBSDK 2.0 --- Expense By Vendor Summary Report"
		Me.ClientSize = New System.Drawing.Size(828, 539)
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
		Me.showResButton.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.showResButton.Text = "Show Response qbXML"
		Me.showResButton.Size = New System.Drawing.Size(145, 25)
		Me.showResButton.Location = New System.Drawing.Point(420, 84)
		Me.showResButton.TabIndex = 10
		Me.showResButton.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.showResButton.BackColor = System.Drawing.SystemColors.Control
		Me.showResButton.CausesValidation = True
		Me.showResButton.Enabled = True
		Me.showResButton.ForeColor = System.Drawing.SystemColors.ControlText
		Me.showResButton.Cursor = System.Windows.Forms.Cursors.Default
		Me.showResButton.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.showResButton.TabStop = True
		Me.showResButton.Name = "showResButton"
		Me.showReqButton.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.showReqButton.Text = "Show Request qbXML"
		Me.showReqButton.Size = New System.Drawing.Size(145, 25)
		Me.showReqButton.Location = New System.Drawing.Point(248, 84)
		Me.showReqButton.TabIndex = 9
		Me.showReqButton.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.showReqButton.BackColor = System.Drawing.SystemColors.Control
		Me.showReqButton.CausesValidation = True
		Me.showReqButton.Enabled = True
		Me.showReqButton.ForeColor = System.Drawing.SystemColors.ControlText
		Me.showReqButton.Cursor = System.Windows.Forms.Cursors.Default
		Me.showReqButton.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.showReqButton.TabStop = True
		Me.showReqButton.Name = "showReqButton"
		Me.exitButton.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.exitButton.Text = "&Exit Program"
		Me.exitButton.Size = New System.Drawing.Size(145, 25)
		Me.exitButton.Location = New System.Drawing.Point(592, 84)
		Me.exitButton.TabIndex = 8
		Me.exitButton.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.exitButton.BackColor = System.Drawing.SystemColors.Control
		Me.exitButton.CausesValidation = True
		Me.exitButton.Enabled = True
		Me.exitButton.ForeColor = System.Drawing.SystemColors.ControlText
		Me.exitButton.Cursor = System.Windows.Forms.Cursors.Default
		Me.exitButton.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.exitButton.TabStop = True
		Me.exitButton.Name = "exitButton"
		Me.goButton.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.goButton.Text = "&Generate Report"
		Me.goButton.Size = New System.Drawing.Size(145, 25)
		Me.goButton.Location = New System.Drawing.Point(76, 84)
		Me.goButton.TabIndex = 4
		Me.goButton.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.goButton.BackColor = System.Drawing.SystemColors.Control
		Me.goButton.CausesValidation = True
		Me.goButton.Enabled = True
		Me.goButton.ForeColor = System.Drawing.SystemColors.ControlText
		Me.goButton.Cursor = System.Windows.Forms.Cursors.Default
		Me.goButton.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.goButton.TabStop = True
		Me.goButton.Name = "goButton"
		Me.qbFile.AutoSize = False
		Me.qbFile.Size = New System.Drawing.Size(241, 25)
		Me.qbFile.Location = New System.Drawing.Point(292, 16)
		Me.qbFile.TabIndex = 3
		Me.qbFile.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.qbFile.AcceptsReturn = True
		Me.qbFile.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.qbFile.BackColor = System.Drawing.SystemColors.Window
		Me.qbFile.CausesValidation = True
		Me.qbFile.Enabled = True
		Me.qbFile.ForeColor = System.Drawing.SystemColors.WindowText
		Me.qbFile.HideSelection = True
		Me.qbFile.ReadOnly = False
		Me.qbFile.Maxlength = 0
		Me.qbFile.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.qbFile.MultiLine = False
		Me.qbFile.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.qbFile.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.qbFile.TabStop = True
		Me.qbFile.Visible = True
		Me.qbFile.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.qbFile.Name = "qbFile"
		Me.browseButton.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.browseButton.Text = "&Browse"
		Me.browseButton.Size = New System.Drawing.Size(105, 25)
		Me.browseButton.Location = New System.Drawing.Point(536, 16)
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
		Me.htmlFile.AutoSize = False
		Me.htmlFile.Size = New System.Drawing.Size(349, 25)
		Me.htmlFile.Location = New System.Drawing.Point(292, 48)
		Me.htmlFile.TabIndex = 1
		Me.htmlFile.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.htmlFile.AcceptsReturn = True
		Me.htmlFile.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.htmlFile.BackColor = System.Drawing.SystemColors.Window
		Me.htmlFile.CausesValidation = True
		Me.htmlFile.Enabled = True
		Me.htmlFile.ForeColor = System.Drawing.SystemColors.WindowText
		Me.htmlFile.HideSelection = True
		Me.htmlFile.ReadOnly = False
		Me.htmlFile.Maxlength = 0
		Me.htmlFile.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.htmlFile.MultiLine = False
		Me.htmlFile.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.htmlFile.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.htmlFile.TabStop = True
		Me.htmlFile.Visible = True
		Me.htmlFile.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.htmlFile.Name = "htmlFile"
		Me.embedBrowser.Size = New System.Drawing.Size(789, 409)
		Me.embedBrowser.Location = New System.Drawing.Point(20, 116)
		Me.embedBrowser.TabIndex = 0
		Me.embedBrowser.AllowWebBrowserDrop = True
		Me.embedBrowser.Name = "embedBrowser"
		Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.Label3.Text = "(blank to use currently open file)"
		Me.Label3.Size = New System.Drawing.Size(165, 13)
		Me.Label3.Location = New System.Drawing.Point(124, 28)
		Me.Label3.TabIndex = 7
		Me.Label3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label1.Text = "Select QuickBooks Company File:"
		Me.Label1.Size = New System.Drawing.Size(165, 21)
		Me.Label1.Location = New System.Drawing.Point(124, 16)
		Me.Label1.TabIndex = 6
		Me.Label1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label2.Text = "Select Output Html File:"
		Me.Label2.Size = New System.Drawing.Size(165, 17)
		Me.Label2.Location = New System.Drawing.Point(124, 52)
		Me.Label2.TabIndex = 5
		Me.Label2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
		Me.Controls.Add(showResButton)
		Me.Controls.Add(showReqButton)
		Me.Controls.Add(exitButton)
		Me.Controls.Add(goButton)
		Me.Controls.Add(qbFile)
		Me.Controls.Add(browseButton)
		Me.Controls.Add(htmlFile)
		Me.Controls.Add(embedBrowser)
		Me.Controls.Add(Label3)
		Me.Controls.Add(Label1)
		Me.Controls.Add(Label2)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class