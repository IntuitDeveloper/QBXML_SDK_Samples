<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class CommPrefsDlg
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
	Public WithEvents QBOE As System.Windows.Forms.RadioButton
	Public WithEvents QBDT As System.Windows.Forms.RadioButton
	Public WithEvents W As System.Windows.Forms.GroupBox
	Public WithEvents CancelButton_Renamed As System.Windows.Forms.Button
	Public WithEvents OKButton As System.Windows.Forms.Button
	Public WithEvents CompanyFile As System.Windows.Forms.TextBox
	Public WithEvents Interactive As System.Windows.Forms.RadioButton
	Public WithEvents UseCurrentCompany As System.Windows.Forms.RadioButton
	Public WithEvents CompanyFileFrame As System.Windows.Forms.GroupBox
	Public WithEvents QBConnect As System.Windows.Forms.Button
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents DTComm As System.Windows.Forms.GroupBox
	Public WithEvents ConnectionTicket As System.Windows.Forms.TextBox
	Public WithEvents Connect As System.Windows.Forms.Button
	Public WithEvents GoToOE As System.Windows.Forms.Button
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents OEComm As System.Windows.Forms.GroupBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(CommPrefsDlg))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.W = New System.Windows.Forms.GroupBox
		Me.QBOE = New System.Windows.Forms.RadioButton
		Me.QBDT = New System.Windows.Forms.RadioButton
		Me.CancelButton_Renamed = New System.Windows.Forms.Button
		Me.OKButton = New System.Windows.Forms.Button
		Me.DTComm = New System.Windows.Forms.GroupBox
		Me.CompanyFileFrame = New System.Windows.Forms.GroupBox
		Me.CompanyFile = New System.Windows.Forms.TextBox
		Me.Interactive = New System.Windows.Forms.RadioButton
		Me.UseCurrentCompany = New System.Windows.Forms.RadioButton
		Me.QBConnect = New System.Windows.Forms.Button
		Me.Label4 = New System.Windows.Forms.Label
		Me.OEComm = New System.Windows.Forms.GroupBox
		Me.ConnectionTicket = New System.Windows.Forms.TextBox
		Me.Connect = New System.Windows.Forms.Button
		Me.GoToOE = New System.Windows.Forms.Button
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.W.SuspendLayout()
		Me.DTComm.SuspendLayout()
		Me.CompanyFileFrame.SuspendLayout()
		Me.OEComm.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "Dialog Caption"
		Me.ClientSize = New System.Drawing.Size(465, 389)
		Me.Location = New System.Drawing.Point(184, 250)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "CommPrefsDlg"
		Me.W.Text = "Communicate with QuickBooks"
		Me.W.Size = New System.Drawing.Size(329, 81)
		Me.W.Location = New System.Drawing.Point(16, 8)
		Me.W.TabIndex = 2
		Me.W.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.W.BackColor = System.Drawing.SystemColors.Control
		Me.W.Enabled = True
		Me.W.ForeColor = System.Drawing.SystemColors.ControlText
		Me.W.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.W.Visible = True
		Me.W.Padding = New System.Windows.Forms.Padding(0)
		Me.W.Name = "W"
		Me.QBOE.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.QBOE.Text = "Online Edition"
		Me.QBOE.Size = New System.Drawing.Size(209, 17)
		Me.QBOE.Location = New System.Drawing.Point(16, 48)
		Me.QBOE.TabIndex = 5
		Me.QBOE.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.QBOE.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.QBOE.BackColor = System.Drawing.SystemColors.Control
		Me.QBOE.CausesValidation = True
		Me.QBOE.Enabled = True
		Me.QBOE.ForeColor = System.Drawing.SystemColors.ControlText
		Me.QBOE.Cursor = System.Windows.Forms.Cursors.Default
		Me.QBOE.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.QBOE.Appearance = System.Windows.Forms.Appearance.Normal
		Me.QBOE.TabStop = True
		Me.QBOE.Checked = False
		Me.QBOE.Visible = True
		Me.QBOE.Name = "QBOE"
		Me.QBDT.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.QBDT.Text = "On My Desktop"
		Me.QBDT.Size = New System.Drawing.Size(209, 17)
		Me.QBDT.Location = New System.Drawing.Point(16, 24)
		Me.QBDT.TabIndex = 3
		Me.QBDT.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.QBDT.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.QBDT.BackColor = System.Drawing.SystemColors.Control
		Me.QBDT.CausesValidation = True
		Me.QBDT.Enabled = True
		Me.QBDT.ForeColor = System.Drawing.SystemColors.ControlText
		Me.QBDT.Cursor = System.Windows.Forms.Cursors.Default
		Me.QBDT.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.QBDT.Appearance = System.Windows.Forms.Appearance.Normal
		Me.QBDT.TabStop = True
		Me.QBDT.Checked = False
		Me.QBDT.Visible = True
		Me.QBDT.Name = "QBDT"
		Me.CancelButton_Renamed.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton_Renamed.Text = "Cancel"
		Me.CancelButton_Renamed.Size = New System.Drawing.Size(81, 25)
		Me.CancelButton_Renamed.Location = New System.Drawing.Point(368, 40)
		Me.CancelButton_Renamed.TabIndex = 1
		Me.CancelButton_Renamed.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.CancelButton_Renamed.BackColor = System.Drawing.SystemColors.Control
		Me.CancelButton_Renamed.CausesValidation = True
		Me.CancelButton_Renamed.Enabled = True
		Me.CancelButton_Renamed.ForeColor = System.Drawing.SystemColors.ControlText
		Me.CancelButton_Renamed.Cursor = System.Windows.Forms.Cursors.Default
		Me.CancelButton_Renamed.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.CancelButton_Renamed.TabStop = True
		Me.CancelButton_Renamed.Name = "CancelButton_Renamed"
		Me.OKButton.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.OKButton.Text = "OK"
		Me.OKButton.Size = New System.Drawing.Size(81, 25)
		Me.OKButton.Location = New System.Drawing.Point(368, 8)
		Me.OKButton.TabIndex = 0
		Me.OKButton.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.OKButton.BackColor = System.Drawing.SystemColors.Control
		Me.OKButton.CausesValidation = True
		Me.OKButton.Enabled = True
		Me.OKButton.ForeColor = System.Drawing.SystemColors.ControlText
		Me.OKButton.Cursor = System.Windows.Forms.Cursors.Default
		Me.OKButton.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.OKButton.TabStop = True
		Me.OKButton.Name = "OKButton"
		Me.DTComm.Text = "Setup Desktop Connection Preferences"
		Me.DTComm.Size = New System.Drawing.Size(425, 265)
		Me.DTComm.Location = New System.Drawing.Point(16, 104)
		Me.DTComm.TabIndex = 4
		Me.DTComm.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.DTComm.BackColor = System.Drawing.SystemColors.Control
		Me.DTComm.Enabled = True
		Me.DTComm.ForeColor = System.Drawing.SystemColors.ControlText
		Me.DTComm.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.DTComm.Visible = True
		Me.DTComm.Padding = New System.Windows.Forms.Padding(0)
		Me.DTComm.Name = "DTComm"
		Me.CompanyFileFrame.Text = "Company File Info"
		Me.CompanyFileFrame.Size = New System.Drawing.Size(393, 97)
		Me.CompanyFileFrame.Location = New System.Drawing.Point(16, 160)
		Me.CompanyFileFrame.TabIndex = 15
		Me.CompanyFileFrame.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.CompanyFileFrame.BackColor = System.Drawing.SystemColors.Control
		Me.CompanyFileFrame.Enabled = True
		Me.CompanyFileFrame.ForeColor = System.Drawing.SystemColors.ControlText
		Me.CompanyFileFrame.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.CompanyFileFrame.Visible = True
		Me.CompanyFileFrame.Padding = New System.Windows.Forms.Padding(0)
		Me.CompanyFileFrame.Name = "CompanyFileFrame"
		Me.CompanyFile.AutoSize = False
		Me.CompanyFile.Size = New System.Drawing.Size(377, 25)
		Me.CompanyFile.Location = New System.Drawing.Point(8, 64)
		Me.CompanyFile.TabIndex = 18
		Me.CompanyFile.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.CompanyFile.AcceptsReturn = True
		Me.CompanyFile.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.CompanyFile.BackColor = System.Drawing.SystemColors.Window
		Me.CompanyFile.CausesValidation = True
		Me.CompanyFile.Enabled = True
		Me.CompanyFile.ForeColor = System.Drawing.SystemColors.WindowText
		Me.CompanyFile.HideSelection = True
		Me.CompanyFile.ReadOnly = False
		Me.CompanyFile.Maxlength = 0
		Me.CompanyFile.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.CompanyFile.MultiLine = False
		Me.CompanyFile.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.CompanyFile.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.CompanyFile.TabStop = True
		Me.CompanyFile.Visible = True
		Me.CompanyFile.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.CompanyFile.Name = "CompanyFile"
		Me.Interactive.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Interactive.Text = "Use whatever file is open"
		Me.Interactive.Size = New System.Drawing.Size(161, 17)
		Me.Interactive.Location = New System.Drawing.Point(200, 32)
		Me.Interactive.TabIndex = 17
		Me.Interactive.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Interactive.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Interactive.BackColor = System.Drawing.SystemColors.Control
		Me.Interactive.CausesValidation = True
		Me.Interactive.Enabled = True
		Me.Interactive.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Interactive.Cursor = System.Windows.Forms.Cursors.Default
		Me.Interactive.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Interactive.Appearance = System.Windows.Forms.Appearance.Normal
		Me.Interactive.TabStop = True
		Me.Interactive.Checked = False
		Me.Interactive.Visible = True
		Me.Interactive.Name = "Interactive"
		Me.UseCurrentCompany.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.UseCurrentCompany.Text = "Use this company file (QuickBooks can be closed):"
		Me.UseCurrentCompany.Size = New System.Drawing.Size(161, 25)
		Me.UseCurrentCompany.Location = New System.Drawing.Point(16, 24)
		Me.UseCurrentCompany.TabIndex = 16
		Me.UseCurrentCompany.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.UseCurrentCompany.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.UseCurrentCompany.BackColor = System.Drawing.SystemColors.Control
		Me.UseCurrentCompany.CausesValidation = True
		Me.UseCurrentCompany.Enabled = True
		Me.UseCurrentCompany.ForeColor = System.Drawing.SystemColors.ControlText
		Me.UseCurrentCompany.Cursor = System.Windows.Forms.Cursors.Default
		Me.UseCurrentCompany.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.UseCurrentCompany.Appearance = System.Windows.Forms.Appearance.Normal
		Me.UseCurrentCompany.TabStop = True
		Me.UseCurrentCompany.Checked = False
		Me.UseCurrentCompany.Visible = True
		Me.UseCurrentCompany.Name = "UseCurrentCompany"
		Me.QBConnect.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.QBConnect.Text = "Connect To QuickBooks!"
		Me.QBConnect.Size = New System.Drawing.Size(225, 33)
		Me.QBConnect.Location = New System.Drawing.Point(16, 96)
		Me.QBConnect.TabIndex = 14
		Me.QBConnect.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.QBConnect.BackColor = System.Drawing.SystemColors.Control
		Me.QBConnect.CausesValidation = True
		Me.QBConnect.Enabled = True
		Me.QBConnect.ForeColor = System.Drawing.SystemColors.ControlText
		Me.QBConnect.Cursor = System.Windows.Forms.Cursors.Default
		Me.QBConnect.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.QBConnect.TabStop = True
		Me.QBConnect.Name = "QBConnect"
		Me.Label4.Text = "To setup a connection with a QuickBooks Desktop Company file for the first time, QuickBooks must be running with the desired company file open.  When you click the ""Connect"" button below, we will establish a connection to QuickBooks to obtain permission to connect to the company file that is open in QuickBooks."
		Me.Label4.Size = New System.Drawing.Size(401, 57)
		Me.Label4.Location = New System.Drawing.Point(8, 24)
		Me.Label4.TabIndex = 13
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
		Me.OEComm.Text = "Set up Online Edition communications"
		Me.OEComm.Size = New System.Drawing.Size(425, 265)
		Me.OEComm.Location = New System.Drawing.Point(16, 104)
		Me.OEComm.TabIndex = 6
		Me.OEComm.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.OEComm.BackColor = System.Drawing.SystemColors.Control
		Me.OEComm.Enabled = True
		Me.OEComm.ForeColor = System.Drawing.SystemColors.ControlText
		Me.OEComm.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.OEComm.Visible = True
		Me.OEComm.Padding = New System.Windows.Forms.Padding(0)
		Me.OEComm.Name = "OEComm"
		Me.ConnectionTicket.AutoSize = False
		Me.ConnectionTicket.Size = New System.Drawing.Size(257, 25)
		Me.ConnectionTicket.Location = New System.Drawing.Point(152, 232)
		Me.ConnectionTicket.TabIndex = 12
		Me.ConnectionTicket.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ConnectionTicket.AcceptsReturn = True
		Me.ConnectionTicket.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.ConnectionTicket.BackColor = System.Drawing.SystemColors.Window
		Me.ConnectionTicket.CausesValidation = True
		Me.ConnectionTicket.Enabled = True
		Me.ConnectionTicket.ForeColor = System.Drawing.SystemColors.WindowText
		Me.ConnectionTicket.HideSelection = True
		Me.ConnectionTicket.ReadOnly = False
		Me.ConnectionTicket.Maxlength = 0
		Me.ConnectionTicket.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.ConnectionTicket.MultiLine = False
		Me.ConnectionTicket.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ConnectionTicket.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.ConnectionTicket.TabStop = True
		Me.ConnectionTicket.Visible = True
		Me.ConnectionTicket.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.ConnectionTicket.Name = "ConnectionTicket"
		Me.Connect.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Connect.Text = "Set up connection!"
		Me.Connect.Size = New System.Drawing.Size(137, 25)
		Me.Connect.Location = New System.Drawing.Point(16, 192)
		Me.Connect.TabIndex = 10
		Me.Connect.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Connect.BackColor = System.Drawing.SystemColors.Control
		Me.Connect.CausesValidation = True
		Me.Connect.Enabled = True
		Me.Connect.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Connect.Cursor = System.Windows.Forms.Cursors.Default
		Me.Connect.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Connect.TabStop = True
		Me.Connect.Name = "Connect"
		Me.GoToOE.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.GoToOE.Text = "Let me set up a company!"
		Me.GoToOE.Size = New System.Drawing.Size(137, 25)
		Me.GoToOE.Location = New System.Drawing.Point(16, 80)
		Me.GoToOE.TabIndex = 8
		Me.GoToOE.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.GoToOE.BackColor = System.Drawing.SystemColors.Control
		Me.GoToOE.CausesValidation = True
		Me.GoToOE.Enabled = True
		Me.GoToOE.ForeColor = System.Drawing.SystemColors.ControlText
		Me.GoToOE.Cursor = System.Windows.Forms.Cursors.Default
		Me.GoToOE.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.GoToOE.TabStop = True
		Me.GoToOE.Name = "GoToOE"
		Me.Label3.Text = "Enter connection ticket here:"
		Me.Label3.Size = New System.Drawing.Size(153, 17)
		Me.Label3.Location = New System.Drawing.Point(8, 240)
		Me.Label3.TabIndex = 11
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
		Me.Label2.Text = "Next you have to give this application permission to access your online edition company.  You do that by clicking the ""Connect"" button below and following the instructions provided by QuickBooks Online Edition in your browser.  When the interview completes you will be given a ""connection ticket"" that must be pasted into the text box below."
		Me.Label2.Size = New System.Drawing.Size(409, 65)
		Me.Label2.Location = New System.Drawing.Point(8, 120)
		Me.Label2.TabIndex = 9
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
		Me.Label1.Text = "Communicating with the Online Edition of QuickBooks is easy.  Before you begin, make sure you have set up  a company in QuickBooks Online Edition by visiting http://oe.quickbooks.com."
		Me.Label1.Size = New System.Drawing.Size(401, 41)
		Me.Label1.Location = New System.Drawing.Point(8, 24)
		Me.Label1.TabIndex = 7
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
		Me.Controls.Add(W)
		Me.Controls.Add(CancelButton_Renamed)
		Me.Controls.Add(OKButton)
		Me.Controls.Add(DTComm)
		Me.Controls.Add(OEComm)
		Me.W.Controls.Add(QBOE)
		Me.W.Controls.Add(QBDT)
		Me.DTComm.Controls.Add(CompanyFileFrame)
		Me.DTComm.Controls.Add(QBConnect)
		Me.DTComm.Controls.Add(Label4)
		Me.CompanyFileFrame.Controls.Add(CompanyFile)
		Me.CompanyFileFrame.Controls.Add(Interactive)
		Me.CompanyFileFrame.Controls.Add(UseCurrentCompany)
		Me.OEComm.Controls.Add(ConnectionTicket)
		Me.OEComm.Controls.Add(Connect)
		Me.OEComm.Controls.Add(GoToOE)
		Me.OEComm.Controls.Add(Label3)
		Me.OEComm.Controls.Add(Label2)
		Me.OEComm.Controls.Add(Label1)
		Me.W.ResumeLayout(False)
		Me.DTComm.ResumeLayout(False)
		Me.CompanyFileFrame.ResumeLayout(False)
		Me.OEComm.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class