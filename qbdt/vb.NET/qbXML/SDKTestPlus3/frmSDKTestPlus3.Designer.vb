<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmSDKTestPlus3
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
	Public WithEvents lblStatus As System.Windows.Forms.TextBox
	Public WithEvents qbForceAuthDialog As System.Windows.Forms.CheckBox
	Public WithEvents qbSimple As System.Windows.Forms.CheckBox
	Public WithEvents qbPro As System.Windows.Forms.CheckBox
	Public WithEvents qbPremier As System.Windows.Forms.CheckBox
	Public WithEvents qbEnterprise As System.Windows.Forms.CheckBox
	Public WithEvents AuthFlagsFrame As System.Windows.Forms.GroupBox
	Public WithEvents optDontCare As System.Windows.Forms.RadioButton
	Public WithEvents optMultiUser As System.Windows.Forms.RadioButton
	Public WithEvents optSingleUser As System.Windows.Forms.RadioButton
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents pdNotNeeded As System.Windows.Forms.RadioButton
	Public WithEvents pdOptional As System.Windows.Forms.RadioButton
	Public WithEvents pdRequired As System.Windows.Forms.RadioButton
	Public WithEvents pdFrame As System.Windows.Forms.GroupBox
	Public WithEvents ReadOnly_Renamed As System.Windows.Forms.CheckBox
	Public WithEvents Unattended As System.Windows.Forms.CheckBox
	Public WithEvents SessionPrefsFrame As System.Windows.Forms.GroupBox
	Public WithEvents ConnLocalUI As System.Windows.Forms.RadioButton
	Public WithEvents ConnQBOE As System.Windows.Forms.RadioButton
	Public WithEvents ConnRemote As System.Windows.Forms.RadioButton
	Public WithEvents ConnLocal As System.Windows.Forms.RadioButton
	Public WithEvents ConnNone As System.Windows.Forms.RadioButton
	Public WithEvents ConnFrame As System.Windows.Forms.GroupBox
	Public WithEvents RequestBrowse As System.Windows.Forms.Button
	Public WithEvents RequestFile As System.Windows.Forms.TextBox
	Public OpenFileDialogOpen As System.Windows.Forms.OpenFileDialog
	Public WithEvents CompanyBrowse As System.Windows.Forms.Button
	Public WithEvents CompanyFile As System.Windows.Forms.TextBox
	Public WithEvents cmdProcessSubscription As System.Windows.Forms.Button
	Public WithEvents cmdSendXML As System.Windows.Forms.Button
	Public WithEvents cmdViewInput As System.Windows.Forms.Button
	Public WithEvents cmdViewOutput As System.Windows.Forms.Button
	Public WithEvents cmdExit As System.Windows.Forms.Button
	Public WithEvents cmdOpenConnection As System.Windows.Forms.Button
	Public WithEvents cmdBeginSession As System.Windows.Forms.Button
	Public WithEvents cmdEndSession As System.Windows.Forms.Button
	Public WithEvents cmdCloseConnection As System.Windows.Forms.Button
	Public WithEvents txtApplicationName As System.Windows.Forms.TextBox
	Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.lblStatus = New System.Windows.Forms.TextBox()
        Me.AuthFlagsFrame = New System.Windows.Forms.GroupBox()
        Me.qbForceAuthDialog = New System.Windows.Forms.CheckBox()
        Me.qbSimple = New System.Windows.Forms.CheckBox()
        Me.qbPro = New System.Windows.Forms.CheckBox()
        Me.qbPremier = New System.Windows.Forms.CheckBox()
        Me.qbEnterprise = New System.Windows.Forms.CheckBox()
        Me.SessionPrefsFrame = New System.Windows.Forms.GroupBox()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.optDontCare = New System.Windows.Forms.RadioButton()
        Me.optMultiUser = New System.Windows.Forms.RadioButton()
        Me.optSingleUser = New System.Windows.Forms.RadioButton()
        Me.pdFrame = New System.Windows.Forms.GroupBox()
        Me.pdNotNeeded = New System.Windows.Forms.RadioButton()
        Me.pdOptional = New System.Windows.Forms.RadioButton()
        Me.pdRequired = New System.Windows.Forms.RadioButton()
        Me.ReadOnly_Renamed = New System.Windows.Forms.CheckBox()
        Me.Unattended = New System.Windows.Forms.CheckBox()
        Me.ConnFrame = New System.Windows.Forms.GroupBox()
        Me.ConnLocalUI = New System.Windows.Forms.RadioButton()
        Me.ConnQBOE = New System.Windows.Forms.RadioButton()
        Me.ConnRemote = New System.Windows.Forms.RadioButton()
        Me.ConnLocal = New System.Windows.Forms.RadioButton()
        Me.ConnNone = New System.Windows.Forms.RadioButton()
        Me.RequestBrowse = New System.Windows.Forms.Button()
        Me.RequestFile = New System.Windows.Forms.TextBox()
        Me.OpenFileDialogOpen = New System.Windows.Forms.OpenFileDialog()
        Me.CompanyBrowse = New System.Windows.Forms.Button()
        Me.CompanyFile = New System.Windows.Forms.TextBox()
        Me.cmdProcessSubscription = New System.Windows.Forms.Button()
        Me.cmdSendXML = New System.Windows.Forms.Button()
        Me.cmdViewInput = New System.Windows.Forms.Button()
        Me.cmdViewOutput = New System.Windows.Forms.Button()
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.cmdOpenConnection = New System.Windows.Forms.Button()
        Me.cmdBeginSession = New System.Windows.Forms.Button()
        Me.cmdEndSession = New System.Windows.Forms.Button()
        Me.cmdCloseConnection = New System.Windows.Forms.Button()
        Me.txtApplicationName = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.AuthFlagsFrame.SuspendLayout()
        Me.SessionPrefsFrame.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me.pdFrame.SuspendLayout()
        Me.ConnFrame.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblStatus
        '
        Me.lblStatus.AcceptsReturn = True
        Me.lblStatus.BackColor = System.Drawing.SystemColors.Window
        Me.lblStatus.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.lblStatus.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStatus.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblStatus.Location = New System.Drawing.Point(16, 400)
        Me.lblStatus.MaxLength = 0
        Me.lblStatus.Multiline = True
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.ReadOnly = True
        Me.lblStatus.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblStatus.Size = New System.Drawing.Size(529, 121)
        Me.lblStatus.TabIndex = 41
        '
        'AuthFlagsFrame
        '
        Me.AuthFlagsFrame.BackColor = System.Drawing.SystemColors.Control
        Me.AuthFlagsFrame.Controls.Add(Me.qbForceAuthDialog)
        Me.AuthFlagsFrame.Controls.Add(Me.qbSimple)
        Me.AuthFlagsFrame.Controls.Add(Me.qbPro)
        Me.AuthFlagsFrame.Controls.Add(Me.qbPremier)
        Me.AuthFlagsFrame.Controls.Add(Me.qbEnterprise)
        Me.AuthFlagsFrame.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AuthFlagsFrame.ForeColor = System.Drawing.SystemColors.ControlText
        Me.AuthFlagsFrame.Location = New System.Drawing.Point(16, 120)
        Me.AuthFlagsFrame.Name = "AuthFlagsFrame"
        Me.AuthFlagsFrame.Padding = New System.Windows.Forms.Padding(0)
        Me.AuthFlagsFrame.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.AuthFlagsFrame.Size = New System.Drawing.Size(129, 137)
        Me.AuthFlagsFrame.TabIndex = 35
        Me.AuthFlagsFrame.TabStop = False
        Me.AuthFlagsFrame.Text = "AuthFlags"
        '
        'qbForceAuthDialog
        '
        Me.qbForceAuthDialog.BackColor = System.Drawing.SystemColors.Control
        Me.qbForceAuthDialog.Cursor = System.Windows.Forms.Cursors.Default
        Me.qbForceAuthDialog.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.qbForceAuthDialog.ForeColor = System.Drawing.SystemColors.ControlText
        Me.qbForceAuthDialog.Location = New System.Drawing.Point(8, 112)
        Me.qbForceAuthDialog.Name = "qbForceAuthDialog"
        Me.qbForceAuthDialog.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.qbForceAuthDialog.Size = New System.Drawing.Size(113, 17)
        Me.qbForceAuthDialog.TabIndex = 40
        Me.qbForceAuthDialog.Text = "Force auth dialog"
        Me.qbForceAuthDialog.UseVisualStyleBackColor = False
        '
        'qbSimple
        '
        Me.qbSimple.BackColor = System.Drawing.SystemColors.Control
        Me.qbSimple.Cursor = System.Windows.Forms.Cursors.Default
        Me.qbSimple.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.qbSimple.ForeColor = System.Drawing.SystemColors.ControlText
        Me.qbSimple.Location = New System.Drawing.Point(8, 88)
        Me.qbSimple.Name = "qbSimple"
        Me.qbSimple.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.qbSimple.Size = New System.Drawing.Size(105, 17)
        Me.qbSimple.TabIndex = 39
        Me.qbSimple.Text = "QB Simple Start"
        Me.qbSimple.UseVisualStyleBackColor = False
        '
        'qbPro
        '
        Me.qbPro.BackColor = System.Drawing.SystemColors.Control
        Me.qbPro.Checked = True
        Me.qbPro.CheckState = System.Windows.Forms.CheckState.Checked
        Me.qbPro.Cursor = System.Windows.Forms.Cursors.Default
        Me.qbPro.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.qbPro.ForeColor = System.Drawing.SystemColors.ControlText
        Me.qbPro.Location = New System.Drawing.Point(8, 64)
        Me.qbPro.Name = "qbPro"
        Me.qbPro.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.qbPro.Size = New System.Drawing.Size(105, 17)
        Me.qbPro.TabIndex = 38
        Me.qbPro.Text = "QB Pro"
        Me.qbPro.UseVisualStyleBackColor = False
        '
        'qbPremier
        '
        Me.qbPremier.BackColor = System.Drawing.SystemColors.Control
        Me.qbPremier.Checked = True
        Me.qbPremier.CheckState = System.Windows.Forms.CheckState.Checked
        Me.qbPremier.Cursor = System.Windows.Forms.Cursors.Default
        Me.qbPremier.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.qbPremier.ForeColor = System.Drawing.SystemColors.ControlText
        Me.qbPremier.Location = New System.Drawing.Point(8, 40)
        Me.qbPremier.Name = "qbPremier"
        Me.qbPremier.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.qbPremier.Size = New System.Drawing.Size(105, 17)
        Me.qbPremier.TabIndex = 37
        Me.qbPremier.Text = "QB Premier"
        Me.qbPremier.UseVisualStyleBackColor = False
        '
        'qbEnterprise
        '
        Me.qbEnterprise.BackColor = System.Drawing.SystemColors.Control
        Me.qbEnterprise.Checked = True
        Me.qbEnterprise.CheckState = System.Windows.Forms.CheckState.Checked
        Me.qbEnterprise.Cursor = System.Windows.Forms.Cursors.Default
        Me.qbEnterprise.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.qbEnterprise.ForeColor = System.Drawing.SystemColors.ControlText
        Me.qbEnterprise.Location = New System.Drawing.Point(8, 16)
        Me.qbEnterprise.Name = "qbEnterprise"
        Me.qbEnterprise.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.qbEnterprise.Size = New System.Drawing.Size(105, 17)
        Me.qbEnterprise.TabIndex = 36
        Me.qbEnterprise.Text = "QB Enterprise"
        Me.qbEnterprise.UseVisualStyleBackColor = False
        '
        'SessionPrefsFrame
        '
        Me.SessionPrefsFrame.BackColor = System.Drawing.SystemColors.Control
        Me.SessionPrefsFrame.Controls.Add(Me.Frame1)
        Me.SessionPrefsFrame.Controls.Add(Me.pdFrame)
        Me.SessionPrefsFrame.Controls.Add(Me.ReadOnly_Renamed)
        Me.SessionPrefsFrame.Controls.Add(Me.Unattended)
        Me.SessionPrefsFrame.Enabled = False
        Me.SessionPrefsFrame.Font = New System.Drawing.Font("Arial", 8.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SessionPrefsFrame.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.SessionPrefsFrame.Location = New System.Drawing.Point(296, 120)
        Me.SessionPrefsFrame.Name = "SessionPrefsFrame"
        Me.SessionPrefsFrame.Padding = New System.Windows.Forms.Padding(0)
        Me.SessionPrefsFrame.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.SessionPrefsFrame.Size = New System.Drawing.Size(217, 137)
        Me.SessionPrefsFrame.TabIndex = 23
        Me.SessionPrefsFrame.TabStop = False
        Me.SessionPrefsFrame.Text = "Session  Prefs"
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.optDontCare)
        Me.Frame1.Controls.Add(Me.optMultiUser)
        Me.Frame1.Controls.Add(Me.optSingleUser)
        Me.Frame1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(112, 40)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(97, 89)
        Me.Frame1.TabIndex = 30
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "File Mode"
        '
        'optDontCare
        '
        Me.optDontCare.BackColor = System.Drawing.SystemColors.Control
        Me.optDontCare.Checked = True
        Me.optDontCare.Cursor = System.Windows.Forms.Cursors.Default
        Me.optDontCare.Enabled = False
        Me.optDontCare.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optDontCare.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optDontCare.Location = New System.Drawing.Point(8, 64)
        Me.optDontCare.Name = "optDontCare"
        Me.optDontCare.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optDontCare.Size = New System.Drawing.Size(73, 17)
        Me.optDontCare.TabIndex = 33
        Me.optDontCare.TabStop = True
        Me.optDontCare.Text = "Don't Care"
        Me.optDontCare.UseVisualStyleBackColor = False
        '
        'optMultiUser
        '
        Me.optMultiUser.BackColor = System.Drawing.SystemColors.Control
        Me.optMultiUser.Cursor = System.Windows.Forms.Cursors.Default
        Me.optMultiUser.Enabled = False
        Me.optMultiUser.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optMultiUser.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optMultiUser.Location = New System.Drawing.Point(8, 16)
        Me.optMultiUser.Name = "optMultiUser"
        Me.optMultiUser.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optMultiUser.Size = New System.Drawing.Size(73, 17)
        Me.optMultiUser.TabIndex = 32
        Me.optMultiUser.TabStop = True
        Me.optMultiUser.Text = "Multi-User"
        Me.optMultiUser.UseVisualStyleBackColor = False
        '
        'optSingleUser
        '
        Me.optSingleUser.BackColor = System.Drawing.SystemColors.Control
        Me.optSingleUser.Cursor = System.Windows.Forms.Cursors.Default
        Me.optSingleUser.Enabled = False
        Me.optSingleUser.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optSingleUser.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optSingleUser.Location = New System.Drawing.Point(8, 40)
        Me.optSingleUser.Name = "optSingleUser"
        Me.optSingleUser.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optSingleUser.Size = New System.Drawing.Size(81, 17)
        Me.optSingleUser.TabIndex = 31
        Me.optSingleUser.TabStop = True
        Me.optSingleUser.Text = "Single User"
        Me.optSingleUser.UseVisualStyleBackColor = False
        '
        'pdFrame
        '
        Me.pdFrame.BackColor = System.Drawing.SystemColors.Control
        Me.pdFrame.Controls.Add(Me.pdNotNeeded)
        Me.pdFrame.Controls.Add(Me.pdOptional)
        Me.pdFrame.Controls.Add(Me.pdRequired)
        Me.pdFrame.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pdFrame.ForeColor = System.Drawing.Color.Black
        Me.pdFrame.Location = New System.Drawing.Point(8, 40)
        Me.pdFrame.Name = "pdFrame"
        Me.pdFrame.Padding = New System.Windows.Forms.Padding(0)
        Me.pdFrame.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.pdFrame.Size = New System.Drawing.Size(97, 89)
        Me.pdFrame.TabIndex = 26
        Me.pdFrame.TabStop = False
        Me.pdFrame.Text = "Personal Data"
        '
        'pdNotNeeded
        '
        Me.pdNotNeeded.BackColor = System.Drawing.SystemColors.Control
        Me.pdNotNeeded.Cursor = System.Windows.Forms.Cursors.Default
        Me.pdNotNeeded.Enabled = False
        Me.pdNotNeeded.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pdNotNeeded.ForeColor = System.Drawing.SystemColors.ControlText
        Me.pdNotNeeded.Location = New System.Drawing.Point(8, 64)
        Me.pdNotNeeded.Name = "pdNotNeeded"
        Me.pdNotNeeded.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.pdNotNeeded.Size = New System.Drawing.Size(81, 17)
        Me.pdNotNeeded.TabIndex = 29
        Me.pdNotNeeded.TabStop = True
        Me.pdNotNeeded.Text = "Not Needed"
        Me.pdNotNeeded.UseVisualStyleBackColor = False
        '
        'pdOptional
        '
        Me.pdOptional.BackColor = System.Drawing.SystemColors.Control
        Me.pdOptional.Checked = True
        Me.pdOptional.Cursor = System.Windows.Forms.Cursors.Default
        Me.pdOptional.Enabled = False
        Me.pdOptional.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pdOptional.ForeColor = System.Drawing.SystemColors.ControlText
        Me.pdOptional.Location = New System.Drawing.Point(8, 40)
        Me.pdOptional.Name = "pdOptional"
        Me.pdOptional.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.pdOptional.Size = New System.Drawing.Size(65, 17)
        Me.pdOptional.TabIndex = 28
        Me.pdOptional.TabStop = True
        Me.pdOptional.Text = "Optional"
        Me.pdOptional.UseVisualStyleBackColor = False
        '
        'pdRequired
        '
        Me.pdRequired.BackColor = System.Drawing.SystemColors.Control
        Me.pdRequired.Cursor = System.Windows.Forms.Cursors.Default
        Me.pdRequired.Enabled = False
        Me.pdRequired.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pdRequired.ForeColor = System.Drawing.SystemColors.ControlText
        Me.pdRequired.Location = New System.Drawing.Point(8, 16)
        Me.pdRequired.Name = "pdRequired"
        Me.pdRequired.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.pdRequired.Size = New System.Drawing.Size(73, 17)
        Me.pdRequired.TabIndex = 27
        Me.pdRequired.TabStop = True
        Me.pdRequired.Text = "Required"
        Me.pdRequired.UseVisualStyleBackColor = False
        '
        'ReadOnly_Renamed
        '
        Me.ReadOnly_Renamed.BackColor = System.Drawing.SystemColors.Control
        Me.ReadOnly_Renamed.Cursor = System.Windows.Forms.Cursors.Default
        Me.ReadOnly_Renamed.Enabled = False
        Me.ReadOnly_Renamed.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ReadOnly_Renamed.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ReadOnly_Renamed.Location = New System.Drawing.Point(112, 16)
        Me.ReadOnly_Renamed.Name = "ReadOnly_Renamed"
        Me.ReadOnly_Renamed.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ReadOnly_Renamed.Size = New System.Drawing.Size(81, 17)
        Me.ReadOnly_Renamed.TabIndex = 25
        Me.ReadOnly_Renamed.Text = "Read-only"
        Me.ReadOnly_Renamed.UseVisualStyleBackColor = False
        '
        'Unattended
        '
        Me.Unattended.BackColor = System.Drawing.SystemColors.Control
        Me.Unattended.Cursor = System.Windows.Forms.Cursors.Default
        Me.Unattended.Enabled = False
        Me.Unattended.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Unattended.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Unattended.Location = New System.Drawing.Point(8, 16)
        Me.Unattended.Name = "Unattended"
        Me.Unattended.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Unattended.Size = New System.Drawing.Size(97, 17)
        Me.Unattended.TabIndex = 24
        Me.Unattended.Text = "Unattended"
        Me.Unattended.UseVisualStyleBackColor = False
        '
        'ConnFrame
        '
        Me.ConnFrame.BackColor = System.Drawing.SystemColors.Control
        Me.ConnFrame.Controls.Add(Me.ConnLocalUI)
        Me.ConnFrame.Controls.Add(Me.ConnQBOE)
        Me.ConnFrame.Controls.Add(Me.ConnRemote)
        Me.ConnFrame.Controls.Add(Me.ConnLocal)
        Me.ConnFrame.Controls.Add(Me.ConnNone)
        Me.ConnFrame.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ConnFrame.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ConnFrame.Location = New System.Drawing.Point(168, 120)
        Me.ConnFrame.Name = "ConnFrame"
        Me.ConnFrame.Padding = New System.Windows.Forms.Padding(0)
        Me.ConnFrame.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ConnFrame.Size = New System.Drawing.Size(105, 137)
        Me.ConnFrame.TabIndex = 17
        Me.ConnFrame.TabStop = False
        Me.ConnFrame.Text = "Connection Type"
        '
        'ConnLocalUI
        '
        Me.ConnLocalUI.BackColor = System.Drawing.SystemColors.Control
        Me.ConnLocalUI.Cursor = System.Windows.Forms.Cursors.Default
        Me.ConnLocalUI.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ConnLocalUI.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ConnLocalUI.Location = New System.Drawing.Point(8, 88)
        Me.ConnLocalUI.Name = "ConnLocalUI"
        Me.ConnLocalUI.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ConnLocalUI.Size = New System.Drawing.Size(89, 17)
        Me.ConnLocalUI.TabIndex = 22
        Me.ConnLocalUI.TabStop = True
        Me.ConnLocalUI.Text = "Local with UI"
        Me.ConnLocalUI.UseVisualStyleBackColor = False
        '
        'ConnQBOE
        '
        Me.ConnQBOE.BackColor = System.Drawing.SystemColors.Control
        Me.ConnQBOE.Cursor = System.Windows.Forms.Cursors.Default
        Me.ConnQBOE.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ConnQBOE.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ConnQBOE.Location = New System.Drawing.Point(8, 112)
        Me.ConnQBOE.Name = "ConnQBOE"
        Me.ConnQBOE.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ConnQBOE.Size = New System.Drawing.Size(89, 17)
        Me.ConnQBOE.TabIndex = 21
        Me.ConnQBOE.TabStop = True
        Me.ConnQBOE.Text = "QBOE"
        Me.ConnQBOE.UseVisualStyleBackColor = False
        '
        'ConnRemote
        '
        Me.ConnRemote.BackColor = System.Drawing.SystemColors.Control
        Me.ConnRemote.Cursor = System.Windows.Forms.Cursors.Default
        Me.ConnRemote.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ConnRemote.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ConnRemote.Location = New System.Drawing.Point(8, 64)
        Me.ConnRemote.Name = "ConnRemote"
        Me.ConnRemote.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ConnRemote.Size = New System.Drawing.Size(89, 17)
        Me.ConnRemote.TabIndex = 20
        Me.ConnRemote.TabStop = True
        Me.ConnRemote.Text = "Remote QB"
        Me.ConnRemote.UseVisualStyleBackColor = False
        '
        'ConnLocal
        '
        Me.ConnLocal.BackColor = System.Drawing.SystemColors.Control
        Me.ConnLocal.Cursor = System.Windows.Forms.Cursors.Default
        Me.ConnLocal.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ConnLocal.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ConnLocal.Location = New System.Drawing.Point(8, 40)
        Me.ConnLocal.Name = "ConnLocal"
        Me.ConnLocal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ConnLocal.Size = New System.Drawing.Size(89, 17)
        Me.ConnLocal.TabIndex = 19
        Me.ConnLocal.TabStop = True
        Me.ConnLocal.Text = "Local QB"
        Me.ConnLocal.UseVisualStyleBackColor = False
        '
        'ConnNone
        '
        Me.ConnNone.BackColor = System.Drawing.SystemColors.Control
        Me.ConnNone.Checked = True
        Me.ConnNone.Cursor = System.Windows.Forms.Cursors.Default
        Me.ConnNone.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ConnNone.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ConnNone.Location = New System.Drawing.Point(8, 16)
        Me.ConnNone.Name = "ConnNone"
        Me.ConnNone.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ConnNone.Size = New System.Drawing.Size(89, 17)
        Me.ConnNone.TabIndex = 18
        Me.ConnNone.TabStop = True
        Me.ConnNone.Text = "No Pref"
        Me.ConnNone.UseVisualStyleBackColor = False
        '
        'RequestBrowse
        '
        Me.RequestBrowse.BackColor = System.Drawing.SystemColors.Control
        Me.RequestBrowse.Cursor = System.Windows.Forms.Cursors.Default
        Me.RequestBrowse.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RequestBrowse.ForeColor = System.Drawing.SystemColors.ControlText
        Me.RequestBrowse.Location = New System.Drawing.Point(520, 80)
        Me.RequestBrowse.Name = "RequestBrowse"
        Me.RequestBrowse.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.RequestBrowse.Size = New System.Drawing.Size(113, 25)
        Me.RequestBrowse.TabIndex = 16
        Me.RequestBrowse.Text = "Browse..."
        Me.RequestBrowse.UseVisualStyleBackColor = False
        '
        'RequestFile
        '
        Me.RequestFile.AcceptsReturn = True
        Me.RequestFile.BackColor = System.Drawing.SystemColors.Window
        Me.RequestFile.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.RequestFile.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RequestFile.ForeColor = System.Drawing.SystemColors.WindowText
        Me.RequestFile.Location = New System.Drawing.Point(16, 88)
        Me.RequestFile.MaxLength = 0
        Me.RequestFile.Name = "RequestFile"
        Me.RequestFile.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.RequestFile.Size = New System.Drawing.Size(497, 20)
        Me.RequestFile.TabIndex = 14
        '
        'CompanyBrowse
        '
        Me.CompanyBrowse.BackColor = System.Drawing.SystemColors.Control
        Me.CompanyBrowse.Cursor = System.Windows.Forms.Cursors.Default
        Me.CompanyBrowse.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CompanyBrowse.ForeColor = System.Drawing.SystemColors.ControlText
        Me.CompanyBrowse.Location = New System.Drawing.Point(520, 40)
        Me.CompanyBrowse.Name = "CompanyBrowse"
        Me.CompanyBrowse.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.CompanyBrowse.Size = New System.Drawing.Size(113, 25)
        Me.CompanyBrowse.TabIndex = 13
        Me.CompanyBrowse.Text = "Browse..."
        Me.CompanyBrowse.UseVisualStyleBackColor = False
        '
        'CompanyFile
        '
        Me.CompanyFile.AcceptsReturn = True
        Me.CompanyFile.BackColor = System.Drawing.SystemColors.Window
        Me.CompanyFile.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.CompanyFile.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CompanyFile.ForeColor = System.Drawing.SystemColors.WindowText
        Me.CompanyFile.Location = New System.Drawing.Point(16, 48)
        Me.CompanyFile.MaxLength = 0
        Me.CompanyFile.Name = "CompanyFile"
        Me.CompanyFile.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.CompanyFile.Size = New System.Drawing.Size(497, 20)
        Me.CompanyFile.TabIndex = 11
        '
        'cmdProcessSubscription
        '
        Me.cmdProcessSubscription.BackColor = System.Drawing.SystemColors.Control
        Me.cmdProcessSubscription.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdProcessSubscription.Enabled = False
        Me.cmdProcessSubscription.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdProcessSubscription.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdProcessSubscription.Location = New System.Drawing.Point(184, 336)
        Me.cmdProcessSubscription.Name = "cmdProcessSubscription"
        Me.cmdProcessSubscription.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdProcessSubscription.Size = New System.Drawing.Size(121, 33)
        Me.cmdProcessSubscription.TabIndex = 10
        Me.cmdProcessSubscription.Text = "Send XML to Subscription Processor"
        Me.cmdProcessSubscription.UseVisualStyleBackColor = False
        '
        'cmdSendXML
        '
        Me.cmdSendXML.BackColor = System.Drawing.SystemColors.Control
        Me.cmdSendXML.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdSendXML.Enabled = False
        Me.cmdSendXML.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSendXML.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdSendXML.Location = New System.Drawing.Point(416, 280)
        Me.cmdSendXML.Name = "cmdSendXML"
        Me.cmdSendXML.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdSendXML.Size = New System.Drawing.Size(105, 33)
        Me.cmdSendXML.TabIndex = 8
        Me.cmdSendXML.Text = "Send XML to Request Processor"
        Me.cmdSendXML.UseVisualStyleBackColor = False
        '
        'cmdViewInput
        '
        Me.cmdViewInput.BackColor = System.Drawing.SystemColors.Control
        Me.cmdViewInput.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdViewInput.Enabled = False
        Me.cmdViewInput.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdViewInput.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdViewInput.Location = New System.Drawing.Point(528, 200)
        Me.cmdViewInput.Name = "cmdViewInput"
        Me.cmdViewInput.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdViewInput.Size = New System.Drawing.Size(105, 25)
        Me.cmdViewInput.TabIndex = 7
        Me.cmdViewInput.Text = "View Input"
        Me.cmdViewInput.UseVisualStyleBackColor = False
        '
        'cmdViewOutput
        '
        Me.cmdViewOutput.BackColor = System.Drawing.SystemColors.Control
        Me.cmdViewOutput.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdViewOutput.Enabled = False
        Me.cmdViewOutput.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdViewOutput.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdViewOutput.Location = New System.Drawing.Point(528, 232)
        Me.cmdViewOutput.Name = "cmdViewOutput"
        Me.cmdViewOutput.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdViewOutput.Size = New System.Drawing.Size(105, 25)
        Me.cmdViewOutput.TabIndex = 6
        Me.cmdViewOutput.Text = "View Output"
        Me.cmdViewOutput.UseVisualStyleBackColor = False
        '
        'cmdExit
        '
        Me.cmdExit.BackColor = System.Drawing.SystemColors.Control
        Me.cmdExit.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdExit.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdExit.Location = New System.Drawing.Point(552, 490)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdExit.Size = New System.Drawing.Size(89, 33)
        Me.cmdExit.TabIndex = 5
        Me.cmdExit.Text = "Exit"
        Me.cmdExit.UseVisualStyleBackColor = False
        '
        'cmdOpenConnection
        '
        Me.cmdOpenConnection.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOpenConnection.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOpenConnection.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOpenConnection.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOpenConnection.Location = New System.Drawing.Point(16, 280)
        Me.cmdOpenConnection.Name = "cmdOpenConnection"
        Me.cmdOpenConnection.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOpenConnection.Size = New System.Drawing.Size(105, 33)
        Me.cmdOpenConnection.TabIndex = 4
        Me.cmdOpenConnection.Text = "Open Connection"
        Me.cmdOpenConnection.UseVisualStyleBackColor = False
        '
        'cmdBeginSession
        '
        Me.cmdBeginSession.BackColor = System.Drawing.SystemColors.Control
        Me.cmdBeginSession.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdBeginSession.Enabled = False
        Me.cmdBeginSession.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBeginSession.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdBeginSession.Location = New System.Drawing.Point(184, 280)
        Me.cmdBeginSession.Name = "cmdBeginSession"
        Me.cmdBeginSession.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdBeginSession.Size = New System.Drawing.Size(121, 33)
        Me.cmdBeginSession.TabIndex = 3
        Me.cmdBeginSession.Text = "Begin Session"
        Me.cmdBeginSession.UseVisualStyleBackColor = False
        '
        'cmdEndSession
        '
        Me.cmdEndSession.BackColor = System.Drawing.SystemColors.Control
        Me.cmdEndSession.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdEndSession.Enabled = False
        Me.cmdEndSession.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdEndSession.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdEndSession.Location = New System.Drawing.Point(552, 280)
        Me.cmdEndSession.Name = "cmdEndSession"
        Me.cmdEndSession.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdEndSession.Size = New System.Drawing.Size(81, 33)
        Me.cmdEndSession.TabIndex = 2
        Me.cmdEndSession.Text = "End Session"
        Me.cmdEndSession.UseVisualStyleBackColor = False
        '
        'cmdCloseConnection
        '
        Me.cmdCloseConnection.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCloseConnection.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCloseConnection.Enabled = False
        Me.cmdCloseConnection.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCloseConnection.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCloseConnection.Location = New System.Drawing.Point(544, 336)
        Me.cmdCloseConnection.Name = "cmdCloseConnection"
        Me.cmdCloseConnection.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCloseConnection.Size = New System.Drawing.Size(89, 33)
        Me.cmdCloseConnection.TabIndex = 1
        Me.cmdCloseConnection.Text = "Close Connection"
        Me.cmdCloseConnection.UseVisualStyleBackColor = False
        '
        'txtApplicationName
        '
        Me.txtApplicationName.AcceptsReturn = True
        Me.txtApplicationName.BackColor = System.Drawing.SystemColors.Window
        Me.txtApplicationName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtApplicationName.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtApplicationName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtApplicationName.Location = New System.Drawing.Point(127, 8)
        Me.txtApplicationName.MaxLength = 0
        Me.txtApplicationName.Name = "txtApplicationName"
        Me.txtApplicationName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtApplicationName.Size = New System.Drawing.Size(138, 20)
        Me.txtApplicationName.TabIndex = 0
        Me.txtApplicationName.Text = "SDKTest Plus 3"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(16, 376)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(57, 17)
        Me.Label4.TabIndex = 34
        Me.Label4.Text = "Status:"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(16, 72)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(153, 17)
        Me.Label3.TabIndex = 15
        Me.Label3.Text = "Request File:"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(16, 32)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(265, 17)
        Me.Label1.TabIndex = 12
        Me.Label1.Text = "Company File: (leave blank to use current open file)"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(16, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(105, 17)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "Application Name:"
        '
        'frmSDKTestPlus3
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(651, 531)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.AuthFlagsFrame)
        Me.Controls.Add(Me.SessionPrefsFrame)
        Me.Controls.Add(Me.ConnFrame)
        Me.Controls.Add(Me.RequestBrowse)
        Me.Controls.Add(Me.RequestFile)
        Me.Controls.Add(Me.CompanyBrowse)
        Me.Controls.Add(Me.CompanyFile)
        Me.Controls.Add(Me.cmdProcessSubscription)
        Me.Controls.Add(Me.cmdSendXML)
        Me.Controls.Add(Me.cmdViewInput)
        Me.Controls.Add(Me.cmdViewOutput)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.cmdOpenConnection)
        Me.Controls.Add(Me.cmdBeginSession)
        Me.Controls.Add(Me.cmdEndSession)
        Me.Controls.Add(Me.cmdCloseConnection)
        Me.Controls.Add(Me.txtApplicationName)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label2)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(194, 146)
        Me.Name = "frmSDKTestPlus3"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "SDKTestPlus3"
        Me.AuthFlagsFrame.ResumeLayout(False)
        Me.SessionPrefsFrame.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.pdFrame.ResumeLayout(False)
        Me.ConnFrame.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
End Class