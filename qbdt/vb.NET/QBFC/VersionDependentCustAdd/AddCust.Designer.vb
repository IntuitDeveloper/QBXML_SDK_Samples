<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class AddCust
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
	Public WithEvents Text_DataExtValue As System.Windows.Forms.TextBox
	Public WithEvents Text_DataExtName As System.Windows.Forms.TextBox
	Public WithEvents Label14 As System.Windows.Forms.Label
	Public WithEvents Label13 As System.Windows.Forms.Label
	Public WithEvents Frame3 As System.Windows.Forms.GroupBox
	Public WithEvents Comm_Exit As System.Windows.Forms.Button
	Public WithEvents Comm_View_Res As System.Windows.Forms.Button
	Public WithEvents Comm_View_Req As System.Windows.Forms.Button
	Public WithEvents Text_CustomerName As System.Windows.Forms.TextBox
	Public WithEvents Text_Phone As System.Windows.Forms.TextBox
	Public WithEvents Comm_Submit As System.Windows.Forms.Button
	Public WithEvents Text_LastName As System.Windows.Forms.TextBox
	Public WithEvents Text_FirstName As System.Windows.Forms.TextBox
	Public WithEvents Text_Country As System.Windows.Forms.TextBox
	Public WithEvents Text_PostalCode As System.Windows.Forms.TextBox
	Public WithEvents Text_State As System.Windows.Forms.TextBox
	Public WithEvents Text_City As System.Windows.Forms.TextBox
	Public WithEvents Text_Addr4 As System.Windows.Forms.TextBox
	Public WithEvents Text_Addr3 As System.Windows.Forms.TextBox
	Public WithEvents Text_Addr2 As System.Windows.Forms.TextBox
	Public WithEvents Text_Addr1 As System.Windows.Forms.TextBox
	Public WithEvents Label12 As System.Windows.Forms.Label
	Public WithEvents Label11 As System.Windows.Forms.Label
	Public WithEvents Label10 As System.Windows.Forms.Label
	Public WithEvents Label9 As System.Windows.Forms.Label
	Public WithEvents Label8 As System.Windows.Forms.Label
	Public WithEvents Label6 As System.Windows.Forms.Label
	Public WithEvents Label5 As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Frame2 As System.Windows.Forms.GroupBox
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents Label7 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(AddCust))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.UseQBOE = New System.Windows.Forms.CheckBox
		Me.Frame3 = New System.Windows.Forms.GroupBox
		Me.Text_DataExtValue = New System.Windows.Forms.TextBox
		Me.Text_DataExtName = New System.Windows.Forms.TextBox
		Me.Label14 = New System.Windows.Forms.Label
		Me.Label13 = New System.Windows.Forms.Label
		Me.Comm_Exit = New System.Windows.Forms.Button
		Me.Comm_View_Res = New System.Windows.Forms.Button
		Me.Comm_View_Req = New System.Windows.Forms.Button
		Me.Text_CustomerName = New System.Windows.Forms.TextBox
		Me.Text_Phone = New System.Windows.Forms.TextBox
		Me.Comm_Submit = New System.Windows.Forms.Button
		Me.Text_LastName = New System.Windows.Forms.TextBox
		Me.Text_FirstName = New System.Windows.Forms.TextBox
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me.Frame2 = New System.Windows.Forms.GroupBox
		Me.Text_Country = New System.Windows.Forms.TextBox
		Me.Text_PostalCode = New System.Windows.Forms.TextBox
		Me.Text_State = New System.Windows.Forms.TextBox
		Me.Text_City = New System.Windows.Forms.TextBox
		Me.Text_Addr4 = New System.Windows.Forms.TextBox
		Me.Text_Addr3 = New System.Windows.Forms.TextBox
		Me.Text_Addr2 = New System.Windows.Forms.TextBox
		Me.Text_Addr1 = New System.Windows.Forms.TextBox
		Me.Label12 = New System.Windows.Forms.Label
		Me.Label11 = New System.Windows.Forms.Label
		Me.Label10 = New System.Windows.Forms.Label
		Me.Label9 = New System.Windows.Forms.Label
		Me.Label8 = New System.Windows.Forms.Label
		Me.Label6 = New System.Windows.Forms.Label
		Me.Label5 = New System.Windows.Forms.Label
		Me.Label4 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.Label7 = New System.Windows.Forms.Label
		Me.Frame3.SuspendLayout()
		Me.Frame1.SuspendLayout()
		Me.Frame2.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "QBFC Sample: Create Requests dependent on the qbxml version"
		Me.ClientSize = New System.Drawing.Size(438, 559)
		Me.Location = New System.Drawing.Point(3, 22)
		Me.ControlBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MaximizeBox = True
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "AddCust"
		Me.UseQBOE.Text = "Communicate with QuickBooks Online Edition"
		Me.UseQBOE.Size = New System.Drawing.Size(241, 25)
		Me.UseQBOE.Location = New System.Drawing.Point(16, 8)
		Me.UseQBOE.TabIndex = 35
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
		Me.Frame3.Text = "Data Extension Information (optional)"
		Me.Frame3.Size = New System.Drawing.Size(409, 73)
		Me.Frame3.Location = New System.Drawing.Point(16, 480)
		Me.Frame3.TabIndex = 26
		Me.Frame3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame3.BackColor = System.Drawing.SystemColors.Control
		Me.Frame3.Enabled = True
		Me.Frame3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame3.Visible = True
		Me.Frame3.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame3.Name = "Frame3"
		Me.Text_DataExtValue.AutoSize = False
		Me.Text_DataExtValue.Size = New System.Drawing.Size(249, 19)
		Me.Text_DataExtValue.Location = New System.Drawing.Point(144, 48)
		Me.Text_DataExtValue.TabIndex = 30
		Me.Text_DataExtValue.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text_DataExtValue.AcceptsReturn = True
		Me.Text_DataExtValue.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text_DataExtValue.BackColor = System.Drawing.SystemColors.Window
		Me.Text_DataExtValue.CausesValidation = True
		Me.Text_DataExtValue.Enabled = True
		Me.Text_DataExtValue.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text_DataExtValue.HideSelection = True
		Me.Text_DataExtValue.ReadOnly = False
		Me.Text_DataExtValue.Maxlength = 0
		Me.Text_DataExtValue.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text_DataExtValue.MultiLine = False
		Me.Text_DataExtValue.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text_DataExtValue.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text_DataExtValue.TabStop = True
		Me.Text_DataExtValue.Visible = True
		Me.Text_DataExtValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text_DataExtValue.Name = "Text_DataExtValue"
		Me.Text_DataExtName.AutoSize = False
		Me.Text_DataExtName.Size = New System.Drawing.Size(249, 19)
		Me.Text_DataExtName.Location = New System.Drawing.Point(144, 24)
		Me.Text_DataExtName.TabIndex = 28
		Me.Text_DataExtName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text_DataExtName.AcceptsReturn = True
		Me.Text_DataExtName.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text_DataExtName.BackColor = System.Drawing.SystemColors.Window
		Me.Text_DataExtName.CausesValidation = True
		Me.Text_DataExtName.Enabled = True
		Me.Text_DataExtName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text_DataExtName.HideSelection = True
		Me.Text_DataExtName.ReadOnly = False
		Me.Text_DataExtName.Maxlength = 0
		Me.Text_DataExtName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text_DataExtName.MultiLine = False
		Me.Text_DataExtName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text_DataExtName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text_DataExtName.TabStop = True
		Me.Text_DataExtName.Visible = True
		Me.Text_DataExtName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text_DataExtName.Name = "Text_DataExtName"
		Me.Label14.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label14.Text = "Data Extension Val&ue:"
		Me.Label14.Size = New System.Drawing.Size(121, 17)
		Me.Label14.Location = New System.Drawing.Point(16, 48)
		Me.Label14.TabIndex = 29
		Me.Label14.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label14.BackColor = System.Drawing.SystemColors.Control
		Me.Label14.Enabled = True
		Me.Label14.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label14.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label14.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label14.UseMnemonic = True
		Me.Label14.Visible = True
		Me.Label14.AutoSize = False
		Me.Label14.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label14.Name = "Label14"
		Me.Label13.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label13.Text = "Data E&xtension Name:"
		Me.Label13.Size = New System.Drawing.Size(121, 17)
		Me.Label13.Location = New System.Drawing.Point(16, 24)
		Me.Label13.TabIndex = 27
		Me.Label13.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label13.BackColor = System.Drawing.SystemColors.Control
		Me.Label13.Enabled = True
		Me.Label13.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label13.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label13.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label13.UseMnemonic = True
		Me.Label13.Visible = True
		Me.Label13.AutoSize = False
		Me.Label13.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label13.Name = "Label13"
		Me.Comm_Exit.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Comm_Exit.Text = "&Done"
		Me.Comm_Exit.Size = New System.Drawing.Size(145, 25)
		Me.Comm_Exit.Location = New System.Drawing.Point(280, 40)
		Me.Comm_Exit.TabIndex = 34
		Me.Comm_Exit.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Comm_Exit.BackColor = System.Drawing.SystemColors.Control
		Me.Comm_Exit.CausesValidation = True
		Me.Comm_Exit.Enabled = True
		Me.Comm_Exit.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Comm_Exit.Cursor = System.Windows.Forms.Cursors.Default
		Me.Comm_Exit.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Comm_Exit.TabStop = True
		Me.Comm_Exit.Name = "Comm_Exit"
		Me.Comm_View_Res.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Comm_View_Res.Text = "Vi&ew qbXML Response"
		Me.Comm_View_Res.Size = New System.Drawing.Size(145, 25)
		Me.Comm_View_Res.Location = New System.Drawing.Point(280, 104)
		Me.Comm_View_Res.TabIndex = 33
		Me.Comm_View_Res.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Comm_View_Res.BackColor = System.Drawing.SystemColors.Control
		Me.Comm_View_Res.CausesValidation = True
		Me.Comm_View_Res.Enabled = True
		Me.Comm_View_Res.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Comm_View_Res.Cursor = System.Windows.Forms.Cursors.Default
		Me.Comm_View_Res.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Comm_View_Res.TabStop = True
		Me.Comm_View_Res.Name = "Comm_View_Res"
		Me.Comm_View_Req.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Comm_View_Req.Text = "&View qbXML Request"
		Me.Comm_View_Req.Size = New System.Drawing.Size(145, 25)
		Me.Comm_View_Req.Location = New System.Drawing.Point(280, 72)
		Me.Comm_View_Req.TabIndex = 32
		Me.Comm_View_Req.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Comm_View_Req.BackColor = System.Drawing.SystemColors.Control
		Me.Comm_View_Req.CausesValidation = True
		Me.Comm_View_Req.Enabled = True
		Me.Comm_View_Req.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Comm_View_Req.Cursor = System.Windows.Forms.Cursors.Default
		Me.Comm_View_Req.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Comm_View_Req.TabStop = True
		Me.Comm_View_Req.Name = "Comm_View_Req"
		Me.Text_CustomerName.AutoSize = False
		Me.Text_CustomerName.Size = New System.Drawing.Size(169, 20)
		Me.Text_CustomerName.Location = New System.Drawing.Point(24, 96)
		Me.Text_CustomerName.TabIndex = 1
		Me.Text_CustomerName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text_CustomerName.AcceptsReturn = True
		Me.Text_CustomerName.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text_CustomerName.BackColor = System.Drawing.SystemColors.Window
		Me.Text_CustomerName.CausesValidation = True
		Me.Text_CustomerName.Enabled = True
		Me.Text_CustomerName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text_CustomerName.HideSelection = True
		Me.Text_CustomerName.ReadOnly = False
		Me.Text_CustomerName.Maxlength = 0
		Me.Text_CustomerName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text_CustomerName.MultiLine = False
		Me.Text_CustomerName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text_CustomerName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text_CustomerName.TabStop = True
		Me.Text_CustomerName.Visible = True
		Me.Text_CustomerName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text_CustomerName.Name = "Text_CustomerName"
		Me.Text_Phone.AutoSize = False
		Me.Text_Phone.Size = New System.Drawing.Size(265, 20)
		Me.Text_Phone.Location = New System.Drawing.Point(128, 216)
		Me.Text_Phone.TabIndex = 8
		Me.Text_Phone.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text_Phone.AcceptsReturn = True
		Me.Text_Phone.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text_Phone.BackColor = System.Drawing.SystemColors.Window
		Me.Text_Phone.CausesValidation = True
		Me.Text_Phone.Enabled = True
		Me.Text_Phone.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text_Phone.HideSelection = True
		Me.Text_Phone.ReadOnly = False
		Me.Text_Phone.Maxlength = 0
		Me.Text_Phone.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text_Phone.MultiLine = False
		Me.Text_Phone.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text_Phone.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text_Phone.TabStop = True
		Me.Text_Phone.Visible = True
		Me.Text_Phone.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text_Phone.Name = "Text_Phone"
		Me.Comm_Submit.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Comm_Submit.Text = "&Add Customer"
		Me.AcceptButton = Me.Comm_Submit
		Me.Comm_Submit.Size = New System.Drawing.Size(145, 25)
		Me.Comm_Submit.Location = New System.Drawing.Point(280, 8)
		Me.Comm_Submit.TabIndex = 31
		Me.Comm_Submit.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Comm_Submit.BackColor = System.Drawing.SystemColors.Control
		Me.Comm_Submit.CausesValidation = True
		Me.Comm_Submit.Enabled = True
		Me.Comm_Submit.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Comm_Submit.Cursor = System.Windows.Forms.Cursors.Default
		Me.Comm_Submit.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Comm_Submit.TabStop = True
		Me.Comm_Submit.Name = "Comm_Submit"
		Me.Text_LastName.AutoSize = False
		Me.Text_LastName.Size = New System.Drawing.Size(265, 20)
		Me.Text_LastName.Location = New System.Drawing.Point(128, 184)
		Me.Text_LastName.TabIndex = 6
		Me.Text_LastName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text_LastName.AcceptsReturn = True
		Me.Text_LastName.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text_LastName.BackColor = System.Drawing.SystemColors.Window
		Me.Text_LastName.CausesValidation = True
		Me.Text_LastName.Enabled = True
		Me.Text_LastName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text_LastName.HideSelection = True
		Me.Text_LastName.ReadOnly = False
		Me.Text_LastName.Maxlength = 0
		Me.Text_LastName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text_LastName.MultiLine = False
		Me.Text_LastName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text_LastName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text_LastName.TabStop = True
		Me.Text_LastName.Visible = True
		Me.Text_LastName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text_LastName.Name = "Text_LastName"
		Me.Text_FirstName.AutoSize = False
		Me.Text_FirstName.Size = New System.Drawing.Size(265, 20)
		Me.Text_FirstName.Location = New System.Drawing.Point(128, 152)
		Me.Text_FirstName.TabIndex = 4
		Me.Text_FirstName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text_FirstName.AcceptsReturn = True
		Me.Text_FirstName.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text_FirstName.BackColor = System.Drawing.SystemColors.Window
		Me.Text_FirstName.CausesValidation = True
		Me.Text_FirstName.Enabled = True
		Me.Text_FirstName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text_FirstName.HideSelection = True
		Me.Text_FirstName.ReadOnly = False
		Me.Text_FirstName.Maxlength = 0
		Me.Text_FirstName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text_FirstName.MultiLine = False
		Me.Text_FirstName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text_FirstName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text_FirstName.TabStop = True
		Me.Text_FirstName.Visible = True
		Me.Text_FirstName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text_FirstName.Name = "Text_FirstName"
		Me.Frame1.Text = " Contact Information (optional) "
		Me.Frame1.Size = New System.Drawing.Size(409, 345)
		Me.Frame1.Location = New System.Drawing.Point(16, 128)
		Me.Frame1.TabIndex = 2
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.BackColor = System.Drawing.SystemColors.Control
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame1.Name = "Frame1"
		Me.Frame2.Text = "Billing Address"
		Me.Frame2.Size = New System.Drawing.Size(377, 225)
		Me.Frame2.Location = New System.Drawing.Point(16, 112)
		Me.Frame2.TabIndex = 9
		Me.Frame2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame2.BackColor = System.Drawing.SystemColors.Control
		Me.Frame2.Enabled = True
		Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame2.Visible = True
		Me.Frame2.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame2.Name = "Frame2"
		Me.Text_Country.AutoSize = False
		Me.Text_Country.Size = New System.Drawing.Size(265, 19)
		Me.Text_Country.Location = New System.Drawing.Point(96, 192)
		Me.Text_Country.TabIndex = 25
		Me.Text_Country.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text_Country.AcceptsReturn = True
		Me.Text_Country.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text_Country.BackColor = System.Drawing.SystemColors.Window
		Me.Text_Country.CausesValidation = True
		Me.Text_Country.Enabled = True
		Me.Text_Country.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text_Country.HideSelection = True
		Me.Text_Country.ReadOnly = False
		Me.Text_Country.Maxlength = 0
		Me.Text_Country.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text_Country.MultiLine = False
		Me.Text_Country.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text_Country.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text_Country.TabStop = True
		Me.Text_Country.Visible = True
		Me.Text_Country.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text_Country.Name = "Text_Country"
		Me.Text_PostalCode.AutoSize = False
		Me.Text_PostalCode.Size = New System.Drawing.Size(265, 19)
		Me.Text_PostalCode.Location = New System.Drawing.Point(96, 168)
		Me.Text_PostalCode.TabIndex = 23
		Me.Text_PostalCode.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text_PostalCode.AcceptsReturn = True
		Me.Text_PostalCode.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text_PostalCode.BackColor = System.Drawing.SystemColors.Window
		Me.Text_PostalCode.CausesValidation = True
		Me.Text_PostalCode.Enabled = True
		Me.Text_PostalCode.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text_PostalCode.HideSelection = True
		Me.Text_PostalCode.ReadOnly = False
		Me.Text_PostalCode.Maxlength = 0
		Me.Text_PostalCode.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text_PostalCode.MultiLine = False
		Me.Text_PostalCode.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text_PostalCode.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text_PostalCode.TabStop = True
		Me.Text_PostalCode.Visible = True
		Me.Text_PostalCode.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text_PostalCode.Name = "Text_PostalCode"
		Me.Text_State.AutoSize = False
		Me.Text_State.Size = New System.Drawing.Size(265, 19)
		Me.Text_State.Location = New System.Drawing.Point(96, 144)
		Me.Text_State.TabIndex = 21
		Me.Text_State.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text_State.AcceptsReturn = True
		Me.Text_State.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text_State.BackColor = System.Drawing.SystemColors.Window
		Me.Text_State.CausesValidation = True
		Me.Text_State.Enabled = True
		Me.Text_State.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text_State.HideSelection = True
		Me.Text_State.ReadOnly = False
		Me.Text_State.Maxlength = 0
		Me.Text_State.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text_State.MultiLine = False
		Me.Text_State.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text_State.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text_State.TabStop = True
		Me.Text_State.Visible = True
		Me.Text_State.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text_State.Name = "Text_State"
		Me.Text_City.AutoSize = False
		Me.Text_City.Size = New System.Drawing.Size(265, 19)
		Me.Text_City.Location = New System.Drawing.Point(96, 120)
		Me.Text_City.TabIndex = 19
		Me.Text_City.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text_City.AcceptsReturn = True
		Me.Text_City.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text_City.BackColor = System.Drawing.SystemColors.Window
		Me.Text_City.CausesValidation = True
		Me.Text_City.Enabled = True
		Me.Text_City.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text_City.HideSelection = True
		Me.Text_City.ReadOnly = False
		Me.Text_City.Maxlength = 0
		Me.Text_City.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text_City.MultiLine = False
		Me.Text_City.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text_City.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text_City.TabStop = True
		Me.Text_City.Visible = True
		Me.Text_City.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text_City.Name = "Text_City"
		Me.Text_Addr4.AutoSize = False
		Me.Text_Addr4.Size = New System.Drawing.Size(265, 19)
		Me.Text_Addr4.Location = New System.Drawing.Point(96, 96)
		Me.Text_Addr4.TabIndex = 17
		Me.Text_Addr4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text_Addr4.AcceptsReturn = True
		Me.Text_Addr4.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text_Addr4.BackColor = System.Drawing.SystemColors.Window
		Me.Text_Addr4.CausesValidation = True
		Me.Text_Addr4.Enabled = True
		Me.Text_Addr4.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text_Addr4.HideSelection = True
		Me.Text_Addr4.ReadOnly = False
		Me.Text_Addr4.Maxlength = 0
		Me.Text_Addr4.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text_Addr4.MultiLine = False
		Me.Text_Addr4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text_Addr4.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text_Addr4.TabStop = True
		Me.Text_Addr4.Visible = True
		Me.Text_Addr4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text_Addr4.Name = "Text_Addr4"
		Me.Text_Addr3.AutoSize = False
		Me.Text_Addr3.Size = New System.Drawing.Size(265, 19)
		Me.Text_Addr3.Location = New System.Drawing.Point(96, 72)
		Me.Text_Addr3.TabIndex = 15
		Me.Text_Addr3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text_Addr3.AcceptsReturn = True
		Me.Text_Addr3.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text_Addr3.BackColor = System.Drawing.SystemColors.Window
		Me.Text_Addr3.CausesValidation = True
		Me.Text_Addr3.Enabled = True
		Me.Text_Addr3.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text_Addr3.HideSelection = True
		Me.Text_Addr3.ReadOnly = False
		Me.Text_Addr3.Maxlength = 0
		Me.Text_Addr3.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text_Addr3.MultiLine = False
		Me.Text_Addr3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text_Addr3.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text_Addr3.TabStop = True
		Me.Text_Addr3.Visible = True
		Me.Text_Addr3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text_Addr3.Name = "Text_Addr3"
		Me.Text_Addr2.AutoSize = False
		Me.Text_Addr2.Size = New System.Drawing.Size(265, 19)
		Me.Text_Addr2.Location = New System.Drawing.Point(96, 48)
		Me.Text_Addr2.TabIndex = 13
		Me.Text_Addr2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text_Addr2.AcceptsReturn = True
		Me.Text_Addr2.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text_Addr2.BackColor = System.Drawing.SystemColors.Window
		Me.Text_Addr2.CausesValidation = True
		Me.Text_Addr2.Enabled = True
		Me.Text_Addr2.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text_Addr2.HideSelection = True
		Me.Text_Addr2.ReadOnly = False
		Me.Text_Addr2.Maxlength = 0
		Me.Text_Addr2.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text_Addr2.MultiLine = False
		Me.Text_Addr2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text_Addr2.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text_Addr2.TabStop = True
		Me.Text_Addr2.Visible = True
		Me.Text_Addr2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text_Addr2.Name = "Text_Addr2"
		Me.Text_Addr1.AutoSize = False
		Me.Text_Addr1.Size = New System.Drawing.Size(265, 19)
		Me.Text_Addr1.Location = New System.Drawing.Point(96, 24)
		Me.Text_Addr1.TabIndex = 11
		Me.Text_Addr1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text_Addr1.AcceptsReturn = True
		Me.Text_Addr1.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text_Addr1.BackColor = System.Drawing.SystemColors.Window
		Me.Text_Addr1.CausesValidation = True
		Me.Text_Addr1.Enabled = True
		Me.Text_Addr1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text_Addr1.HideSelection = True
		Me.Text_Addr1.ReadOnly = False
		Me.Text_Addr1.Maxlength = 0
		Me.Text_Addr1.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text_Addr1.MultiLine = False
		Me.Text_Addr1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text_Addr1.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text_Addr1.TabStop = True
		Me.Text_Addr1.Visible = True
		Me.Text_Addr1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text_Addr1.Name = "Text_Addr1"
		Me.Label12.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label12.Text = "Countr&y:"
		Me.Label12.Size = New System.Drawing.Size(73, 17)
		Me.Label12.Location = New System.Drawing.Point(16, 192)
		Me.Label12.TabIndex = 24
		Me.Label12.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label12.BackColor = System.Drawing.SystemColors.Control
		Me.Label12.Enabled = True
		Me.Label12.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label12.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label12.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label12.UseMnemonic = True
		Me.Label12.Visible = True
		Me.Label12.AutoSize = False
		Me.Label12.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label12.Name = "Label12"
		Me.Label11.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label11.Text = "P&ostal Code:"
		Me.Label11.Size = New System.Drawing.Size(73, 17)
		Me.Label11.Location = New System.Drawing.Point(16, 168)
		Me.Label11.TabIndex = 22
		Me.Label11.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label11.BackColor = System.Drawing.SystemColors.Control
		Me.Label11.Enabled = True
		Me.Label11.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label11.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label11.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label11.UseMnemonic = True
		Me.Label11.Visible = True
		Me.Label11.AutoSize = False
		Me.Label11.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label11.Name = "Label11"
		Me.Label10.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label10.Text = "&State:"
		Me.Label10.Size = New System.Drawing.Size(73, 17)
		Me.Label10.Location = New System.Drawing.Point(16, 144)
		Me.Label10.TabIndex = 20
		Me.Label10.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label10.BackColor = System.Drawing.SystemColors.Control
		Me.Label10.Enabled = True
		Me.Label10.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label10.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label10.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label10.UseMnemonic = True
		Me.Label10.Visible = True
		Me.Label10.AutoSize = False
		Me.Label10.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label10.Name = "Label10"
		Me.Label9.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label9.Text = "Ci&ty:"
		Me.Label9.Size = New System.Drawing.Size(73, 17)
		Me.Label9.Location = New System.Drawing.Point(16, 120)
		Me.Label9.TabIndex = 18
		Me.Label9.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label9.BackColor = System.Drawing.SystemColors.Control
		Me.Label9.Enabled = True
		Me.Label9.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label9.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label9.UseMnemonic = True
		Me.Label9.Visible = True
		Me.Label9.AutoSize = False
		Me.Label9.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label9.Name = "Label9"
		Me.Label8.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label8.Text = "Address Line &4:"
		Me.Label8.Size = New System.Drawing.Size(73, 17)
		Me.Label8.Location = New System.Drawing.Point(16, 96)
		Me.Label8.TabIndex = 16
		Me.Label8.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
		Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label6.Text = "Address Line &3:"
		Me.Label6.Size = New System.Drawing.Size(73, 17)
		Me.Label6.Location = New System.Drawing.Point(16, 72)
		Me.Label6.TabIndex = 14
		Me.Label6.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
		Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label5.Text = "Address Line &2:"
		Me.Label5.Size = New System.Drawing.Size(73, 17)
		Me.Label5.Location = New System.Drawing.Point(16, 48)
		Me.Label5.TabIndex = 12
		Me.Label5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
		Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label4.Text = "Address Line &1:"
		Me.Label4.Size = New System.Drawing.Size(73, 17)
		Me.Label4.Location = New System.Drawing.Point(16, 24)
		Me.Label4.TabIndex = 10
		Me.Label4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
		Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label3.Text = "&Phone:"
		Me.Label3.Size = New System.Drawing.Size(73, 17)
		Me.Label3.Location = New System.Drawing.Point(32, 88)
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
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label2.Text = "&Last Name:"
		Me.Label2.Size = New System.Drawing.Size(73, 17)
		Me.Label2.Location = New System.Drawing.Point(32, 56)
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
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label1.Text = "&First Name:"
		Me.Label1.Size = New System.Drawing.Size(73, 17)
		Me.Label1.Location = New System.Drawing.Point(32, 24)
		Me.Label1.TabIndex = 3
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
		Me.Label7.Text = "&Customer to add:"
		Me.Label7.Size = New System.Drawing.Size(89, 17)
		Me.Label7.Location = New System.Drawing.Point(24, 80)
		Me.Label7.TabIndex = 0
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
		Me.Controls.Add(UseQBOE)
		Me.Controls.Add(Frame3)
		Me.Controls.Add(Comm_Exit)
		Me.Controls.Add(Comm_View_Res)
		Me.Controls.Add(Comm_View_Req)
		Me.Controls.Add(Text_CustomerName)
		Me.Controls.Add(Text_Phone)
		Me.Controls.Add(Comm_Submit)
		Me.Controls.Add(Text_LastName)
		Me.Controls.Add(Text_FirstName)
		Me.Controls.Add(Frame1)
		Me.Controls.Add(Label7)
		Me.Frame3.Controls.Add(Text_DataExtValue)
		Me.Frame3.Controls.Add(Text_DataExtName)
		Me.Frame3.Controls.Add(Label14)
		Me.Frame3.Controls.Add(Label13)
		Me.Frame1.Controls.Add(Frame2)
		Me.Frame1.Controls.Add(Label3)
		Me.Frame1.Controls.Add(Label2)
		Me.Frame1.Controls.Add(Label1)
		Me.Frame2.Controls.Add(Text_Country)
		Me.Frame2.Controls.Add(Text_PostalCode)
		Me.Frame2.Controls.Add(Text_State)
		Me.Frame2.Controls.Add(Text_City)
		Me.Frame2.Controls.Add(Text_Addr4)
		Me.Frame2.Controls.Add(Text_Addr3)
		Me.Frame2.Controls.Add(Text_Addr2)
		Me.Frame2.Controls.Add(Text_Addr1)
		Me.Frame2.Controls.Add(Label12)
		Me.Frame2.Controls.Add(Label11)
		Me.Frame2.Controls.Add(Label10)
		Me.Frame2.Controls.Add(Label9)
		Me.Frame2.Controls.Add(Label8)
		Me.Frame2.Controls.Add(Label6)
		Me.Frame2.Controls.Add(Label5)
		Me.Frame2.Controls.Add(Label4)
		Me.Frame3.ResumeLayout(False)
		Me.Frame1.ResumeLayout(False)
		Me.Frame2.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class