<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class CustomerMod
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
	Public WithEvents Text_Phone As System.Windows.Forms.TextBox
	Public WithEvents Text_LastName As System.Windows.Forms.TextBox
	Public WithEvents Text_FirstName As System.Windows.Forms.TextBox
	Public WithEvents Comm_Exit As System.Windows.Forms.Button
	Public WithEvents Comm_View_Res As System.Windows.Forms.Button
	Public WithEvents Comm_View_Req As System.Windows.Forms.Button
	Public WithEvents Comm_Submit As System.Windows.Forms.Button
	Public WithEvents Text_CompanyName As System.Windows.Forms.TextBox
	Public WithEvents Text_Name As System.Windows.Forms.TextBox
	Public WithEvents Label10 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents L_ModifiedTime As System.Windows.Forms.Label
	Public WithEvents L_FullName As System.Windows.Forms.Label
	Public WithEvents L_ListID As System.Windows.Forms.Label
	Public WithEvents L_CreatedTime As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents Label9 As System.Windows.Forms.Label
	Public WithEvents Label7 As System.Windows.Forms.Label
	Public WithEvents Label8 As System.Windows.Forms.Label
	Public WithEvents Label6 As System.Windows.Forms.Label
	Public WithEvents Label5 As System.Windows.Forms.Label
	Public WithEvents Frame2 As System.Windows.Forms.GroupBox
	Public WithEvents Image_QBBANNER As System.Windows.Forms.PictureBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(CustomerMod))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Text_Phone = New System.Windows.Forms.TextBox()
        Me.Text_LastName = New System.Windows.Forms.TextBox()
        Me.Text_FirstName = New System.Windows.Forms.TextBox()
        Me.Comm_Exit = New System.Windows.Forms.Button()
        Me.Comm_View_Res = New System.Windows.Forms.Button()
        Me.Comm_View_Req = New System.Windows.Forms.Button()
        Me.Comm_Submit = New System.Windows.Forms.Button()
        Me.Text_CompanyName = New System.Windows.Forms.TextBox()
        Me.Text_Name = New System.Windows.Forms.TextBox()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.L_ModifiedTime = New System.Windows.Forms.Label()
        Me.L_FullName = New System.Windows.Forms.Label()
        Me.L_ListID = New System.Windows.Forms.Label()
        Me.L_CreatedTime = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Image_QBBANNER = New System.Windows.Forms.PictureBox()
        Me.Frame1.SuspendLayout()
        Me.Frame2.SuspendLayout()
        CType(Me.Image_QBBANNER, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Text_Phone
        '
        Me.Text_Phone.AcceptsReturn = True
        Me.Text_Phone.BackColor = System.Drawing.SystemColors.Window
        Me.Text_Phone.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text_Phone.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text_Phone.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text_Phone.Location = New System.Drawing.Point(128, 232)
        Me.Text_Phone.MaxLength = 0
        Me.Text_Phone.Name = "Text_Phone"
        Me.Text_Phone.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text_Phone.Size = New System.Drawing.Size(81, 23)
        Me.Text_Phone.TabIndex = 4
        '
        'Text_LastName
        '
        Me.Text_LastName.AcceptsReturn = True
        Me.Text_LastName.BackColor = System.Drawing.SystemColors.Window
        Me.Text_LastName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text_LastName.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text_LastName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text_LastName.Location = New System.Drawing.Point(128, 208)
        Me.Text_LastName.MaxLength = 0
        Me.Text_LastName.Name = "Text_LastName"
        Me.Text_LastName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text_LastName.Size = New System.Drawing.Size(161, 23)
        Me.Text_LastName.TabIndex = 3
        '
        'Text_FirstName
        '
        Me.Text_FirstName.AcceptsReturn = True
        Me.Text_FirstName.BackColor = System.Drawing.SystemColors.Window
        Me.Text_FirstName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text_FirstName.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text_FirstName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text_FirstName.Location = New System.Drawing.Point(128, 184)
        Me.Text_FirstName.MaxLength = 0
        Me.Text_FirstName.Name = "Text_FirstName"
        Me.Text_FirstName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text_FirstName.Size = New System.Drawing.Size(161, 23)
        Me.Text_FirstName.TabIndex = 2
        '
        'Comm_Exit
        '
        Me.Comm_Exit.BackColor = System.Drawing.SystemColors.Control
        Me.Comm_Exit.Cursor = System.Windows.Forms.Cursors.Default
        Me.Comm_Exit.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Comm_Exit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Comm_Exit.Location = New System.Drawing.Point(376, 184)
        Me.Comm_Exit.Name = "Comm_Exit"
        Me.Comm_Exit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Comm_Exit.Size = New System.Drawing.Size(129, 25)
        Me.Comm_Exit.TabIndex = 8
        Me.Comm_Exit.Text = "&Done"
        Me.Comm_Exit.UseVisualStyleBackColor = False
        '
        'Comm_View_Res
        '
        Me.Comm_View_Res.BackColor = System.Drawing.SystemColors.Control
        Me.Comm_View_Res.Cursor = System.Windows.Forms.Cursors.Default
        Me.Comm_View_Res.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Comm_View_Res.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Comm_View_Res.Location = New System.Drawing.Point(376, 264)
        Me.Comm_View_Res.Name = "Comm_View_Res"
        Me.Comm_View_Res.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Comm_View_Res.Size = New System.Drawing.Size(129, 25)
        Me.Comm_View_Res.TabIndex = 7
        Me.Comm_View_Res.Text = "View qbXML  Re&sponse"
        Me.Comm_View_Res.UseVisualStyleBackColor = False
        '
        'Comm_View_Req
        '
        Me.Comm_View_Req.BackColor = System.Drawing.SystemColors.Control
        Me.Comm_View_Req.Cursor = System.Windows.Forms.Cursors.Default
        Me.Comm_View_Req.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Comm_View_Req.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Comm_View_Req.Location = New System.Drawing.Point(376, 224)
        Me.Comm_View_Req.Name = "Comm_View_Req"
        Me.Comm_View_Req.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Comm_View_Req.Size = New System.Drawing.Size(129, 25)
        Me.Comm_View_Req.TabIndex = 6
        Me.Comm_View_Req.Text = "View qbXML Re&quest"
        Me.Comm_View_Req.UseVisualStyleBackColor = False
        '
        'Comm_Submit
        '
        Me.Comm_Submit.BackColor = System.Drawing.SystemColors.Control
        Me.Comm_Submit.Cursor = System.Windows.Forms.Cursors.Default
        Me.Comm_Submit.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Comm_Submit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Comm_Submit.Location = New System.Drawing.Point(376, 144)
        Me.Comm_Submit.Name = "Comm_Submit"
        Me.Comm_Submit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Comm_Submit.Size = New System.Drawing.Size(129, 25)
        Me.Comm_Submit.TabIndex = 5
        Me.Comm_Submit.Text = "&Modify Customer"
        Me.Comm_Submit.UseVisualStyleBackColor = False
        '
        'Text_CompanyName
        '
        Me.Text_CompanyName.AcceptsReturn = True
        Me.Text_CompanyName.BackColor = System.Drawing.SystemColors.Window
        Me.Text_CompanyName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text_CompanyName.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text_CompanyName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text_CompanyName.Location = New System.Drawing.Point(128, 256)
        Me.Text_CompanyName.MaxLength = 0
        Me.Text_CompanyName.Name = "Text_CompanyName"
        Me.Text_CompanyName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text_CompanyName.Size = New System.Drawing.Size(161, 23)
        Me.Text_CompanyName.TabIndex = 1
        '
        'Text_Name
        '
        Me.Text_Name.AcceptsReturn = True
        Me.Text_Name.BackColor = System.Drawing.SystemColors.Window
        Me.Text_Name.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text_Name.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text_Name.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text_Name.Location = New System.Drawing.Point(128, 160)
        Me.Text_Name.MaxLength = 0
        Me.Text_Name.Name = "Text_Name"
        Me.Text_Name.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text_Name.Size = New System.Drawing.Size(161, 23)
        Me.Text_Name.TabIndex = 0
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.Label10)
        Me.Frame1.Controls.Add(Me.Label2)
        Me.Frame1.Controls.Add(Me.L_ModifiedTime)
        Me.Frame1.Controls.Add(Me.L_FullName)
        Me.Frame1.Controls.Add(Me.L_ListID)
        Me.Frame1.Controls.Add(Me.L_CreatedTime)
        Me.Frame1.Controls.Add(Me.Label1)
        Me.Frame1.Controls.Add(Me.Label3)
        Me.Frame1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(24, 16)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(321, 97)
        Me.Frame1.TabIndex = 9
        Me.Frame1.TabStop = False
        Me.Frame1.Text = " &Customer Info: "
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.SystemColors.Control
        Me.Label10.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label10.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label10.Location = New System.Drawing.Point(8, 40)
        Me.Label10.Name = "Label10"
        Me.Label10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label10.Size = New System.Drawing.Size(57, 17)
        Me.Label10.TabIndex = 17
        Me.Label10.Text = "&Full Name:"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(8, 72)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(81, 20)
        Me.Label2.TabIndex = 16
        Me.Label2.Text = "M&odified Time:"
        '
        'L_ModifiedTime
        '
        Me.L_ModifiedTime.BackColor = System.Drawing.SystemColors.Control
        Me.L_ModifiedTime.Cursor = System.Windows.Forms.Cursors.Default
        Me.L_ModifiedTime.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.L_ModifiedTime.ForeColor = System.Drawing.SystemColors.ControlText
        Me.L_ModifiedTime.Location = New System.Drawing.Point(96, 72)
        Me.L_ModifiedTime.Name = "L_ModifiedTime"
        Me.L_ModifiedTime.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.L_ModifiedTime.Size = New System.Drawing.Size(145, 17)
        Me.L_ModifiedTime.TabIndex = 15
        Me.L_ModifiedTime.Text = "Mtime"
        '
        'L_FullName
        '
        Me.L_FullName.BackColor = System.Drawing.SystemColors.Control
        Me.L_FullName.Cursor = System.Windows.Forms.Cursors.Default
        Me.L_FullName.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.L_FullName.ForeColor = System.Drawing.SystemColors.ControlText
        Me.L_FullName.Location = New System.Drawing.Point(96, 40)
        Me.L_FullName.Name = "L_FullName"
        Me.L_FullName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.L_FullName.Size = New System.Drawing.Size(169, 17)
        Me.L_FullName.TabIndex = 14
        Me.L_FullName.Text = "Fullname"
        '
        'L_ListID
        '
        Me.L_ListID.BackColor = System.Drawing.SystemColors.Control
        Me.L_ListID.Cursor = System.Windows.Forms.Cursors.Default
        Me.L_ListID.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.L_ListID.ForeColor = System.Drawing.SystemColors.ControlText
        Me.L_ListID.Location = New System.Drawing.Point(96, 24)
        Me.L_ListID.Name = "L_ListID"
        Me.L_ListID.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.L_ListID.Size = New System.Drawing.Size(161, 17)
        Me.L_ListID.TabIndex = 13
        Me.L_ListID.Text = "ListID"
        '
        'L_CreatedTime
        '
        Me.L_CreatedTime.BackColor = System.Drawing.SystemColors.Control
        Me.L_CreatedTime.Cursor = System.Windows.Forms.Cursors.Default
        Me.L_CreatedTime.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.L_CreatedTime.ForeColor = System.Drawing.SystemColors.ControlText
        Me.L_CreatedTime.Location = New System.Drawing.Point(96, 56)
        Me.L_CreatedTime.Name = "L_CreatedTime"
        Me.L_CreatedTime.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.L_CreatedTime.Size = New System.Drawing.Size(145, 17)
        Me.L_CreatedTime.TabIndex = 12
        Me.L_CreatedTime.Text = "Ctime"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(8, 56)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(73, 17)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "C&reated Time:"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(8, 24)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(50, 17)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "&List ID:"
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.Label9)
        Me.Frame2.Controls.Add(Me.Label7)
        Me.Frame2.Controls.Add(Me.Label8)
        Me.Frame2.Controls.Add(Me.Label6)
        Me.Frame2.Controls.Add(Me.Label5)
        Me.Frame2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame2.Location = New System.Drawing.Point(24, 136)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(321, 153)
        Me.Frame2.TabIndex = 18
        Me.Frame2.TabStop = False
        Me.Frame2.Text = " &You can modify the following fields: "
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.SystemColors.Control
        Me.Label9.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label9.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label9.Location = New System.Drawing.Point(8, 96)
        Me.Label9.Name = "Label9"
        Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label9.Size = New System.Drawing.Size(49, 17)
        Me.Label9.TabIndex = 23
        Me.Label9.Text = "&Phone:"
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(8, 72)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(65, 17)
        Me.Label7.TabIndex = 22
        Me.Label7.Text = "&Last Name:"
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.SystemColors.Control
        Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label8.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label8.Location = New System.Drawing.Point(8, 120)
        Me.Label8.Name = "Label8"
        Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label8.Size = New System.Drawing.Size(89, 17)
        Me.Label8.TabIndex = 21
        Me.Label8.Text = "&Company Name:"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(8, 48)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(80, 17)
        Me.Label6.TabIndex = 20
        Me.Label6.Text = "&First Name:"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(8, 24)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(41, 17)
        Me.Label5.TabIndex = 19
        Me.Label5.Text = "&Name:"
        '
        'Image_QBBANNER
        '
        Me.Image_QBBANNER.Cursor = System.Windows.Forms.Cursors.Default
        Me.Image_QBBANNER.Image = CType(resources.GetObject("Image_QBBANNER.Image"), System.Drawing.Image)
        Me.Image_QBBANNER.Location = New System.Drawing.Point(24, 320)
        Me.Image_QBBANNER.Name = "Image_QBBANNER"
        Me.Image_QBBANNER.Size = New System.Drawing.Size(429, 28)
        Me.Image_QBBANNER.TabIndex = 19
        Me.Image_QBBANNER.TabStop = False
        '
        'CustomerMod
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(528, 391)
        Me.Controls.Add(Me.Text_Phone)
        Me.Controls.Add(Me.Text_LastName)
        Me.Controls.Add(Me.Text_FirstName)
        Me.Controls.Add(Me.Comm_Exit)
        Me.Controls.Add(Me.Comm_View_Res)
        Me.Controls.Add(Me.Comm_View_Req)
        Me.Controls.Add(Me.Comm_Submit)
        Me.Controls.Add(Me.Text_CompanyName)
        Me.Controls.Add(Me.Text_Name)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.Image_QBBANNER)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(3, 22)
        Me.Name = "CustomerMod"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "qbXML Sample: Modify Customer"
        Me.Frame1.ResumeLayout(False)
        Me.Frame2.ResumeLayout(False)
        CType(Me.Image_QBBANNER, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
End Class