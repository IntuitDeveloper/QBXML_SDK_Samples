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
	Public WithEvents Comm_Exit As System.Windows.Forms.Button
	Public WithEvents Comm_View_Res As System.Windows.Forms.Button
	Public WithEvents Comm_View_Req As System.Windows.Forms.Button
	Public WithEvents Text_CustomerName As System.Windows.Forms.TextBox
	Public WithEvents Text_Phone As System.Windows.Forms.TextBox
	Public WithEvents Comm_Submit As System.Windows.Forms.Button
	Public WithEvents Text_LastName As System.Windows.Forms.TextBox
	Public WithEvents Text_FirstName As System.Windows.Forms.TextBox
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents Image_QBBANNER As System.Windows.Forms.PictureBox
	Public WithEvents Image_QBCUST As System.Windows.Forms.PictureBox
	Public WithEvents Label7 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(AddCust))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Comm_Exit = New System.Windows.Forms.Button
		Me.Comm_View_Res = New System.Windows.Forms.Button
		Me.Comm_View_Req = New System.Windows.Forms.Button
		Me.Text_CustomerName = New System.Windows.Forms.TextBox
		Me.Text_Phone = New System.Windows.Forms.TextBox
		Me.Comm_Submit = New System.Windows.Forms.Button
		Me.Text_LastName = New System.Windows.Forms.TextBox
		Me.Text_FirstName = New System.Windows.Forms.TextBox
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.Image_QBBANNER = New System.Windows.Forms.PictureBox
		Me.Image_QBCUST = New System.Windows.Forms.PictureBox
		Me.Label7 = New System.Windows.Forms.Label
		Me.Frame1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "qbXML Sample: Create a New Customer"
		Me.ClientSize = New System.Drawing.Size(503, 341)
		Me.Location = New System.Drawing.Point(3, 22)
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
		Me.Name = "AddCust"
		Me.Comm_Exit.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Comm_Exit.Text = "&Done"
		Me.Comm_Exit.Size = New System.Drawing.Size(145, 25)
		Me.Comm_Exit.Location = New System.Drawing.Point(336, 80)
		Me.Comm_Exit.TabIndex = 8
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
		Me.Comm_View_Res.Location = New System.Drawing.Point(336, 184)
		Me.Comm_View_Res.TabIndex = 7
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
		Me.Comm_View_Req.Location = New System.Drawing.Point(336, 152)
		Me.Comm_View_Req.TabIndex = 6
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
		Me.Text_CustomerName.Location = New System.Drawing.Point(32, 72)
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
		Me.Text_Phone.Size = New System.Drawing.Size(89, 20)
		Me.Text_Phone.Location = New System.Drawing.Point(112, 200)
		Me.Text_Phone.TabIndex = 4
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
		Me.Comm_Submit.Location = New System.Drawing.Point(336, 48)
		Me.Comm_Submit.TabIndex = 5
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
		Me.Text_LastName.Size = New System.Drawing.Size(129, 20)
		Me.Text_LastName.Location = New System.Drawing.Point(112, 168)
		Me.Text_LastName.TabIndex = 3
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
		Me.Text_FirstName.Size = New System.Drawing.Size(129, 20)
		Me.Text_FirstName.Location = New System.Drawing.Point(112, 136)
		Me.Text_FirstName.TabIndex = 2
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
		Me.Frame1.Size = New System.Drawing.Size(265, 145)
		Me.Frame1.Location = New System.Drawing.Point(32, 112)
		Me.Frame1.TabIndex = 9
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.BackColor = System.Drawing.SystemColors.Control
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame1.Name = "Frame1"
		Me.Label3.Text = "&Phone:"
		Me.Label3.Size = New System.Drawing.Size(57, 17)
		Me.Label3.Location = New System.Drawing.Point(8, 88)
		Me.Label3.TabIndex = 12
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
		Me.Label2.Text = "&Last Name:"
		Me.Label2.Size = New System.Drawing.Size(65, 17)
		Me.Label2.Location = New System.Drawing.Point(8, 56)
		Me.Label2.TabIndex = 11
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
		Me.Label1.Text = "&First Name:"
		Me.Label1.Size = New System.Drawing.Size(65, 17)
		Me.Label1.Location = New System.Drawing.Point(8, 24)
		Me.Label1.TabIndex = 10
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
		Me.Image_QBBANNER.Size = New System.Drawing.Size(430, 29)
		Me.Image_QBBANNER.Location = New System.Drawing.Point(40, 280)
		Me.Image_QBBANNER.Image = CType(resources.GetObject("Image_QBBANNER.Image"), System.Drawing.Image)
		Me.Image_QBBANNER.Enabled = True
		Me.Image_QBBANNER.Cursor = System.Windows.Forms.Cursors.Default
		Me.Image_QBBANNER.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.Image_QBBANNER.Visible = True
		Me.Image_QBBANNER.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Image_QBBANNER.Name = "Image_QBBANNER"
		Me.Image_QBCUST.Size = New System.Drawing.Size(30, 33)
		Me.Image_QBCUST.Location = New System.Drawing.Point(224, 64)
		Me.Image_QBCUST.Image = CType(resources.GetObject("Image_QBCUST.Image"), System.Drawing.Image)
		Me.Image_QBCUST.Enabled = True
		Me.Image_QBCUST.Cursor = System.Windows.Forms.Cursors.Default
		Me.Image_QBCUST.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.Image_QBCUST.Visible = True
		Me.Image_QBCUST.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Image_QBCUST.Name = "Image_QBCUST"
		Me.Label7.Text = "&Customer to add:"
		Me.Label7.Size = New System.Drawing.Size(89, 17)
		Me.Label7.Location = New System.Drawing.Point(32, 48)
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
		Me.Controls.Add(Comm_Exit)
		Me.Controls.Add(Comm_View_Res)
		Me.Controls.Add(Comm_View_Req)
		Me.Controls.Add(Text_CustomerName)
		Me.Controls.Add(Text_Phone)
		Me.Controls.Add(Comm_Submit)
		Me.Controls.Add(Text_LastName)
		Me.Controls.Add(Text_FirstName)
		Me.Controls.Add(Frame1)
		Me.Controls.Add(Image_QBBANNER)
		Me.Controls.Add(Image_QBCUST)
		Me.Controls.Add(Label7)
		Me.Frame1.Controls.Add(Label3)
		Me.Frame1.Controls.Add(Label2)
		Me.Frame1.Controls.Add(Label1)
		Me.Frame1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class