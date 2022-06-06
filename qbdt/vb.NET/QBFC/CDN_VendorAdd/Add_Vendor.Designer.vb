<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class AddVendor
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
	Public WithEvents Command_Clear As System.Windows.Forms.Button
	Public WithEvents Text_Note_MultiCurrency As System.Windows.Forms.TextBox
	Public WithEvents Combo_Currency As System.Windows.Forms.ComboBox
	Public WithEvents Combo_Vendor_Type As System.Windows.Forms.ComboBox
	Public WithEvents Combo_State_Prov As System.Windows.Forms.ComboBox
	Public WithEvents Text_Postal_Code As System.Windows.Forms.TextBox
	Public WithEvents Text_EMail As System.Windows.Forms.TextBox
	Public WithEvents Text_Alt_Phone As System.Windows.Forms.TextBox
	Public WithEvents Text_Phone As System.Windows.Forms.TextBox
	Public WithEvents Text_Name_On_Check As System.Windows.Forms.TextBox
	Public WithEvents Text_First_Name As System.Windows.Forms.TextBox
	Public WithEvents Text_Last_Name As System.Windows.Forms.TextBox
	Public WithEvents Text_Salutation As System.Windows.Forms.TextBox
	Public WithEvents Text_Address_1 As System.Windows.Forms.TextBox
	Public WithEvents Text_Address_2 As System.Windows.Forms.TextBox
	Public WithEvents Text_City As System.Windows.Forms.TextBox
	Public WithEvents Text_Account_Number As System.Windows.Forms.TextBox
	Public WithEvents Command_Add_Vendor As System.Windows.Forms.Button
	Public WithEvents Command_Exit As System.Windows.Forms.Button
	Public WithEvents Text_Middle_Name As System.Windows.Forms.TextBox
	Public WithEvents Label_Note As System.Windows.Forms.Label
	Public WithEvents Label_Multicurrency As System.Windows.Forms.Label
	Public WithEvents Label9 As System.Windows.Forms.Label
	Public WithEvents Image_QBBANNER As System.Windows.Forms.PictureBox
	Public WithEvents Label7 As System.Windows.Forms.Label
	Public WithEvents Label8 As System.Windows.Forms.Label
	Public WithEvents Label16 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label5 As System.Windows.Forms.Label
	Public WithEvents Label6 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label11 As System.Windows.Forms.Label
	Public WithEvents Label12 As System.Windows.Forms.Label
	Public WithEvents Label13 As System.Windows.Forms.Label
	Public WithEvents Label10 As System.Windows.Forms.Label
	Public WithEvents Label17 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(AddVendor))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Command_Clear = New System.Windows.Forms.Button
		Me.Text_Note_MultiCurrency = New System.Windows.Forms.TextBox
		Me.Combo_Currency = New System.Windows.Forms.ComboBox
		Me.Combo_Vendor_Type = New System.Windows.Forms.ComboBox
		Me.Combo_State_Prov = New System.Windows.Forms.ComboBox
		Me.Text_Postal_Code = New System.Windows.Forms.TextBox
		Me.Text_EMail = New System.Windows.Forms.TextBox
		Me.Text_Alt_Phone = New System.Windows.Forms.TextBox
		Me.Text_Phone = New System.Windows.Forms.TextBox
		Me.Text_Name_On_Check = New System.Windows.Forms.TextBox
		Me.Text_First_Name = New System.Windows.Forms.TextBox
		Me.Text_Last_Name = New System.Windows.Forms.TextBox
		Me.Text_Salutation = New System.Windows.Forms.TextBox
		Me.Text_Address_1 = New System.Windows.Forms.TextBox
		Me.Text_Address_2 = New System.Windows.Forms.TextBox
		Me.Text_City = New System.Windows.Forms.TextBox
		Me.Text_Account_Number = New System.Windows.Forms.TextBox
		Me.Command_Add_Vendor = New System.Windows.Forms.Button
		Me.Command_Exit = New System.Windows.Forms.Button
		Me.Text_Middle_Name = New System.Windows.Forms.TextBox
		Me.Label_Note = New System.Windows.Forms.Label
		Me.Label_Multicurrency = New System.Windows.Forms.Label
		Me.Label9 = New System.Windows.Forms.Label
		Me.Image_QBBANNER = New System.Windows.Forms.PictureBox
		Me.Label7 = New System.Windows.Forms.Label
		Me.Label8 = New System.Windows.Forms.Label
		Me.Label16 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label4 = New System.Windows.Forms.Label
		Me.Label5 = New System.Windows.Forms.Label
		Me.Label6 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label11 = New System.Windows.Forms.Label
		Me.Label12 = New System.Windows.Forms.Label
		Me.Label13 = New System.Windows.Forms.Label
		Me.Label10 = New System.Windows.Forms.Label
		Me.Label17 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "Add a vendor"
		Me.ClientSize = New System.Drawing.Size(704, 445)
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
		Me.Name = "AddVendor"
		Me.Command_Clear.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command_Clear.Text = "Clear"
		Me.Command_Clear.Size = New System.Drawing.Size(113, 33)
		Me.Command_Clear.Location = New System.Drawing.Point(536, 104)
		Me.Command_Clear.TabIndex = 36
		Me.Command_Clear.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Command_Clear.BackColor = System.Drawing.SystemColors.Control
		Me.Command_Clear.CausesValidation = True
		Me.Command_Clear.Enabled = True
		Me.Command_Clear.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command_Clear.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command_Clear.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command_Clear.TabStop = True
		Me.Command_Clear.Name = "Command_Clear"
		Me.Text_Note_MultiCurrency.AutoSize = False
		Me.Text_Note_MultiCurrency.Enabled = False
		Me.Text_Note_MultiCurrency.Size = New System.Drawing.Size(257, 25)
		Me.Text_Note_MultiCurrency.Location = New System.Drawing.Point(360, 296)
		Me.Text_Note_MultiCurrency.TabIndex = 34
		Me.Text_Note_MultiCurrency.Text = "This data file doesn't have multicurrency turned on"
		Me.Text_Note_MultiCurrency.Visible = False
		Me.Text_Note_MultiCurrency.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text_Note_MultiCurrency.AcceptsReturn = True
		Me.Text_Note_MultiCurrency.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text_Note_MultiCurrency.BackColor = System.Drawing.SystemColors.Window
		Me.Text_Note_MultiCurrency.CausesValidation = True
		Me.Text_Note_MultiCurrency.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text_Note_MultiCurrency.HideSelection = True
		Me.Text_Note_MultiCurrency.ReadOnly = False
		Me.Text_Note_MultiCurrency.Maxlength = 0
		Me.Text_Note_MultiCurrency.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text_Note_MultiCurrency.MultiLine = False
		Me.Text_Note_MultiCurrency.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text_Note_MultiCurrency.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text_Note_MultiCurrency.TabStop = True
		Me.Text_Note_MultiCurrency.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text_Note_MultiCurrency.Name = "Text_Note_MultiCurrency"
		Me.Combo_Currency.Size = New System.Drawing.Size(177, 21)
		Me.Combo_Currency.Location = New System.Drawing.Point(360, 256)
		Me.Combo_Currency.TabIndex = 15
		Me.Combo_Currency.Visible = False
		Me.Combo_Currency.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Combo_Currency.BackColor = System.Drawing.SystemColors.Window
		Me.Combo_Currency.CausesValidation = True
		Me.Combo_Currency.Enabled = True
		Me.Combo_Currency.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Combo_Currency.IntegralHeight = True
		Me.Combo_Currency.Cursor = System.Windows.Forms.Cursors.Default
		Me.Combo_Currency.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Combo_Currency.Sorted = False
		Me.Combo_Currency.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.Combo_Currency.TabStop = True
		Me.Combo_Currency.Name = "Combo_Currency"
		Me.Combo_Vendor_Type.Size = New System.Drawing.Size(137, 21)
		Me.Combo_Vendor_Type.Location = New System.Drawing.Point(360, 216)
		Me.Combo_Vendor_Type.TabIndex = 14
		Me.Combo_Vendor_Type.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Combo_Vendor_Type.BackColor = System.Drawing.SystemColors.Window
		Me.Combo_Vendor_Type.CausesValidation = True
		Me.Combo_Vendor_Type.Enabled = True
		Me.Combo_Vendor_Type.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Combo_Vendor_Type.IntegralHeight = True
		Me.Combo_Vendor_Type.Cursor = System.Windows.Forms.Cursors.Default
		Me.Combo_Vendor_Type.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Combo_Vendor_Type.Sorted = False
		Me.Combo_Vendor_Type.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.Combo_Vendor_Type.TabStop = True
		Me.Combo_Vendor_Type.Visible = True
		Me.Combo_Vendor_Type.Name = "Combo_Vendor_Type"
		Me.Combo_State_Prov.Size = New System.Drawing.Size(57, 21)
		Me.Combo_State_Prov.Location = New System.Drawing.Point(104, 296)
		Me.Combo_State_Prov.TabIndex = 7
		Me.Combo_State_Prov.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Combo_State_Prov.BackColor = System.Drawing.SystemColors.Window
		Me.Combo_State_Prov.CausesValidation = True
		Me.Combo_State_Prov.Enabled = True
		Me.Combo_State_Prov.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Combo_State_Prov.IntegralHeight = True
		Me.Combo_State_Prov.Cursor = System.Windows.Forms.Cursors.Default
		Me.Combo_State_Prov.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Combo_State_Prov.Sorted = False
		Me.Combo_State_Prov.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.Combo_State_Prov.TabStop = True
		Me.Combo_State_Prov.Visible = True
		Me.Combo_State_Prov.Name = "Combo_State_Prov"
		Me.Text_Postal_Code.AutoSize = False
		Me.Text_Postal_Code.Size = New System.Drawing.Size(65, 19)
		Me.Text_Postal_Code.Location = New System.Drawing.Point(176, 296)
		Me.Text_Postal_Code.TabIndex = 8
		Me.Text_Postal_Code.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text_Postal_Code.AcceptsReturn = True
		Me.Text_Postal_Code.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text_Postal_Code.BackColor = System.Drawing.SystemColors.Window
		Me.Text_Postal_Code.CausesValidation = True
		Me.Text_Postal_Code.Enabled = True
		Me.Text_Postal_Code.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text_Postal_Code.HideSelection = True
		Me.Text_Postal_Code.ReadOnly = False
		Me.Text_Postal_Code.Maxlength = 0
		Me.Text_Postal_Code.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text_Postal_Code.MultiLine = False
		Me.Text_Postal_Code.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text_Postal_Code.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text_Postal_Code.TabStop = True
		Me.Text_Postal_Code.Visible = True
		Me.Text_Postal_Code.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text_Postal_Code.Name = "Text_Postal_Code"
		Me.Text_EMail.AutoSize = False
		Me.Text_EMail.Size = New System.Drawing.Size(137, 19)
		Me.Text_EMail.Location = New System.Drawing.Point(360, 114)
		Me.Text_EMail.TabIndex = 11
		Me.Text_EMail.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text_EMail.AcceptsReturn = True
		Me.Text_EMail.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text_EMail.BackColor = System.Drawing.SystemColors.Window
		Me.Text_EMail.CausesValidation = True
		Me.Text_EMail.Enabled = True
		Me.Text_EMail.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text_EMail.HideSelection = True
		Me.Text_EMail.ReadOnly = False
		Me.Text_EMail.Maxlength = 0
		Me.Text_EMail.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text_EMail.MultiLine = False
		Me.Text_EMail.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text_EMail.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text_EMail.TabStop = True
		Me.Text_EMail.Visible = True
		Me.Text_EMail.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text_EMail.Name = "Text_EMail"
		Me.Text_Alt_Phone.AutoSize = False
		Me.Text_Alt_Phone.Size = New System.Drawing.Size(137, 19)
		Me.Text_Alt_Phone.Location = New System.Drawing.Point(360, 84)
		Me.Text_Alt_Phone.TabIndex = 10
		Me.Text_Alt_Phone.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text_Alt_Phone.AcceptsReturn = True
		Me.Text_Alt_Phone.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text_Alt_Phone.BackColor = System.Drawing.SystemColors.Window
		Me.Text_Alt_Phone.CausesValidation = True
		Me.Text_Alt_Phone.Enabled = True
		Me.Text_Alt_Phone.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text_Alt_Phone.HideSelection = True
		Me.Text_Alt_Phone.ReadOnly = False
		Me.Text_Alt_Phone.Maxlength = 0
		Me.Text_Alt_Phone.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text_Alt_Phone.MultiLine = False
		Me.Text_Alt_Phone.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text_Alt_Phone.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text_Alt_Phone.TabStop = True
		Me.Text_Alt_Phone.Visible = True
		Me.Text_Alt_Phone.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text_Alt_Phone.Name = "Text_Alt_Phone"
		Me.Text_Phone.AutoSize = False
		Me.Text_Phone.Size = New System.Drawing.Size(137, 19)
		Me.Text_Phone.Location = New System.Drawing.Point(360, 54)
		Me.Text_Phone.TabIndex = 9
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
		Me.Text_Name_On_Check.AutoSize = False
		Me.Text_Name_On_Check.Size = New System.Drawing.Size(137, 19)
		Me.Text_Name_On_Check.Location = New System.Drawing.Point(360, 144)
		Me.Text_Name_On_Check.TabIndex = 12
		Me.Text_Name_On_Check.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text_Name_On_Check.AcceptsReturn = True
		Me.Text_Name_On_Check.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text_Name_On_Check.BackColor = System.Drawing.SystemColors.Window
		Me.Text_Name_On_Check.CausesValidation = True
		Me.Text_Name_On_Check.Enabled = True
		Me.Text_Name_On_Check.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text_Name_On_Check.HideSelection = True
		Me.Text_Name_On_Check.ReadOnly = False
		Me.Text_Name_On_Check.Maxlength = 0
		Me.Text_Name_On_Check.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text_Name_On_Check.MultiLine = False
		Me.Text_Name_On_Check.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text_Name_On_Check.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text_Name_On_Check.TabStop = True
		Me.Text_Name_On_Check.Visible = True
		Me.Text_Name_On_Check.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text_Name_On_Check.Name = "Text_Name_On_Check"
		Me.Text_First_Name.AutoSize = False
		Me.Text_First_Name.Size = New System.Drawing.Size(81, 20)
		Me.Text_First_Name.Location = New System.Drawing.Point(104, 91)
		Me.Text_First_Name.TabIndex = 1
		Me.Text_First_Name.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text_First_Name.AcceptsReturn = True
		Me.Text_First_Name.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text_First_Name.BackColor = System.Drawing.SystemColors.Window
		Me.Text_First_Name.CausesValidation = True
		Me.Text_First_Name.Enabled = True
		Me.Text_First_Name.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text_First_Name.HideSelection = True
		Me.Text_First_Name.ReadOnly = False
		Me.Text_First_Name.Maxlength = 0
		Me.Text_First_Name.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text_First_Name.MultiLine = False
		Me.Text_First_Name.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text_First_Name.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text_First_Name.TabStop = True
		Me.Text_First_Name.Visible = True
		Me.Text_First_Name.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text_First_Name.Name = "Text_First_Name"
		Me.Text_Last_Name.AutoSize = False
		Me.Text_Last_Name.Size = New System.Drawing.Size(137, 20)
		Me.Text_Last_Name.Location = New System.Drawing.Point(104, 127)
		Me.Text_Last_Name.TabIndex = 3
		Me.Text_Last_Name.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text_Last_Name.AcceptsReturn = True
		Me.Text_Last_Name.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text_Last_Name.BackColor = System.Drawing.SystemColors.Window
		Me.Text_Last_Name.CausesValidation = True
		Me.Text_Last_Name.Enabled = True
		Me.Text_Last_Name.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text_Last_Name.HideSelection = True
		Me.Text_Last_Name.ReadOnly = False
		Me.Text_Last_Name.Maxlength = 0
		Me.Text_Last_Name.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text_Last_Name.MultiLine = False
		Me.Text_Last_Name.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text_Last_Name.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text_Last_Name.TabStop = True
		Me.Text_Last_Name.Visible = True
		Me.Text_Last_Name.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text_Last_Name.Name = "Text_Last_Name"
		Me.Text_Salutation.AutoSize = False
		Me.Text_Salutation.Size = New System.Drawing.Size(73, 20)
		Me.Text_Salutation.Location = New System.Drawing.Point(104, 56)
		Me.Text_Salutation.TabIndex = 0
		Me.Text_Salutation.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text_Salutation.AcceptsReturn = True
		Me.Text_Salutation.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text_Salutation.BackColor = System.Drawing.SystemColors.Window
		Me.Text_Salutation.CausesValidation = True
		Me.Text_Salutation.Enabled = True
		Me.Text_Salutation.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text_Salutation.HideSelection = True
		Me.Text_Salutation.ReadOnly = False
		Me.Text_Salutation.Maxlength = 0
		Me.Text_Salutation.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text_Salutation.MultiLine = False
		Me.Text_Salutation.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text_Salutation.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text_Salutation.TabStop = True
		Me.Text_Salutation.Visible = True
		Me.Text_Salutation.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text_Salutation.Name = "Text_Salutation"
		Me.Text_Address_1.AutoSize = False
		Me.Text_Address_1.Size = New System.Drawing.Size(137, 19)
		Me.Text_Address_1.Location = New System.Drawing.Point(104, 178)
		Me.Text_Address_1.TabIndex = 4
		Me.Text_Address_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text_Address_1.AcceptsReturn = True
		Me.Text_Address_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text_Address_1.BackColor = System.Drawing.SystemColors.Window
		Me.Text_Address_1.CausesValidation = True
		Me.Text_Address_1.Enabled = True
		Me.Text_Address_1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text_Address_1.HideSelection = True
		Me.Text_Address_1.ReadOnly = False
		Me.Text_Address_1.Maxlength = 0
		Me.Text_Address_1.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text_Address_1.MultiLine = False
		Me.Text_Address_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text_Address_1.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text_Address_1.TabStop = True
		Me.Text_Address_1.Visible = True
		Me.Text_Address_1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text_Address_1.Name = "Text_Address_1"
		Me.Text_Address_2.AutoSize = False
		Me.Text_Address_2.Size = New System.Drawing.Size(137, 19)
		Me.Text_Address_2.Location = New System.Drawing.Point(104, 213)
		Me.Text_Address_2.TabIndex = 5
		Me.Text_Address_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text_Address_2.AcceptsReturn = True
		Me.Text_Address_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text_Address_2.BackColor = System.Drawing.SystemColors.Window
		Me.Text_Address_2.CausesValidation = True
		Me.Text_Address_2.Enabled = True
		Me.Text_Address_2.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text_Address_2.HideSelection = True
		Me.Text_Address_2.ReadOnly = False
		Me.Text_Address_2.Maxlength = 0
		Me.Text_Address_2.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text_Address_2.MultiLine = False
		Me.Text_Address_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text_Address_2.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text_Address_2.TabStop = True
		Me.Text_Address_2.Visible = True
		Me.Text_Address_2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text_Address_2.Name = "Text_Address_2"
		Me.Text_City.AutoSize = False
		Me.Text_City.Size = New System.Drawing.Size(137, 19)
		Me.Text_City.Location = New System.Drawing.Point(104, 248)
		Me.Text_City.TabIndex = 6
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
		Me.Text_Account_Number.AutoSize = False
		Me.Text_Account_Number.Size = New System.Drawing.Size(137, 19)
		Me.Text_Account_Number.Location = New System.Drawing.Point(360, 176)
		Me.Text_Account_Number.TabIndex = 13
		Me.Text_Account_Number.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text_Account_Number.AcceptsReturn = True
		Me.Text_Account_Number.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text_Account_Number.BackColor = System.Drawing.SystemColors.Window
		Me.Text_Account_Number.CausesValidation = True
		Me.Text_Account_Number.Enabled = True
		Me.Text_Account_Number.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text_Account_Number.HideSelection = True
		Me.Text_Account_Number.ReadOnly = False
		Me.Text_Account_Number.Maxlength = 0
		Me.Text_Account_Number.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text_Account_Number.MultiLine = False
		Me.Text_Account_Number.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text_Account_Number.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text_Account_Number.TabStop = True
		Me.Text_Account_Number.Visible = True
		Me.Text_Account_Number.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text_Account_Number.Name = "Text_Account_Number"
		Me.Command_Add_Vendor.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command_Add_Vendor.Text = "Add Vendor"
		Me.Command_Add_Vendor.Size = New System.Drawing.Size(113, 33)
		Me.Command_Add_Vendor.Location = New System.Drawing.Point(536, 56)
		Me.Command_Add_Vendor.TabIndex = 17
		Me.Command_Add_Vendor.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Command_Add_Vendor.BackColor = System.Drawing.SystemColors.Control
		Me.Command_Add_Vendor.CausesValidation = True
		Me.Command_Add_Vendor.Enabled = True
		Me.Command_Add_Vendor.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command_Add_Vendor.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command_Add_Vendor.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command_Add_Vendor.TabStop = True
		Me.Command_Add_Vendor.Name = "Command_Add_Vendor"
		Me.Command_Exit.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command_Exit.Text = "Exit"
		Me.Command_Exit.Size = New System.Drawing.Size(113, 33)
		Me.Command_Exit.Location = New System.Drawing.Point(536, 152)
		Me.Command_Exit.TabIndex = 16
		Me.Command_Exit.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Command_Exit.BackColor = System.Drawing.SystemColors.Control
		Me.Command_Exit.CausesValidation = True
		Me.Command_Exit.Enabled = True
		Me.Command_Exit.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command_Exit.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command_Exit.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command_Exit.TabStop = True
		Me.Command_Exit.Name = "Command_Exit"
		Me.Text_Middle_Name.AutoSize = False
		Me.Text_Middle_Name.Size = New System.Drawing.Size(25, 19)
		Me.Text_Middle_Name.Location = New System.Drawing.Point(216, 92)
		Me.Text_Middle_Name.TabIndex = 2
		Me.Text_Middle_Name.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text_Middle_Name.AcceptsReturn = True
		Me.Text_Middle_Name.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text_Middle_Name.BackColor = System.Drawing.SystemColors.Window
		Me.Text_Middle_Name.CausesValidation = True
		Me.Text_Middle_Name.Enabled = True
		Me.Text_Middle_Name.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text_Middle_Name.HideSelection = True
		Me.Text_Middle_Name.ReadOnly = False
		Me.Text_Middle_Name.Maxlength = 0
		Me.Text_Middle_Name.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text_Middle_Name.MultiLine = False
		Me.Text_Middle_Name.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text_Middle_Name.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text_Middle_Name.TabStop = True
		Me.Text_Middle_Name.Visible = True
		Me.Text_Middle_Name.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text_Middle_Name.Name = "Text_Middle_Name"
		Me.Label_Note.Text = "Note on data file"
		Me.Label_Note.Size = New System.Drawing.Size(73, 25)
		Me.Label_Note.Location = New System.Drawing.Point(288, 296)
		Me.Label_Note.TabIndex = 35
		Me.Label_Note.Visible = False
		Me.Label_Note.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label_Note.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label_Note.BackColor = System.Drawing.SystemColors.Control
		Me.Label_Note.Enabled = True
		Me.Label_Note.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label_Note.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label_Note.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label_Note.UseMnemonic = True
		Me.Label_Note.AutoSize = False
		Me.Label_Note.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label_Note.Name = "Label_Note"
		Me.Label_Multicurrency.Text = "Currency"
		Me.Label_Multicurrency.Size = New System.Drawing.Size(57, 25)
		Me.Label_Multicurrency.Location = New System.Drawing.Point(288, 256)
		Me.Label_Multicurrency.TabIndex = 33
		Me.Label_Multicurrency.Visible = False
		Me.Label_Multicurrency.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label_Multicurrency.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label_Multicurrency.BackColor = System.Drawing.SystemColors.Control
		Me.Label_Multicurrency.Enabled = True
		Me.Label_Multicurrency.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label_Multicurrency.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label_Multicurrency.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label_Multicurrency.UseMnemonic = True
		Me.Label_Multicurrency.AutoSize = False
		Me.Label_Multicurrency.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label_Multicurrency.Name = "Label_Multicurrency"
		Me.Label9.Text = "Vendor Type"
		Me.Label9.Size = New System.Drawing.Size(57, 25)
		Me.Label9.Location = New System.Drawing.Point(288, 216)
		Me.Label9.TabIndex = 32
		Me.Label9.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label9.TextAlign = System.Drawing.ContentAlignment.TopLeft
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
		Me.Image_QBBANNER.Size = New System.Drawing.Size(201, 41)
		Me.Image_QBBANNER.Location = New System.Drawing.Point(96, 376)
		Me.Image_QBBANNER.Enabled = True
		Me.Image_QBBANNER.Cursor = System.Windows.Forms.Cursors.Default
		Me.Image_QBBANNER.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.Image_QBBANNER.Visible = True
		Me.Image_QBBANNER.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Image_QBBANNER.Name = "Image_QBBANNER"
		Me.Label7.Text = "Postal Code"
		Me.Label7.Size = New System.Drawing.Size(97, 17)
		Me.Label7.Location = New System.Drawing.Point(176, 280)
		Me.Label7.TabIndex = 31
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
		Me.Label8.Text = "State/Province"
		Me.Label8.Size = New System.Drawing.Size(73, 25)
		Me.Label8.Location = New System.Drawing.Point(96, 280)
		Me.Label8.TabIndex = 30
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
		Me.Label16.Text = "Add Vendor"
		Me.Label16.Font = New System.Drawing.Font("Verdana", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label16.Size = New System.Drawing.Size(309, 25)
		Me.Label16.Location = New System.Drawing.Point(40, 16)
		Me.Label16.TabIndex = 29
		Me.Label16.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label16.BackColor = System.Drawing.SystemColors.Control
		Me.Label16.Enabled = True
		Me.Label16.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label16.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label16.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label16.UseMnemonic = True
		Me.Label16.Visible = True
		Me.Label16.AutoSize = False
		Me.Label16.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label16.Name = "Label16"
		Me.Label1.Text = "First Name"
		Me.Label1.Size = New System.Drawing.Size(65, 17)
		Me.Label1.Location = New System.Drawing.Point(40, 92)
		Me.Label1.TabIndex = 28
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
		Me.Label2.Text = "Last Name"
		Me.Label2.Size = New System.Drawing.Size(65, 17)
		Me.Label2.Location = New System.Drawing.Point(40, 128)
		Me.Label2.TabIndex = 27
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
		Me.Label4.Text = "Salutation"
		Me.Label4.Size = New System.Drawing.Size(57, 17)
		Me.Label4.Location = New System.Drawing.Point(40, 56)
		Me.Label4.TabIndex = 26
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
		Me.Label5.Text = "Address"
		Me.Label5.Size = New System.Drawing.Size(57, 25)
		Me.Label5.Location = New System.Drawing.Point(40, 184)
		Me.Label5.TabIndex = 25
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
		Me.Label6.Text = "City"
		Me.Label6.Size = New System.Drawing.Size(57, 25)
		Me.Label6.Location = New System.Drawing.Point(40, 248)
		Me.Label6.TabIndex = 24
		Me.Label6.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopLeft
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
		Me.Label3.Text = "Phone"
		Me.Label3.Size = New System.Drawing.Size(57, 25)
		Me.Label3.Location = New System.Drawing.Point(288, 56)
		Me.Label3.TabIndex = 23
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
		Me.Label11.Text = "E-mail"
		Me.Label11.Size = New System.Drawing.Size(57, 25)
		Me.Label11.Location = New System.Drawing.Point(288, 120)
		Me.Label11.TabIndex = 22
		Me.Label11.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label11.TextAlign = System.Drawing.ContentAlignment.TopLeft
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
		Me.Label12.Text = "Alt. Ph."
		Me.Label12.Size = New System.Drawing.Size(57, 25)
		Me.Label12.Location = New System.Drawing.Point(288, 88)
		Me.Label12.TabIndex = 21
		Me.Label12.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label12.TextAlign = System.Drawing.ContentAlignment.TopLeft
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
		Me.Label13.Text = "Name on check"
		Me.Label13.Size = New System.Drawing.Size(57, 25)
		Me.Label13.Location = New System.Drawing.Point(288, 144)
		Me.Label13.TabIndex = 20
		Me.Label13.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label13.TextAlign = System.Drawing.ContentAlignment.TopLeft
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
		Me.Label10.Text = "Account Number"
		Me.Label10.Size = New System.Drawing.Size(57, 25)
		Me.Label10.Location = New System.Drawing.Point(288, 176)
		Me.Label10.TabIndex = 19
		Me.Label10.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label10.TextAlign = System.Drawing.ContentAlignment.TopLeft
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
		Me.Label17.Text = "M.I."
		Me.Label17.Size = New System.Drawing.Size(33, 17)
		Me.Label17.Location = New System.Drawing.Point(192, 96)
		Me.Label17.TabIndex = 18
		Me.Label17.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label17.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label17.BackColor = System.Drawing.SystemColors.Control
		Me.Label17.Enabled = True
		Me.Label17.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label17.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label17.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label17.UseMnemonic = True
		Me.Label17.Visible = True
		Me.Label17.AutoSize = False
		Me.Label17.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label17.Name = "Label17"
		Me.Controls.Add(Command_Clear)
		Me.Controls.Add(Text_Note_MultiCurrency)
		Me.Controls.Add(Combo_Currency)
		Me.Controls.Add(Combo_Vendor_Type)
		Me.Controls.Add(Combo_State_Prov)
		Me.Controls.Add(Text_Postal_Code)
		Me.Controls.Add(Text_EMail)
		Me.Controls.Add(Text_Alt_Phone)
		Me.Controls.Add(Text_Phone)
		Me.Controls.Add(Text_Name_On_Check)
		Me.Controls.Add(Text_First_Name)
		Me.Controls.Add(Text_Last_Name)
		Me.Controls.Add(Text_Salutation)
		Me.Controls.Add(Text_Address_1)
		Me.Controls.Add(Text_Address_2)
		Me.Controls.Add(Text_City)
		Me.Controls.Add(Text_Account_Number)
		Me.Controls.Add(Command_Add_Vendor)
		Me.Controls.Add(Command_Exit)
		Me.Controls.Add(Text_Middle_Name)
		Me.Controls.Add(Label_Note)
		Me.Controls.Add(Label_Multicurrency)
		Me.Controls.Add(Label9)
		Me.Controls.Add(Image_QBBANNER)
		Me.Controls.Add(Label7)
		Me.Controls.Add(Label8)
		Me.Controls.Add(Label16)
		Me.Controls.Add(Label1)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label4)
		Me.Controls.Add(Label5)
		Me.Controls.Add(Label6)
		Me.Controls.Add(Label3)
		Me.Controls.Add(Label11)
		Me.Controls.Add(Label12)
		Me.Controls.Add(Label13)
		Me.Controls.Add(Label10)
		Me.Controls.Add(Label17)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class