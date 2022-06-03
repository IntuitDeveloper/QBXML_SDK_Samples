<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class Invoice
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
	Public WithEvents Command_Save As System.Windows.Forms.Button
	Public WithEvents Command_Close As System.Windows.Forms.Button
	Public WithEvents Command_Clear As System.Windows.Forms.Button
	Public WithEvents Text_ExRate As System.Windows.Forms.TextBox
	Public WithEvents Text_Total As System.Windows.Forms.TextBox
	Public WithEvents Text_PST As System.Windows.Forms.TextBox
	Public WithEvents Text_GST As System.Windows.Forms.TextBox
	Public WithEvents Label11 As System.Windows.Forms.Label
	Public WithEvents Frame6 As System.Windows.Forms.GroupBox
	Public WithEvents Label10 As System.Windows.Forms.Label
	Public WithEvents Frame5 As System.Windows.Forms.GroupBox
	Public WithEvents Label9 As System.Windows.Forms.Label
	Public WithEvents Frame4 As System.Windows.Forms.GroupBox
	Public WithEvents Label8 As System.Windows.Forms.Label
	Public WithEvents Frame3 As System.Windows.Forms.GroupBox
	Public WithEvents Label6 As System.Windows.Forms.Label
	Public WithEvents Frame2 As System.Windows.Forms.GroupBox
	Public WithEvents Label7 As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents Combo_Second_TaxCode As System.Windows.Forms.ComboBox
	Public WithEvents Combo_Third_TaxCode As System.Windows.Forms.ComboBox
	Public WithEvents Text_Second_Amount As System.Windows.Forms.TextBox
	Public WithEvents Text_Third_Amount As System.Windows.Forms.TextBox
	Public WithEvents Combo_Second_Item As System.Windows.Forms.ComboBox
	Public WithEvents Combo_Third_Item As System.Windows.Forms.ComboBox
	Public WithEvents Text_Third_Rate As System.Windows.Forms.TextBox
	Public WithEvents Text_Third_QTY As System.Windows.Forms.TextBox
	Public WithEvents Text_Desc_Third_Item As System.Windows.Forms.TextBox
	Public WithEvents Text_Second_Rate As System.Windows.Forms.TextBox
	Public WithEvents Text_Second_QTY As System.Windows.Forms.TextBox
	Public WithEvents Text_Desc_Second_Item As System.Windows.Forms.TextBox
	Public WithEvents Combo_First_TaxCode As System.Windows.Forms.ComboBox
	Public WithEvents Combo_First_Item As System.Windows.Forms.ComboBox
	Public WithEvents Text_First_Amount As System.Windows.Forms.TextBox
	Public WithEvents Text_First_Rate As System.Windows.Forms.TextBox
	Public WithEvents Text_First_QTY As System.Windows.Forms.TextBox
	Public WithEvents Text_Desc_First_Item As System.Windows.Forms.TextBox
    Public WithEvents Frame As System.Windows.Forms.GroupBox
    Public WithEvents Text_PO_Number As System.Windows.Forms.TextBox
    Public WithEvents Combo_Terms As System.Windows.Forms.ComboBox
    Public WithEvents Text_Date As System.Windows.Forms.TextBox
    Public WithEvents Combo_ArAccount As System.Windows.Forms.ComboBox
    Public WithEvents Combo_Customer As System.Windows.Forms.ComboBox
    Public WithEvents Image_QBBANNER As System.Windows.Forms.PictureBox
    Public WithEvents Label20 As System.Windows.Forms.Label
    Public WithEvents Label14 As System.Windows.Forms.Label
    Public WithEvents Label13 As System.Windows.Forms.Label
    Public WithEvents Label12 As System.Windows.Forms.Label
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Invoice))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Command_Save = New System.Windows.Forms.Button()
        Me.Command_Close = New System.Windows.Forms.Button()
        Me.Command_Clear = New System.Windows.Forms.Button()
        Me.Text_ExRate = New System.Windows.Forms.TextBox()
        Me.Text_Total = New System.Windows.Forms.TextBox()
        Me.Text_PST = New System.Windows.Forms.TextBox()
        Me.Text_GST = New System.Windows.Forms.TextBox()
        Me.Frame6 = New System.Windows.Forms.GroupBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Frame5 = New System.Windows.Forms.GroupBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Frame4 = New System.Windows.Forms.GroupBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Combo_Second_TaxCode = New System.Windows.Forms.ComboBox()
        Me.Combo_Third_TaxCode = New System.Windows.Forms.ComboBox()
        Me.Text_Second_Amount = New System.Windows.Forms.TextBox()
        Me.Text_Third_Amount = New System.Windows.Forms.TextBox()
        Me.Combo_Second_Item = New System.Windows.Forms.ComboBox()
        Me.Combo_Third_Item = New System.Windows.Forms.ComboBox()
        Me.Text_Third_Rate = New System.Windows.Forms.TextBox()
        Me.Text_Third_QTY = New System.Windows.Forms.TextBox()
        Me.Text_Desc_Third_Item = New System.Windows.Forms.TextBox()
        Me.Text_Second_Rate = New System.Windows.Forms.TextBox()
        Me.Text_Second_QTY = New System.Windows.Forms.TextBox()
        Me.Text_Desc_Second_Item = New System.Windows.Forms.TextBox()
        Me.Combo_First_TaxCode = New System.Windows.Forms.ComboBox()
        Me.Combo_First_Item = New System.Windows.Forms.ComboBox()
        Me.Frame = New System.Windows.Forms.GroupBox()
        Me.Text_First_Amount = New System.Windows.Forms.TextBox()
        Me.Text_First_Rate = New System.Windows.Forms.TextBox()
        Me.Text_First_QTY = New System.Windows.Forms.TextBox()
        Me.Text_Desc_First_Item = New System.Windows.Forms.TextBox()
        Me.Text_PO_Number = New System.Windows.Forms.TextBox()
        Me.Combo_Terms = New System.Windows.Forms.ComboBox()
        Me.Text_Date = New System.Windows.Forms.TextBox()
        Me.Combo_ArAccount = New System.Windows.Forms.ComboBox()
        Me.Combo_Customer = New System.Windows.Forms.ComboBox()
        Me.Image_QBBANNER = New System.Windows.Forms.PictureBox()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Frame6.SuspendLayout()
        Me.Frame5.SuspendLayout()
        Me.Frame4.SuspendLayout()
        Me.Frame3.SuspendLayout()
        Me.Frame2.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me.Frame.SuspendLayout()
        CType(Me.Image_QBBANNER, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Command_Save
        '
        Me.Command_Save.BackColor = System.Drawing.SystemColors.Control
        Me.Command_Save.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command_Save.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command_Save.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command_Save.Location = New System.Drawing.Point(272, 416)
        Me.Command_Save.Name = "Command_Save"
        Me.Command_Save.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command_Save.Size = New System.Drawing.Size(105, 33)
        Me.Command_Save.TabIndex = 25
        Me.Command_Save.Text = "Save"
        Me.Command_Save.UseVisualStyleBackColor = False
        '
        'Command_Close
        '
        Me.Command_Close.BackColor = System.Drawing.SystemColors.Control
        Me.Command_Close.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command_Close.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command_Close.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command_Close.Location = New System.Drawing.Point(528, 416)
        Me.Command_Close.Name = "Command_Close"
        Me.Command_Close.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command_Close.Size = New System.Drawing.Size(105, 33)
        Me.Command_Close.TabIndex = 27
        Me.Command_Close.Text = "Close"
        Me.Command_Close.UseVisualStyleBackColor = False
        '
        'Command_Clear
        '
        Me.Command_Clear.BackColor = System.Drawing.SystemColors.Control
        Me.Command_Clear.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command_Clear.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command_Clear.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command_Clear.Location = New System.Drawing.Point(400, 416)
        Me.Command_Clear.Name = "Command_Clear"
        Me.Command_Clear.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command_Clear.Size = New System.Drawing.Size(105, 33)
        Me.Command_Clear.TabIndex = 26
        Me.Command_Clear.Text = "Clear"
        Me.Command_Clear.UseVisualStyleBackColor = False
        '
        'Text_ExRate
        '
        Me.Text_ExRate.AcceptsReturn = True
        Me.Text_ExRate.BackColor = System.Drawing.SystemColors.Window
        Me.Text_ExRate.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text_ExRate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text_ExRate.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text_ExRate.Location = New System.Drawing.Point(120, 336)
        Me.Text_ExRate.MaxLength = 0
        Me.Text_ExRate.Name = "Text_ExRate"
        Me.Text_ExRate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text_ExRate.Size = New System.Drawing.Size(65, 20)
        Me.Text_ExRate.TabIndex = 21
        '
        'Text_Total
        '
        Me.Text_Total.AcceptsReturn = True
        Me.Text_Total.BackColor = System.Drawing.SystemColors.Window
        Me.Text_Total.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text_Total.Enabled = False
        Me.Text_Total.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text_Total.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text_Total.Location = New System.Drawing.Point(552, 360)
        Me.Text_Total.MaxLength = 0
        Me.Text_Total.Name = "Text_Total"
        Me.Text_Total.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text_Total.Size = New System.Drawing.Size(97, 20)
        Me.Text_Total.TabIndex = 24
        '
        'Text_PST
        '
        Me.Text_PST.AcceptsReturn = True
        Me.Text_PST.BackColor = System.Drawing.SystemColors.Window
        Me.Text_PST.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text_PST.Enabled = False
        Me.Text_PST.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text_PST.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text_PST.Location = New System.Drawing.Point(592, 328)
        Me.Text_PST.MaxLength = 0
        Me.Text_PST.Name = "Text_PST"
        Me.Text_PST.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text_PST.Size = New System.Drawing.Size(57, 20)
        Me.Text_PST.TabIndex = 23
        '
        'Text_GST
        '
        Me.Text_GST.AcceptsReturn = True
        Me.Text_GST.BackColor = System.Drawing.SystemColors.Window
        Me.Text_GST.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text_GST.Enabled = False
        Me.Text_GST.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text_GST.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text_GST.Location = New System.Drawing.Point(592, 304)
        Me.Text_GST.MaxLength = 0
        Me.Text_GST.Name = "Text_GST"
        Me.Text_GST.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text_GST.Size = New System.Drawing.Size(57, 20)
        Me.Text_GST.TabIndex = 22
        '
        'Frame6
        '
        Me.Frame6.BackColor = System.Drawing.SystemColors.Control
        Me.Frame6.Controls.Add(Me.Label11)
        Me.Frame6.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame6.Location = New System.Drawing.Point(656, 104)
        Me.Frame6.Name = "Frame6"
        Me.Frame6.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame6.Size = New System.Drawing.Size(89, 33)
        Me.Frame6.TabIndex = 41
        Me.Frame6.TabStop = False
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.SystemColors.Control
        Me.Label11.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label11.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label11.Location = New System.Drawing.Point(8, 8)
        Me.Label11.Name = "Label11"
        Me.Label11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label11.Size = New System.Drawing.Size(73, 17)
        Me.Label11.TabIndex = 47
        Me.Label11.Text = "Tax Code"
        '
        'Frame5
        '
        Me.Frame5.BackColor = System.Drawing.SystemColors.Control
        Me.Frame5.Controls.Add(Me.Label10)
        Me.Frame5.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame5.Location = New System.Drawing.Point(552, 104)
        Me.Frame5.Name = "Frame5"
        Me.Frame5.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame5.Size = New System.Drawing.Size(105, 33)
        Me.Frame5.TabIndex = 40
        Me.Frame5.TabStop = False
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.SystemColors.Control
        Me.Label10.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label10.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label10.Location = New System.Drawing.Point(24, 8)
        Me.Label10.Name = "Label10"
        Me.Label10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label10.Size = New System.Drawing.Size(57, 17)
        Me.Label10.TabIndex = 46
        Me.Label10.Text = "Amount"
        '
        'Frame4
        '
        Me.Frame4.BackColor = System.Drawing.SystemColors.Control
        Me.Frame4.Controls.Add(Me.Label9)
        Me.Frame4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame4.Location = New System.Drawing.Point(480, 104)
        Me.Frame4.Name = "Frame4"
        Me.Frame4.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame4.Size = New System.Drawing.Size(73, 33)
        Me.Frame4.TabIndex = 39
        Me.Frame4.TabStop = False
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.SystemColors.Control
        Me.Label9.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label9.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label9.Location = New System.Drawing.Point(16, 8)
        Me.Label9.Name = "Label9"
        Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label9.Size = New System.Drawing.Size(41, 17)
        Me.Label9.TabIndex = 45
        Me.Label9.Text = "Rate"
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me.Label8)
        Me.Frame3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame3.Location = New System.Drawing.Point(424, 104)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(57, 33)
        Me.Frame3.TabIndex = 38
        Me.Frame3.TabStop = False
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.SystemColors.Control
        Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label8.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label8.Location = New System.Drawing.Point(16, 8)
        Me.Label8.Name = "Label8"
        Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label8.Size = New System.Drawing.Size(33, 17)
        Me.Label8.TabIndex = 44
        Me.Label8.Text = "Qty"
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.Label6)
        Me.Frame2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame2.Location = New System.Drawing.Point(176, 104)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(249, 33)
        Me.Frame2.TabIndex = 37
        Me.Frame2.TabStop = False
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(80, 8)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(105, 17)
        Me.Label6.TabIndex = 43
        Me.Label6.Text = "Description"
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.Label7)
        Me.Frame1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(32, 104)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(145, 33)
        Me.Frame1.TabIndex = 36
        Me.Frame1.TabStop = False
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(56, 8)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(52, 17)
        Me.Label7.TabIndex = 42
        Me.Label7.Text = "Item"
        '
        'Combo_Second_TaxCode
        '
        Me.Combo_Second_TaxCode.BackColor = System.Drawing.SystemColors.Window
        Me.Combo_Second_TaxCode.Cursor = System.Windows.Forms.Cursors.Default
        Me.Combo_Second_TaxCode.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Combo_Second_TaxCode.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Combo_Second_TaxCode.Location = New System.Drawing.Point(656, 200)
        Me.Combo_Second_TaxCode.Name = "Combo_Second_TaxCode"
        Me.Combo_Second_TaxCode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Combo_Second_TaxCode.Size = New System.Drawing.Size(73, 22)
        Me.Combo_Second_TaxCode.TabIndex = 14
        '
        'Combo_Third_TaxCode
        '
        Me.Combo_Third_TaxCode.BackColor = System.Drawing.SystemColors.Window
        Me.Combo_Third_TaxCode.Cursor = System.Windows.Forms.Cursors.Default
        Me.Combo_Third_TaxCode.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Combo_Third_TaxCode.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Combo_Third_TaxCode.Location = New System.Drawing.Point(656, 240)
        Me.Combo_Third_TaxCode.Name = "Combo_Third_TaxCode"
        Me.Combo_Third_TaxCode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Combo_Third_TaxCode.Size = New System.Drawing.Size(73, 22)
        Me.Combo_Third_TaxCode.TabIndex = 20
        '
        'Text_Second_Amount
        '
        Me.Text_Second_Amount.AcceptsReturn = True
        Me.Text_Second_Amount.BackColor = System.Drawing.SystemColors.Window
        Me.Text_Second_Amount.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text_Second_Amount.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text_Second_Amount.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text_Second_Amount.Location = New System.Drawing.Point(552, 200)
        Me.Text_Second_Amount.MaxLength = 0
        Me.Text_Second_Amount.Name = "Text_Second_Amount"
        Me.Text_Second_Amount.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text_Second_Amount.Size = New System.Drawing.Size(105, 20)
        Me.Text_Second_Amount.TabIndex = 13
        '
        'Text_Third_Amount
        '
        Me.Text_Third_Amount.AcceptsReturn = True
        Me.Text_Third_Amount.BackColor = System.Drawing.SystemColors.Window
        Me.Text_Third_Amount.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text_Third_Amount.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text_Third_Amount.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text_Third_Amount.Location = New System.Drawing.Point(552, 240)
        Me.Text_Third_Amount.MaxLength = 0
        Me.Text_Third_Amount.Name = "Text_Third_Amount"
        Me.Text_Third_Amount.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text_Third_Amount.Size = New System.Drawing.Size(105, 20)
        Me.Text_Third_Amount.TabIndex = 19
        '
        'Combo_Second_Item
        '
        Me.Combo_Second_Item.BackColor = System.Drawing.SystemColors.Window
        Me.Combo_Second_Item.Cursor = System.Windows.Forms.Cursors.Default
        Me.Combo_Second_Item.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Combo_Second_Item.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Combo_Second_Item.Location = New System.Drawing.Point(40, 200)
        Me.Combo_Second_Item.Name = "Combo_Second_Item"
        Me.Combo_Second_Item.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Combo_Second_Item.Size = New System.Drawing.Size(137, 22)
        Me.Combo_Second_Item.TabIndex = 9
        '
        'Combo_Third_Item
        '
        Me.Combo_Third_Item.BackColor = System.Drawing.SystemColors.Window
        Me.Combo_Third_Item.Cursor = System.Windows.Forms.Cursors.Default
        Me.Combo_Third_Item.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Combo_Third_Item.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Combo_Third_Item.Location = New System.Drawing.Point(40, 240)
        Me.Combo_Third_Item.Name = "Combo_Third_Item"
        Me.Combo_Third_Item.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Combo_Third_Item.Size = New System.Drawing.Size(137, 22)
        Me.Combo_Third_Item.TabIndex = 15
        '
        'Text_Third_Rate
        '
        Me.Text_Third_Rate.AcceptsReturn = True
        Me.Text_Third_Rate.BackColor = System.Drawing.SystemColors.Window
        Me.Text_Third_Rate.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text_Third_Rate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text_Third_Rate.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text_Third_Rate.Location = New System.Drawing.Point(480, 240)
        Me.Text_Third_Rate.MaxLength = 0
        Me.Text_Third_Rate.Name = "Text_Third_Rate"
        Me.Text_Third_Rate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text_Third_Rate.Size = New System.Drawing.Size(73, 20)
        Me.Text_Third_Rate.TabIndex = 18
        '
        'Text_Third_QTY
        '
        Me.Text_Third_QTY.AcceptsReturn = True
        Me.Text_Third_QTY.BackColor = System.Drawing.SystemColors.Window
        Me.Text_Third_QTY.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text_Third_QTY.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text_Third_QTY.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text_Third_QTY.Location = New System.Drawing.Point(424, 240)
        Me.Text_Third_QTY.MaxLength = 0
        Me.Text_Third_QTY.Name = "Text_Third_QTY"
        Me.Text_Third_QTY.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text_Third_QTY.Size = New System.Drawing.Size(57, 20)
        Me.Text_Third_QTY.TabIndex = 17
        '
        'Text_Desc_Third_Item
        '
        Me.Text_Desc_Third_Item.AcceptsReturn = True
        Me.Text_Desc_Third_Item.BackColor = System.Drawing.SystemColors.Window
        Me.Text_Desc_Third_Item.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text_Desc_Third_Item.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text_Desc_Third_Item.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text_Desc_Third_Item.Location = New System.Drawing.Point(176, 240)
        Me.Text_Desc_Third_Item.MaxLength = 0
        Me.Text_Desc_Third_Item.Name = "Text_Desc_Third_Item"
        Me.Text_Desc_Third_Item.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text_Desc_Third_Item.Size = New System.Drawing.Size(249, 20)
        Me.Text_Desc_Third_Item.TabIndex = 16
        '
        'Text_Second_Rate
        '
        Me.Text_Second_Rate.AcceptsReturn = True
        Me.Text_Second_Rate.BackColor = System.Drawing.SystemColors.Window
        Me.Text_Second_Rate.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text_Second_Rate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text_Second_Rate.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text_Second_Rate.Location = New System.Drawing.Point(480, 200)
        Me.Text_Second_Rate.MaxLength = 0
        Me.Text_Second_Rate.Name = "Text_Second_Rate"
        Me.Text_Second_Rate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text_Second_Rate.Size = New System.Drawing.Size(73, 20)
        Me.Text_Second_Rate.TabIndex = 12
        '
        'Text_Second_QTY
        '
        Me.Text_Second_QTY.AcceptsReturn = True
        Me.Text_Second_QTY.BackColor = System.Drawing.SystemColors.Window
        Me.Text_Second_QTY.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text_Second_QTY.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text_Second_QTY.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text_Second_QTY.Location = New System.Drawing.Point(424, 200)
        Me.Text_Second_QTY.MaxLength = 0
        Me.Text_Second_QTY.Name = "Text_Second_QTY"
        Me.Text_Second_QTY.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text_Second_QTY.Size = New System.Drawing.Size(57, 20)
        Me.Text_Second_QTY.TabIndex = 11
        '
        'Text_Desc_Second_Item
        '
        Me.Text_Desc_Second_Item.AcceptsReturn = True
        Me.Text_Desc_Second_Item.BackColor = System.Drawing.SystemColors.Window
        Me.Text_Desc_Second_Item.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text_Desc_Second_Item.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text_Desc_Second_Item.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text_Desc_Second_Item.Location = New System.Drawing.Point(176, 200)
        Me.Text_Desc_Second_Item.MaxLength = 0
        Me.Text_Desc_Second_Item.Name = "Text_Desc_Second_Item"
        Me.Text_Desc_Second_Item.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text_Desc_Second_Item.Size = New System.Drawing.Size(249, 20)
        Me.Text_Desc_Second_Item.TabIndex = 10
        '
        'Combo_First_TaxCode
        '
        Me.Combo_First_TaxCode.BackColor = System.Drawing.SystemColors.Window
        Me.Combo_First_TaxCode.Cursor = System.Windows.Forms.Cursors.Default
        Me.Combo_First_TaxCode.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Combo_First_TaxCode.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Combo_First_TaxCode.Location = New System.Drawing.Point(656, 160)
        Me.Combo_First_TaxCode.Name = "Combo_First_TaxCode"
        Me.Combo_First_TaxCode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Combo_First_TaxCode.Size = New System.Drawing.Size(73, 22)
        Me.Combo_First_TaxCode.TabIndex = 8
        '
        'Combo_First_Item
        '
        Me.Combo_First_Item.BackColor = System.Drawing.SystemColors.Window
        Me.Combo_First_Item.Cursor = System.Windows.Forms.Cursors.Default
        Me.Combo_First_Item.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Combo_First_Item.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Combo_First_Item.Location = New System.Drawing.Point(40, 160)
        Me.Combo_First_Item.Name = "Combo_First_Item"
        Me.Combo_First_Item.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Combo_First_Item.Size = New System.Drawing.Size(137, 22)
        Me.Combo_First_Item.TabIndex = 3
        '
        'Frame
        '
        Me.Frame.BackColor = System.Drawing.SystemColors.Control
        Me.Frame.Controls.Add(Me.Text_First_Amount)
        Me.Frame.Controls.Add(Me.Text_First_Rate)
        Me.Frame.Controls.Add(Me.Text_First_QTY)
        Me.Frame.Controls.Add(Me.Text_Desc_First_Item)
        Me.Frame.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame.Location = New System.Drawing.Point(32, 128)
        Me.Frame.Name = "Frame"
        Me.Frame.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame.Size = New System.Drawing.Size(713, 161)
        Me.Frame.TabIndex = 35
        Me.Frame.TabStop = False
        '
        'Text_First_Amount
        '
        Me.Text_First_Amount.AcceptsReturn = True
        Me.Text_First_Amount.BackColor = System.Drawing.SystemColors.Window
        Me.Text_First_Amount.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text_First_Amount.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text_First_Amount.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text_First_Amount.Location = New System.Drawing.Point(520, 32)
        Me.Text_First_Amount.MaxLength = 0
        Me.Text_First_Amount.Name = "Text_First_Amount"
        Me.Text_First_Amount.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text_First_Amount.Size = New System.Drawing.Size(105, 20)
        Me.Text_First_Amount.TabIndex = 7
        '
        'Text_First_Rate
        '
        Me.Text_First_Rate.AcceptsReturn = True
        Me.Text_First_Rate.BackColor = System.Drawing.SystemColors.Window
        Me.Text_First_Rate.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text_First_Rate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text_First_Rate.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text_First_Rate.Location = New System.Drawing.Point(448, 32)
        Me.Text_First_Rate.MaxLength = 0
        Me.Text_First_Rate.Name = "Text_First_Rate"
        Me.Text_First_Rate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text_First_Rate.Size = New System.Drawing.Size(73, 20)
        Me.Text_First_Rate.TabIndex = 6
        '
        'Text_First_QTY
        '
        Me.Text_First_QTY.AcceptsReturn = True
        Me.Text_First_QTY.BackColor = System.Drawing.SystemColors.Window
        Me.Text_First_QTY.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text_First_QTY.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text_First_QTY.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text_First_QTY.Location = New System.Drawing.Point(392, 32)
        Me.Text_First_QTY.MaxLength = 0
        Me.Text_First_QTY.Name = "Text_First_QTY"
        Me.Text_First_QTY.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text_First_QTY.Size = New System.Drawing.Size(57, 20)
        Me.Text_First_QTY.TabIndex = 5
        '
        'Text_Desc_First_Item
        '
        Me.Text_Desc_First_Item.AcceptsReturn = True
        Me.Text_Desc_First_Item.BackColor = System.Drawing.SystemColors.Window
        Me.Text_Desc_First_Item.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text_Desc_First_Item.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text_Desc_First_Item.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text_Desc_First_Item.Location = New System.Drawing.Point(144, 32)
        Me.Text_Desc_First_Item.MaxLength = 0
        Me.Text_Desc_First_Item.Name = "Text_Desc_First_Item"
        Me.Text_Desc_First_Item.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text_Desc_First_Item.Size = New System.Drawing.Size(249, 20)
        Me.Text_Desc_First_Item.TabIndex = 4
        '
        'Text_PO_Number
        '
        Me.Text_PO_Number.AcceptsReturn = True
        Me.Text_PO_Number.BackColor = System.Drawing.SystemColors.Window
        Me.Text_PO_Number.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text_PO_Number.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text_PO_Number.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text_PO_Number.Location = New System.Drawing.Point(648, 208)
        Me.Text_PO_Number.MaxLength = 0
        Me.Text_PO_Number.Name = "Text_PO_Number"
        Me.Text_PO_Number.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text_PO_Number.Size = New System.Drawing.Size(89, 20)
        Me.Text_PO_Number.TabIndex = 33
        '
        'Combo_Terms
        '
        Me.Combo_Terms.BackColor = System.Drawing.SystemColors.Window
        Me.Combo_Terms.Cursor = System.Windows.Forms.Cursors.Default
        Me.Combo_Terms.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Combo_Terms.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Combo_Terms.Location = New System.Drawing.Point(544, 208)
        Me.Combo_Terms.Name = "Combo_Terms"
        Me.Combo_Terms.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Combo_Terms.Size = New System.Drawing.Size(97, 22)
        Me.Combo_Terms.TabIndex = 31
        '
        'Text_Date
        '
        Me.Text_Date.AcceptsReturn = True
        Me.Text_Date.BackColor = System.Drawing.SystemColors.Window
        Me.Text_Date.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text_Date.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text_Date.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text_Date.Location = New System.Drawing.Point(456, 40)
        Me.Text_Date.MaxLength = 0
        Me.Text_Date.Name = "Text_Date"
        Me.Text_Date.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text_Date.Size = New System.Drawing.Size(113, 20)
        Me.Text_Date.TabIndex = 2
        '
        'Combo_ArAccount
        '
        Me.Combo_ArAccount.BackColor = System.Drawing.SystemColors.Window
        Me.Combo_ArAccount.Cursor = System.Windows.Forms.Cursors.Default
        Me.Combo_ArAccount.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Combo_ArAccount.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Combo_ArAccount.Location = New System.Drawing.Point(240, 40)
        Me.Combo_ArAccount.Name = "Combo_ArAccount"
        Me.Combo_ArAccount.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Combo_ArAccount.Size = New System.Drawing.Size(169, 22)
        Me.Combo_ArAccount.TabIndex = 1
        '
        'Combo_Customer
        '
        Me.Combo_Customer.BackColor = System.Drawing.SystemColors.Window
        Me.Combo_Customer.Cursor = System.Windows.Forms.Cursors.Default
        Me.Combo_Customer.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Combo_Customer.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Combo_Customer.Location = New System.Drawing.Point(32, 40)
        Me.Combo_Customer.Name = "Combo_Customer"
        Me.Combo_Customer.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Combo_Customer.Size = New System.Drawing.Size(161, 22)
        Me.Combo_Customer.TabIndex = 0
        '
        'Image_QBBANNER
        '
        Me.Image_QBBANNER.Cursor = System.Windows.Forms.Cursors.Default
        Me.Image_QBBANNER.Image = CType(resources.GetObject("Image_QBBANNER.Image"), System.Drawing.Image)
        Me.Image_QBBANNER.Location = New System.Drawing.Point(40, 416)
        Me.Image_QBBANNER.Name = "Image_QBBANNER"
        Me.Image_QBBANNER.Size = New System.Drawing.Size(137, 33)
        Me.Image_QBBANNER.TabIndex = 42
        Me.Image_QBBANNER.TabStop = False
        '
        'Label20
        '
        Me.Label20.BackColor = System.Drawing.SystemColors.Control
        Me.Label20.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label20.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label20.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label20.Location = New System.Drawing.Point(24, 336)
        Me.Label20.Name = "Label20"
        Me.Label20.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label20.Size = New System.Drawing.Size(137, 17)
        Me.Label20.TabIndex = 51
        Me.Label20.Text = "Exchange Rate:"
        '
        'Label14
        '
        Me.Label14.BackColor = System.Drawing.SystemColors.Control
        Me.Label14.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label14.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label14.Location = New System.Drawing.Point(512, 360)
        Me.Label14.Name = "Label14"
        Me.Label14.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label14.Size = New System.Drawing.Size(41, 25)
        Me.Label14.TabIndex = 50
        Me.Label14.Text = "Total"
        '
        'Label13
        '
        Me.Label13.BackColor = System.Drawing.SystemColors.Control
        Me.Label13.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label13.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label13.Location = New System.Drawing.Point(552, 328)
        Me.Label13.Name = "Label13"
        Me.Label13.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label13.Size = New System.Drawing.Size(25, 17)
        Me.Label13.TabIndex = 49
        Me.Label13.Text = "PST"
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.SystemColors.Control
        Me.Label12.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label12.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label12.Location = New System.Drawing.Point(552, 304)
        Me.Label12.Name = "Label12"
        Me.Label12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label12.Size = New System.Drawing.Size(33, 17)
        Me.Label12.TabIndex = 48
        Me.Label12.Text = "GST"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(648, 184)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(89, 17)
        Me.Label5.TabIndex = 34
        Me.Label5.Text = "P.O. No."
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(544, 184)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(89, 17)
        Me.Label4.TabIndex = 32
        Me.Label4.Text = "Terms"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(456, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(113, 17)
        Me.Label3.TabIndex = 30
        Me.Label3.Text = "Date"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(240, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(113, 17)
        Me.Label2.TabIndex = 29
        Me.Label2.Text = "Account"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(32, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(129, 17)
        Me.Label1.TabIndex = 28
        Me.Label1.Text = "Customer:Job"
        '
        'Invoice
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(778, 466)
        Me.Controls.Add(Me.Command_Save)
        Me.Controls.Add(Me.Command_Close)
        Me.Controls.Add(Me.Command_Clear)
        Me.Controls.Add(Me.Text_ExRate)
        Me.Controls.Add(Me.Text_Total)
        Me.Controls.Add(Me.Text_PST)
        Me.Controls.Add(Me.Text_GST)
        Me.Controls.Add(Me.Frame6)
        Me.Controls.Add(Me.Frame5)
        Me.Controls.Add(Me.Frame4)
        Me.Controls.Add(Me.Frame3)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.Combo_Second_TaxCode)
        Me.Controls.Add(Me.Combo_Third_TaxCode)
        Me.Controls.Add(Me.Text_Second_Amount)
        Me.Controls.Add(Me.Text_Third_Amount)
        Me.Controls.Add(Me.Combo_Second_Item)
        Me.Controls.Add(Me.Combo_Third_Item)
        Me.Controls.Add(Me.Text_Third_Rate)
        Me.Controls.Add(Me.Text_Third_QTY)
        Me.Controls.Add(Me.Text_Desc_Third_Item)
        Me.Controls.Add(Me.Text_Second_Rate)
        Me.Controls.Add(Me.Text_Second_QTY)
        Me.Controls.Add(Me.Text_Desc_Second_Item)
        Me.Controls.Add(Me.Combo_First_TaxCode)
        Me.Controls.Add(Me.Combo_First_Item)
        Me.Controls.Add(Me.Frame)
        Me.Controls.Add(Me.Text_PO_Number)
        Me.Controls.Add(Me.Combo_Terms)
        Me.Controls.Add(Me.Text_Date)
        Me.Controls.Add(Me.Combo_ArAccount)
        Me.Controls.Add(Me.Combo_Customer)
        Me.Controls.Add(Me.Image_QBBANNER)
        Me.Controls.Add(Me.Label20)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(4, 23)
        Me.Name = "Invoice"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "Add Invoice Sample"
        Me.Frame6.ResumeLayout(False)
        Me.Frame5.ResumeLayout(False)
        Me.Frame4.ResumeLayout(False)
        Me.Frame3.ResumeLayout(False)
        Me.Frame2.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.Frame.ResumeLayout(False)
        Me.Frame.PerformLayout()
        CType(Me.Image_QBBANNER, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    'Private WithEvents Line5 As PowerPacks.LineShape
    'Private WithEvents Line4 As PowerPacks.LineShape
    'Private WithEvents Line3 As PowerPacks.LineShape
    'Private WithEvents Line2 As PowerPacks.LineShape
    'Private WithEvents Line1 As PowerPacks.LineShape
    'Private WithEvents Line6 As PowerPacks.LineShape
    'Private WithEvents ShapeContainer2 As PowerPacks.ShapeContainer
    'Private WithEvents ShapeContainer1 As PowerPacks.ShapeContainer
#End Region
End Class