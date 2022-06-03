<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class MainForm
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
	Public WithEvents Combo_Credit_Memos As System.Windows.Forms.ComboBox
	Public WithEvents Combo_Invoices As System.Windows.Forms.ComboBox
	Public WithEvents Label_Credit_Memo As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Frame3 As System.Windows.Forms.GroupBox
	Public WithEvents Command_Apply_Payment As System.Windows.Forms.Button
	Public WithEvents Text_Ref_Number As System.Windows.Forms.TextBox
	Public WithEvents Combo_Customer As System.Windows.Forms.ComboBox
	Public WithEvents Combo_Pay_Method As System.Windows.Forms.ComboBox
	Public WithEvents Command_Display_Invoices As System.Windows.Forms.Button
	Public WithEvents Label8 As System.Windows.Forms.Label
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents Text_Pay_Date As System.Windows.Forms.TextBox
    Public WithEvents Text_Amt_Paid As System.Windows.Forms.TextBox
    Public WithEvents Text_Credit_Memo As System.Windows.Forms.TextBox
    Public WithEvents Text_Discount As System.Windows.Forms.TextBox
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Label6 As System.Windows.Forms.Label
    Public WithEvents Label_Credit_Amt As System.Windows.Forms.Label
    Public WithEvents Label11 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents Label7 As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainForm))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.Combo_Credit_Memos = New System.Windows.Forms.ComboBox()
        Me.Combo_Invoices = New System.Windows.Forms.ComboBox()
        Me.Label_Credit_Memo = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Command_Apply_Payment = New System.Windows.Forms.Button()
        Me.Text_Ref_Number = New System.Windows.Forms.TextBox()
        Me.Combo_Customer = New System.Windows.Forms.ComboBox()
        Me.Combo_Pay_Method = New System.Windows.Forms.ComboBox()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.Command_Display_Invoices = New System.Windows.Forms.Button()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.Text_Pay_Date = New System.Windows.Forms.TextBox()
        Me.Text_Amt_Paid = New System.Windows.Forms.TextBox()
        Me.Text_Credit_Memo = New System.Windows.Forms.TextBox()
        Me.Text_Discount = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label_Credit_Amt = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Image_QBCUST = New System.Windows.Forms.PictureBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Frame3.SuspendLayout()
        Me.Frame2.SuspendLayout()
        Me.Frame1.SuspendLayout()
        CType(Me.Image_QBCUST, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me.Combo_Credit_Memos)
        Me.Frame3.Controls.Add(Me.Combo_Invoices)
        Me.Frame3.Controls.Add(Me.Label_Credit_Memo)
        Me.Frame3.Controls.Add(Me.Label4)
        Me.Frame3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame3.Location = New System.Drawing.Point(8, 192)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(573, 129)
        Me.Frame3.TabIndex = 13
        Me.Frame3.TabStop = False
        Me.Frame3.Text = "Step 2: Select Payment Options"
        '
        'Combo_Credit_Memos
        '
        Me.Combo_Credit_Memos.BackColor = System.Drawing.SystemColors.Window
        Me.Combo_Credit_Memos.Cursor = System.Windows.Forms.Cursors.Default
        Me.Combo_Credit_Memos.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Combo_Credit_Memos.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Combo_Credit_Memos.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Combo_Credit_Memos.Location = New System.Drawing.Point(196, 84)
        Me.Combo_Credit_Memos.Name = "Combo_Credit_Memos"
        Me.Combo_Credit_Memos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Combo_Credit_Memos.Size = New System.Drawing.Size(361, 22)
        Me.Combo_Credit_Memos.TabIndex = 3
        '
        'Combo_Invoices
        '
        Me.Combo_Invoices.BackColor = System.Drawing.SystemColors.Window
        Me.Combo_Invoices.Cursor = System.Windows.Forms.Cursors.Default
        Me.Combo_Invoices.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Combo_Invoices.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Combo_Invoices.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Combo_Invoices.Location = New System.Drawing.Point(112, 48)
        Me.Combo_Invoices.Name = "Combo_Invoices"
        Me.Combo_Invoices.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Combo_Invoices.Size = New System.Drawing.Size(445, 22)
        Me.Combo_Invoices.TabIndex = 2
        '
        'Label_Credit_Memo
        '
        Me.Label_Credit_Memo.BackColor = System.Drawing.SystemColors.Control
        Me.Label_Credit_Memo.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label_Credit_Memo.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_Credit_Memo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label_Credit_Memo.Location = New System.Drawing.Point(12, 84)
        Me.Label_Credit_Memo.Name = "Label_Credit_Memo"
        Me.Label_Credit_Memo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label_Credit_Memo.Size = New System.Drawing.Size(173, 21)
        Me.Label_Credit_Memo.TabIndex = 19
        Me.Label_Credit_Memo.Text = "Choose a Credit Memo (Optional):"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(12, 52)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(129, 25)
        Me.Label4.TabIndex = 18
        Me.Label4.Text = "Choose an Invoice:"
        '
        'Command_Apply_Payment
        '
        Me.Command_Apply_Payment.BackColor = System.Drawing.SystemColors.Control
        Me.Command_Apply_Payment.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command_Apply_Payment.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command_Apply_Payment.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command_Apply_Payment.Location = New System.Drawing.Point(416, 464)
        Me.Command_Apply_Payment.Name = "Command_Apply_Payment"
        Me.Command_Apply_Payment.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command_Apply_Payment.Size = New System.Drawing.Size(145, 25)
        Me.Command_Apply_Payment.TabIndex = 10
        Me.Command_Apply_Payment.Text = "Apply Payment"
        Me.Command_Apply_Payment.UseVisualStyleBackColor = False
        '
        'Text_Ref_Number
        '
        Me.Text_Ref_Number.AcceptsReturn = True
        Me.Text_Ref_Number.BackColor = System.Drawing.SystemColors.Window
        Me.Text_Ref_Number.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text_Ref_Number.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text_Ref_Number.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text_Ref_Number.Location = New System.Drawing.Point(72, 472)
        Me.Text_Ref_Number.MaxLength = 0
        Me.Text_Ref_Number.Name = "Text_Ref_Number"
        Me.Text_Ref_Number.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text_Ref_Number.Size = New System.Drawing.Size(167, 19)
        Me.Text_Ref_Number.TabIndex = 5
        '
        'Combo_Customer
        '
        Me.Combo_Customer.BackColor = System.Drawing.SystemColors.Window
        Me.Combo_Customer.Cursor = System.Windows.Forms.Cursors.Default
        Me.Combo_Customer.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Combo_Customer.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Combo_Customer.Location = New System.Drawing.Point(112, 136)
        Me.Combo_Customer.Name = "Combo_Customer"
        Me.Combo_Customer.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Combo_Customer.Size = New System.Drawing.Size(167, 22)
        Me.Combo_Customer.Sorted = True
        Me.Combo_Customer.TabIndex = 0
        '
        'Combo_Pay_Method
        '
        Me.Combo_Pay_Method.BackColor = System.Drawing.SystemColors.Window
        Me.Combo_Pay_Method.Cursor = System.Windows.Forms.Cursors.Default
        Me.Combo_Pay_Method.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Combo_Pay_Method.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Combo_Pay_Method.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Combo_Pay_Method.Location = New System.Drawing.Point(72, 424)
        Me.Combo_Pay_Method.Name = "Combo_Pay_Method"
        Me.Combo_Pay_Method.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Combo_Pay_Method.Size = New System.Drawing.Size(167, 22)
        Me.Combo_Pay_Method.TabIndex = 6
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.Image_QBCUST)
        Me.Frame2.Controls.Add(Me.Command_Display_Invoices)
        Me.Frame2.Controls.Add(Me.Label8)
        Me.Frame2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame2.Location = New System.Drawing.Point(8, 56)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(573, 121)
        Me.Frame2.TabIndex = 12
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Step 1: Choose a Customer"
        '
        'Command_Display_Invoices
        '
        Me.Command_Display_Invoices.BackColor = System.Drawing.SystemColors.Control
        Me.Command_Display_Invoices.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command_Display_Invoices.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command_Display_Invoices.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command_Display_Invoices.Location = New System.Drawing.Point(288, 72)
        Me.Command_Display_Invoices.Name = "Command_Display_Invoices"
        Me.Command_Display_Invoices.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command_Display_Invoices.Size = New System.Drawing.Size(145, 29)
        Me.Command_Display_Invoices.TabIndex = 1
        Me.Command_Display_Invoices.Text = "Fill In Customer Info Below"
        Me.Command_Display_Invoices.UseVisualStyleBackColor = False
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.SystemColors.Control
        Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label8.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label8.Location = New System.Drawing.Point(88, 36)
        Me.Label8.Name = "Label8"
        Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label8.Size = New System.Drawing.Size(433, 41)
        Me.Label8.TabIndex = 21
        Me.Label8.Text = "Start by choosing the customer whose payment is being received.  After you click " &
    """Fill In Customer Info Below,"" you will be able to choose a specific invoice to " &
    "make a payment on."
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.Text_Pay_Date)
        Me.Frame1.Controls.Add(Me.Text_Amt_Paid)
        Me.Frame1.Controls.Add(Me.Text_Credit_Memo)
        Me.Frame1.Controls.Add(Me.Text_Discount)
        Me.Frame1.Controls.Add(Me.Label3)
        Me.Frame1.Controls.Add(Me.Label2)
        Me.Frame1.Controls.Add(Me.Label1)
        Me.Frame1.Controls.Add(Me.Label6)
        Me.Frame1.Controls.Add(Me.Label_Credit_Amt)
        Me.Frame1.Controls.Add(Me.Label11)
        Me.Frame1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(8, 344)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(573, 205)
        Me.Frame1.TabIndex = 14
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Step 3: Apply Payment"
        '
        'Text_Pay_Date
        '
        Me.Text_Pay_Date.AcceptsReturn = True
        Me.Text_Pay_Date.BackColor = System.Drawing.SystemColors.Window
        Me.Text_Pay_Date.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text_Pay_Date.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text_Pay_Date.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text_Pay_Date.Location = New System.Drawing.Point(68, 40)
        Me.Text_Pay_Date.MaxLength = 0
        Me.Text_Pay_Date.Name = "Text_Pay_Date"
        Me.Text_Pay_Date.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text_Pay_Date.Size = New System.Drawing.Size(167, 19)
        Me.Text_Pay_Date.TabIndex = 4
        '
        'Text_Amt_Paid
        '
        Me.Text_Amt_Paid.AcceptsReturn = True
        Me.Text_Amt_Paid.BackColor = System.Drawing.SystemColors.Window
        Me.Text_Amt_Paid.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text_Amt_Paid.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text_Amt_Paid.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text_Amt_Paid.Location = New System.Drawing.Point(152, 168)
        Me.Text_Amt_Paid.MaxLength = 0
        Me.Text_Amt_Paid.Name = "Text_Amt_Paid"
        Me.Text_Amt_Paid.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text_Amt_Paid.Size = New System.Drawing.Size(80, 19)
        Me.Text_Amt_Paid.TabIndex = 7
        '
        'Text_Credit_Memo
        '
        Me.Text_Credit_Memo.AcceptsReturn = True
        Me.Text_Credit_Memo.BackColor = System.Drawing.SystemColors.Window
        Me.Text_Credit_Memo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text_Credit_Memo.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text_Credit_Memo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text_Credit_Memo.Location = New System.Drawing.Point(472, 40)
        Me.Text_Credit_Memo.MaxLength = 0
        Me.Text_Credit_Memo.Name = "Text_Credit_Memo"
        Me.Text_Credit_Memo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text_Credit_Memo.Size = New System.Drawing.Size(80, 19)
        Me.Text_Credit_Memo.TabIndex = 8
        '
        'Text_Discount
        '
        Me.Text_Discount.AcceptsReturn = True
        Me.Text_Discount.BackColor = System.Drawing.SystemColors.Window
        Me.Text_Discount.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text_Discount.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text_Discount.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text_Discount.Location = New System.Drawing.Point(472, 72)
        Me.Text_Discount.MaxLength = 0
        Me.Text_Discount.Name = "Text_Discount"
        Me.Text_Discount.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text_Discount.Size = New System.Drawing.Size(80, 19)
        Me.Text_Discount.TabIndex = 9
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(28, 24)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(193, 21)
        Me.Label3.TabIndex = 24
        Me.Label3.Text = "Payment Date (MM/DD/YYYY):"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(24, 112)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(141, 25)
        Me.Label2.TabIndex = 23
        Me.Label2.Text = "Ref/Check Number:"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(24, 64)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(117, 21)
        Me.Label1.TabIndex = 22
        Me.Label1.Text = "Payment Method:"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(32, 168)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(149, 21)
        Me.Label6.TabIndex = 17
        Me.Label6.Text = "Amount of Payment:   $"
        '
        'Label_Credit_Amt
        '
        Me.Label_Credit_Amt.BackColor = System.Drawing.SystemColors.Control
        Me.Label_Credit_Amt.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label_Credit_Amt.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_Credit_Amt.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label_Credit_Amt.Location = New System.Drawing.Point(288, 40)
        Me.Label_Credit_Amt.Name = "Label_Credit_Amt"
        Me.Label_Credit_Amt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label_Credit_Amt.Size = New System.Drawing.Size(173, 25)
        Me.Label_Credit_Amt.TabIndex = 16
        Me.Label_Credit_Amt.Text = "Credit Memo (If Selected Above):   $"
        Me.Label_Credit_Amt.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.SystemColors.Control
        Me.Label11.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label11.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label11.Location = New System.Drawing.Point(288, 72)
        Me.Label11.Name = "Label11"
        Me.Label11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label11.Size = New System.Drawing.Size(169, 21)
        Me.Label11.TabIndex = 15
        Me.Label11.Text = "Discount (Optional):   $"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.Font = New System.Drawing.Font("Verdana", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(8, 8)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(541, 25)
        Me.Label7.TabIndex = 11
        Me.Label7.Text = " Receive Payment"
        '
        'Image_QBCUST
        '
        Me.Image_QBCUST.Image = CType(resources.GetObject("Image_QBCUST.Image"), System.Drawing.Image)
        Me.Image_QBCUST.Location = New System.Drawing.Point(51, 36)
        Me.Image_QBCUST.Name = "Image_QBCUST"
        Me.Image_QBCUST.Size = New System.Drawing.Size(30, 33)
        Me.Image_QBCUST.TabIndex = 22
        Me.Image_QBCUST.TabStop = False
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(72, 564)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(429, 35)
        Me.PictureBox1.TabIndex = 15
        Me.PictureBox1.TabStop = False
        '
        'MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(591, 610)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Frame3)
        Me.Controls.Add(Me.Command_Apply_Payment)
        Me.Controls.Add(Me.Text_Ref_Number)
        Me.Controls.Add(Me.Combo_Customer)
        Me.Controls.Add(Me.Combo_Pay_Method)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.Label7)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(282, 129)
        Me.Name = "MainForm"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Enter a Payment"
        Me.Frame3.ResumeLayout(False)
        Me.Frame2.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        CType(Me.Image_QBCUST, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Image_QBCUST As PictureBox
    Friend WithEvents PictureBox1 As PictureBox
#End Region
End Class