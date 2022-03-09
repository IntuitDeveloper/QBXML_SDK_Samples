<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmAddDataExtDef
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
	Public WithEvents cmdShowResponse As System.Windows.Forms.Button
	Public WithEvents cmdShowRequest As System.Windows.Forms.Button
	Public WithEvents Text1 As System.Windows.Forms.TextBox
	Public WithEvents cmdClearAll As System.Windows.Forms.Button
	Public WithEvents cmdAllTransactions As System.Windows.Forms.Button
	Public WithEvents cmdAllEntities As System.Windows.Forms.Button
	Public WithEvents cmdAll As System.Windows.Forms.Button
	Public WithEvents cmdQuit As System.Windows.Forms.Button
	Public WithEvents cmdAdd As System.Windows.Forms.Button
	Public WithEvents chkSalesTaxPaymentCheck As System.Windows.Forms.CheckBox
	Public WithEvents chkSalesReceipt As System.Windows.Forms.CheckBox
	Public WithEvents chkReceivePayment As System.Windows.Forms.CheckBox
	Public WithEvents chkPurchaseOrder As System.Windows.Forms.CheckBox
	Public WithEvents chkInvoice As System.Windows.Forms.CheckBox
	Public WithEvents chkJournalEntry As System.Windows.Forms.CheckBox
	Public WithEvents chkInventoryAdjustment As System.Windows.Forms.CheckBox
	Public WithEvents chkEstimate As System.Windows.Forms.CheckBox
	Public WithEvents chkDeposit As System.Windows.Forms.CheckBox
	Public WithEvents chkCreditMemo As System.Windows.Forms.CheckBox
	Public WithEvents chkCreditCardCredit As System.Windows.Forms.CheckBox
	Public WithEvents chkCreditCardCharge As System.Windows.Forms.CheckBox
	Public WithEvents chkCheck As System.Windows.Forms.CheckBox
	Public WithEvents chkCharge As System.Windows.Forms.CheckBox
	Public WithEvents chkBillPaymentCreditCard As System.Windows.Forms.CheckBox
	Public WithEvents chkItem As System.Windows.Forms.CheckBox
	Public WithEvents chkBill As System.Windows.Forms.CheckBox
	Public WithEvents chkAccount As System.Windows.Forms.CheckBox
	Public WithEvents chkBillPaymentCheck As System.Windows.Forms.CheckBox
	Public WithEvents chkOtherName As System.Windows.Forms.CheckBox
	Public WithEvents chkEmployee As System.Windows.Forms.CheckBox
	Public WithEvents chkVendor As System.Windows.Forms.CheckBox
	Public WithEvents chkVendorCredit As System.Windows.Forms.CheckBox
	Public WithEvents chkCustomer As System.Windows.Forms.CheckBox
	Public WithEvents lstDataExtType As System.Windows.Forms.ListBox
	Public WithEvents txtDataExtName As System.Windows.Forms.TextBox
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAddDataExtDef))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdShowResponse = New System.Windows.Forms.Button()
        Me.cmdShowRequest = New System.Windows.Forms.Button()
        Me.Text1 = New System.Windows.Forms.TextBox()
        Me.cmdClearAll = New System.Windows.Forms.Button()
        Me.cmdAllTransactions = New System.Windows.Forms.Button()
        Me.cmdAllEntities = New System.Windows.Forms.Button()
        Me.cmdAll = New System.Windows.Forms.Button()
        Me.cmdQuit = New System.Windows.Forms.Button()
        Me.cmdAdd = New System.Windows.Forms.Button()
        Me.chkSalesTaxPaymentCheck = New System.Windows.Forms.CheckBox()
        Me.chkSalesReceipt = New System.Windows.Forms.CheckBox()
        Me.chkReceivePayment = New System.Windows.Forms.CheckBox()
        Me.chkPurchaseOrder = New System.Windows.Forms.CheckBox()
        Me.chkInvoice = New System.Windows.Forms.CheckBox()
        Me.chkJournalEntry = New System.Windows.Forms.CheckBox()
        Me.chkInventoryAdjustment = New System.Windows.Forms.CheckBox()
        Me.chkEstimate = New System.Windows.Forms.CheckBox()
        Me.chkDeposit = New System.Windows.Forms.CheckBox()
        Me.chkCreditMemo = New System.Windows.Forms.CheckBox()
        Me.chkCreditCardCredit = New System.Windows.Forms.CheckBox()
        Me.chkCreditCardCharge = New System.Windows.Forms.CheckBox()
        Me.chkCheck = New System.Windows.Forms.CheckBox()
        Me.chkCharge = New System.Windows.Forms.CheckBox()
        Me.chkBillPaymentCreditCard = New System.Windows.Forms.CheckBox()
        Me.chkItem = New System.Windows.Forms.CheckBox()
        Me.chkBill = New System.Windows.Forms.CheckBox()
        Me.chkAccount = New System.Windows.Forms.CheckBox()
        Me.chkBillPaymentCheck = New System.Windows.Forms.CheckBox()
        Me.chkOtherName = New System.Windows.Forms.CheckBox()
        Me.chkEmployee = New System.Windows.Forms.CheckBox()
        Me.chkVendor = New System.Windows.Forms.CheckBox()
        Me.chkVendorCredit = New System.Windows.Forms.CheckBox()
        Me.chkCustomer = New System.Windows.Forms.CheckBox()
        Me.lstDataExtType = New System.Windows.Forms.ListBox()
        Me.txtDataExtName = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'cmdShowResponse
        '
        Me.cmdShowResponse.BackColor = System.Drawing.SystemColors.Control
        Me.cmdShowResponse.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdShowResponse.Enabled = False
        Me.cmdShowResponse.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdShowResponse.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdShowResponse.Location = New System.Drawing.Point(248, 528)
        Me.cmdShowResponse.Name = "cmdShowResponse"
        Me.cmdShowResponse.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdShowResponse.Size = New System.Drawing.Size(73, 41)
        Me.cmdShowResponse.TabIndex = 38
        Me.cmdShowResponse.Text = "Show Add Response"
        Me.cmdShowResponse.UseVisualStyleBackColor = False
        '
        'cmdShowRequest
        '
        Me.cmdShowRequest.BackColor = System.Drawing.SystemColors.Control
        Me.cmdShowRequest.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdShowRequest.Enabled = False
        Me.cmdShowRequest.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdShowRequest.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdShowRequest.Location = New System.Drawing.Point(168, 528)
        Me.cmdShowRequest.Name = "cmdShowRequest"
        Me.cmdShowRequest.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdShowRequest.Size = New System.Drawing.Size(73, 41)
        Me.cmdShowRequest.TabIndex = 37
        Me.cmdShowRequest.Text = "Show Add Request"
        Me.cmdShowRequest.UseVisualStyleBackColor = False
        '
        'Text1
        '
        Me.Text1.AcceptsReturn = True
        Me.Text1.BackColor = System.Drawing.SystemColors.Menu
        Me.Text1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Text1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text1.Location = New System.Drawing.Point(8, 40)
        Me.Text1.MaxLength = 0
        Me.Text1.Multiline = True
        Me.Text1.Name = "Text1"
        Me.Text1.ReadOnly = True
        Me.Text1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text1.Size = New System.Drawing.Size(401, 81)
        Me.Text1.TabIndex = 36
        Me.Text1.TabStop = False
        Me.Text1.Text = resources.GetString("Text1.Text")
        '
        'cmdClearAll
        '
        Me.cmdClearAll.BackColor = System.Drawing.SystemColors.Control
        Me.cmdClearAll.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdClearAll.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdClearAll.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdClearAll.Location = New System.Drawing.Point(152, 480)
        Me.cmdClearAll.Name = "cmdClearAll"
        Me.cmdClearAll.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdClearAll.Size = New System.Drawing.Size(121, 25)
        Me.cmdClearAll.TabIndex = 35
        Me.cmdClearAll.Text = "Clear All Checkboxes"
        Me.cmdClearAll.UseVisualStyleBackColor = False
        '
        'cmdAllTransactions
        '
        Me.cmdAllTransactions.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAllTransactions.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAllTransactions.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAllTransactions.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAllTransactions.Location = New System.Drawing.Point(304, 192)
        Me.cmdAllTransactions.Name = "cmdAllTransactions"
        Me.cmdAllTransactions.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAllTransactions.Size = New System.Drawing.Size(113, 25)
        Me.cmdAllTransactions.TabIndex = 34
        Me.cmdAllTransactions.Text = "All Transactions"
        Me.cmdAllTransactions.UseVisualStyleBackColor = False
        '
        'cmdAllEntities
        '
        Me.cmdAllEntities.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAllEntities.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAllEntities.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAllEntities.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAllEntities.Location = New System.Drawing.Point(200, 192)
        Me.cmdAllEntities.Name = "cmdAllEntities"
        Me.cmdAllEntities.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAllEntities.Size = New System.Drawing.Size(89, 25)
        Me.cmdAllEntities.TabIndex = 33
        Me.cmdAllEntities.Text = "All Entities"
        Me.cmdAllEntities.UseVisualStyleBackColor = False
        '
        'cmdAll
        '
        Me.cmdAll.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAll.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAll.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAll.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAll.Location = New System.Drawing.Point(144, 192)
        Me.cmdAll.Name = "cmdAll"
        Me.cmdAll.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAll.Size = New System.Drawing.Size(41, 25)
        Me.cmdAll.TabIndex = 32
        Me.cmdAll.Text = "All"
        Me.cmdAll.UseVisualStyleBackColor = False
        '
        'cmdQuit
        '
        Me.cmdQuit.BackColor = System.Drawing.SystemColors.Control
        Me.cmdQuit.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdQuit.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdQuit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdQuit.Location = New System.Drawing.Point(328, 528)
        Me.cmdQuit.Name = "cmdQuit"
        Me.cmdQuit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdQuit.Size = New System.Drawing.Size(97, 41)
        Me.cmdQuit.TabIndex = 31
        Me.cmdQuit.Text = "Close Window"
        Me.cmdQuit.UseVisualStyleBackColor = False
        '
        'cmdAdd
        '
        Me.cmdAdd.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAdd.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAdd.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAdd.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAdd.Location = New System.Drawing.Point(8, 528)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAdd.Size = New System.Drawing.Size(153, 41)
        Me.cmdAdd.TabIndex = 30
        Me.cmdAdd.Text = "Add Data Extension Definition"
        Me.cmdAdd.UseVisualStyleBackColor = False
        '
        'chkSalesTaxPaymentCheck
        '
        Me.chkSalesTaxPaymentCheck.BackColor = System.Drawing.SystemColors.Control
        Me.chkSalesTaxPaymentCheck.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkSalesTaxPaymentCheck.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkSalesTaxPaymentCheck.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkSalesTaxPaymentCheck.Location = New System.Drawing.Point(144, 440)
        Me.chkSalesTaxPaymentCheck.Name = "chkSalesTaxPaymentCheck"
        Me.chkSalesTaxPaymentCheck.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkSalesTaxPaymentCheck.Size = New System.Drawing.Size(145, 17)
        Me.chkSalesTaxPaymentCheck.TabIndex = 29
        Me.chkSalesTaxPaymentCheck.Text = "Sales Tax Payment Check"
        Me.chkSalesTaxPaymentCheck.UseVisualStyleBackColor = False
        '
        'chkSalesReceipt
        '
        Me.chkSalesReceipt.BackColor = System.Drawing.SystemColors.Control
        Me.chkSalesReceipt.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkSalesReceipt.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkSalesReceipt.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkSalesReceipt.Location = New System.Drawing.Point(8, 440)
        Me.chkSalesReceipt.Name = "chkSalesReceipt"
        Me.chkSalesReceipt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkSalesReceipt.Size = New System.Drawing.Size(121, 17)
        Me.chkSalesReceipt.TabIndex = 28
        Me.chkSalesReceipt.Text = "Sales Receipt"
        Me.chkSalesReceipt.UseVisualStyleBackColor = False
        '
        'chkReceivePayment
        '
        Me.chkReceivePayment.BackColor = System.Drawing.SystemColors.Control
        Me.chkReceivePayment.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkReceivePayment.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkReceivePayment.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkReceivePayment.Location = New System.Drawing.Point(280, 416)
        Me.chkReceivePayment.Name = "chkReceivePayment"
        Me.chkReceivePayment.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkReceivePayment.Size = New System.Drawing.Size(137, 17)
        Me.chkReceivePayment.TabIndex = 27
        Me.chkReceivePayment.Text = "Receive Payment"
        Me.chkReceivePayment.UseVisualStyleBackColor = False
        '
        'chkPurchaseOrder
        '
        Me.chkPurchaseOrder.BackColor = System.Drawing.SystemColors.Control
        Me.chkPurchaseOrder.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkPurchaseOrder.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkPurchaseOrder.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkPurchaseOrder.Location = New System.Drawing.Point(144, 416)
        Me.chkPurchaseOrder.Name = "chkPurchaseOrder"
        Me.chkPurchaseOrder.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkPurchaseOrder.Size = New System.Drawing.Size(121, 17)
        Me.chkPurchaseOrder.TabIndex = 26
        Me.chkPurchaseOrder.Text = "Purchase Order"
        Me.chkPurchaseOrder.UseVisualStyleBackColor = False
        '
        'chkInvoice
        '
        Me.chkInvoice.BackColor = System.Drawing.SystemColors.Control
        Me.chkInvoice.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkInvoice.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkInvoice.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkInvoice.Location = New System.Drawing.Point(280, 392)
        Me.chkInvoice.Name = "chkInvoice"
        Me.chkInvoice.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkInvoice.Size = New System.Drawing.Size(137, 17)
        Me.chkInvoice.TabIndex = 25
        Me.chkInvoice.Text = "Invoice"
        Me.chkInvoice.UseVisualStyleBackColor = False
        '
        'chkJournalEntry
        '
        Me.chkJournalEntry.BackColor = System.Drawing.SystemColors.Control
        Me.chkJournalEntry.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkJournalEntry.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkJournalEntry.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkJournalEntry.Location = New System.Drawing.Point(8, 416)
        Me.chkJournalEntry.Name = "chkJournalEntry"
        Me.chkJournalEntry.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkJournalEntry.Size = New System.Drawing.Size(121, 17)
        Me.chkJournalEntry.TabIndex = 24
        Me.chkJournalEntry.Text = "Journal Entry"
        Me.chkJournalEntry.UseVisualStyleBackColor = False
        '
        'chkInventoryAdjustment
        '
        Me.chkInventoryAdjustment.BackColor = System.Drawing.SystemColors.Control
        Me.chkInventoryAdjustment.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkInventoryAdjustment.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkInventoryAdjustment.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkInventoryAdjustment.Location = New System.Drawing.Point(144, 392)
        Me.chkInventoryAdjustment.Name = "chkInventoryAdjustment"
        Me.chkInventoryAdjustment.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkInventoryAdjustment.Size = New System.Drawing.Size(129, 17)
        Me.chkInventoryAdjustment.TabIndex = 23
        Me.chkInventoryAdjustment.Text = "Inventory Adjustment"
        Me.chkInventoryAdjustment.UseVisualStyleBackColor = False
        '
        'chkEstimate
        '
        Me.chkEstimate.BackColor = System.Drawing.SystemColors.Control
        Me.chkEstimate.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkEstimate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkEstimate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkEstimate.Location = New System.Drawing.Point(8, 392)
        Me.chkEstimate.Name = "chkEstimate"
        Me.chkEstimate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkEstimate.Size = New System.Drawing.Size(121, 17)
        Me.chkEstimate.TabIndex = 22
        Me.chkEstimate.Text = "Estimate"
        Me.chkEstimate.UseVisualStyleBackColor = False
        '
        'chkDeposit
        '
        Me.chkDeposit.BackColor = System.Drawing.SystemColors.Control
        Me.chkDeposit.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkDeposit.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkDeposit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkDeposit.Location = New System.Drawing.Point(280, 368)
        Me.chkDeposit.Name = "chkDeposit"
        Me.chkDeposit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkDeposit.Size = New System.Drawing.Size(137, 17)
        Me.chkDeposit.TabIndex = 21
        Me.chkDeposit.Text = "Deposit"
        Me.chkDeposit.UseVisualStyleBackColor = False
        '
        'chkCreditMemo
        '
        Me.chkCreditMemo.BackColor = System.Drawing.SystemColors.Control
        Me.chkCreditMemo.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkCreditMemo.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkCreditMemo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkCreditMemo.Location = New System.Drawing.Point(144, 368)
        Me.chkCreditMemo.Name = "chkCreditMemo"
        Me.chkCreditMemo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkCreditMemo.Size = New System.Drawing.Size(121, 17)
        Me.chkCreditMemo.TabIndex = 20
        Me.chkCreditMemo.Text = "Credit Memo"
        Me.chkCreditMemo.UseVisualStyleBackColor = False
        '
        'chkCreditCardCredit
        '
        Me.chkCreditCardCredit.BackColor = System.Drawing.SystemColors.Control
        Me.chkCreditCardCredit.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkCreditCardCredit.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkCreditCardCredit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkCreditCardCredit.Location = New System.Drawing.Point(8, 368)
        Me.chkCreditCardCredit.Name = "chkCreditCardCredit"
        Me.chkCreditCardCredit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkCreditCardCredit.Size = New System.Drawing.Size(129, 17)
        Me.chkCreditCardCredit.TabIndex = 19
        Me.chkCreditCardCredit.Text = "Credit Card Credit"
        Me.chkCreditCardCredit.UseVisualStyleBackColor = False
        '
        'chkCreditCardCharge
        '
        Me.chkCreditCardCharge.BackColor = System.Drawing.SystemColors.Control
        Me.chkCreditCardCharge.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkCreditCardCharge.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkCreditCardCharge.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkCreditCardCharge.Location = New System.Drawing.Point(280, 344)
        Me.chkCreditCardCharge.Name = "chkCreditCardCharge"
        Me.chkCreditCardCharge.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkCreditCardCharge.Size = New System.Drawing.Size(121, 17)
        Me.chkCreditCardCharge.TabIndex = 18
        Me.chkCreditCardCharge.Text = "Credit Card Charge"
        Me.chkCreditCardCharge.UseVisualStyleBackColor = False
        '
        'chkCheck
        '
        Me.chkCheck.BackColor = System.Drawing.SystemColors.Control
        Me.chkCheck.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkCheck.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkCheck.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkCheck.Location = New System.Drawing.Point(144, 344)
        Me.chkCheck.Name = "chkCheck"
        Me.chkCheck.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkCheck.Size = New System.Drawing.Size(129, 17)
        Me.chkCheck.TabIndex = 17
        Me.chkCheck.Text = "Check"
        Me.chkCheck.UseVisualStyleBackColor = False
        '
        'chkCharge
        '
        Me.chkCharge.BackColor = System.Drawing.SystemColors.Control
        Me.chkCharge.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkCharge.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkCharge.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkCharge.Location = New System.Drawing.Point(8, 344)
        Me.chkCharge.Name = "chkCharge"
        Me.chkCharge.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkCharge.Size = New System.Drawing.Size(121, 17)
        Me.chkCharge.TabIndex = 16
        Me.chkCharge.Text = "Charge"
        Me.chkCharge.UseVisualStyleBackColor = False
        '
        'chkBillPaymentCreditCard
        '
        Me.chkBillPaymentCreditCard.BackColor = System.Drawing.SystemColors.Control
        Me.chkBillPaymentCreditCard.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkBillPaymentCreditCard.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkBillPaymentCreditCard.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkBillPaymentCreditCard.Location = New System.Drawing.Point(280, 320)
        Me.chkBillPaymentCreditCard.Name = "chkBillPaymentCreditCard"
        Me.chkBillPaymentCreditCard.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkBillPaymentCreditCard.Size = New System.Drawing.Size(137, 17)
        Me.chkBillPaymentCreditCard.TabIndex = 15
        Me.chkBillPaymentCreditCard.Text = "Bill Payment Credit Card"
        Me.chkBillPaymentCreditCard.UseVisualStyleBackColor = False
        '
        'chkItem
        '
        Me.chkItem.BackColor = System.Drawing.SystemColors.Control
        Me.chkItem.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkItem.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkItem.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkItem.Location = New System.Drawing.Point(8, 280)
        Me.chkItem.Name = "chkItem"
        Me.chkItem.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkItem.Size = New System.Drawing.Size(41, 17)
        Me.chkItem.TabIndex = 14
        Me.chkItem.Text = "Item"
        Me.chkItem.UseVisualStyleBackColor = False
        '
        'chkBill
        '
        Me.chkBill.BackColor = System.Drawing.SystemColors.Control
        Me.chkBill.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkBill.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkBill.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkBill.Location = New System.Drawing.Point(8, 320)
        Me.chkBill.Name = "chkBill"
        Me.chkBill.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkBill.Size = New System.Drawing.Size(41, 17)
        Me.chkBill.TabIndex = 13
        Me.chkBill.Text = "Bill"
        Me.chkBill.UseVisualStyleBackColor = False
        '
        'chkAccount
        '
        Me.chkAccount.BackColor = System.Drawing.SystemColors.Control
        Me.chkAccount.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkAccount.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkAccount.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkAccount.Location = New System.Drawing.Point(80, 280)
        Me.chkAccount.Name = "chkAccount"
        Me.chkAccount.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkAccount.Size = New System.Drawing.Size(121, 17)
        Me.chkAccount.TabIndex = 12
        Me.chkAccount.Text = "Account"
        Me.chkAccount.UseVisualStyleBackColor = False
        '
        'chkBillPaymentCheck
        '
        Me.chkBillPaymentCheck.BackColor = System.Drawing.SystemColors.Control
        Me.chkBillPaymentCheck.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkBillPaymentCheck.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkBillPaymentCheck.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkBillPaymentCheck.Location = New System.Drawing.Point(144, 320)
        Me.chkBillPaymentCheck.Name = "chkBillPaymentCheck"
        Me.chkBillPaymentCheck.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkBillPaymentCheck.Size = New System.Drawing.Size(113, 17)
        Me.chkBillPaymentCheck.TabIndex = 11
        Me.chkBillPaymentCheck.Text = "Bill Payment Check"
        Me.chkBillPaymentCheck.UseVisualStyleBackColor = False
        '
        'chkOtherName
        '
        Me.chkOtherName.BackColor = System.Drawing.SystemColors.Control
        Me.chkOtherName.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkOtherName.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkOtherName.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkOtherName.Location = New System.Drawing.Point(224, 240)
        Me.chkOtherName.Name = "chkOtherName"
        Me.chkOtherName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkOtherName.Size = New System.Drawing.Size(81, 17)
        Me.chkOtherName.TabIndex = 10
        Me.chkOtherName.Text = "Other Name"
        Me.chkOtherName.UseVisualStyleBackColor = False
        '
        'chkEmployee
        '
        Me.chkEmployee.BackColor = System.Drawing.SystemColors.Control
        Me.chkEmployee.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkEmployee.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkEmployee.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkEmployee.Location = New System.Drawing.Point(144, 240)
        Me.chkEmployee.Name = "chkEmployee"
        Me.chkEmployee.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkEmployee.Size = New System.Drawing.Size(73, 17)
        Me.chkEmployee.TabIndex = 9
        Me.chkEmployee.Text = "Employee"
        Me.chkEmployee.UseVisualStyleBackColor = False
        '
        'chkVendor
        '
        Me.chkVendor.BackColor = System.Drawing.SystemColors.Control
        Me.chkVendor.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkVendor.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkVendor.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkVendor.Location = New System.Drawing.Point(80, 240)
        Me.chkVendor.Name = "chkVendor"
        Me.chkVendor.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkVendor.Size = New System.Drawing.Size(57, 17)
        Me.chkVendor.TabIndex = 8
        Me.chkVendor.Text = "Vendor"
        Me.chkVendor.UseVisualStyleBackColor = False
        '
        'chkVendorCredit
        '
        Me.chkVendorCredit.BackColor = System.Drawing.SystemColors.Control
        Me.chkVendorCredit.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkVendorCredit.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkVendorCredit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkVendorCredit.Location = New System.Drawing.Point(8, 464)
        Me.chkVendorCredit.Name = "chkVendorCredit"
        Me.chkVendorCredit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkVendorCredit.Size = New System.Drawing.Size(137, 17)
        Me.chkVendorCredit.TabIndex = 7
        Me.chkVendorCredit.Text = "Vendor Credit"
        Me.chkVendorCredit.UseVisualStyleBackColor = False
        '
        'chkCustomer
        '
        Me.chkCustomer.BackColor = System.Drawing.SystemColors.Control
        Me.chkCustomer.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkCustomer.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkCustomer.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkCustomer.Location = New System.Drawing.Point(8, 240)
        Me.chkCustomer.Name = "chkCustomer"
        Me.chkCustomer.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkCustomer.Size = New System.Drawing.Size(73, 17)
        Me.chkCustomer.TabIndex = 6
        Me.chkCustomer.Text = "Customer"
        Me.chkCustomer.UseVisualStyleBackColor = False
        '
        'lstDataExtType
        '
        Me.lstDataExtType.BackColor = System.Drawing.SystemColors.Window
        Me.lstDataExtType.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstDataExtType.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstDataExtType.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstDataExtType.ItemHeight = 14
        Me.lstDataExtType.Location = New System.Drawing.Point(304, 128)
        Me.lstDataExtType.Name = "lstDataExtType"
        Me.lstDataExtType.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstDataExtType.Size = New System.Drawing.Size(113, 46)
        Me.lstDataExtType.TabIndex = 4
        '
        'txtDataExtName
        '
        Me.txtDataExtName.AcceptsReturn = True
        Me.txtDataExtName.BackColor = System.Drawing.SystemColors.Window
        Me.txtDataExtName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDataExtName.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDataExtName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDataExtName.Location = New System.Drawing.Point(96, 128)
        Me.txtDataExtName.MaxLength = 0
        Me.txtDataExtName.Name = "txtDataExtName"
        Me.txtDataExtName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDataExtName.Size = New System.Drawing.Size(105, 19)
        Me.txtDataExtName.TabIndex = 2
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(8, 200)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(129, 17)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "Apply Data Extension To:"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(216, 128)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(81, 17)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Extension Type"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(8, 128)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(81, 22)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Extension Name"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(409, 25)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "New Data Extensions for OwnerID {E09C86CF-9D6E-4EF2-BCBE-4D66B6B0F754}"
        '
        'frmAddDataExtDef
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(431, 586)
        Me.Controls.Add(Me.cmdShowResponse)
        Me.Controls.Add(Me.cmdShowRequest)
        Me.Controls.Add(Me.Text1)
        Me.Controls.Add(Me.cmdClearAll)
        Me.Controls.Add(Me.cmdAllTransactions)
        Me.Controls.Add(Me.cmdAllEntities)
        Me.Controls.Add(Me.cmdAll)
        Me.Controls.Add(Me.cmdQuit)
        Me.Controls.Add(Me.cmdAdd)
        Me.Controls.Add(Me.chkSalesTaxPaymentCheck)
        Me.Controls.Add(Me.chkSalesReceipt)
        Me.Controls.Add(Me.chkReceivePayment)
        Me.Controls.Add(Me.chkPurchaseOrder)
        Me.Controls.Add(Me.chkInvoice)
        Me.Controls.Add(Me.chkJournalEntry)
        Me.Controls.Add(Me.chkInventoryAdjustment)
        Me.Controls.Add(Me.chkEstimate)
        Me.Controls.Add(Me.chkDeposit)
        Me.Controls.Add(Me.chkCreditMemo)
        Me.Controls.Add(Me.chkCreditCardCredit)
        Me.Controls.Add(Me.chkCreditCardCharge)
        Me.Controls.Add(Me.chkCheck)
        Me.Controls.Add(Me.chkCharge)
        Me.Controls.Add(Me.chkBillPaymentCreditCard)
        Me.Controls.Add(Me.chkItem)
        Me.Controls.Add(Me.chkBill)
        Me.Controls.Add(Me.chkAccount)
        Me.Controls.Add(Me.chkBillPaymentCheck)
        Me.Controls.Add(Me.chkOtherName)
        Me.Controls.Add(Me.chkEmployee)
        Me.Controls.Add(Me.chkVendor)
        Me.Controls.Add(Me.chkVendorCredit)
        Me.Controls.Add(Me.chkCustomer)
        Me.Controls.Add(Me.lstDataExtType)
        Me.Controls.Add(Me.txtDataExtName)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Location = New System.Drawing.Point(262, 124)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmAddDataExtDef"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Add Data Extension Definition"
        Me.ResumeLayout(False)

    End Sub
#End Region
End Class