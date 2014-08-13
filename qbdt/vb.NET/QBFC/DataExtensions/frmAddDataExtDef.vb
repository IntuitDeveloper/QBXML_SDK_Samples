Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmAddDataExtDef
	Inherits System.Windows.Forms.Form
#Region "Windows Form Designer generated code "
	Public Sub New()
		MyBase.New()
		If m_vb6FormDefInstance Is Nothing Then
			If m_InitializingDefInstance Then
				m_vb6FormDefInstance = Me
			Else
				Try 
					'For the start-up form, the first instance created is the default instance.
					If System.Reflection.Assembly.GetExecutingAssembly.EntryPoint.DeclaringType Is Me.GetType Then
						m_vb6FormDefInstance = Me
					End If
				Catch
				End Try
			End If
		End If
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
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
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
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
        'Text1
        '
        Me.Text1.AcceptsReturn = True
        Me.Text1.AutoSize = False
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
        Me.Text1.Text = "1) Enter a Data Extension Name" & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10) & "2) Choose a Data Extension Type" & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10) & "3) Select the li" & _
        "st items and/or transactions you wish to associate with " & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10) & "the Data Extension" & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10) & "4)" & _
        " Click on the Add Data Extension Definition button" & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10)
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
        '
        'cmdQuit
        '
        Me.cmdQuit.BackColor = System.Drawing.SystemColors.Control
        Me.cmdQuit.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdQuit.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdQuit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdQuit.Location = New System.Drawing.Point(272, 528)
        Me.cmdQuit.Name = "cmdQuit"
        Me.cmdQuit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdQuit.Size = New System.Drawing.Size(97, 41)
        Me.cmdQuit.TabIndex = 31
        Me.cmdQuit.Text = "Close Window"
        '
        'cmdAdd
        '
        Me.cmdAdd.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAdd.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAdd.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAdd.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAdd.Location = New System.Drawing.Point(56, 528)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAdd.Size = New System.Drawing.Size(193, 41)
        Me.cmdAdd.TabIndex = 30
        Me.cmdAdd.Text = "Add Data Extension Definition"
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
        Me.chkSalesTaxPaymentCheck.Size = New System.Drawing.Size(160, 17)
        Me.chkSalesTaxPaymentCheck.TabIndex = 29
        Me.chkSalesTaxPaymentCheck.Text = "Sales Tax Payment Check"
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
        Me.chkBillPaymentCreditCard.Size = New System.Drawing.Size(144, 17)
        Me.chkBillPaymentCreditCard.TabIndex = 15
        Me.chkBillPaymentCreditCard.Text = "Bill Payment Credit Card"
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
        Me.chkItem.Size = New System.Drawing.Size(48, 17)
        Me.chkItem.TabIndex = 14
        Me.chkItem.Text = "Item"
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
        Me.chkAccount.Size = New System.Drawing.Size(128, 17)
        Me.chkAccount.TabIndex = 12
        Me.chkAccount.Text = "Account"
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
        Me.chkBillPaymentCheck.Size = New System.Drawing.Size(120, 17)
        Me.chkBillPaymentCheck.TabIndex = 11
        Me.chkBillPaymentCheck.Text = "Bill Payment Check"
        '
        'chkOtherName
        '
        Me.chkOtherName.BackColor = System.Drawing.SystemColors.Control
        Me.chkOtherName.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkOtherName.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkOtherName.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkOtherName.Location = New System.Drawing.Point(240, 240)
        Me.chkOtherName.Name = "chkOtherName"
        Me.chkOtherName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkOtherName.Size = New System.Drawing.Size(88, 17)
        Me.chkOtherName.TabIndex = 10
        Me.chkOtherName.Text = "Other Name"
        '
        'chkEmployee
        '
        Me.chkEmployee.BackColor = System.Drawing.SystemColors.Control
        Me.chkEmployee.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkEmployee.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkEmployee.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkEmployee.Location = New System.Drawing.Point(160, 240)
        Me.chkEmployee.Name = "chkEmployee"
        Me.chkEmployee.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkEmployee.Size = New System.Drawing.Size(73, 17)
        Me.chkEmployee.TabIndex = 9
        Me.chkEmployee.Text = "Employee"
        '
        'chkVendor
        '
        Me.chkVendor.BackColor = System.Drawing.SystemColors.Control
        Me.chkVendor.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkVendor.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkVendor.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkVendor.Location = New System.Drawing.Point(81, 240)
        Me.chkVendor.Name = "chkVendor"
        Me.chkVendor.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkVendor.Size = New System.Drawing.Size(63, 17)
        Me.chkVendor.TabIndex = 8
        Me.chkVendor.Text = "Vendor"
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
        '
        'lstDataExtType
        '
        Me.lstDataExtType.BackColor = System.Drawing.SystemColors.Window
        Me.lstDataExtType.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstDataExtType.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstDataExtType.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstDataExtType.ItemHeight = 14
        Me.lstDataExtType.Location = New System.Drawing.Point(312, 128)
        Me.lstDataExtType.Name = "lstDataExtType"
        Me.lstDataExtType.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstDataExtType.Size = New System.Drawing.Size(113, 46)
        Me.lstDataExtType.TabIndex = 4
        '
        'txtDataExtName
        '
        Me.txtDataExtName.AcceptsReturn = True
        Me.txtDataExtName.AutoSize = False
        Me.txtDataExtName.BackColor = System.Drawing.SystemColors.Window
        Me.txtDataExtName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDataExtName.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDataExtName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDataExtName.Location = New System.Drawing.Point(104, 128)
        Me.txtDataExtName.MaxLength = 0
        Me.txtDataExtName.Name = "txtDataExtName"
        Me.txtDataExtName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDataExtName.Size = New System.Drawing.Size(105, 19)
        Me.txtDataExtName.TabIndex = 2
        Me.txtDataExtName.Text = ""
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
        Me.Label3.Location = New System.Drawing.Point(224, 128)
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
        Me.Label2.Size = New System.Drawing.Size(88, 17)
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
        Me.Label1.Size = New System.Drawing.Size(432, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "New Data Extensions for OwnerID {E09C86CF-9D6E-4EF2-BCBE-4D66B6B0F754}"
        '
        'frmAddDataExtDef
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(450, 586)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Text1, Me.cmdClearAll, Me.cmdAllTransactions, Me.cmdAllEntities, Me.cmdAll, Me.cmdQuit, Me.cmdAdd, Me.chkSalesTaxPaymentCheck, Me.chkSalesReceipt, Me.chkReceivePayment, Me.chkPurchaseOrder, Me.chkInvoice, Me.chkJournalEntry, Me.chkInventoryAdjustment, Me.chkEstimate, Me.chkDeposit, Me.chkCreditMemo, Me.chkCreditCardCredit, Me.chkCreditCardCharge, Me.chkCheck, Me.chkCharge, Me.chkBillPaymentCreditCard, Me.chkItem, Me.chkBill, Me.chkAccount, Me.chkBillPaymentCheck, Me.chkOtherName, Me.chkEmployee, Me.chkVendor, Me.chkVendorCredit, Me.chkCustomer, Me.lstDataExtType, Me.txtDataExtName, Me.Label4, Me.Label3, Me.Label2, Me.Label1})
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Location = New System.Drawing.Point(262, 124)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmAddDataExtDef"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Add Data Extension Definition"
        Me.ResumeLayout(False)

    End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As frmAddDataExtDef
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As frmAddDataExtDef
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New frmAddDataExtDef()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	'----------------------------------------------------------
	' Form: frmAddDataExtDef
	'
	' Description: This form allows the user to type in the name of a
	'              new data extension definition, select the type for it,
	'              select the items and transactions to associate it with
	'              and the add it to the currently open QuickBooks company
	'              file.
	'
	' Copyright © 2002-2013 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	'----------------------------------------------------------
	
	Private Sub cmdAdd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAdd.Click
		
		Dim strObjects As String
		
		If txtDataExtName.Text = "" Then
			MsgBox("You must enter a Data Extension Name")
			Exit Sub
		End If
		
		If lstDataExtType.SelectedIndex < 0 Then
			MsgBox("You must select a Data Extension Type")
			Exit Sub
		End If
		
		strObjects = ""
		If chkCustomer.CheckState = 1 Then strObjects = strObjects & " Customer"
		If chkVendor.CheckState = 1 Then strObjects = strObjects & " Vendor"
		If chkEmployee.CheckState = 1 Then strObjects = strObjects & " Employee"
		If chkOtherName.CheckState = 1 Then strObjects = strObjects & " OtherName"
		If chkItem.CheckState = 1 Then strObjects = strObjects & " Item"
		If chkAccount.CheckState = 1 Then strObjects = strObjects & " Account"
		If chkBill.CheckState = 1 Then strObjects = strObjects & " Bill"
		If chkBillPaymentCheck.CheckState = 1 Then strObjects = strObjects & " BillPaymentCheck"
		If chkBillPaymentCreditCard.CheckState = 1 Then strObjects = strObjects & " BillPaymentCreditCard"
		If chkCharge.CheckState = 1 Then strObjects = strObjects & " Charge"
		If chkCheck.CheckState = 1 Then strObjects = strObjects & " Check"
		If chkCreditCardCharge.CheckState = 1 Then strObjects = strObjects & " CreditCardCharge"
		If chkCreditCardCredit.CheckState = 1 Then strObjects = strObjects & " CreditCardCredit"
		If chkCreditMemo.CheckState = 1 Then strObjects = strObjects & " CreditMemo"
		If chkDeposit.CheckState = 1 Then strObjects = strObjects & " Deposit"
		If chkEstimate.CheckState = 1 Then strObjects = strObjects & " Estimate"
		If chkInventoryAdjustment.CheckState = 1 Then strObjects = strObjects & " InventoryAdjustment"
		If chkInvoice.CheckState = 1 Then strObjects = strObjects & " Invoice"
		If chkJournalEntry.CheckState = 1 Then strObjects = strObjects & " JournalEntry"
		If chkPurchaseOrder.CheckState = 1 Then strObjects = strObjects & " PurchaseOrder"
		If chkReceivePayment.CheckState = 1 Then strObjects = strObjects & " ReceivePayment"
		If chkSalesReceipt.CheckState = 1 Then strObjects = strObjects & " SalesReceipt"
		If chkSalesTaxPaymentCheck.CheckState = 1 Then strObjects = strObjects & " SalesTaxPaymentCheck"
		If chkVendorCredit.CheckState = 1 Then strObjects = strObjects & " VendorCredit"
		
		If strObjects = "" Then
			MsgBox("You need to pick at least one target for the data extension")
			Exit Sub
		End If
		
		'Remove the leading space from strObjects
		strObjects = VB.Right(strObjects, Len(strObjects) - 1)
		
		AddDataExtDef((txtDataExtName.Text), VB6.GetItemString(lstDataExtType, lstDataExtType.SelectedIndex), strObjects)
		
		frmAddDataExtDef.DefInstance.Close()
	End Sub
	
	Private Sub cmdAll_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAll.Click
		chkCustomer.CheckState = System.Windows.Forms.CheckState.Checked
		chkVendor.CheckState = System.Windows.Forms.CheckState.Checked
		chkEmployee.CheckState = System.Windows.Forms.CheckState.Checked
		chkOtherName.CheckState = System.Windows.Forms.CheckState.Checked
		chkItem.CheckState = System.Windows.Forms.CheckState.Checked
		chkAccount.CheckState = System.Windows.Forms.CheckState.Checked
		chkBill.CheckState = System.Windows.Forms.CheckState.Checked
		chkBillPaymentCheck.CheckState = System.Windows.Forms.CheckState.Checked
		chkBillPaymentCreditCard.CheckState = System.Windows.Forms.CheckState.Checked
		chkCharge.CheckState = System.Windows.Forms.CheckState.Checked
		chkCheck.CheckState = System.Windows.Forms.CheckState.Checked
		chkCreditCardCharge.CheckState = System.Windows.Forms.CheckState.Checked
		chkCreditCardCredit.CheckState = System.Windows.Forms.CheckState.Checked
		chkCreditMemo.CheckState = System.Windows.Forms.CheckState.Checked
		chkDeposit.CheckState = System.Windows.Forms.CheckState.Checked
		chkEstimate.CheckState = System.Windows.Forms.CheckState.Checked
		chkInventoryAdjustment.CheckState = System.Windows.Forms.CheckState.Checked
		chkInvoice.CheckState = System.Windows.Forms.CheckState.Checked
		chkJournalEntry.CheckState = System.Windows.Forms.CheckState.Checked
		chkPurchaseOrder.CheckState = System.Windows.Forms.CheckState.Checked
		chkReceivePayment.CheckState = System.Windows.Forms.CheckState.Checked
		chkSalesReceipt.CheckState = System.Windows.Forms.CheckState.Checked
		chkSalesTaxPaymentCheck.CheckState = System.Windows.Forms.CheckState.Checked
		chkVendorCredit.CheckState = System.Windows.Forms.CheckState.Checked
	End Sub
	
	Private Sub cmdAllEntities_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAllEntities.Click
		chkCustomer.CheckState = System.Windows.Forms.CheckState.Checked
		chkVendor.CheckState = System.Windows.Forms.CheckState.Checked
		chkEmployee.CheckState = System.Windows.Forms.CheckState.Checked
		chkOtherName.CheckState = System.Windows.Forms.CheckState.Checked
	End Sub
	
	Private Sub cmdAllTransactions_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAllTransactions.Click
		chkBill.CheckState = System.Windows.Forms.CheckState.Checked
		chkBillPaymentCheck.CheckState = System.Windows.Forms.CheckState.Checked
		chkBillPaymentCreditCard.CheckState = System.Windows.Forms.CheckState.Checked
		chkCharge.CheckState = System.Windows.Forms.CheckState.Checked
		chkCheck.CheckState = System.Windows.Forms.CheckState.Checked
		chkCreditCardCharge.CheckState = System.Windows.Forms.CheckState.Checked
		chkCreditCardCredit.CheckState = System.Windows.Forms.CheckState.Checked
		chkCreditMemo.CheckState = System.Windows.Forms.CheckState.Checked
		chkDeposit.CheckState = System.Windows.Forms.CheckState.Checked
		chkEstimate.CheckState = System.Windows.Forms.CheckState.Checked
		chkInventoryAdjustment.CheckState = System.Windows.Forms.CheckState.Checked
		chkInvoice.CheckState = System.Windows.Forms.CheckState.Checked
		chkJournalEntry.CheckState = System.Windows.Forms.CheckState.Checked
		chkPurchaseOrder.CheckState = System.Windows.Forms.CheckState.Checked
		chkReceivePayment.CheckState = System.Windows.Forms.CheckState.Checked
		chkSalesReceipt.CheckState = System.Windows.Forms.CheckState.Checked
		chkSalesTaxPaymentCheck.CheckState = System.Windows.Forms.CheckState.Checked
		chkVendorCredit.CheckState = System.Windows.Forms.CheckState.Checked
	End Sub
	
	Private Sub cmdClearAll_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClearAll.Click
		chkCustomer.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkVendor.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkEmployee.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkOtherName.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkItem.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkAccount.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkBill.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkBillPaymentCheck.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkBillPaymentCreditCard.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkCharge.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkCheck.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkCreditCardCharge.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkCreditCardCredit.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkCreditMemo.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkDeposit.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkEstimate.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkInventoryAdjustment.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkInvoice.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkJournalEntry.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkPurchaseOrder.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkReceivePayment.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkSalesReceipt.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkSalesTaxPaymentCheck.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkVendorCredit.CheckState = System.Windows.Forms.CheckState.Unchecked
	End Sub
	
	Private Sub cmdQuit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdQuit.Click
		frmAddDataExtDef.DefInstance.Close()
	End Sub
	
	Private Sub frmAddDataExtDef_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'Set up the data extension type list box with the accepted types
		lstDataExtType.Items.Clear()
        lstDataExtType.Items.Add("INTTYPE")
        lstDataExtType.Items.Add("AMTTYPE")
        lstDataExtType.Items.Add("PRICETYPE")
        lstDataExtType.Items.Add("QUANTYPE")
        lstDataExtType.Items.Add("PERCENTTYPE")
        lstDataExtType.Items.Add("DATETIMETYPE")
        lstDataExtType.Items.Add("STR255TYPE")
        lstDataExtType.Items.Add("STR1024TYPE")
	End Sub
End Class