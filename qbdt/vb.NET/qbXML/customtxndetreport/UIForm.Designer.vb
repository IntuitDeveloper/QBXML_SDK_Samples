<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class UIForm
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
	Public WithEvents cmdClearFilters As System.Windows.Forms.Button
	Public WithEvents RowsList As System.Windows.Forms.ListBox
	Public WithEvents showResButton As System.Windows.Forms.Button
	Public WithEvents showReqButton As System.Windows.Forms.Button
	Public WithEvents exitButton As System.Windows.Forms.Button
	Public WithEvents ColumnsList As System.Windows.Forms.ListBox
	Public WithEvents toDateText As System.Windows.Forms.TextBox
	Public WithEvents fromDateText As System.Windows.Forms.TextBox
	Public WithEvents TxnTypeFilterList As System.Windows.Forms.ListBox
	Public WithEvents ItemTypeFilterList As System.Windows.Forms.ListBox
	Public WithEvents EntityFilterList As System.Windows.Forms.ListBox
	Public WithEvents AccountFilterList As System.Windows.Forms.ListBox
	Public WithEvents goButton As System.Windows.Forms.Button
	Public WithEvents qbFile As System.Windows.Forms.TextBox
	Public WithEvents browseButton As System.Windows.Forms.Button
	Public WithEvents htmlFile As System.Windows.Forms.TextBox
	Public browseDialogOpen As System.Windows.Forms.OpenFileDialog
	Public WithEvents Label14 As System.Windows.Forms.Label
	Public WithEvents Label13 As System.Windows.Forms.Label
	Public WithEvents Label12 As System.Windows.Forms.Label
	Public WithEvents Label11 As System.Windows.Forms.Label
	Public WithEvents Label10 As System.Windows.Forms.Label
	Public WithEvents Label9 As System.Windows.Forms.Label
	Public WithEvents _Label8_0 As System.Windows.Forms.Label
	Public WithEvents Label7 As System.Windows.Forms.Label
	Public WithEvents Label6 As System.Windows.Forms.Label
	Public WithEvents Label5 As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label8 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(UIForm))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdClearFilters = New System.Windows.Forms.Button
		Me.RowsList = New System.Windows.Forms.ListBox
		Me.showResButton = New System.Windows.Forms.Button
		Me.showReqButton = New System.Windows.Forms.Button
		Me.exitButton = New System.Windows.Forms.Button
		Me.ColumnsList = New System.Windows.Forms.ListBox
		Me.toDateText = New System.Windows.Forms.TextBox
		Me.fromDateText = New System.Windows.Forms.TextBox
		Me.TxnTypeFilterList = New System.Windows.Forms.ListBox
		Me.ItemTypeFilterList = New System.Windows.Forms.ListBox
		Me.EntityFilterList = New System.Windows.Forms.ListBox
		Me.AccountFilterList = New System.Windows.Forms.ListBox
		Me.goButton = New System.Windows.Forms.Button
		Me.qbFile = New System.Windows.Forms.TextBox
		Me.browseButton = New System.Windows.Forms.Button
		Me.htmlFile = New System.Windows.Forms.TextBox
		Me.browseDialogOpen = New System.Windows.Forms.OpenFileDialog
		Me.Label14 = New System.Windows.Forms.Label
		Me.Label13 = New System.Windows.Forms.Label
		Me.Label12 = New System.Windows.Forms.Label
		Me.Label11 = New System.Windows.Forms.Label
		Me.Label10 = New System.Windows.Forms.Label
		Me.Label9 = New System.Windows.Forms.Label
		Me._Label8_0 = New System.Windows.Forms.Label
		Me.Label7 = New System.Windows.Forms.Label
		Me.Label6 = New System.Windows.Forms.Label
		Me.Label5 = New System.Windows.Forms.Label
		Me.Label4 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label8 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.Label8, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.Text = "QBSDK 2.0 -- Custom Transaction Detail Report"
		Me.ClientSize = New System.Drawing.Size(672, 511)
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
		Me.Name = "UIForm"
		Me.cmdClearFilters.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdClearFilters.Text = "Clear Filters"
		Me.cmdClearFilters.Size = New System.Drawing.Size(145, 25)
		Me.cmdClearFilters.Location = New System.Drawing.Point(480, 328)
		Me.cmdClearFilters.TabIndex = 29
		Me.cmdClearFilters.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdClearFilters.BackColor = System.Drawing.SystemColors.Control
		Me.cmdClearFilters.CausesValidation = True
		Me.cmdClearFilters.Enabled = True
		Me.cmdClearFilters.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdClearFilters.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdClearFilters.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdClearFilters.TabStop = True
		Me.cmdClearFilters.Name = "cmdClearFilters"
		Me.RowsList.Size = New System.Drawing.Size(201, 137)
		Me.RowsList.Location = New System.Drawing.Point(240, 356)
		Me.RowsList.Items.AddRange(New Object(){"Account", "BalanceSheet", "Class", "Customer", "CustomerType", "Day", "Employee", "FourWeek", "HalfMonth", "IncomeStatement", "ItemDetail", "ItemType", "Month", "Payee", "PaymentMethod", "PayrollItemDetail", "PayrollYtdDetail", "Quarter", "SalesRep", "SalesTaxCode", "ShipMethod", "TaxLine", "Terms", "TotalOnly", "TwoWeek", "Vendor", "VendorType", "Week", "Year"})
		Me.RowsList.TabIndex = 28
		Me.RowsList.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.RowsList.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.RowsList.BackColor = System.Drawing.SystemColors.Window
		Me.RowsList.CausesValidation = True
		Me.RowsList.Enabled = True
		Me.RowsList.ForeColor = System.Drawing.SystemColors.WindowText
		Me.RowsList.IntegralHeight = True
		Me.RowsList.Cursor = System.Windows.Forms.Cursors.Default
		Me.RowsList.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.RowsList.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.RowsList.Sorted = False
		Me.RowsList.TabStop = True
		Me.RowsList.Visible = True
		Me.RowsList.MultiColumn = False
		Me.RowsList.Name = "RowsList"
		Me.showResButton.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.showResButton.Text = "Show Response qbXML"
		Me.showResButton.Size = New System.Drawing.Size(145, 25)
		Me.showResButton.Location = New System.Drawing.Point(344, 108)
		Me.showResButton.TabIndex = 26
		Me.showResButton.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.showResButton.BackColor = System.Drawing.SystemColors.Control
		Me.showResButton.CausesValidation = True
		Me.showResButton.Enabled = True
		Me.showResButton.ForeColor = System.Drawing.SystemColors.ControlText
		Me.showResButton.Cursor = System.Windows.Forms.Cursors.Default
		Me.showResButton.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.showResButton.TabStop = True
		Me.showResButton.Name = "showResButton"
		Me.showReqButton.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.showReqButton.Text = "Show Request qbXML"
		Me.showReqButton.Size = New System.Drawing.Size(145, 25)
		Me.showReqButton.Location = New System.Drawing.Point(180, 108)
		Me.showReqButton.TabIndex = 25
		Me.showReqButton.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.showReqButton.BackColor = System.Drawing.SystemColors.Control
		Me.showReqButton.CausesValidation = True
		Me.showReqButton.Enabled = True
		Me.showReqButton.ForeColor = System.Drawing.SystemColors.ControlText
		Me.showReqButton.Cursor = System.Windows.Forms.Cursors.Default
		Me.showReqButton.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.showReqButton.TabStop = True
		Me.showReqButton.Name = "showReqButton"
		Me.exitButton.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.exitButton.Text = "&Exit Program"
		Me.exitButton.Size = New System.Drawing.Size(145, 25)
		Me.exitButton.Location = New System.Drawing.Point(508, 108)
		Me.exitButton.TabIndex = 12
		Me.exitButton.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.exitButton.BackColor = System.Drawing.SystemColors.Control
		Me.exitButton.CausesValidation = True
		Me.exitButton.Enabled = True
		Me.exitButton.ForeColor = System.Drawing.SystemColors.ControlText
		Me.exitButton.Cursor = System.Windows.Forms.Cursors.Default
		Me.exitButton.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.exitButton.TabStop = True
		Me.exitButton.Name = "exitButton"
		Me.ColumnsList.Size = New System.Drawing.Size(209, 137)
		Me.ColumnsList.Location = New System.Drawing.Point(16, 356)
		Me.ColumnsList.Items.AddRange(New Object(){"Account", "Aging", "Amount", "AmountDifference", "AverageColumns", "BilledDate", "BillingStatus", "CalculatedAmount", "Class", "ClearedStatus", "CostPrice", "Credit", "Date", "Debit", "DeliveryDate", "DueDate", "EstimateActive", "FOB", "IncomeSubjectToTax", "Item", "ItemDesc", "LastModifiedBy", "Memo", "ModifiedTime", "Name", "NameAccountNumber", "NameAddress", "NameCity", "NameContact", "NameEmail", "NameFax", "NamePhone", "NameState", "NameZip", "OpenBalance", "OriginalAmount", "PaidStatus", "PaidThroughDate", "PaymentMethod", "PayrollItem", "PONumber", "PrintStatus", "Quantity", "QuantityAvailable", "QuantityOnHand", "RefNumber", "RunningBalance", "SalesPrice", "SalesRep", "SalesTaxCode", "ShipDate", "ShipMethod", "SourceName", "SplitAccount", "SSNOrTaxID", "Terms", "TxnNumber", "TxnType", "UnitPrice", "ValueOnHand", "WageBase", "WageBaseTips"})
		Me.ColumnsList.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
		Me.ColumnsList.Sorted = True
		Me.ColumnsList.TabIndex = 7
		Me.ColumnsList.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ColumnsList.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.ColumnsList.BackColor = System.Drawing.SystemColors.Window
		Me.ColumnsList.CausesValidation = True
		Me.ColumnsList.Enabled = True
		Me.ColumnsList.ForeColor = System.Drawing.SystemColors.WindowText
		Me.ColumnsList.IntegralHeight = True
		Me.ColumnsList.Cursor = System.Windows.Forms.Cursors.Default
		Me.ColumnsList.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ColumnsList.TabStop = True
		Me.ColumnsList.Visible = True
		Me.ColumnsList.MultiColumn = False
		Me.ColumnsList.Name = "ColumnsList"
		Me.toDateText.AutoSize = False
		Me.toDateText.Size = New System.Drawing.Size(97, 21)
		Me.toDateText.Location = New System.Drawing.Point(520, 408)
		Me.toDateText.TabIndex = 9
		Me.toDateText.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.toDateText.AcceptsReturn = True
		Me.toDateText.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.toDateText.BackColor = System.Drawing.SystemColors.Window
		Me.toDateText.CausesValidation = True
		Me.toDateText.Enabled = True
		Me.toDateText.ForeColor = System.Drawing.SystemColors.WindowText
		Me.toDateText.HideSelection = True
		Me.toDateText.ReadOnly = False
		Me.toDateText.Maxlength = 0
		Me.toDateText.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.toDateText.MultiLine = False
		Me.toDateText.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.toDateText.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.toDateText.TabStop = True
		Me.toDateText.Visible = True
		Me.toDateText.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.toDateText.Name = "toDateText"
		Me.fromDateText.AutoSize = False
		Me.fromDateText.Size = New System.Drawing.Size(97, 21)
		Me.fromDateText.Location = New System.Drawing.Point(520, 372)
		Me.fromDateText.TabIndex = 8
		Me.fromDateText.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fromDateText.AcceptsReturn = True
		Me.fromDateText.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.fromDateText.BackColor = System.Drawing.SystemColors.Window
		Me.fromDateText.CausesValidation = True
		Me.fromDateText.Enabled = True
		Me.fromDateText.ForeColor = System.Drawing.SystemColors.WindowText
		Me.fromDateText.HideSelection = True
		Me.fromDateText.ReadOnly = False
		Me.fromDateText.Maxlength = 0
		Me.fromDateText.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.fromDateText.MultiLine = False
		Me.fromDateText.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fromDateText.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.fromDateText.TabStop = True
		Me.fromDateText.Visible = True
		Me.fromDateText.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.fromDateText.Name = "fromDateText"
		Me.TxnTypeFilterList.Size = New System.Drawing.Size(137, 137)
		Me.TxnTypeFilterList.Location = New System.Drawing.Point(505, 169)
		Me.TxnTypeFilterList.Items.AddRange(New Object(){"All", "Bill", "BillPaymentCheck", "BillPaymentCreditCard", "Charge", "Check", "CreditCardCharge", "CreditCardCredit", "CreditMemo", "Deposit", "Estimate", "InventoryAdjustment", "Invoice", "ItemReceipt", "JournalEntry", "LiabilityAdjustment", "PurchaseOrder", "ReceivePayment", "SalesOrder", "SalesReceipt", "SalesTaxPaymentCheck", "Transfer", "VendorCredit"})
		Me.TxnTypeFilterList.TabIndex = 6
		Me.TxnTypeFilterList.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.TxnTypeFilterList.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.TxnTypeFilterList.BackColor = System.Drawing.SystemColors.Window
		Me.TxnTypeFilterList.CausesValidation = True
		Me.TxnTypeFilterList.Enabled = True
		Me.TxnTypeFilterList.ForeColor = System.Drawing.SystemColors.WindowText
		Me.TxnTypeFilterList.IntegralHeight = True
		Me.TxnTypeFilterList.Cursor = System.Windows.Forms.Cursors.Default
		Me.TxnTypeFilterList.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.TxnTypeFilterList.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.TxnTypeFilterList.Sorted = False
		Me.TxnTypeFilterList.TabStop = True
		Me.TxnTypeFilterList.Visible = True
		Me.TxnTypeFilterList.MultiColumn = False
		Me.TxnTypeFilterList.Name = "TxnTypeFilterList"
		Me.ItemTypeFilterList.Size = New System.Drawing.Size(137, 137)
		Me.ItemTypeFilterList.Location = New System.Drawing.Point(344, 169)
		Me.ItemTypeFilterList.Items.AddRange(New Object(){"Assembly", "Discount", "Inventory", "InventoryAndAssembly", "NonInventory", "OtherCharge", "Payment", "Sales", "SalesTax", "Service"})
		Me.ItemTypeFilterList.Sorted = True
		Me.ItemTypeFilterList.TabIndex = 5
		Me.ItemTypeFilterList.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ItemTypeFilterList.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.ItemTypeFilterList.BackColor = System.Drawing.SystemColors.Window
		Me.ItemTypeFilterList.CausesValidation = True
		Me.ItemTypeFilterList.Enabled = True
		Me.ItemTypeFilterList.ForeColor = System.Drawing.SystemColors.WindowText
		Me.ItemTypeFilterList.IntegralHeight = True
		Me.ItemTypeFilterList.Cursor = System.Windows.Forms.Cursors.Default
		Me.ItemTypeFilterList.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.ItemTypeFilterList.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ItemTypeFilterList.TabStop = True
		Me.ItemTypeFilterList.Visible = True
		Me.ItemTypeFilterList.MultiColumn = False
		Me.ItemTypeFilterList.Name = "ItemTypeFilterList"
		Me.EntityFilterList.Size = New System.Drawing.Size(137, 124)
		Me.EntityFilterList.Location = New System.Drawing.Point(180, 169)
		Me.EntityFilterList.Items.AddRange(New Object(){"Customer", "Employee", "OtherName", "Vendor"})
		Me.EntityFilterList.TabIndex = 4
		Me.EntityFilterList.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.EntityFilterList.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.EntityFilterList.BackColor = System.Drawing.SystemColors.Window
		Me.EntityFilterList.CausesValidation = True
		Me.EntityFilterList.Enabled = True
		Me.EntityFilterList.ForeColor = System.Drawing.SystemColors.WindowText
		Me.EntityFilterList.IntegralHeight = True
		Me.EntityFilterList.Cursor = System.Windows.Forms.Cursors.Default
		Me.EntityFilterList.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.EntityFilterList.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.EntityFilterList.Sorted = False
		Me.EntityFilterList.TabStop = True
		Me.EntityFilterList.Visible = True
		Me.EntityFilterList.MultiColumn = False
		Me.EntityFilterList.Name = "EntityFilterList"
		Me.AccountFilterList.Size = New System.Drawing.Size(137, 137)
		Me.AccountFilterList.Location = New System.Drawing.Point(16, 169)
		Me.AccountFilterList.Items.AddRange(New Object(){"AccountsPayable", "AccountsReceivable", "AllowedFor1099", "APAndSalesTax", "APOrCreditCard", "ARAndAP", "Asset", "BalanceSheet", "Bank", "BankAndARAndAPAndUF", "BankAndUF", "CostOfSales", "CreditCard", "CurrentAsset", "CurrentAssetAndExpense", "CurrentLiability", "Equity", "EquityAndIncomeAndExpense", "ExpenseAndOtherExpense", "FixedAsset", "IncomeAndExpense", "IncomeAndOtherIncome", "Liability", "LiabilityAndEquity", "LongTermLiability", "NonPosting", "OrdinaryExpense", "OrdinaryIncome", "OrdinaryIncomeAndCOGS", "OrdinaryIncomeAndExpense", "OtherAsset", "OtherCurrentAsset", "OtherCurrentLiability", "OtherExpense", "OtherIncome", "OtherIncomeOrExpense"})
		Me.AccountFilterList.TabIndex = 3
		Me.AccountFilterList.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AccountFilterList.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.AccountFilterList.BackColor = System.Drawing.SystemColors.Window
		Me.AccountFilterList.CausesValidation = True
		Me.AccountFilterList.Enabled = True
		Me.AccountFilterList.ForeColor = System.Drawing.SystemColors.WindowText
		Me.AccountFilterList.IntegralHeight = True
		Me.AccountFilterList.Cursor = System.Windows.Forms.Cursors.Default
		Me.AccountFilterList.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.AccountFilterList.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.AccountFilterList.Sorted = False
		Me.AccountFilterList.TabStop = True
		Me.AccountFilterList.Visible = True
		Me.AccountFilterList.MultiColumn = False
		Me.AccountFilterList.Name = "AccountFilterList"
		Me.goButton.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.goButton.Text = "&Generate Report"
		Me.goButton.Size = New System.Drawing.Size(145, 25)
		Me.goButton.Location = New System.Drawing.Point(16, 108)
		Me.goButton.TabIndex = 11
		Me.goButton.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.goButton.BackColor = System.Drawing.SystemColors.Control
		Me.goButton.CausesValidation = True
		Me.goButton.Enabled = True
		Me.goButton.ForeColor = System.Drawing.SystemColors.ControlText
		Me.goButton.Cursor = System.Windows.Forms.Cursors.Default
		Me.goButton.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.goButton.TabStop = True
		Me.goButton.Name = "goButton"
		Me.qbFile.AutoSize = False
		Me.qbFile.Size = New System.Drawing.Size(241, 25)
		Me.qbFile.Location = New System.Drawing.Point(176, 24)
		Me.qbFile.TabIndex = 0
		Me.qbFile.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.qbFile.AcceptsReturn = True
		Me.qbFile.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.qbFile.BackColor = System.Drawing.SystemColors.Window
		Me.qbFile.CausesValidation = True
		Me.qbFile.Enabled = True
		Me.qbFile.ForeColor = System.Drawing.SystemColors.WindowText
		Me.qbFile.HideSelection = True
		Me.qbFile.ReadOnly = False
		Me.qbFile.Maxlength = 0
		Me.qbFile.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.qbFile.MultiLine = False
		Me.qbFile.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.qbFile.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.qbFile.TabStop = True
		Me.qbFile.Visible = True
		Me.qbFile.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.qbFile.Name = "qbFile"
		Me.browseButton.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.browseButton.Text = "&Browse"
		Me.browseButton.Size = New System.Drawing.Size(105, 25)
		Me.browseButton.Location = New System.Drawing.Point(424, 24)
		Me.browseButton.TabIndex = 1
		Me.browseButton.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.browseButton.BackColor = System.Drawing.SystemColors.Control
		Me.browseButton.CausesValidation = True
		Me.browseButton.Enabled = True
		Me.browseButton.ForeColor = System.Drawing.SystemColors.ControlText
		Me.browseButton.Cursor = System.Windows.Forms.Cursors.Default
		Me.browseButton.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.browseButton.TabStop = True
		Me.browseButton.Name = "browseButton"
		Me.htmlFile.AutoSize = False
		Me.htmlFile.Size = New System.Drawing.Size(281, 25)
		Me.htmlFile.Location = New System.Drawing.Point(176, 68)
		Me.htmlFile.TabIndex = 2
		Me.htmlFile.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.htmlFile.AcceptsReturn = True
		Me.htmlFile.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.htmlFile.BackColor = System.Drawing.SystemColors.Window
		Me.htmlFile.CausesValidation = True
		Me.htmlFile.Enabled = True
		Me.htmlFile.ForeColor = System.Drawing.SystemColors.WindowText
		Me.htmlFile.HideSelection = True
		Me.htmlFile.ReadOnly = False
		Me.htmlFile.Maxlength = 0
		Me.htmlFile.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.htmlFile.MultiLine = False
		Me.htmlFile.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.htmlFile.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.htmlFile.TabStop = True
		Me.htmlFile.Visible = True
		Me.htmlFile.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.htmlFile.Name = "htmlFile"
		Me.Label14.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.Label14.Text = "Summarize Data By (required)"
		Me.Label14.Size = New System.Drawing.Size(201, 33)
		Me.Label14.Location = New System.Drawing.Point(240, 328)
		Me.Label14.TabIndex = 27
		Me.Label14.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label14.BackColor = System.Drawing.SystemColors.Control
		Me.Label14.Enabled = True
		Me.Label14.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label14.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label14.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label14.UseMnemonic = True
		Me.Label14.Visible = True
		Me.Label14.AutoSize = False
		Me.Label14.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Label14.Name = "Label14"
		Me.Label13.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.Label13.Text = "(blank to use currently open file)"
		Me.Label13.Size = New System.Drawing.Size(165, 13)
		Me.Label13.Location = New System.Drawing.Point(12, 40)
		Me.Label13.TabIndex = 24
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
		Me.Label12.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label12.Text = "(yyyy-mm-dd)"
		Me.Label12.Size = New System.Drawing.Size(80, 17)
		Me.Label12.Location = New System.Drawing.Point(436, 420)
		Me.Label12.TabIndex = 23
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
		Me.Label11.Text = "(yyyy-mm-dd)"
		Me.Label11.Size = New System.Drawing.Size(80, 17)
		Me.Label11.Location = New System.Drawing.Point(436, 384)
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
		Me.Label10.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.Label10.Text = "[hold Ctrl to select more than one.]"
		Me.Label10.Size = New System.Drawing.Size(201, 17)
		Me.Label10.Location = New System.Drawing.Point(20, 344)
		Me.Label10.TabIndex = 21
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
		Me.Label9.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.Label9.Text = "Columns (required)"
		Me.Label9.Size = New System.Drawing.Size(209, 33)
		Me.Label9.Location = New System.Drawing.Point(16, 328)
		Me.Label9.TabIndex = 20
		Me.Label9.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label9.BackColor = System.Drawing.SystemColors.Control
		Me.Label9.Enabled = True
		Me.Label9.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label9.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label9.UseMnemonic = True
		Me.Label9.Visible = True
		Me.Label9.AutoSize = False
		Me.Label9.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Label9.Name = "Label9"
		Me._Label8_0.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label8_0.Text = "To Date:"
		Me._Label8_0.Size = New System.Drawing.Size(80, 17)
		Me._Label8_0.Location = New System.Drawing.Point(436, 408)
		Me._Label8_0.TabIndex = 19
		Me._Label8_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label8_0.BackColor = System.Drawing.SystemColors.Control
		Me._Label8_0.Enabled = True
		Me._Label8_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label8_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label8_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label8_0.UseMnemonic = True
		Me._Label8_0.Visible = True
		Me._Label8_0.AutoSize = False
		Me._Label8_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label8_0.Name = "_Label8_0"
		Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label7.Text = "From Date:"
		Me.Label7.Size = New System.Drawing.Size(80, 17)
		Me.Label7.Location = New System.Drawing.Point(436, 372)
		Me.Label7.TabIndex = 18
		Me.Label7.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
		Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.Label6.Text = "Transaction Type Filter"
		Me.Label6.Size = New System.Drawing.Size(137, 17)
		Me.Label6.Location = New System.Drawing.Point(504, 152)
		Me.Label6.TabIndex = 17
		Me.Label6.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label6.BackColor = System.Drawing.SystemColors.Control
		Me.Label6.Enabled = True
		Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label6.UseMnemonic = True
		Me.Label6.Visible = True
		Me.Label6.AutoSize = False
		Me.Label6.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Label6.Name = "Label6"
		Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.Label5.Text = "Item Type Filter"
		Me.Label5.Size = New System.Drawing.Size(137, 17)
		Me.Label5.Location = New System.Drawing.Point(344, 152)
		Me.Label5.TabIndex = 16
		Me.Label5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label5.BackColor = System.Drawing.SystemColors.Control
		Me.Label5.Enabled = True
		Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label5.UseMnemonic = True
		Me.Label5.Visible = True
		Me.Label5.AutoSize = False
		Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Label5.Name = "Label5"
		Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.Label4.Text = "Entity Type Filter"
		Me.Label4.Size = New System.Drawing.Size(137, 17)
		Me.Label4.Location = New System.Drawing.Point(180, 152)
		Me.Label4.TabIndex = 15
		Me.Label4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label4.BackColor = System.Drawing.SystemColors.Control
		Me.Label4.Enabled = True
		Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label4.UseMnemonic = True
		Me.Label4.Visible = True
		Me.Label4.AutoSize = False
		Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Label4.Name = "Label4"
		Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.Label3.Text = "Account Type Filter"
		Me.Label3.Size = New System.Drawing.Size(137, 17)
		Me.Label3.Location = New System.Drawing.Point(16, 152)
		Me.Label3.TabIndex = 14
		Me.Label3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.BackColor = System.Drawing.SystemColors.Control
		Me.Label3.Enabled = True
		Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label3.UseMnemonic = True
		Me.Label3.Visible = True
		Me.Label3.AutoSize = False
		Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Label3.Name = "Label3"
		Me.Label1.Text = "Select QuickBooks Company File:"
		Me.Label1.Size = New System.Drawing.Size(165, 21)
		Me.Label1.Location = New System.Drawing.Point(12, 28)
		Me.Label1.TabIndex = 13
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
		Me.Label2.Text = "Select Output Html File:"
		Me.Label2.Size = New System.Drawing.Size(165, 21)
		Me.Label2.Location = New System.Drawing.Point(60, 72)
		Me.Label2.TabIndex = 10
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
		Me.Controls.Add(cmdClearFilters)
		Me.Controls.Add(RowsList)
		Me.Controls.Add(showResButton)
		Me.Controls.Add(showReqButton)
		Me.Controls.Add(exitButton)
		Me.Controls.Add(ColumnsList)
		Me.Controls.Add(toDateText)
		Me.Controls.Add(fromDateText)
		Me.Controls.Add(TxnTypeFilterList)
		Me.Controls.Add(ItemTypeFilterList)
		Me.Controls.Add(EntityFilterList)
		Me.Controls.Add(AccountFilterList)
		Me.Controls.Add(goButton)
		Me.Controls.Add(qbFile)
		Me.Controls.Add(browseButton)
		Me.Controls.Add(htmlFile)
		Me.Controls.Add(Label14)
		Me.Controls.Add(Label13)
		Me.Controls.Add(Label12)
		Me.Controls.Add(Label11)
		Me.Controls.Add(Label10)
		Me.Controls.Add(Label9)
		Me.Controls.Add(_Label8_0)
		Me.Controls.Add(Label7)
		Me.Controls.Add(Label6)
		Me.Controls.Add(Label5)
		Me.Controls.Add(Label4)
		Me.Controls.Add(Label3)
		Me.Controls.Add(Label1)
		Me.Controls.Add(Label2)
		Me.Label8.SetIndex(_Label8_0, CType(0, Short))
		CType(Me.Label8, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class