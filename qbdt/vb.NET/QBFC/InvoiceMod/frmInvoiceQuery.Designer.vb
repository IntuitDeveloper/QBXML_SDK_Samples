<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmInvoiceQuery
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
	Public WithEvents chkShow As System.Windows.Forms.CheckBox
	Public WithEvents cmdQuery As System.Windows.Forms.Button
	Public WithEvents cmbAccounts As System.Windows.Forms.ComboBox
	Public WithEvents cmbCustomers As System.Windows.Forms.ComboBox
	Public WithEvents cmbRefNumberCriteria As System.Windows.Forms.ComboBox
	Public WithEvents txtToRefNumber As System.Windows.Forms.TextBox
	Public WithEvents txtFromRefNumber As System.Windows.Forms.TextBox
	Public WithEvents txtRefNumberPart As System.Windows.Forms.TextBox
	Public WithEvents optRefNumberFilter As System.Windows.Forms.RadioButton
	Public WithEvents optRefNumberRange As System.Windows.Forms.RadioButton
	Public WithEvents lblToRefNumber As System.Windows.Forms.Label
	Public WithEvents lblFromRefNumber As System.Windows.Forms.Label
	Public WithEvents Frame4 As System.Windows.Forms.Panel
	Public WithEvents cmdClearRefNumber As System.Windows.Forms.Button
	Public WithEvents frameRefNumberArea As System.Windows.Forms.GroupBox
	Public WithEvents txtToTime As System.Windows.Forms.TextBox
	Public WithEvents txtFromTime As System.Windows.Forms.TextBox
	Public WithEvents txtToDate As System.Windows.Forms.TextBox
	Public WithEvents txtFromDate As System.Windows.Forms.TextBox
	Public WithEvents optFromAM As System.Windows.Forms.RadioButton
	Public WithEvents optFromPM As System.Windows.Forms.RadioButton
	Public WithEvents frameFromAMPM As System.Windows.Forms.Panel
	Public WithEvents optToPM As System.Windows.Forms.RadioButton
	Public WithEvents optToAM As System.Windows.Forms.RadioButton
	Public WithEvents frameToAMPM As System.Windows.Forms.Panel
	Public WithEvents optModified As System.Windows.Forms.RadioButton
	Public WithEvents optTxnDate As System.Windows.Forms.RadioButton
	Public WithEvents optMacroDate As System.Windows.Forms.RadioButton
	Public WithEvents frameDateCriteria As System.Windows.Forms.Panel
	Public WithEvents cmdClearDateSelection As System.Windows.Forms.Button
	Public WithEvents optThis As System.Windows.Forms.RadioButton
	Public WithEvents optLast As System.Windows.Forms.RadioButton
	Public WithEvents frmThisLast As System.Windows.Forms.GroupBox
	Public WithEvents optWeek As System.Windows.Forms.RadioButton
	Public WithEvents optMonth As System.Windows.Forms.RadioButton
	Public WithEvents optFiscalQuarter As System.Windows.Forms.RadioButton
	Public WithEvents optFiscalYear As System.Windows.Forms.RadioButton
	Public WithEvents frameMacroPeriod As System.Windows.Forms.GroupBox
	Public WithEvents lblToTime As System.Windows.Forms.Label
	Public WithEvents lblFromTime As System.Windows.Forms.Label
	Public WithEvents lblAfter As System.Windows.Forms.Label
	Public WithEvents lblBefore As System.Windows.Forms.Label
	Public WithEvents Frame3 As System.Windows.Forms.GroupBox
	Public WithEvents optPaidOnly As System.Windows.Forms.RadioButton
	Public WithEvents optUnpaidOnly As System.Windows.Forms.RadioButton
	Public WithEvents optAll As System.Windows.Forms.RadioButton
	Public WithEvents Frame2 As System.Windows.Forms.Panel
	Public WithEvents chkAccountChildren As System.Windows.Forms.CheckBox
	Public WithEvents chkCustomerChildren As System.Windows.Forms.CheckBox
	Public WithEvents txtInvoiceNumber As System.Windows.Forms.TextBox
	Public WithEvents cmdQuit As System.Windows.Forms.Button
	Public WithEvents cmdModifyInvoice As System.Windows.Forms.Button
	Public WithEvents lstInvoices As System.Windows.Forms.ListBox
	Public WithEvents Label8 As System.Windows.Forms.Label
	Public WithEvents Label7 As System.Windows.Forms.Label
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
        Me.chkShow = New System.Windows.Forms.CheckBox()
        Me.cmdQuery = New System.Windows.Forms.Button()
        Me.cmbAccounts = New System.Windows.Forms.ComboBox()
        Me.cmbCustomers = New System.Windows.Forms.ComboBox()
        Me.frameRefNumberArea = New System.Windows.Forms.GroupBox()
        Me.Frame4 = New System.Windows.Forms.Panel()
        Me.cmbRefNumberCriteria = New System.Windows.Forms.ComboBox()
        Me.txtToRefNumber = New System.Windows.Forms.TextBox()
        Me.txtFromRefNumber = New System.Windows.Forms.TextBox()
        Me.txtRefNumberPart = New System.Windows.Forms.TextBox()
        Me.optRefNumberFilter = New System.Windows.Forms.RadioButton()
        Me.optRefNumberRange = New System.Windows.Forms.RadioButton()
        Me.lblToRefNumber = New System.Windows.Forms.Label()
        Me.lblFromRefNumber = New System.Windows.Forms.Label()
        Me.cmdClearRefNumber = New System.Windows.Forms.Button()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.txtToTime = New System.Windows.Forms.TextBox()
        Me.txtFromTime = New System.Windows.Forms.TextBox()
        Me.txtToDate = New System.Windows.Forms.TextBox()
        Me.txtFromDate = New System.Windows.Forms.TextBox()
        Me.frameFromAMPM = New System.Windows.Forms.Panel()
        Me.optFromAM = New System.Windows.Forms.RadioButton()
        Me.optFromPM = New System.Windows.Forms.RadioButton()
        Me.frameToAMPM = New System.Windows.Forms.Panel()
        Me.optToPM = New System.Windows.Forms.RadioButton()
        Me.optToAM = New System.Windows.Forms.RadioButton()
        Me.frameDateCriteria = New System.Windows.Forms.Panel()
        Me.optModified = New System.Windows.Forms.RadioButton()
        Me.optTxnDate = New System.Windows.Forms.RadioButton()
        Me.optMacroDate = New System.Windows.Forms.RadioButton()
        Me.cmdClearDateSelection = New System.Windows.Forms.Button()
        Me.frmThisLast = New System.Windows.Forms.GroupBox()
        Me.optThis = New System.Windows.Forms.RadioButton()
        Me.optLast = New System.Windows.Forms.RadioButton()
        Me.frameMacroPeriod = New System.Windows.Forms.GroupBox()
        Me.optWeek = New System.Windows.Forms.RadioButton()
        Me.optMonth = New System.Windows.Forms.RadioButton()
        Me.optFiscalQuarter = New System.Windows.Forms.RadioButton()
        Me.optFiscalYear = New System.Windows.Forms.RadioButton()
        Me.lblToTime = New System.Windows.Forms.Label()
        Me.lblFromTime = New System.Windows.Forms.Label()
        Me.lblAfter = New System.Windows.Forms.Label()
        Me.lblBefore = New System.Windows.Forms.Label()
        Me.Frame2 = New System.Windows.Forms.Panel()
        Me.optPaidOnly = New System.Windows.Forms.RadioButton()
        Me.optUnpaidOnly = New System.Windows.Forms.RadioButton()
        Me.optAll = New System.Windows.Forms.RadioButton()
        Me.chkAccountChildren = New System.Windows.Forms.CheckBox()
        Me.chkCustomerChildren = New System.Windows.Forms.CheckBox()
        Me.txtInvoiceNumber = New System.Windows.Forms.TextBox()
        Me.cmdQuit = New System.Windows.Forms.Button()
        Me.cmdModifyInvoice = New System.Windows.Forms.Button()
        Me.lstInvoices = New System.Windows.Forms.ListBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.frameRefNumberArea.SuspendLayout()
        Me.Frame4.SuspendLayout()
        Me.Frame3.SuspendLayout()
        Me.frameFromAMPM.SuspendLayout()
        Me.frameToAMPM.SuspendLayout()
        Me.frameDateCriteria.SuspendLayout()
        Me.frmThisLast.SuspendLayout()
        Me.frameMacroPeriod.SuspendLayout()
        Me.Frame2.SuspendLayout()
        Me.SuspendLayout()
        '
        'chkShow
        '
        Me.chkShow.BackColor = System.Drawing.SystemColors.Control
        Me.chkShow.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkShow.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkShow.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkShow.Location = New System.Drawing.Point(16, 552)
        Me.chkShow.Name = "chkShow"
        Me.chkShow.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkShow.Size = New System.Drawing.Size(113, 17)
        Me.chkShow.TabIndex = 58
        Me.chkShow.Text = "Show invoice query"
        Me.chkShow.UseVisualStyleBackColor = False
        '
        'cmdQuery
        '
        Me.cmdQuery.BackColor = System.Drawing.SystemColors.Control
        Me.cmdQuery.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdQuery.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdQuery.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdQuery.Location = New System.Drawing.Point(144, 544)
        Me.cmdQuery.Name = "cmdQuery"
        Me.cmdQuery.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdQuery.Size = New System.Drawing.Size(105, 41)
        Me.cmdQuery.TabIndex = 27
        Me.cmdQuery.Text = "Query for Invoices"
        Me.cmdQuery.UseVisualStyleBackColor = False
        '
        'cmbAccounts
        '
        Me.cmbAccounts.BackColor = System.Drawing.SystemColors.Window
        Me.cmbAccounts.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbAccounts.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbAccounts.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbAccounts.Location = New System.Drawing.Point(96, 200)
        Me.cmbAccounts.Name = "cmbAccounts"
        Me.cmbAccounts.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbAccounts.Size = New System.Drawing.Size(321, 22)
        Me.cmbAccounts.Sorted = True
        Me.cmbAccounts.TabIndex = 16
        '
        'cmbCustomers
        '
        Me.cmbCustomers.BackColor = System.Drawing.SystemColors.Window
        Me.cmbCustomers.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbCustomers.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbCustomers.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbCustomers.Location = New System.Drawing.Point(96, 168)
        Me.cmbCustomers.Name = "cmbCustomers"
        Me.cmbCustomers.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbCustomers.Size = New System.Drawing.Size(321, 22)
        Me.cmbCustomers.TabIndex = 14
        '
        'frameRefNumberArea
        '
        Me.frameRefNumberArea.BackColor = System.Drawing.SystemColors.Control
        Me.frameRefNumberArea.Controls.Add(Me.Frame4)
        Me.frameRefNumberArea.Controls.Add(Me.cmdClearRefNumber)
        Me.frameRefNumberArea.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.frameRefNumberArea.ForeColor = System.Drawing.SystemColors.ControlText
        Me.frameRefNumberArea.Location = New System.Drawing.Point(8, 224)
        Me.frameRefNumberArea.Name = "frameRefNumberArea"
        Me.frameRefNumberArea.Padding = New System.Windows.Forms.Padding(0)
        Me.frameRefNumberArea.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.frameRefNumberArea.Size = New System.Drawing.Size(489, 73)
        Me.frameRefNumberArea.TabIndex = 47
        Me.frameRefNumberArea.TabStop = False
        '
        'Frame4
        '
        Me.Frame4.BackColor = System.Drawing.SystemColors.Control
        Me.Frame4.Controls.Add(Me.cmbRefNumberCriteria)
        Me.Frame4.Controls.Add(Me.txtToRefNumber)
        Me.Frame4.Controls.Add(Me.txtFromRefNumber)
        Me.Frame4.Controls.Add(Me.txtRefNumberPart)
        Me.Frame4.Controls.Add(Me.optRefNumberFilter)
        Me.Frame4.Controls.Add(Me.optRefNumberRange)
        Me.Frame4.Controls.Add(Me.lblToRefNumber)
        Me.Frame4.Controls.Add(Me.lblFromRefNumber)
        Me.Frame4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Frame4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame4.Location = New System.Drawing.Point(8, 8)
        Me.Frame4.Name = "Frame4"
        Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame4.Size = New System.Drawing.Size(377, 57)
        Me.Frame4.TabIndex = 49
        Me.Frame4.Text = "Frame1"
        '
        'cmbRefNumberCriteria
        '
        Me.cmbRefNumberCriteria.BackColor = System.Drawing.SystemColors.Menu
        Me.cmbRefNumberCriteria.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbRefNumberCriteria.Enabled = False
        Me.cmbRefNumberCriteria.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbRefNumberCriteria.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbRefNumberCriteria.Items.AddRange(New Object() {"Starts With", "Contains", "Ends With"})
        Me.cmbRefNumberCriteria.Location = New System.Drawing.Point(72, 32)
        Me.cmbRefNumberCriteria.Name = "cmbRefNumberCriteria"
        Me.cmbRefNumberCriteria.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbRefNumberCriteria.Size = New System.Drawing.Size(89, 22)
        Me.cmbRefNumberCriteria.TabIndex = 22
        Me.cmbRefNumberCriteria.Text = "Starts With"
        '
        'txtToRefNumber
        '
        Me.txtToRefNumber.AcceptsReturn = True
        Me.txtToRefNumber.BackColor = System.Drawing.SystemColors.Menu
        Me.txtToRefNumber.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtToRefNumber.Enabled = False
        Me.txtToRefNumber.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtToRefNumber.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtToRefNumber.Location = New System.Drawing.Point(264, 8)
        Me.txtToRefNumber.MaxLength = 0
        Me.txtToRefNumber.Name = "txtToRefNumber"
        Me.txtToRefNumber.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtToRefNumber.Size = New System.Drawing.Size(105, 20)
        Me.txtToRefNumber.TabIndex = 20
        '
        'txtFromRefNumber
        '
        Me.txtFromRefNumber.AcceptsReturn = True
        Me.txtFromRefNumber.BackColor = System.Drawing.SystemColors.Menu
        Me.txtFromRefNumber.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFromRefNumber.Enabled = False
        Me.txtFromRefNumber.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFromRefNumber.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFromRefNumber.Location = New System.Drawing.Point(136, 8)
        Me.txtFromRefNumber.MaxLength = 0
        Me.txtFromRefNumber.Name = "txtFromRefNumber"
        Me.txtFromRefNumber.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFromRefNumber.Size = New System.Drawing.Size(105, 20)
        Me.txtFromRefNumber.TabIndex = 19
        '
        'txtRefNumberPart
        '
        Me.txtRefNumberPart.AcceptsReturn = True
        Me.txtRefNumberPart.BackColor = System.Drawing.SystemColors.Menu
        Me.txtRefNumberPart.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRefNumberPart.Enabled = False
        Me.txtRefNumberPart.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRefNumberPart.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRefNumberPart.Location = New System.Drawing.Point(176, 32)
        Me.txtRefNumberPart.MaxLength = 0
        Me.txtRefNumberPart.Name = "txtRefNumberPart"
        Me.txtRefNumberPart.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRefNumberPart.Size = New System.Drawing.Size(105, 20)
        Me.txtRefNumberPart.TabIndex = 23
        '
        'optRefNumberFilter
        '
        Me.optRefNumberFilter.BackColor = System.Drawing.SystemColors.Control
        Me.optRefNumberFilter.Cursor = System.Windows.Forms.Cursors.Default
        Me.optRefNumberFilter.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optRefNumberFilter.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optRefNumberFilter.Location = New System.Drawing.Point(0, 32)
        Me.optRefNumberFilter.Name = "optRefNumberFilter"
        Me.optRefNumberFilter.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optRefNumberFilter.Size = New System.Drawing.Size(69, 20)
        Me.optRefNumberFilter.TabIndex = 21
        Me.optRefNumberFilter.TabStop = True
        Me.optRefNumberFilter.Text = "Invoice #"
        Me.optRefNumberFilter.UseVisualStyleBackColor = False
        '
        'optRefNumberRange
        '
        Me.optRefNumberRange.BackColor = System.Drawing.SystemColors.Control
        Me.optRefNumberRange.Cursor = System.Windows.Forms.Cursors.Default
        Me.optRefNumberRange.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optRefNumberRange.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optRefNumberRange.Location = New System.Drawing.Point(0, 8)
        Me.optRefNumberRange.Name = "optRefNumberRange"
        Me.optRefNumberRange.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optRefNumberRange.Size = New System.Drawing.Size(105, 17)
        Me.optRefNumberRange.TabIndex = 18
        Me.optRefNumberRange.TabStop = True
        Me.optRefNumberRange.Text = "Invoice # Range"
        Me.optRefNumberRange.UseVisualStyleBackColor = False
        '
        'lblToRefNumber
        '
        Me.lblToRefNumber.BackColor = System.Drawing.SystemColors.Control
        Me.lblToRefNumber.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblToRefNumber.Enabled = False
        Me.lblToRefNumber.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblToRefNumber.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblToRefNumber.Location = New System.Drawing.Point(248, 8)
        Me.lblToRefNumber.Name = "lblToRefNumber"
        Me.lblToRefNumber.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblToRefNumber.Size = New System.Drawing.Size(17, 17)
        Me.lblToRefNumber.TabIndex = 51
        Me.lblToRefNumber.Text = "To"
        '
        'lblFromRefNumber
        '
        Me.lblFromRefNumber.BackColor = System.Drawing.SystemColors.Control
        Me.lblFromRefNumber.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFromRefNumber.Enabled = False
        Me.lblFromRefNumber.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFromRefNumber.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFromRefNumber.Location = New System.Drawing.Point(104, 8)
        Me.lblFromRefNumber.Name = "lblFromRefNumber"
        Me.lblFromRefNumber.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFromRefNumber.Size = New System.Drawing.Size(32, 17)
        Me.lblFromRefNumber.TabIndex = 50
        Me.lblFromRefNumber.Text = "From"
        '
        'cmdClearRefNumber
        '
        Me.cmdClearRefNumber.BackColor = System.Drawing.SystemColors.Control
        Me.cmdClearRefNumber.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdClearRefNumber.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdClearRefNumber.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdClearRefNumber.Location = New System.Drawing.Point(408, 16)
        Me.cmdClearRefNumber.Name = "cmdClearRefNumber"
        Me.cmdClearRefNumber.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdClearRefNumber.Size = New System.Drawing.Size(65, 49)
        Me.cmdClearRefNumber.TabIndex = 48
        Me.cmdClearRefNumber.TabStop = False
        Me.cmdClearRefNumber.Text = "Clear Invoice # Filtering"
        Me.cmdClearRefNumber.UseVisualStyleBackColor = False
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me.txtToTime)
        Me.Frame3.Controls.Add(Me.txtFromTime)
        Me.Frame3.Controls.Add(Me.txtToDate)
        Me.Frame3.Controls.Add(Me.txtFromDate)
        Me.Frame3.Controls.Add(Me.frameFromAMPM)
        Me.Frame3.Controls.Add(Me.frameToAMPM)
        Me.Frame3.Controls.Add(Me.frameDateCriteria)
        Me.Frame3.Controls.Add(Me.cmdClearDateSelection)
        Me.Frame3.Controls.Add(Me.frmThisLast)
        Me.Frame3.Controls.Add(Me.frameMacroPeriod)
        Me.Frame3.Controls.Add(Me.lblToTime)
        Me.Frame3.Controls.Add(Me.lblFromTime)
        Me.Frame3.Controls.Add(Me.lblAfter)
        Me.Frame3.Controls.Add(Me.lblBefore)
        Me.Frame3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame3.Location = New System.Drawing.Point(8, 32)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(489, 129)
        Me.Frame3.TabIndex = 38
        Me.Frame3.TabStop = False
        '
        'txtToTime
        '
        Me.txtToTime.AcceptsReturn = True
        Me.txtToTime.BackColor = System.Drawing.SystemColors.Menu
        Me.txtToTime.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtToTime.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtToTime.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtToTime.Location = New System.Drawing.Point(232, 64)
        Me.txtToTime.MaxLength = 0
        Me.txtToTime.Name = "txtToTime"
        Me.txtToTime.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtToTime.Size = New System.Drawing.Size(41, 20)
        Me.txtToTime.TabIndex = 55
        '
        'txtFromTime
        '
        Me.txtFromTime.AcceptsReturn = True
        Me.txtFromTime.BackColor = System.Drawing.SystemColors.Menu
        Me.txtFromTime.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFromTime.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFromTime.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFromTime.Location = New System.Drawing.Point(232, 40)
        Me.txtFromTime.MaxLength = 0
        Me.txtFromTime.Name = "txtFromTime"
        Me.txtFromTime.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFromTime.Size = New System.Drawing.Size(41, 20)
        Me.txtFromTime.TabIndex = 54
        '
        'txtToDate
        '
        Me.txtToDate.AcceptsReturn = True
        Me.txtToDate.BackColor = System.Drawing.SystemColors.Menu
        Me.txtToDate.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtToDate.Enabled = False
        Me.txtToDate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtToDate.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtToDate.Location = New System.Drawing.Point(104, 64)
        Me.txtToDate.MaxLength = 0
        Me.txtToDate.Name = "txtToDate"
        Me.txtToDate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtToDate.Size = New System.Drawing.Size(73, 20)
        Me.txtToDate.TabIndex = 53
        '
        'txtFromDate
        '
        Me.txtFromDate.AcceptsReturn = True
        Me.txtFromDate.BackColor = System.Drawing.SystemColors.Menu
        Me.txtFromDate.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFromDate.Enabled = False
        Me.txtFromDate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFromDate.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFromDate.Location = New System.Drawing.Point(104, 40)
        Me.txtFromDate.MaxLength = 0
        Me.txtFromDate.Name = "txtFromDate"
        Me.txtFromDate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFromDate.Size = New System.Drawing.Size(73, 20)
        Me.txtFromDate.TabIndex = 52
        '
        'frameFromAMPM
        '
        Me.frameFromAMPM.BackColor = System.Drawing.SystemColors.Control
        Me.frameFromAMPM.Controls.Add(Me.optFromAM)
        Me.frameFromAMPM.Controls.Add(Me.optFromPM)
        Me.frameFromAMPM.Cursor = System.Windows.Forms.Cursors.Default
        Me.frameFromAMPM.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.frameFromAMPM.ForeColor = System.Drawing.SystemColors.ControlText
        Me.frameFromAMPM.Location = New System.Drawing.Point(280, 32)
        Me.frameFromAMPM.Name = "frameFromAMPM"
        Me.frameFromAMPM.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.frameFromAMPM.Size = New System.Drawing.Size(89, 33)
        Me.frameFromAMPM.TabIndex = 44
        '
        'optFromAM
        '
        Me.optFromAM.BackColor = System.Drawing.SystemColors.Control
        Me.optFromAM.Checked = True
        Me.optFromAM.Cursor = System.Windows.Forms.Cursors.Default
        Me.optFromAM.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optFromAM.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optFromAM.Location = New System.Drawing.Point(0, 8)
        Me.optFromAM.Name = "optFromAM"
        Me.optFromAM.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optFromAM.Size = New System.Drawing.Size(41, 17)
        Me.optFromAM.TabIndex = 4
        Me.optFromAM.TabStop = True
        Me.optFromAM.Text = "AM"
        Me.optFromAM.UseVisualStyleBackColor = False
        '
        'optFromPM
        '
        Me.optFromPM.BackColor = System.Drawing.SystemColors.Control
        Me.optFromPM.Cursor = System.Windows.Forms.Cursors.Default
        Me.optFromPM.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optFromPM.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optFromPM.Location = New System.Drawing.Point(40, 8)
        Me.optFromPM.Name = "optFromPM"
        Me.optFromPM.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optFromPM.Size = New System.Drawing.Size(41, 17)
        Me.optFromPM.TabIndex = 5
        Me.optFromPM.TabStop = True
        Me.optFromPM.Text = "PM"
        Me.optFromPM.UseVisualStyleBackColor = False
        '
        'frameToAMPM
        '
        Me.frameToAMPM.BackColor = System.Drawing.SystemColors.Control
        Me.frameToAMPM.Controls.Add(Me.optToPM)
        Me.frameToAMPM.Controls.Add(Me.optToAM)
        Me.frameToAMPM.Cursor = System.Windows.Forms.Cursors.Default
        Me.frameToAMPM.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.frameToAMPM.ForeColor = System.Drawing.SystemColors.ControlText
        Me.frameToAMPM.Location = New System.Drawing.Point(280, 56)
        Me.frameToAMPM.Name = "frameToAMPM"
        Me.frameToAMPM.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.frameToAMPM.Size = New System.Drawing.Size(89, 33)
        Me.frameToAMPM.TabIndex = 43
        '
        'optToPM
        '
        Me.optToPM.BackColor = System.Drawing.SystemColors.Control
        Me.optToPM.Cursor = System.Windows.Forms.Cursors.Default
        Me.optToPM.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optToPM.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optToPM.Location = New System.Drawing.Point(40, 8)
        Me.optToPM.Name = "optToPM"
        Me.optToPM.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optToPM.Size = New System.Drawing.Size(41, 17)
        Me.optToPM.TabIndex = 7
        Me.optToPM.TabStop = True
        Me.optToPM.Text = "PM"
        Me.optToPM.UseVisualStyleBackColor = False
        '
        'optToAM
        '
        Me.optToAM.BackColor = System.Drawing.SystemColors.Control
        Me.optToAM.Checked = True
        Me.optToAM.Cursor = System.Windows.Forms.Cursors.Default
        Me.optToAM.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optToAM.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optToAM.Location = New System.Drawing.Point(0, 8)
        Me.optToAM.Name = "optToAM"
        Me.optToAM.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optToAM.Size = New System.Drawing.Size(41, 17)
        Me.optToAM.TabIndex = 6
        Me.optToAM.TabStop = True
        Me.optToAM.Text = "AM"
        Me.optToAM.UseVisualStyleBackColor = False
        '
        'frameDateCriteria
        '
        Me.frameDateCriteria.BackColor = System.Drawing.SystemColors.Control
        Me.frameDateCriteria.Controls.Add(Me.optModified)
        Me.frameDateCriteria.Controls.Add(Me.optTxnDate)
        Me.frameDateCriteria.Controls.Add(Me.optMacroDate)
        Me.frameDateCriteria.Cursor = System.Windows.Forms.Cursors.Default
        Me.frameDateCriteria.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.frameDateCriteria.ForeColor = System.Drawing.SystemColors.ControlText
        Me.frameDateCriteria.Location = New System.Drawing.Point(8, 8)
        Me.frameDateCriteria.Name = "frameDateCriteria"
        Me.frameDateCriteria.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.frameDateCriteria.Size = New System.Drawing.Size(385, 33)
        Me.frameDateCriteria.TabIndex = 42
        Me.frameDateCriteria.Text = "Frame1"
        '
        'optModified
        '
        Me.optModified.BackColor = System.Drawing.SystemColors.Control
        Me.optModified.Cursor = System.Windows.Forms.Cursors.Default
        Me.optModified.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optModified.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optModified.Location = New System.Drawing.Point(0, 8)
        Me.optModified.Name = "optModified"
        Me.optModified.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optModified.Size = New System.Drawing.Size(113, 17)
        Me.optModified.TabIndex = 1
        Me.optModified.TabStop = True
        Me.optModified.Text = "Created or Modified"
        Me.optModified.UseVisualStyleBackColor = False
        '
        'optTxnDate
        '
        Me.optTxnDate.BackColor = System.Drawing.SystemColors.Control
        Me.optTxnDate.Cursor = System.Windows.Forms.Cursors.Default
        Me.optTxnDate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optTxnDate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optTxnDate.Location = New System.Drawing.Point(128, 8)
        Me.optTxnDate.Name = "optTxnDate"
        Me.optTxnDate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optTxnDate.Size = New System.Drawing.Size(113, 17)
        Me.optTxnDate.TabIndex = 2
        Me.optTxnDate.TabStop = True
        Me.optTxnDate.Text = "Transaction Date"
        Me.optTxnDate.UseVisualStyleBackColor = False
        '
        'optMacroDate
        '
        Me.optMacroDate.BackColor = System.Drawing.SystemColors.Control
        Me.optMacroDate.Cursor = System.Windows.Forms.Cursors.Default
        Me.optMacroDate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optMacroDate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optMacroDate.Location = New System.Drawing.Point(248, 8)
        Me.optMacroDate.Name = "optMacroDate"
        Me.optMacroDate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optMacroDate.Size = New System.Drawing.Size(137, 17)
        Me.optMacroDate.TabIndex = 3
        Me.optMacroDate.TabStop = True
        Me.optMacroDate.Text = "Transaction Date Macro"
        Me.optMacroDate.UseVisualStyleBackColor = False
        '
        'cmdClearDateSelection
        '
        Me.cmdClearDateSelection.BackColor = System.Drawing.SystemColors.Control
        Me.cmdClearDateSelection.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdClearDateSelection.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdClearDateSelection.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdClearDateSelection.Location = New System.Drawing.Point(408, 16)
        Me.cmdClearDateSelection.Name = "cmdClearDateSelection"
        Me.cmdClearDateSelection.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdClearDateSelection.Size = New System.Drawing.Size(65, 33)
        Me.cmdClearDateSelection.TabIndex = 41
        Me.cmdClearDateSelection.TabStop = False
        Me.cmdClearDateSelection.Text = "Clear Date Filtering"
        Me.cmdClearDateSelection.UseVisualStyleBackColor = False
        '
        'frmThisLast
        '
        Me.frmThisLast.BackColor = System.Drawing.SystemColors.Control
        Me.frmThisLast.Controls.Add(Me.optThis)
        Me.frmThisLast.Controls.Add(Me.optLast)
        Me.frmThisLast.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.frmThisLast.ForeColor = System.Drawing.SystemColors.ControlText
        Me.frmThisLast.Location = New System.Drawing.Point(8, 88)
        Me.frmThisLast.Name = "frmThisLast"
        Me.frmThisLast.Padding = New System.Windows.Forms.Padding(0)
        Me.frmThisLast.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.frmThisLast.Size = New System.Drawing.Size(145, 33)
        Me.frmThisLast.TabIndex = 40
        Me.frmThisLast.TabStop = False
        '
        'optThis
        '
        Me.optThis.BackColor = System.Drawing.SystemColors.Control
        Me.optThis.Cursor = System.Windows.Forms.Cursors.Default
        Me.optThis.Enabled = False
        Me.optThis.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optThis.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optThis.Location = New System.Drawing.Point(8, 8)
        Me.optThis.Name = "optThis"
        Me.optThis.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optThis.Size = New System.Drawing.Size(89, 17)
        Me.optThis.TabIndex = 8
        Me.optThis.TabStop = True
        Me.optThis.Text = "This (to date)"
        Me.optThis.UseVisualStyleBackColor = False
        '
        'optLast
        '
        Me.optLast.BackColor = System.Drawing.SystemColors.Control
        Me.optLast.Cursor = System.Windows.Forms.Cursors.Default
        Me.optLast.Enabled = False
        Me.optLast.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optLast.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optLast.Location = New System.Drawing.Point(96, 8)
        Me.optLast.Name = "optLast"
        Me.optLast.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optLast.Size = New System.Drawing.Size(41, 17)
        Me.optLast.TabIndex = 9
        Me.optLast.TabStop = True
        Me.optLast.Text = "Last"
        Me.optLast.UseVisualStyleBackColor = False
        '
        'frameMacroPeriod
        '
        Me.frameMacroPeriod.BackColor = System.Drawing.SystemColors.Control
        Me.frameMacroPeriod.Controls.Add(Me.optWeek)
        Me.frameMacroPeriod.Controls.Add(Me.optMonth)
        Me.frameMacroPeriod.Controls.Add(Me.optFiscalQuarter)
        Me.frameMacroPeriod.Controls.Add(Me.optFiscalYear)
        Me.frameMacroPeriod.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.frameMacroPeriod.ForeColor = System.Drawing.SystemColors.ControlText
        Me.frameMacroPeriod.Location = New System.Drawing.Point(160, 88)
        Me.frameMacroPeriod.Name = "frameMacroPeriod"
        Me.frameMacroPeriod.Padding = New System.Windows.Forms.Padding(0)
        Me.frameMacroPeriod.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.frameMacroPeriod.Size = New System.Drawing.Size(321, 33)
        Me.frameMacroPeriod.TabIndex = 39
        Me.frameMacroPeriod.TabStop = False
        '
        'optWeek
        '
        Me.optWeek.BackColor = System.Drawing.SystemColors.Control
        Me.optWeek.Cursor = System.Windows.Forms.Cursors.Default
        Me.optWeek.Enabled = False
        Me.optWeek.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optWeek.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optWeek.Location = New System.Drawing.Point(8, 8)
        Me.optWeek.Name = "optWeek"
        Me.optWeek.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optWeek.Size = New System.Drawing.Size(49, 17)
        Me.optWeek.TabIndex = 10
        Me.optWeek.TabStop = True
        Me.optWeek.Text = "Week"
        Me.optWeek.UseVisualStyleBackColor = False
        '
        'optMonth
        '
        Me.optMonth.BackColor = System.Drawing.SystemColors.Control
        Me.optMonth.Cursor = System.Windows.Forms.Cursors.Default
        Me.optMonth.Enabled = False
        Me.optMonth.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optMonth.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optMonth.Location = New System.Drawing.Point(72, 8)
        Me.optMonth.Name = "optMonth"
        Me.optMonth.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optMonth.Size = New System.Drawing.Size(57, 17)
        Me.optMonth.TabIndex = 11
        Me.optMonth.TabStop = True
        Me.optMonth.Text = "Month"
        Me.optMonth.UseVisualStyleBackColor = False
        '
        'optFiscalQuarter
        '
        Me.optFiscalQuarter.BackColor = System.Drawing.SystemColors.Control
        Me.optFiscalQuarter.Cursor = System.Windows.Forms.Cursors.Default
        Me.optFiscalQuarter.Enabled = False
        Me.optFiscalQuarter.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optFiscalQuarter.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optFiscalQuarter.Location = New System.Drawing.Point(136, 8)
        Me.optFiscalQuarter.Name = "optFiscalQuarter"
        Me.optFiscalQuarter.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optFiscalQuarter.Size = New System.Drawing.Size(89, 17)
        Me.optFiscalQuarter.TabIndex = 12
        Me.optFiscalQuarter.TabStop = True
        Me.optFiscalQuarter.Text = "Fiscal Quarter"
        Me.optFiscalQuarter.UseVisualStyleBackColor = False
        '
        'optFiscalYear
        '
        Me.optFiscalYear.BackColor = System.Drawing.SystemColors.Control
        Me.optFiscalYear.Cursor = System.Windows.Forms.Cursors.Default
        Me.optFiscalYear.Enabled = False
        Me.optFiscalYear.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optFiscalYear.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optFiscalYear.Location = New System.Drawing.Point(240, 8)
        Me.optFiscalYear.Name = "optFiscalYear"
        Me.optFiscalYear.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optFiscalYear.Size = New System.Drawing.Size(73, 17)
        Me.optFiscalYear.TabIndex = 13
        Me.optFiscalYear.TabStop = True
        Me.optFiscalYear.Text = "Fiscal Year"
        Me.optFiscalYear.UseVisualStyleBackColor = False
        '
        'lblToTime
        '
        Me.lblToTime.BackColor = System.Drawing.SystemColors.Control
        Me.lblToTime.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblToTime.Enabled = False
        Me.lblToTime.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblToTime.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblToTime.Location = New System.Drawing.Point(192, 64)
        Me.lblToTime.Name = "lblToTime"
        Me.lblToTime.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblToTime.Size = New System.Drawing.Size(38, 17)
        Me.lblToTime.TabIndex = 57
        Me.lblToTime.Text = "hh:mm"
        '
        'lblFromTime
        '
        Me.lblFromTime.BackColor = System.Drawing.SystemColors.Control
        Me.lblFromTime.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFromTime.Enabled = False
        Me.lblFromTime.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFromTime.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFromTime.Location = New System.Drawing.Point(192, 40)
        Me.lblFromTime.Name = "lblFromTime"
        Me.lblFromTime.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFromTime.Size = New System.Drawing.Size(38, 17)
        Me.lblFromTime.TabIndex = 56
        Me.lblFromTime.Text = "hh:mm"
        '
        'lblAfter
        '
        Me.lblAfter.BackColor = System.Drawing.SystemColors.Control
        Me.lblAfter.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblAfter.Enabled = False
        Me.lblAfter.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAfter.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblAfter.Location = New System.Drawing.Point(8, 40)
        Me.lblAfter.Name = "lblAfter"
        Me.lblAfter.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblAfter.Size = New System.Drawing.Size(97, 22)
        Me.lblAfter.TabIndex = 46
        Me.lblAfter.Text = "From (yyyy-mm-dd)"
        '
        'lblBefore
        '
        Me.lblBefore.BackColor = System.Drawing.SystemColors.Control
        Me.lblBefore.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblBefore.Enabled = False
        Me.lblBefore.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBefore.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblBefore.Location = New System.Drawing.Point(8, 64)
        Me.lblBefore.Name = "lblBefore"
        Me.lblBefore.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblBefore.Size = New System.Drawing.Size(105, 17)
        Me.lblBefore.TabIndex = 45
        Me.lblBefore.Text = "To    (yyyy-mm-dd)"
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.optPaidOnly)
        Me.Frame2.Controls.Add(Me.optUnpaidOnly)
        Me.Frame2.Controls.Add(Me.optAll)
        Me.Frame2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Frame2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame2.Location = New System.Drawing.Point(56, 296)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(185, 33)
        Me.Frame2.TabIndex = 35
        '
        'optPaidOnly
        '
        Me.optPaidOnly.BackColor = System.Drawing.SystemColors.Control
        Me.optPaidOnly.Cursor = System.Windows.Forms.Cursors.Default
        Me.optPaidOnly.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optPaidOnly.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optPaidOnly.Location = New System.Drawing.Point(120, 8)
        Me.optPaidOnly.Name = "optPaidOnly"
        Me.optPaidOnly.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optPaidOnly.Size = New System.Drawing.Size(75, 17)
        Me.optPaidOnly.TabIndex = 26
        Me.optPaidOnly.TabStop = True
        Me.optPaidOnly.Text = "only paid"
        Me.optPaidOnly.UseVisualStyleBackColor = False
        '
        'optUnpaidOnly
        '
        Me.optUnpaidOnly.BackColor = System.Drawing.SystemColors.Control
        Me.optUnpaidOnly.Cursor = System.Windows.Forms.Cursors.Default
        Me.optUnpaidOnly.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optUnpaidOnly.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optUnpaidOnly.Location = New System.Drawing.Point(40, 8)
        Me.optUnpaidOnly.Name = "optUnpaidOnly"
        Me.optUnpaidOnly.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optUnpaidOnly.Size = New System.Drawing.Size(81, 17)
        Me.optUnpaidOnly.TabIndex = 25
        Me.optUnpaidOnly.TabStop = True
        Me.optUnpaidOnly.Text = "only unpaid"
        Me.optUnpaidOnly.UseVisualStyleBackColor = False
        '
        'optAll
        '
        Me.optAll.BackColor = System.Drawing.SystemColors.Control
        Me.optAll.Checked = True
        Me.optAll.Cursor = System.Windows.Forms.Cursors.Default
        Me.optAll.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optAll.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optAll.Location = New System.Drawing.Point(0, 8)
        Me.optAll.Name = "optAll"
        Me.optAll.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optAll.Size = New System.Drawing.Size(41, 17)
        Me.optAll.TabIndex = 24
        Me.optAll.TabStop = True
        Me.optAll.Text = "all"
        Me.optAll.UseVisualStyleBackColor = False
        '
        'chkAccountChildren
        '
        Me.chkAccountChildren.BackColor = System.Drawing.SystemColors.Control
        Me.chkAccountChildren.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkAccountChildren.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkAccountChildren.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkAccountChildren.Location = New System.Drawing.Point(424, 200)
        Me.chkAccountChildren.Name = "chkAccountChildren"
        Me.chkAccountChildren.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkAccountChildren.Size = New System.Drawing.Size(89, 17)
        Me.chkAccountChildren.TabIndex = 17
        Me.chkAccountChildren.Text = "With Children"
        Me.chkAccountChildren.UseVisualStyleBackColor = False
        '
        'chkCustomerChildren
        '
        Me.chkCustomerChildren.BackColor = System.Drawing.SystemColors.Control
        Me.chkCustomerChildren.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkCustomerChildren.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkCustomerChildren.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkCustomerChildren.Location = New System.Drawing.Point(424, 168)
        Me.chkCustomerChildren.Name = "chkCustomerChildren"
        Me.chkCustomerChildren.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkCustomerChildren.Size = New System.Drawing.Size(89, 17)
        Me.chkCustomerChildren.TabIndex = 15
        Me.chkCustomerChildren.Text = "With Children"
        Me.chkCustomerChildren.UseVisualStyleBackColor = False
        '
        'txtInvoiceNumber
        '
        Me.txtInvoiceNumber.AcceptsReturn = True
        Me.txtInvoiceNumber.BackColor = System.Drawing.SystemColors.Window
        Me.txtInvoiceNumber.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtInvoiceNumber.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtInvoiceNumber.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtInvoiceNumber.Location = New System.Drawing.Point(96, 8)
        Me.txtInvoiceNumber.MaxLength = 0
        Me.txtInvoiceNumber.Name = "txtInvoiceNumber"
        Me.txtInvoiceNumber.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtInvoiceNumber.Size = New System.Drawing.Size(65, 20)
        Me.txtInvoiceNumber.TabIndex = 0
        '
        'cmdQuit
        '
        Me.cmdQuit.BackColor = System.Drawing.SystemColors.Control
        Me.cmdQuit.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdQuit.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdQuit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdQuit.Location = New System.Drawing.Point(400, 544)
        Me.cmdQuit.Name = "cmdQuit"
        Me.cmdQuit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdQuit.Size = New System.Drawing.Size(105, 41)
        Me.cmdQuit.TabIndex = 30
        Me.cmdQuit.Text = "Quit"
        Me.cmdQuit.UseVisualStyleBackColor = False
        '
        'cmdModifyInvoice
        '
        Me.cmdModifyInvoice.BackColor = System.Drawing.SystemColors.Control
        Me.cmdModifyInvoice.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdModifyInvoice.Enabled = False
        Me.cmdModifyInvoice.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdModifyInvoice.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdModifyInvoice.Location = New System.Drawing.Point(272, 544)
        Me.cmdModifyInvoice.Name = "cmdModifyInvoice"
        Me.cmdModifyInvoice.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdModifyInvoice.Size = New System.Drawing.Size(105, 41)
        Me.cmdModifyInvoice.TabIndex = 29
        Me.cmdModifyInvoice.Text = "Modify Invoice"
        Me.cmdModifyInvoice.UseVisualStyleBackColor = False
        '
        'lstInvoices
        '
        Me.lstInvoices.BackColor = System.Drawing.SystemColors.Window
        Me.lstInvoices.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstInvoices.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstInvoices.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstInvoices.ItemHeight = 14
        Me.lstInvoices.Location = New System.Drawing.Point(16, 336)
        Me.lstInvoices.Name = "lstInvoices"
        Me.lstInvoices.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstInvoices.Size = New System.Drawing.Size(489, 200)
        Me.lstInvoices.TabIndex = 28
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.SystemColors.Control
        Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label8.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label8.Location = New System.Drawing.Point(248, 304)
        Me.Label8.Name = "Label8"
        Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label8.Size = New System.Drawing.Size(48, 17)
        Me.Label8.TabIndex = 37
        Me.Label8.Text = "Invoices"
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(16, 304)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(41, 17)
        Me.Label7.TabIndex = 36
        Me.Label7.Text = "Return"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(16, 200)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(49, 17)
        Me.Label4.TabIndex = 34
        Me.Label4.Text = "Account"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(16, 168)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(73, 17)
        Me.Label3.TabIndex = 33
        Me.Label3.Text = "Customer/Job"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(176, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(249, 23)
        Me.Label2.TabIndex = 32
        Me.Label2.Text = "or find up to 30 invoices that meet the below criteria"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(16, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(73, 17)
        Me.Label1.TabIndex = 31
        Me.Label1.Text = "Find Invoice #"
        '
        'frmInvoiceQuery
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(518, 591)
        Me.Controls.Add(Me.chkShow)
        Me.Controls.Add(Me.cmdQuery)
        Me.Controls.Add(Me.cmbAccounts)
        Me.Controls.Add(Me.cmbCustomers)
        Me.Controls.Add(Me.frameRefNumberArea)
        Me.Controls.Add(Me.Frame3)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.chkAccountChildren)
        Me.Controls.Add(Me.chkCustomerChildren)
        Me.Controls.Add(Me.txtInvoiceNumber)
        Me.Controls.Add(Me.cmdQuit)
        Me.Controls.Add(Me.cmdModifyInvoice)
        Me.Controls.Add(Me.lstInvoices)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(271, 111)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmInvoiceQuery"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Select Invoice for modification"
        Me.frameRefNumberArea.ResumeLayout(False)
        Me.Frame4.ResumeLayout(False)
        Me.Frame4.PerformLayout()
        Me.Frame3.ResumeLayout(False)
        Me.Frame3.PerformLayout()
        Me.frameFromAMPM.ResumeLayout(False)
        Me.frameToAMPM.ResumeLayout(False)
        Me.frameDateCriteria.ResumeLayout(False)
        Me.frmThisLast.ResumeLayout(False)
        Me.frameMacroPeriod.ResumeLayout(False)
        Me.Frame2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
End Class