<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmEditInvoiceLine
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
	Public WithEvents txtSavedDesc As System.Windows.Forms.TextBox
	Public WithEvents chkDisable As System.Windows.Forms.CheckBox
	Public WithEvents cmdFinish As System.Windows.Forms.Button
	Public WithEvents cmdSave As System.Windows.Forms.Button
	Public WithEvents cmdUndo As System.Windows.Forms.Button
	Public WithEvents cmbOverrideItemAccount As System.Windows.Forms.ComboBox
	Public WithEvents cmbSalesTaxCode As System.Windows.Forms.ComboBox
	Public WithEvents txtServiceDate As System.Windows.Forms.TextBox
	Public WithEvents cmbClass As System.Windows.Forms.ComboBox
	Public WithEvents txtAmount As System.Windows.Forms.TextBox
	Public WithEvents txtRate As System.Windows.Forms.TextBox
	Public WithEvents optRatePercent As System.Windows.Forms.RadioButton
	Public WithEvents optRate As System.Windows.Forms.RadioButton
	Public WithEvents txtQuantity As System.Windows.Forms.TextBox
	Public WithEvents txtDesc As System.Windows.Forms.TextBox
	Public WithEvents cmbItem As System.Windows.Forms.ComboBox
	Public WithEvents txtTxnLineID As System.Windows.Forms.TextBox
	Public WithEvents lblOverrideItemAccount As System.Windows.Forms.Label
	Public WithEvents lblSalesTaxCode As System.Windows.Forms.Label
	Public WithEvents lblServiceDate As System.Windows.Forms.Label
	Public WithEvents lblClass As System.Windows.Forms.Label
	Public WithEvents lblAmount As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents lblDesc As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtSavedDesc = New System.Windows.Forms.TextBox()
        Me.chkDisable = New System.Windows.Forms.CheckBox()
        Me.cmdFinish = New System.Windows.Forms.Button()
        Me.cmdSave = New System.Windows.Forms.Button()
        Me.cmdUndo = New System.Windows.Forms.Button()
        Me.cmbOverrideItemAccount = New System.Windows.Forms.ComboBox()
        Me.cmbSalesTaxCode = New System.Windows.Forms.ComboBox()
        Me.txtServiceDate = New System.Windows.Forms.TextBox()
        Me.cmbClass = New System.Windows.Forms.ComboBox()
        Me.txtAmount = New System.Windows.Forms.TextBox()
        Me.txtRate = New System.Windows.Forms.TextBox()
        Me.optRatePercent = New System.Windows.Forms.RadioButton()
        Me.optRate = New System.Windows.Forms.RadioButton()
        Me.txtQuantity = New System.Windows.Forms.TextBox()
        Me.txtDesc = New System.Windows.Forms.TextBox()
        Me.cmbItem = New System.Windows.Forms.ComboBox()
        Me.txtTxnLineID = New System.Windows.Forms.TextBox()
        Me.lblOverrideItemAccount = New System.Windows.Forms.Label()
        Me.lblSalesTaxCode = New System.Windows.Forms.Label()
        Me.lblServiceDate = New System.Windows.Forms.Label()
        Me.lblClass = New System.Windows.Forms.Label()
        Me.lblAmount = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.lblDesc = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'txtSavedDesc
        '
        Me.txtSavedDesc.AcceptsReturn = True
        Me.txtSavedDesc.BackColor = System.Drawing.SystemColors.Menu
        Me.txtSavedDesc.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtSavedDesc.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSavedDesc.Enabled = False
        Me.txtSavedDesc.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSavedDesc.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSavedDesc.Location = New System.Drawing.Point(504, 256)
        Me.txtSavedDesc.MaxLength = 0
        Me.txtSavedDesc.Multiline = True
        Me.txtSavedDesc.Name = "txtSavedDesc"
        Me.txtSavedDesc.ReadOnly = True
        Me.txtSavedDesc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSavedDesc.Size = New System.Drawing.Size(65, 57)
        Me.txtSavedDesc.TabIndex = 25
        Me.txtSavedDesc.Text = "Invisible " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "control" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "used to" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "save desc" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.txtSavedDesc.Visible = False
        '
        'chkDisable
        '
        Me.chkDisable.BackColor = System.Drawing.SystemColors.Control
        Me.chkDisable.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkDisable.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkDisable.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkDisable.Location = New System.Drawing.Point(8, 288)
        Me.chkDisable.Name = "chkDisable"
        Me.chkDisable.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkDisable.Size = New System.Drawing.Size(465, 17)
        Me.chkDisable.TabIndex = 24
        Me.chkDisable.Text = "Disable modify warnings (don't seek approval to clear fields based on changes to " &
    "other fields)"
        Me.chkDisable.UseVisualStyleBackColor = False
        '
        'cmdFinish
        '
        Me.cmdFinish.BackColor = System.Drawing.SystemColors.Control
        Me.cmdFinish.CausesValidation = False
        Me.cmdFinish.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdFinish.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdFinish.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdFinish.Location = New System.Drawing.Point(392, 320)
        Me.cmdFinish.Name = "cmdFinish"
        Me.cmdFinish.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdFinish.Size = New System.Drawing.Size(177, 49)
        Me.cmdFinish.TabIndex = 23
        Me.cmdFinish.Text = "Finish without saving changes"
        Me.cmdFinish.UseVisualStyleBackColor = False
        '
        'cmdSave
        '
        Me.cmdSave.BackColor = System.Drawing.SystemColors.Control
        Me.cmdSave.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdSave.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSave.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdSave.Location = New System.Drawing.Point(216, 320)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdSave.Size = New System.Drawing.Size(161, 49)
        Me.cmdSave.TabIndex = 22
        Me.cmdSave.Text = "Save Invoice Line"
        Me.cmdSave.UseVisualStyleBackColor = False
        '
        'cmdUndo
        '
        Me.cmdUndo.BackColor = System.Drawing.SystemColors.Control
        Me.cmdUndo.CausesValidation = False
        Me.cmdUndo.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdUndo.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdUndo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdUndo.Location = New System.Drawing.Point(8, 320)
        Me.cmdUndo.Name = "cmdUndo"
        Me.cmdUndo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdUndo.Size = New System.Drawing.Size(193, 49)
        Me.cmdUndo.TabIndex = 21
        Me.cmdUndo.Text = "Undo Changes / Get Original Values"
        Me.cmdUndo.UseVisualStyleBackColor = False
        '
        'cmbOverrideItemAccount
        '
        Me.cmbOverrideItemAccount.BackColor = System.Drawing.SystemColors.Window
        Me.cmbOverrideItemAccount.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbOverrideItemAccount.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbOverrideItemAccount.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbOverrideItemAccount.Location = New System.Drawing.Point(120, 256)
        Me.cmbOverrideItemAccount.Name = "cmbOverrideItemAccount"
        Me.cmbOverrideItemAccount.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbOverrideItemAccount.Size = New System.Drawing.Size(338, 22)
        Me.cmbOverrideItemAccount.TabIndex = 20
        '
        'cmbSalesTaxCode
        '
        Me.cmbSalesTaxCode.BackColor = System.Drawing.SystemColors.Window
        Me.cmbSalesTaxCode.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbSalesTaxCode.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbSalesTaxCode.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbSalesTaxCode.Location = New System.Drawing.Point(512, 224)
        Me.cmbSalesTaxCode.Name = "cmbSalesTaxCode"
        Me.cmbSalesTaxCode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbSalesTaxCode.Size = New System.Drawing.Size(58, 22)
        Me.cmbSalesTaxCode.TabIndex = 18
        '
        'txtServiceDate
        '
        Me.txtServiceDate.AcceptsReturn = True
        Me.txtServiceDate.BackColor = System.Drawing.SystemColors.Window
        Me.txtServiceDate.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtServiceDate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtServiceDate.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtServiceDate.Location = New System.Drawing.Point(328, 224)
        Me.txtServiceDate.MaxLength = 0
        Me.txtServiceDate.Name = "txtServiceDate"
        Me.txtServiceDate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtServiceDate.Size = New System.Drawing.Size(65, 19)
        Me.txtServiceDate.TabIndex = 16
        '
        'cmbClass
        '
        Me.cmbClass.BackColor = System.Drawing.SystemColors.Window
        Me.cmbClass.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbClass.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbClass.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbClass.Location = New System.Drawing.Point(40, 224)
        Me.cmbClass.Name = "cmbClass"
        Me.cmbClass.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbClass.Size = New System.Drawing.Size(201, 22)
        Me.cmbClass.TabIndex = 14
        '
        'txtAmount
        '
        Me.txtAmount.AcceptsReturn = True
        Me.txtAmount.BackColor = System.Drawing.SystemColors.Window
        Me.txtAmount.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtAmount.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAmount.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtAmount.Location = New System.Drawing.Point(456, 192)
        Me.txtAmount.MaxLength = 0
        Me.txtAmount.Name = "txtAmount"
        Me.txtAmount.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtAmount.Size = New System.Drawing.Size(113, 19)
        Me.txtAmount.TabIndex = 12
        '
        'txtRate
        '
        Me.txtRate.AcceptsReturn = True
        Me.txtRate.BackColor = System.Drawing.SystemColors.Window
        Me.txtRate.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRate.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRate.Location = New System.Drawing.Point(272, 192)
        Me.txtRate.MaxLength = 0
        Me.txtRate.Name = "txtRate"
        Me.txtRate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRate.Size = New System.Drawing.Size(81, 19)
        Me.txtRate.TabIndex = 10
        '
        'optRatePercent
        '
        Me.optRatePercent.BackColor = System.Drawing.SystemColors.Control
        Me.optRatePercent.Cursor = System.Windows.Forms.Cursors.Default
        Me.optRatePercent.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optRatePercent.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optRatePercent.Location = New System.Drawing.Point(184, 192)
        Me.optRatePercent.Name = "optRatePercent"
        Me.optRatePercent.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optRatePercent.Size = New System.Drawing.Size(89, 17)
        Me.optRatePercent.TabIndex = 9
        Me.optRatePercent.TabStop = True
        Me.optRatePercent.Text = "Rate Percent"
        Me.optRatePercent.UseVisualStyleBackColor = False
        '
        'optRate
        '
        Me.optRate.BackColor = System.Drawing.SystemColors.Control
        Me.optRate.Cursor = System.Windows.Forms.Cursors.Default
        Me.optRate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optRate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optRate.Location = New System.Drawing.Point(128, 192)
        Me.optRate.Name = "optRate"
        Me.optRate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optRate.Size = New System.Drawing.Size(49, 17)
        Me.optRate.TabIndex = 8
        Me.optRate.TabStop = True
        Me.optRate.Text = "Rate"
        Me.optRate.UseVisualStyleBackColor = False
        '
        'txtQuantity
        '
        Me.txtQuantity.AcceptsReturn = True
        Me.txtQuantity.BackColor = System.Drawing.SystemColors.Window
        Me.txtQuantity.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtQuantity.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtQuantity.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtQuantity.Location = New System.Drawing.Point(56, 192)
        Me.txtQuantity.MaxLength = 0
        Me.txtQuantity.Name = "txtQuantity"
        Me.txtQuantity.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtQuantity.Size = New System.Drawing.Size(57, 19)
        Me.txtQuantity.TabIndex = 7
        '
        'txtDesc
        '
        Me.txtDesc.AcceptsReturn = True
        Me.txtDesc.BackColor = System.Drawing.SystemColors.Window
        Me.txtDesc.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDesc.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDesc.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDesc.Location = New System.Drawing.Point(45, 72)
        Me.txtDesc.MaxLength = 0
        Me.txtDesc.Multiline = True
        Me.txtDesc.Name = "txtDesc"
        Me.txtDesc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDesc.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtDesc.Size = New System.Drawing.Size(529, 105)
        Me.txtDesc.TabIndex = 5
        Me.txtDesc.WordWrap = False
        '
        'cmbItem
        '
        Me.cmbItem.BackColor = System.Drawing.SystemColors.Window
        Me.cmbItem.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbItem.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbItem.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbItem.Location = New System.Drawing.Point(40, 40)
        Me.cmbItem.Name = "cmbItem"
        Me.cmbItem.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbItem.Size = New System.Drawing.Size(529, 22)
        Me.cmbItem.TabIndex = 3
        '
        'txtTxnLineID
        '
        Me.txtTxnLineID.AcceptsReturn = True
        Me.txtTxnLineID.BackColor = System.Drawing.SystemColors.Window
        Me.txtTxnLineID.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTxnLineID.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTxnLineID.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTxnLineID.Location = New System.Drawing.Point(64, 8)
        Me.txtTxnLineID.MaxLength = 0
        Me.txtTxnLineID.Name = "txtTxnLineID"
        Me.txtTxnLineID.ReadOnly = True
        Me.txtTxnLineID.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTxnLineID.Size = New System.Drawing.Size(177, 19)
        Me.txtTxnLineID.TabIndex = 1
        '
        'lblOverrideItemAccount
        '
        Me.lblOverrideItemAccount.BackColor = System.Drawing.SystemColors.Control
        Me.lblOverrideItemAccount.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblOverrideItemAccount.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOverrideItemAccount.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblOverrideItemAccount.Location = New System.Drawing.Point(8, 256)
        Me.lblOverrideItemAccount.Name = "lblOverrideItemAccount"
        Me.lblOverrideItemAccount.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblOverrideItemAccount.Size = New System.Drawing.Size(113, 25)
        Me.lblOverrideItemAccount.TabIndex = 19
        Me.lblOverrideItemAccount.Text = "Override Item Account"
        '
        'lblSalesTaxCode
        '
        Me.lblSalesTaxCode.BackColor = System.Drawing.SystemColors.Control
        Me.lblSalesTaxCode.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSalesTaxCode.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSalesTaxCode.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblSalesTaxCode.Location = New System.Drawing.Point(424, 224)
        Me.lblSalesTaxCode.Name = "lblSalesTaxCode"
        Me.lblSalesTaxCode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSalesTaxCode.Size = New System.Drawing.Size(85, 17)
        Me.lblSalesTaxCode.TabIndex = 17
        Me.lblSalesTaxCode.Text = "Sales Tax Code"
        '
        'lblServiceDate
        '
        Me.lblServiceDate.BackColor = System.Drawing.SystemColors.Control
        Me.lblServiceDate.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblServiceDate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblServiceDate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblServiceDate.Location = New System.Drawing.Point(256, 224)
        Me.lblServiceDate.Name = "lblServiceDate"
        Me.lblServiceDate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblServiceDate.Size = New System.Drawing.Size(75, 17)
        Me.lblServiceDate.TabIndex = 15
        Me.lblServiceDate.Text = "Service Date"
        '
        'lblClass
        '
        Me.lblClass.BackColor = System.Drawing.SystemColors.Control
        Me.lblClass.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblClass.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblClass.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblClass.Location = New System.Drawing.Point(8, 224)
        Me.lblClass.Name = "lblClass"
        Me.lblClass.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblClass.Size = New System.Drawing.Size(35, 17)
        Me.lblClass.TabIndex = 13
        Me.lblClass.Text = "Class"
        '
        'lblAmount
        '
        Me.lblAmount.BackColor = System.Drawing.SystemColors.Control
        Me.lblAmount.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblAmount.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAmount.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblAmount.Location = New System.Drawing.Point(408, 192)
        Me.lblAmount.Name = "lblAmount"
        Me.lblAmount.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblAmount.Size = New System.Drawing.Size(45, 17)
        Me.lblAmount.TabIndex = 11
        Me.lblAmount.Text = "Amount"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(8, 192)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(48, 17)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Quantity"
        '
        'lblDesc
        '
        Me.lblDesc.BackColor = System.Drawing.SystemColors.Control
        Me.lblDesc.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDesc.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDesc.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDesc.Location = New System.Drawing.Point(8, 72)
        Me.lblDesc.Name = "lblDesc"
        Me.lblDesc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDesc.Size = New System.Drawing.Size(32, 17)
        Me.lblDesc.TabIndex = 4
        Me.lblDesc.Text = "Desc"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(8, 40)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(30, 17)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Item"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(55, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "TxnLineID"
        '
        'frmEditInvoiceLine
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(577, 383)
        Me.Controls.Add(Me.txtSavedDesc)
        Me.Controls.Add(Me.chkDisable)
        Me.Controls.Add(Me.cmdFinish)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.cmdUndo)
        Me.Controls.Add(Me.cmbOverrideItemAccount)
        Me.Controls.Add(Me.cmbSalesTaxCode)
        Me.Controls.Add(Me.txtServiceDate)
        Me.Controls.Add(Me.cmbClass)
        Me.Controls.Add(Me.txtAmount)
        Me.Controls.Add(Me.txtRate)
        Me.Controls.Add(Me.optRatePercent)
        Me.Controls.Add(Me.optRate)
        Me.Controls.Add(Me.txtQuantity)
        Me.Controls.Add(Me.txtDesc)
        Me.Controls.Add(Me.cmbItem)
        Me.Controls.Add(Me.txtTxnLineID)
        Me.Controls.Add(Me.lblOverrideItemAccount)
        Me.Controls.Add(Me.lblSalesTaxCode)
        Me.Controls.Add(Me.lblServiceDate)
        Me.Controls.Add(Me.lblClass)
        Me.Controls.Add(Me.lblAmount)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.lblDesc)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(243, 230)
        Me.Name = "frmEditInvoiceLine"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Edit Invoice Line"
        Me.ResumeLayout(False)

    End Sub
#End Region
End Class