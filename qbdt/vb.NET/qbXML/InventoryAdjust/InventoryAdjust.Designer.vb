<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class InventoryAdjust
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
	Public WithEvents WarningMsg As System.Windows.Forms.TextBox
	Public WithEvents SendBtn As System.Windows.Forms.Button
	Public WithEvents AddBtn As System.Windows.Forms.Button
	Public WithEvents ValueAdjust As System.Windows.Forms.TextBox
	Public WithEvents QuantityAdjust As System.Windows.Forms.TextBox
	Public WithEvents ItemList As System.Windows.Forms.ComboBox
	Public WithEvents ValueOption As System.Windows.Forms.RadioButton
	Public WithEvents QuantityOption As System.Windows.Forms.RadioButton
	Public WithEvents DiffCheck As System.Windows.Forms.CheckBox
	Public WithEvents ValueLabel As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents _LineItems_ColumnHeader_1 As System.Windows.Forms.ColumnHeader
	Public WithEvents _LineItems_ColumnHeader_2 As System.Windows.Forms.ColumnHeader
	Public WithEvents _LineItems_ColumnHeader_3 As System.Windows.Forms.ColumnHeader
	Public WithEvents _LineItems_ColumnHeader_4 As System.Windows.Forms.ColumnHeader
	Public WithEvents LineItems As System.Windows.Forms.ListView
	Public WithEvents Memo As System.Windows.Forms.TextBox
	Public WithEvents ClassList As System.Windows.Forms.ComboBox
	Public WithEvents CustomerList As System.Windows.Forms.ComboBox
	Public WithEvents AccountList As System.Windows.Forms.ComboBox
	Public WithEvents ClassLabel As System.Windows.Forms.Label
	Public WithEvents Label6 As System.Windows.Forms.Label
	Public WithEvents Label5 As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(InventoryAdjust))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.WarningMsg = New System.Windows.Forms.TextBox()
        Me.SendBtn = New System.Windows.Forms.Button()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.AddBtn = New System.Windows.Forms.Button()
        Me.ValueAdjust = New System.Windows.Forms.TextBox()
        Me.QuantityAdjust = New System.Windows.Forms.TextBox()
        Me.ItemList = New System.Windows.Forms.ComboBox()
        Me.ValueOption = New System.Windows.Forms.RadioButton()
        Me.QuantityOption = New System.Windows.Forms.RadioButton()
        Me.DiffCheck = New System.Windows.Forms.CheckBox()
        Me.ValueLabel = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.LineItems = New System.Windows.Forms.ListView()
        Me._LineItems_ColumnHeader_1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me._LineItems_ColumnHeader_2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me._LineItems_ColumnHeader_3 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me._LineItems_ColumnHeader_4 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Memo = New System.Windows.Forms.TextBox()
        Me.ClassList = New System.Windows.Forms.ComboBox()
        Me.CustomerList = New System.Windows.Forms.ComboBox()
        Me.AccountList = New System.Windows.Forms.ComboBox()
        Me.ClassLabel = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        '
        'WarningMsg
        '
        Me.WarningMsg.AcceptsReturn = True
        Me.WarningMsg.BackColor = System.Drawing.SystemColors.Window
        Me.WarningMsg.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.WarningMsg.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.WarningMsg.ForeColor = System.Drawing.SystemColors.WindowText
        Me.WarningMsg.Location = New System.Drawing.Point(8, 8)
        Me.WarningMsg.MaxLength = 0
        Me.WarningMsg.Multiline = True
        Me.WarningMsg.Name = "WarningMsg"
        Me.WarningMsg.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.WarningMsg.Size = New System.Drawing.Size(553, 49)
        Me.WarningMsg.TabIndex = 21
        Me.WarningMsg.Text = resources.GetString("WarningMsg.Text")
        '
        'SendBtn
        '
        Me.SendBtn.BackColor = System.Drawing.SystemColors.Control
        Me.SendBtn.Cursor = System.Windows.Forms.Cursors.Default
        Me.SendBtn.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SendBtn.ForeColor = System.Drawing.SystemColors.ControlText
        Me.SendBtn.Location = New System.Drawing.Point(207, 536)
        Me.SendBtn.Name = "SendBtn"
        Me.SendBtn.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.SendBtn.Size = New System.Drawing.Size(161, 25)
        Me.SendBtn.TabIndex = 17
        Me.SendBtn.Text = "Send to QuickBooks!"
        Me.SendBtn.UseVisualStyleBackColor = False
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.AddBtn)
        Me.Frame1.Controls.Add(Me.ValueAdjust)
        Me.Frame1.Controls.Add(Me.QuantityAdjust)
        Me.Frame1.Controls.Add(Me.ItemList)
        Me.Frame1.Controls.Add(Me.ValueOption)
        Me.Frame1.Controls.Add(Me.QuantityOption)
        Me.Frame1.Controls.Add(Me.DiffCheck)
        Me.Frame1.Controls.Add(Me.ValueLabel)
        Me.Frame1.Controls.Add(Me.Label2)
        Me.Frame1.Controls.Add(Me.Label1)
        Me.Frame1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(24, 200)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(521, 129)
        Me.Frame1.TabIndex = 5
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Line Item"
        '
        'AddBtn
        '
        Me.AddBtn.BackColor = System.Drawing.SystemColors.Control
        Me.AddBtn.Cursor = System.Windows.Forms.Cursors.Default
        Me.AddBtn.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AddBtn.ForeColor = System.Drawing.SystemColors.ControlText
        Me.AddBtn.Location = New System.Drawing.Point(360, 96)
        Me.AddBtn.Name = "AddBtn"
        Me.AddBtn.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.AddBtn.Size = New System.Drawing.Size(105, 25)
        Me.AddBtn.TabIndex = 15
        Me.AddBtn.Text = "Add Line Item"
        Me.AddBtn.UseVisualStyleBackColor = False
        '
        'ValueAdjust
        '
        Me.ValueAdjust.AcceptsReturn = True
        Me.ValueAdjust.BackColor = System.Drawing.SystemColors.Window
        Me.ValueAdjust.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.ValueAdjust.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ValueAdjust.ForeColor = System.Drawing.SystemColors.WindowText
        Me.ValueAdjust.Location = New System.Drawing.Point(184, 96)
        Me.ValueAdjust.MaxLength = 0
        Me.ValueAdjust.Name = "ValueAdjust"
        Me.ValueAdjust.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ValueAdjust.Size = New System.Drawing.Size(57, 25)
        Me.ValueAdjust.TabIndex = 12
        Me.ValueAdjust.Text = "0"
        Me.ValueAdjust.Visible = False
        '
        'QuantityAdjust
        '
        Me.QuantityAdjust.AcceptsReturn = True
        Me.QuantityAdjust.BackColor = System.Drawing.SystemColors.Window
        Me.QuantityAdjust.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.QuantityAdjust.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.QuantityAdjust.ForeColor = System.Drawing.SystemColors.WindowText
        Me.QuantityAdjust.Location = New System.Drawing.Point(96, 96)
        Me.QuantityAdjust.MaxLength = 0
        Me.QuantityAdjust.Name = "QuantityAdjust"
        Me.QuantityAdjust.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.QuantityAdjust.Size = New System.Drawing.Size(57, 25)
        Me.QuantityAdjust.TabIndex = 10
        Me.QuantityAdjust.Text = "0"
        '
        'ItemList
        '
        Me.ItemList.BackColor = System.Drawing.SystemColors.Window
        Me.ItemList.Cursor = System.Windows.Forms.Cursors.Default
        Me.ItemList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ItemList.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ItemList.ForeColor = System.Drawing.SystemColors.WindowText
        Me.ItemList.Location = New System.Drawing.Point(96, 32)
        Me.ItemList.Name = "ItemList"
        Me.ItemList.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ItemList.Size = New System.Drawing.Size(393, 22)
        Me.ItemList.TabIndex = 8
        '
        'ValueOption
        '
        Me.ValueOption.BackColor = System.Drawing.SystemColors.Control
        Me.ValueOption.Cursor = System.Windows.Forms.Cursors.Default
        Me.ValueOption.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ValueOption.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ValueOption.Location = New System.Drawing.Point(16, 48)
        Me.ValueOption.Name = "ValueOption"
        Me.ValueOption.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ValueOption.Size = New System.Drawing.Size(65, 17)
        Me.ValueOption.TabIndex = 7
        Me.ValueOption.TabStop = True
        Me.ValueOption.Text = "Value"
        Me.ValueOption.UseVisualStyleBackColor = False
        '
        'QuantityOption
        '
        Me.QuantityOption.BackColor = System.Drawing.SystemColors.Control
        Me.QuantityOption.Checked = True
        Me.QuantityOption.Cursor = System.Windows.Forms.Cursors.Default
        Me.QuantityOption.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.QuantityOption.ForeColor = System.Drawing.SystemColors.ControlText
        Me.QuantityOption.Location = New System.Drawing.Point(16, 24)
        Me.QuantityOption.Name = "QuantityOption"
        Me.QuantityOption.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.QuantityOption.Size = New System.Drawing.Size(73, 17)
        Me.QuantityOption.TabIndex = 6
        Me.QuantityOption.TabStop = True
        Me.QuantityOption.Text = "Quantity"
        Me.QuantityOption.UseVisualStyleBackColor = False
        '
        'DiffCheck
        '
        Me.DiffCheck.BackColor = System.Drawing.SystemColors.Control
        Me.DiffCheck.Checked = True
        Me.DiffCheck.CheckState = System.Windows.Forms.CheckState.Checked
        Me.DiffCheck.Cursor = System.Windows.Forms.Cursors.Default
        Me.DiffCheck.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DiffCheck.ForeColor = System.Drawing.SystemColors.ControlText
        Me.DiffCheck.Location = New System.Drawing.Point(176, 96)
        Me.DiffCheck.Name = "DiffCheck"
        Me.DiffCheck.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.DiffCheck.Size = New System.Drawing.Size(105, 25)
        Me.DiffCheck.TabIndex = 14
        Me.DiffCheck.Text = "Difference?"
        Me.DiffCheck.UseVisualStyleBackColor = False
        '
        'ValueLabel
        '
        Me.ValueLabel.BackColor = System.Drawing.SystemColors.Control
        Me.ValueLabel.Cursor = System.Windows.Forms.Cursors.Default
        Me.ValueLabel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ValueLabel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ValueLabel.Location = New System.Drawing.Point(184, 64)
        Me.ValueLabel.Name = "ValueLabel"
        Me.ValueLabel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ValueLabel.Size = New System.Drawing.Size(57, 33)
        Me.ValueLabel.TabIndex = 13
        Me.ValueLabel.Text = "Adjust Value:"
        Me.ValueLabel.Visible = False
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(96, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(57, 33)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "Adjust Quantity:"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(96, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(89, 17)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Select Item:"
        '
        'LineItems
        '
        Me.LineItems.BackColor = System.Drawing.SystemColors.Window
        Me.LineItems.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me._LineItems_ColumnHeader_1, Me._LineItems_ColumnHeader_2, Me._LineItems_ColumnHeader_3, Me._LineItems_ColumnHeader_4})
        Me.LineItems.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LineItems.ForeColor = System.Drawing.SystemColors.WindowText
        Me.LineItems.LabelEdit = True
        Me.LineItems.Location = New System.Drawing.Point(24, 368)
        Me.LineItems.Name = "LineItems"
        Me.LineItems.Size = New System.Drawing.Size(521, 153)
        Me.LineItems.TabIndex = 4
        Me.LineItems.UseCompatibleStateImageBehavior = False
        Me.LineItems.View = System.Windows.Forms.View.Details
        '
        '_LineItems_ColumnHeader_1
        '
        Me._LineItems_ColumnHeader_1.Text = "Item"
        Me._LineItems_ColumnHeader_1.Width = 170
        '
        '_LineItems_ColumnHeader_2
        '
        Me._LineItems_ColumnHeader_2.Text = "Value/Quantity"
        Me._LineItems_ColumnHeader_2.Width = 170
        '
        '_LineItems_ColumnHeader_3
        '
        Me._LineItems_ColumnHeader_3.Text = "Type/Value"
        Me._LineItems_ColumnHeader_3.Width = 170
        '
        '_LineItems_ColumnHeader_4
        '
        Me._LineItems_ColumnHeader_4.Text = "Quantity"
        Me._LineItems_ColumnHeader_4.Width = 170
        '
        'Memo
        '
        Me.Memo.AcceptsReturn = True
        Me.Memo.BackColor = System.Drawing.SystemColors.Window
        Me.Memo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Memo.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Memo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Memo.Location = New System.Drawing.Point(24, 160)
        Me.Memo.MaxLength = 0
        Me.Memo.Name = "Memo"
        Me.Memo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Memo.Size = New System.Drawing.Size(521, 25)
        Me.Memo.TabIndex = 3
        Me.Memo.Text = "Memo"
        '
        'ClassList
        '
        Me.ClassList.BackColor = System.Drawing.SystemColors.Window
        Me.ClassList.Cursor = System.Windows.Forms.Cursors.Default
        Me.ClassList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ClassList.Enabled = False
        Me.ClassList.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ClassList.ForeColor = System.Drawing.SystemColors.WindowText
        Me.ClassList.Location = New System.Drawing.Point(184, 128)
        Me.ClassList.Name = "ClassList"
        Me.ClassList.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ClassList.Size = New System.Drawing.Size(361, 22)
        Me.ClassList.TabIndex = 2
        '
        'CustomerList
        '
        Me.CustomerList.BackColor = System.Drawing.SystemColors.Window
        Me.CustomerList.Cursor = System.Windows.Forms.Cursors.Default
        Me.CustomerList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CustomerList.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomerList.ForeColor = System.Drawing.SystemColors.WindowText
        Me.CustomerList.Location = New System.Drawing.Point(184, 96)
        Me.CustomerList.Name = "CustomerList"
        Me.CustomerList.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.CustomerList.Size = New System.Drawing.Size(361, 22)
        Me.CustomerList.TabIndex = 1
        '
        'AccountList
        '
        Me.AccountList.BackColor = System.Drawing.SystemColors.Window
        Me.AccountList.Cursor = System.Windows.Forms.Cursors.Default
        Me.AccountList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.AccountList.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AccountList.ForeColor = System.Drawing.SystemColors.WindowText
        Me.AccountList.Location = New System.Drawing.Point(184, 64)
        Me.AccountList.Name = "AccountList"
        Me.AccountList.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.AccountList.Size = New System.Drawing.Size(361, 22)
        Me.AccountList.TabIndex = 0
        '
        'ClassLabel
        '
        Me.ClassLabel.BackColor = System.Drawing.SystemColors.Control
        Me.ClassLabel.Cursor = System.Windows.Forms.Cursors.Default
        Me.ClassLabel.Enabled = False
        Me.ClassLabel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ClassLabel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ClassLabel.Location = New System.Drawing.Point(48, 128)
        Me.ClassLabel.Name = "ClassLabel"
        Me.ClassLabel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ClassLabel.Size = New System.Drawing.Size(113, 30)
        Me.ClassLabel.TabIndex = 20
        Me.ClassLabel.Text = "Select Class[Optional]:"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(48, 96)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(129, 30)
        Me.Label6.TabIndex = 19
        Me.Label6.Text = "Select Customer [Optional]:"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(96, 64)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(81, 21)
        Me.Label5.TabIndex = 18
        Me.Label5.Text = "Select Account:"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(24, 352)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(121, 17)
        Me.Label4.TabIndex = 16
        Me.Label4.Text = "Line Items:"
        '
        'InventoryAdjust
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(575, 565)
        Me.Controls.Add(Me.WarningMsg)
        Me.Controls.Add(Me.SendBtn)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.LineItems)
        Me.Controls.Add(Me.Memo)
        Me.Controls.Add(Me.ClassList)
        Me.Controls.Add(Me.CustomerList)
        Me.Controls.Add(Me.AccountList)
        Me.Controls.Add(Me.ClassLabel)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(4, 25)
        Me.Name = "InventoryAdjust"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "InventoryAdjust"
        Me.Frame1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
#End Region
End Class