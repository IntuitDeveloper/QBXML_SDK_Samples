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
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(InventoryAdjust))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.WarningMsg = New System.Windows.Forms.TextBox
		Me.SendBtn = New System.Windows.Forms.Button
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me.AddBtn = New System.Windows.Forms.Button
		Me.ValueAdjust = New System.Windows.Forms.TextBox
		Me.QuantityAdjust = New System.Windows.Forms.TextBox
		Me.ItemList = New System.Windows.Forms.ComboBox
		Me.ValueOption = New System.Windows.Forms.RadioButton
		Me.QuantityOption = New System.Windows.Forms.RadioButton
		Me.DiffCheck = New System.Windows.Forms.CheckBox
		Me.ValueLabel = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.LineItems = New System.Windows.Forms.ListView
		Me._LineItems_ColumnHeader_1 = New System.Windows.Forms.ColumnHeader
		Me._LineItems_ColumnHeader_2 = New System.Windows.Forms.ColumnHeader
		Me._LineItems_ColumnHeader_3 = New System.Windows.Forms.ColumnHeader
		Me._LineItems_ColumnHeader_4 = New System.Windows.Forms.ColumnHeader
		Me.Memo = New System.Windows.Forms.TextBox
		Me.ClassList = New System.Windows.Forms.ComboBox
		Me.CustomerList = New System.Windows.Forms.ComboBox
		Me.AccountList = New System.Windows.Forms.ComboBox
		Me.ClassLabel = New System.Windows.Forms.Label
		Me.Label6 = New System.Windows.Forms.Label
		Me.Label5 = New System.Windows.Forms.Label
		Me.Label4 = New System.Windows.Forms.Label
		Me.Frame1.SuspendLayout()
		Me.LineItems.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "InventoryAdjust"
		Me.ClientSize = New System.Drawing.Size(575, 565)
		Me.Location = New System.Drawing.Point(4, 25)
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
		Me.Name = "InventoryAdjust"
		Me.WarningMsg.AutoSize = False
		Me.WarningMsg.Size = New System.Drawing.Size(553, 49)
		Me.WarningMsg.Location = New System.Drawing.Point(0, 0)
		Me.WarningMsg.MultiLine = True
		Me.WarningMsg.TabIndex = 21
		Me.WarningMsg.Text = "The QuickBooks company file must be open in single-user mode. (This is also true in the user interface: you cannot " & Chr(13) & Chr(10) & "adjust the inventory while the company file is open in multi-user mode.) " & Chr(13) & Chr(10) & "The inventory adjustment form must be closed in the QuickBooks user interface. " & Chr(13) & Chr(10)
		Me.WarningMsg.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.WarningMsg.AcceptsReturn = True
		Me.WarningMsg.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.WarningMsg.BackColor = System.Drawing.SystemColors.Window
		Me.WarningMsg.CausesValidation = True
		Me.WarningMsg.Enabled = True
		Me.WarningMsg.ForeColor = System.Drawing.SystemColors.WindowText
		Me.WarningMsg.HideSelection = True
		Me.WarningMsg.ReadOnly = False
		Me.WarningMsg.Maxlength = 0
		Me.WarningMsg.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.WarningMsg.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.WarningMsg.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.WarningMsg.TabStop = True
		Me.WarningMsg.Visible = True
		Me.WarningMsg.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.WarningMsg.Name = "WarningMsg"
		Me.SendBtn.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.SendBtn.Text = "Send to QuickBooks!"
		Me.SendBtn.Size = New System.Drawing.Size(161, 25)
		Me.SendBtn.Location = New System.Drawing.Point(207, 536)
		Me.SendBtn.TabIndex = 17
		Me.SendBtn.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.SendBtn.BackColor = System.Drawing.SystemColors.Control
		Me.SendBtn.CausesValidation = True
		Me.SendBtn.Enabled = True
		Me.SendBtn.ForeColor = System.Drawing.SystemColors.ControlText
		Me.SendBtn.Cursor = System.Windows.Forms.Cursors.Default
		Me.SendBtn.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.SendBtn.TabStop = True
		Me.SendBtn.Name = "SendBtn"
		Me.Frame1.Text = "Line Item"
		Me.Frame1.Size = New System.Drawing.Size(521, 129)
		Me.Frame1.Location = New System.Drawing.Point(24, 200)
		Me.Frame1.TabIndex = 5
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.BackColor = System.Drawing.SystemColors.Control
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame1.Name = "Frame1"
		Me.AddBtn.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.AddBtn.Text = "Add Line Item"
		Me.AddBtn.Size = New System.Drawing.Size(105, 25)
		Me.AddBtn.Location = New System.Drawing.Point(360, 96)
		Me.AddBtn.TabIndex = 15
		Me.AddBtn.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AddBtn.BackColor = System.Drawing.SystemColors.Control
		Me.AddBtn.CausesValidation = True
		Me.AddBtn.Enabled = True
		Me.AddBtn.ForeColor = System.Drawing.SystemColors.ControlText
		Me.AddBtn.Cursor = System.Windows.Forms.Cursors.Default
		Me.AddBtn.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.AddBtn.TabStop = True
		Me.AddBtn.Name = "AddBtn"
		Me.ValueAdjust.AutoSize = False
		Me.ValueAdjust.Size = New System.Drawing.Size(57, 25)
		Me.ValueAdjust.Location = New System.Drawing.Point(184, 96)
		Me.ValueAdjust.TabIndex = 12
		Me.ValueAdjust.Text = "0"
		Me.ValueAdjust.Visible = False
		Me.ValueAdjust.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ValueAdjust.AcceptsReturn = True
		Me.ValueAdjust.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.ValueAdjust.BackColor = System.Drawing.SystemColors.Window
		Me.ValueAdjust.CausesValidation = True
		Me.ValueAdjust.Enabled = True
		Me.ValueAdjust.ForeColor = System.Drawing.SystemColors.WindowText
		Me.ValueAdjust.HideSelection = True
		Me.ValueAdjust.ReadOnly = False
		Me.ValueAdjust.Maxlength = 0
		Me.ValueAdjust.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.ValueAdjust.MultiLine = False
		Me.ValueAdjust.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ValueAdjust.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.ValueAdjust.TabStop = True
		Me.ValueAdjust.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.ValueAdjust.Name = "ValueAdjust"
		Me.QuantityAdjust.AutoSize = False
		Me.QuantityAdjust.Size = New System.Drawing.Size(57, 25)
		Me.QuantityAdjust.Location = New System.Drawing.Point(96, 96)
		Me.QuantityAdjust.TabIndex = 10
		Me.QuantityAdjust.Text = "0"
		Me.QuantityAdjust.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.QuantityAdjust.AcceptsReturn = True
		Me.QuantityAdjust.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.QuantityAdjust.BackColor = System.Drawing.SystemColors.Window
		Me.QuantityAdjust.CausesValidation = True
		Me.QuantityAdjust.Enabled = True
		Me.QuantityAdjust.ForeColor = System.Drawing.SystemColors.WindowText
		Me.QuantityAdjust.HideSelection = True
		Me.QuantityAdjust.ReadOnly = False
		Me.QuantityAdjust.Maxlength = 0
		Me.QuantityAdjust.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.QuantityAdjust.MultiLine = False
		Me.QuantityAdjust.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.QuantityAdjust.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.QuantityAdjust.TabStop = True
		Me.QuantityAdjust.Visible = True
		Me.QuantityAdjust.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.QuantityAdjust.Name = "QuantityAdjust"
		Me.ItemList.Size = New System.Drawing.Size(393, 21)
		Me.ItemList.Location = New System.Drawing.Point(96, 32)
		Me.ItemList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.ItemList.TabIndex = 8
		Me.ItemList.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ItemList.BackColor = System.Drawing.SystemColors.Window
		Me.ItemList.CausesValidation = True
		Me.ItemList.Enabled = True
		Me.ItemList.ForeColor = System.Drawing.SystemColors.WindowText
		Me.ItemList.IntegralHeight = True
		Me.ItemList.Cursor = System.Windows.Forms.Cursors.Default
		Me.ItemList.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ItemList.Sorted = False
		Me.ItemList.TabStop = True
		Me.ItemList.Visible = True
		Me.ItemList.Name = "ItemList"
		Me.ValueOption.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.ValueOption.Text = "Value"
		Me.ValueOption.Size = New System.Drawing.Size(65, 17)
		Me.ValueOption.Location = New System.Drawing.Point(16, 48)
		Me.ValueOption.TabIndex = 7
		Me.ValueOption.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ValueOption.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.ValueOption.BackColor = System.Drawing.SystemColors.Control
		Me.ValueOption.CausesValidation = True
		Me.ValueOption.Enabled = True
		Me.ValueOption.ForeColor = System.Drawing.SystemColors.ControlText
		Me.ValueOption.Cursor = System.Windows.Forms.Cursors.Default
		Me.ValueOption.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ValueOption.Appearance = System.Windows.Forms.Appearance.Normal
		Me.ValueOption.TabStop = True
		Me.ValueOption.Checked = False
		Me.ValueOption.Visible = True
		Me.ValueOption.Name = "ValueOption"
		Me.QuantityOption.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.QuantityOption.Text = "Quantity"
		Me.QuantityOption.Size = New System.Drawing.Size(73, 17)
		Me.QuantityOption.Location = New System.Drawing.Point(16, 24)
		Me.QuantityOption.TabIndex = 6
		Me.QuantityOption.Checked = True
		Me.QuantityOption.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.QuantityOption.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.QuantityOption.BackColor = System.Drawing.SystemColors.Control
		Me.QuantityOption.CausesValidation = True
		Me.QuantityOption.Enabled = True
		Me.QuantityOption.ForeColor = System.Drawing.SystemColors.ControlText
		Me.QuantityOption.Cursor = System.Windows.Forms.Cursors.Default
		Me.QuantityOption.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.QuantityOption.Appearance = System.Windows.Forms.Appearance.Normal
		Me.QuantityOption.TabStop = True
		Me.QuantityOption.Visible = True
		Me.QuantityOption.Name = "QuantityOption"
		Me.DiffCheck.Text = "Difference?"
		Me.DiffCheck.Size = New System.Drawing.Size(105, 25)
		Me.DiffCheck.Location = New System.Drawing.Point(176, 96)
		Me.DiffCheck.TabIndex = 14
		Me.DiffCheck.CheckState = System.Windows.Forms.CheckState.Checked
		Me.DiffCheck.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.DiffCheck.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.DiffCheck.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.DiffCheck.BackColor = System.Drawing.SystemColors.Control
		Me.DiffCheck.CausesValidation = True
		Me.DiffCheck.Enabled = True
		Me.DiffCheck.ForeColor = System.Drawing.SystemColors.ControlText
		Me.DiffCheck.Cursor = System.Windows.Forms.Cursors.Default
		Me.DiffCheck.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.DiffCheck.Appearance = System.Windows.Forms.Appearance.Normal
		Me.DiffCheck.TabStop = True
		Me.DiffCheck.Visible = True
		Me.DiffCheck.Name = "DiffCheck"
		Me.ValueLabel.Text = "Adjust Value:"
		Me.ValueLabel.Size = New System.Drawing.Size(57, 33)
		Me.ValueLabel.Location = New System.Drawing.Point(184, 64)
		Me.ValueLabel.TabIndex = 13
		Me.ValueLabel.Visible = False
		Me.ValueLabel.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ValueLabel.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.ValueLabel.BackColor = System.Drawing.SystemColors.Control
		Me.ValueLabel.Enabled = True
		Me.ValueLabel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.ValueLabel.Cursor = System.Windows.Forms.Cursors.Default
		Me.ValueLabel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ValueLabel.UseMnemonic = True
		Me.ValueLabel.AutoSize = False
		Me.ValueLabel.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.ValueLabel.Name = "ValueLabel"
		Me.Label2.Text = "Adjust Quantity:"
		Me.Label2.Size = New System.Drawing.Size(57, 33)
		Me.Label2.Location = New System.Drawing.Point(96, 64)
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
		Me.Label1.Text = "Select Item:"
		Me.Label1.Size = New System.Drawing.Size(89, 17)
		Me.Label1.Location = New System.Drawing.Point(96, 16)
		Me.Label1.TabIndex = 9
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
		Me.LineItems.Size = New System.Drawing.Size(521, 153)
		Me.LineItems.Location = New System.Drawing.Point(24, 368)
		Me.LineItems.TabIndex = 4
		Me.LineItems.View = System.Windows.Forms.View.Details
		Me.LineItems.LabelWrap = True
		Me.LineItems.HideSelection = True
		Me.LineItems.ForeColor = System.Drawing.SystemColors.WindowText
		Me.LineItems.BackColor = System.Drawing.SystemColors.Window
		Me.LineItems.Enabled = False
		Me.LineItems.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.LineItems.LabelEdit = True
		Me.LineItems.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.LineItems.Name = "LineItems"
		Me._LineItems_ColumnHeader_1.Text = "Item"
		Me._LineItems_ColumnHeader_1.Width = 170
		Me._LineItems_ColumnHeader_2.Text = "Value/Quantity"
		Me._LineItems_ColumnHeader_2.Width = 170
		Me._LineItems_ColumnHeader_3.Text = "Type/Value"
		Me._LineItems_ColumnHeader_3.Width = 170
		Me._LineItems_ColumnHeader_4.Text = "Quantity"
		Me._LineItems_ColumnHeader_4.Width = 170
		Me.Memo.AutoSize = False
		Me.Memo.Size = New System.Drawing.Size(521, 25)
		Me.Memo.Location = New System.Drawing.Point(24, 160)
		Me.Memo.TabIndex = 3
		Me.Memo.Text = "Memo"
		Me.Memo.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Memo.AcceptsReturn = True
		Me.Memo.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Memo.BackColor = System.Drawing.SystemColors.Window
		Me.Memo.CausesValidation = True
		Me.Memo.Enabled = True
		Me.Memo.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Memo.HideSelection = True
		Me.Memo.ReadOnly = False
		Me.Memo.Maxlength = 0
		Me.Memo.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Memo.MultiLine = False
		Me.Memo.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Memo.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Memo.TabStop = True
		Me.Memo.Visible = True
		Me.Memo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Memo.Name = "Memo"
		Me.ClassList.Enabled = False
		Me.ClassList.Size = New System.Drawing.Size(361, 21)
		Me.ClassList.Location = New System.Drawing.Point(184, 128)
		Me.ClassList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.ClassList.TabIndex = 2
		Me.ClassList.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ClassList.BackColor = System.Drawing.SystemColors.Window
		Me.ClassList.CausesValidation = True
		Me.ClassList.ForeColor = System.Drawing.SystemColors.WindowText
		Me.ClassList.IntegralHeight = True
		Me.ClassList.Cursor = System.Windows.Forms.Cursors.Default
		Me.ClassList.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ClassList.Sorted = False
		Me.ClassList.TabStop = True
		Me.ClassList.Visible = True
		Me.ClassList.Name = "ClassList"
		Me.CustomerList.Size = New System.Drawing.Size(361, 21)
		Me.CustomerList.Location = New System.Drawing.Point(184, 96)
		Me.CustomerList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.CustomerList.TabIndex = 1
		Me.CustomerList.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.CustomerList.BackColor = System.Drawing.SystemColors.Window
		Me.CustomerList.CausesValidation = True
		Me.CustomerList.Enabled = True
		Me.CustomerList.ForeColor = System.Drawing.SystemColors.WindowText
		Me.CustomerList.IntegralHeight = True
		Me.CustomerList.Cursor = System.Windows.Forms.Cursors.Default
		Me.CustomerList.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.CustomerList.Sorted = False
		Me.CustomerList.TabStop = True
		Me.CustomerList.Visible = True
		Me.CustomerList.Name = "CustomerList"
		Me.AccountList.Size = New System.Drawing.Size(361, 21)
		Me.AccountList.Location = New System.Drawing.Point(184, 64)
		Me.AccountList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.AccountList.TabIndex = 0
		Me.AccountList.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AccountList.BackColor = System.Drawing.SystemColors.Window
		Me.AccountList.CausesValidation = True
		Me.AccountList.Enabled = True
		Me.AccountList.ForeColor = System.Drawing.SystemColors.WindowText
		Me.AccountList.IntegralHeight = True
		Me.AccountList.Cursor = System.Windows.Forms.Cursors.Default
		Me.AccountList.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.AccountList.Sorted = False
		Me.AccountList.TabStop = True
		Me.AccountList.Visible = True
		Me.AccountList.Name = "AccountList"
		Me.ClassLabel.Text = "Select Class[Optional]:"
		Me.ClassLabel.Enabled = False
		Me.ClassLabel.Size = New System.Drawing.Size(113, 21)
		Me.ClassLabel.Location = New System.Drawing.Point(64, 128)
		Me.ClassLabel.TabIndex = 20
		Me.ClassLabel.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ClassLabel.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.ClassLabel.BackColor = System.Drawing.SystemColors.Control
		Me.ClassLabel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.ClassLabel.Cursor = System.Windows.Forms.Cursors.Default
		Me.ClassLabel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ClassLabel.UseMnemonic = True
		Me.ClassLabel.Visible = True
		Me.ClassLabel.AutoSize = False
		Me.ClassLabel.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.ClassLabel.Name = "ClassLabel"
		Me.Label6.Text = "Select Customer [Optional]:"
		Me.Label6.Size = New System.Drawing.Size(129, 21)
		Me.Label6.Location = New System.Drawing.Point(48, 96)
		Me.Label6.TabIndex = 19
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
		Me.Label5.Text = "Select Account:"
		Me.Label5.Size = New System.Drawing.Size(81, 21)
		Me.Label5.Location = New System.Drawing.Point(96, 64)
		Me.Label5.TabIndex = 18
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
		Me.Label4.Text = "Line Items:"
		Me.Label4.Size = New System.Drawing.Size(121, 17)
		Me.Label4.Location = New System.Drawing.Point(24, 352)
		Me.Label4.TabIndex = 16
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
		Me.Controls.Add(WarningMsg)
		Me.Controls.Add(SendBtn)
		Me.Controls.Add(Frame1)
		Me.Controls.Add(LineItems)
		Me.Controls.Add(Memo)
		Me.Controls.Add(ClassList)
		Me.Controls.Add(CustomerList)
		Me.Controls.Add(AccountList)
		Me.Controls.Add(ClassLabel)
		Me.Controls.Add(Label6)
		Me.Controls.Add(Label5)
		Me.Controls.Add(Label4)
		Me.Frame1.Controls.Add(AddBtn)
		Me.Frame1.Controls.Add(ValueAdjust)
		Me.Frame1.Controls.Add(QuantityAdjust)
		Me.Frame1.Controls.Add(ItemList)
		Me.Frame1.Controls.Add(ValueOption)
		Me.Frame1.Controls.Add(QuantityOption)
		Me.Frame1.Controls.Add(DiffCheck)
		Me.Frame1.Controls.Add(ValueLabel)
		Me.Frame1.Controls.Add(Label2)
		Me.Frame1.Controls.Add(Label1)
		Me.LineItems.Columns.Add(_LineItems_ColumnHeader_1)
		Me.LineItems.Columns.Add(_LineItems_ColumnHeader_2)
		Me.LineItems.Columns.Add(_LineItems_ColumnHeader_3)
		Me.LineItems.Columns.Add(_LineItems_ColumnHeader_4)
		Me.Frame1.ResumeLayout(False)
		Me.LineItems.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class