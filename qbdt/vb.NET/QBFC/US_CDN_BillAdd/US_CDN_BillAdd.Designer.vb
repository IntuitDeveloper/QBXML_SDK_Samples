<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class BillAdd
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
	Public WithEvents EnterBill As System.Windows.Forms.Button
	Public WithEvents BillAmount As System.Windows.Forms.TextBox
	Public WithEvents AccountList As System.Windows.Forms.ComboBox
	Public WithEvents VendorList As System.Windows.Forms.ComboBox
	Public WithEvents CommPrefs As System.Windows.Forms.Button
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(BillAdd))
		Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
        Me.EnterBill = New System.Windows.Forms.Button
		Me.BillAmount = New System.Windows.Forms.TextBox
		Me.AccountList = New System.Windows.Forms.ComboBox
		Me.VendorList = New System.Windows.Forms.ComboBox
		Me.CommPrefs = New System.Windows.Forms.Button
		Me.Label4 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "BillAdd Sample"
		Me.ClientSize = New System.Drawing.Size(479, 353)
		Me.Location = New System.Drawing.Point(4, 30)
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
		Me.Name = "BillAdd"
		Me.EnterBill.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.EnterBill.Text = "Enter Bill!"
		Me.EnterBill.Size = New System.Drawing.Size(185, 33)
		Me.EnterBill.Location = New System.Drawing.Point(104, 296)
		Me.EnterBill.TabIndex = 8
		Me.EnterBill.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.EnterBill.BackColor = System.Drawing.SystemColors.Control
		Me.EnterBill.CausesValidation = True
		Me.EnterBill.Enabled = True
		Me.EnterBill.ForeColor = System.Drawing.SystemColors.ControlText
		Me.EnterBill.Cursor = System.Windows.Forms.Cursors.Default
		Me.EnterBill.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.EnterBill.TabStop = True
		Me.EnterBill.Name = "EnterBill"
		Me.BillAmount.AutoSize = False
		Me.BillAmount.Size = New System.Drawing.Size(241, 25)
		Me.BillAmount.Location = New System.Drawing.Point(216, 232)
		Me.BillAmount.TabIndex = 6
		Me.BillAmount.Text = "0.00"
		Me.BillAmount.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.BillAmount.AcceptsReturn = True
		Me.BillAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.BillAmount.BackColor = System.Drawing.SystemColors.Window
		Me.BillAmount.CausesValidation = True
		Me.BillAmount.Enabled = True
		Me.BillAmount.ForeColor = System.Drawing.SystemColors.WindowText
		Me.BillAmount.HideSelection = True
		Me.BillAmount.ReadOnly = False
		Me.BillAmount.Maxlength = 0
		Me.BillAmount.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.BillAmount.MultiLine = False
		Me.BillAmount.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.BillAmount.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.BillAmount.TabStop = True
		Me.BillAmount.Visible = True
		Me.BillAmount.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.BillAmount.Name = "BillAmount"
		Me.AccountList.Size = New System.Drawing.Size(241, 21)
		Me.AccountList.Location = New System.Drawing.Point(216, 200)
		Me.AccountList.TabIndex = 4
		Me.AccountList.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AccountList.BackColor = System.Drawing.SystemColors.Window
		Me.AccountList.CausesValidation = True
		Me.AccountList.Enabled = True
		Me.AccountList.ForeColor = System.Drawing.SystemColors.WindowText
		Me.AccountList.IntegralHeight = True
		Me.AccountList.Cursor = System.Windows.Forms.Cursors.Default
		Me.AccountList.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.AccountList.Sorted = False
		Me.AccountList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.AccountList.TabStop = True
		Me.AccountList.Visible = True
		Me.AccountList.Name = "AccountList"
		Me.VendorList.Size = New System.Drawing.Size(241, 21)
		Me.VendorList.Location = New System.Drawing.Point(216, 168)
		Me.VendorList.TabIndex = 2
		Me.VendorList.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.VendorList.BackColor = System.Drawing.SystemColors.Window
		Me.VendorList.CausesValidation = True
		Me.VendorList.Enabled = True
		Me.VendorList.ForeColor = System.Drawing.SystemColors.WindowText
		Me.VendorList.IntegralHeight = True
		Me.VendorList.Cursor = System.Windows.Forms.Cursors.Default
		Me.VendorList.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.VendorList.Sorted = False
		Me.VendorList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.VendorList.TabStop = True
		Me.VendorList.Visible = True
		Me.VendorList.Name = "VendorList"
		Me.CommPrefs.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CommPrefs.Text = "Change Communication Preferences"
		Me.CommPrefs.Size = New System.Drawing.Size(377, 33)
		Me.CommPrefs.Location = New System.Drawing.Point(48, 96)
		Me.CommPrefs.TabIndex = 0
		Me.CommPrefs.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.CommPrefs.BackColor = System.Drawing.SystemColors.Control
		Me.CommPrefs.CausesValidation = True
		Me.CommPrefs.Enabled = True
		Me.CommPrefs.ForeColor = System.Drawing.SystemColors.ControlText
		Me.CommPrefs.Cursor = System.Windows.Forms.Cursors.Default
		Me.CommPrefs.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.CommPrefs.TabStop = True
		Me.CommPrefs.Name = "CommPrefs"
		Me.Label4.Text = "Specify the amount of the bill:"
		Me.Label4.Size = New System.Drawing.Size(153, 17)
		Me.Label4.Location = New System.Drawing.Point(56, 240)
		Me.Label4.TabIndex = 7
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
		Me.Label3.Text = "Select an item for the bill:"
		Me.Label3.Size = New System.Drawing.Size(121, 17)
		Me.Label3.Location = New System.Drawing.Point(88, 200)
		Me.Label3.TabIndex = 5
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
		Me.Label2.Text = "Select a vendor for the bill:"
		Me.Label2.Size = New System.Drawing.Size(129, 17)
		Me.Label2.Location = New System.Drawing.Point(80, 168)
		Me.Label2.TabIndex = 3
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
        Me.Label1.Text = "This example adds a simple (one line item) bill to QuickBooks using the vendor, item, and amount of the bill that you specify below.  Note that the vendor list and item list are populated from data in the QuickBooks file that you are connected to.  Click ""Set Communication Preferences"" to specify whether you wish to communicate with QuickBooks Online Edition or QuickBooks running on the desktop."
		Me.Label1.Size = New System.Drawing.Size(457, 73)
		Me.Label1.Location = New System.Drawing.Point(8, 8)
		Me.Label1.TabIndex = 1
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
		Me.Controls.Add(EnterBill)
		Me.Controls.Add(BillAmount)
		Me.Controls.Add(AccountList)
		Me.Controls.Add(VendorList)
		Me.Controls.Add(CommPrefs)
		Me.Controls.Add(Label4)
		Me.Controls.Add(Label3)
        Me.Controls.Add(Label2)
        Me.Controls.Add(Label1)
        Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class