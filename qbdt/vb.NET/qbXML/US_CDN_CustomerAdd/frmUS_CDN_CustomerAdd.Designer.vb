<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmUS_CDN_CustomerAdd
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
	Public WithEvents cmdExit As System.Windows.Forms.Button
	Public WithEvents cmdClearForm As System.Windows.Forms.Button
	Public WithEvents cmdAddCustomer As System.Windows.Forms.Button
	Public WithEvents txtPostalCode As System.Windows.Forms.TextBox
	Public WithEvents txtStateProvince As System.Windows.Forms.TextBox
	Public WithEvents txtCity As System.Windows.Forms.TextBox
	Public WithEvents txtAddr4 As System.Windows.Forms.TextBox
	Public WithEvents txtAddr3 As System.Windows.Forms.TextBox
	Public WithEvents txtAddr2 As System.Windows.Forms.TextBox
	Public WithEvents txtAddr1 As System.Windows.Forms.TextBox
	Public WithEvents txtCustomer As System.Windows.Forms.TextBox
	Public WithEvents Label8 As System.Windows.Forms.Label
	Public WithEvents Label7 As System.Windows.Forms.Label
	Public WithEvents lblStateProvince As System.Windows.Forms.Label
	Public WithEvents Label6 As System.Windows.Forms.Label
	Public WithEvents lblAddr4 As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents lblNationality As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.cmdClearForm = New System.Windows.Forms.Button()
        Me.cmdAddCustomer = New System.Windows.Forms.Button()
        Me.txtPostalCode = New System.Windows.Forms.TextBox()
        Me.txtStateProvince = New System.Windows.Forms.TextBox()
        Me.txtCity = New System.Windows.Forms.TextBox()
        Me.txtAddr4 = New System.Windows.Forms.TextBox()
        Me.txtAddr3 = New System.Windows.Forms.TextBox()
        Me.txtAddr2 = New System.Windows.Forms.TextBox()
        Me.txtAddr1 = New System.Windows.Forms.TextBox()
        Me.txtCustomer = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.lblStateProvince = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.lblAddr4 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblNationality = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'cmdExit
        '
        Me.cmdExit.BackColor = System.Drawing.SystemColors.Control
        Me.cmdExit.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdExit.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdExit.Location = New System.Drawing.Point(248, 296)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdExit.Size = New System.Drawing.Size(113, 33)
        Me.cmdExit.TabIndex = 20
        Me.cmdExit.Text = "Exit"
        Me.cmdExit.UseVisualStyleBackColor = False
        '
        'cmdClearForm
        '
        Me.cmdClearForm.BackColor = System.Drawing.SystemColors.Control
        Me.cmdClearForm.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdClearForm.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdClearForm.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdClearForm.Location = New System.Drawing.Point(128, 296)
        Me.cmdClearForm.Name = "cmdClearForm"
        Me.cmdClearForm.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdClearForm.Size = New System.Drawing.Size(113, 33)
        Me.cmdClearForm.TabIndex = 19
        Me.cmdClearForm.Text = "Clear Form"
        Me.cmdClearForm.UseVisualStyleBackColor = False
        '
        'cmdAddCustomer
        '
        Me.cmdAddCustomer.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAddCustomer.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAddCustomer.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAddCustomer.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAddCustomer.Location = New System.Drawing.Point(8, 296)
        Me.cmdAddCustomer.Name = "cmdAddCustomer"
        Me.cmdAddCustomer.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAddCustomer.Size = New System.Drawing.Size(113, 33)
        Me.cmdAddCustomer.TabIndex = 18
        Me.cmdAddCustomer.Text = "Add Customer"
        Me.cmdAddCustomer.UseVisualStyleBackColor = False
        '
        'txtPostalCode
        '
        Me.txtPostalCode.AcceptsReturn = True
        Me.txtPostalCode.BackColor = System.Drawing.SystemColors.Window
        Me.txtPostalCode.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPostalCode.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPostalCode.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPostalCode.Location = New System.Drawing.Point(88, 256)
        Me.txtPostalCode.MaxLength = 0
        Me.txtPostalCode.Name = "txtPostalCode"
        Me.txtPostalCode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPostalCode.Size = New System.Drawing.Size(169, 19)
        Me.txtPostalCode.TabIndex = 17
        '
        'txtStateProvince
        '
        Me.txtStateProvince.AcceptsReturn = True
        Me.txtStateProvince.BackColor = System.Drawing.SystemColors.Window
        Me.txtStateProvince.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtStateProvince.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStateProvince.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtStateProvince.Location = New System.Drawing.Point(88, 232)
        Me.txtStateProvince.MaxLength = 0
        Me.txtStateProvince.Name = "txtStateProvince"
        Me.txtStateProvince.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtStateProvince.Size = New System.Drawing.Size(169, 19)
        Me.txtStateProvince.TabIndex = 16
        '
        'txtCity
        '
        Me.txtCity.AcceptsReturn = True
        Me.txtCity.BackColor = System.Drawing.SystemColors.Window
        Me.txtCity.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCity.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCity.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCity.Location = New System.Drawing.Point(88, 208)
        Me.txtCity.MaxLength = 0
        Me.txtCity.Name = "txtCity"
        Me.txtCity.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCity.Size = New System.Drawing.Size(169, 19)
        Me.txtCity.TabIndex = 15
        '
        'txtAddr4
        '
        Me.txtAddr4.AcceptsReturn = True
        Me.txtAddr4.BackColor = System.Drawing.SystemColors.Window
        Me.txtAddr4.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtAddr4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAddr4.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtAddr4.Location = New System.Drawing.Point(88, 184)
        Me.txtAddr4.MaxLength = 0
        Me.txtAddr4.Name = "txtAddr4"
        Me.txtAddr4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtAddr4.Size = New System.Drawing.Size(169, 19)
        Me.txtAddr4.TabIndex = 14
        '
        'txtAddr3
        '
        Me.txtAddr3.AcceptsReturn = True
        Me.txtAddr3.BackColor = System.Drawing.SystemColors.Window
        Me.txtAddr3.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtAddr3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAddr3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtAddr3.Location = New System.Drawing.Point(88, 160)
        Me.txtAddr3.MaxLength = 0
        Me.txtAddr3.Name = "txtAddr3"
        Me.txtAddr3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtAddr3.Size = New System.Drawing.Size(169, 19)
        Me.txtAddr3.TabIndex = 13
        '
        'txtAddr2
        '
        Me.txtAddr2.AcceptsReturn = True
        Me.txtAddr2.BackColor = System.Drawing.SystemColors.Window
        Me.txtAddr2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtAddr2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAddr2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtAddr2.Location = New System.Drawing.Point(88, 136)
        Me.txtAddr2.MaxLength = 0
        Me.txtAddr2.Name = "txtAddr2"
        Me.txtAddr2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtAddr2.Size = New System.Drawing.Size(169, 19)
        Me.txtAddr2.TabIndex = 12
        '
        'txtAddr1
        '
        Me.txtAddr1.AcceptsReturn = True
        Me.txtAddr1.BackColor = System.Drawing.SystemColors.Window
        Me.txtAddr1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtAddr1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAddr1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtAddr1.Location = New System.Drawing.Point(88, 112)
        Me.txtAddr1.MaxLength = 0
        Me.txtAddr1.Name = "txtAddr1"
        Me.txtAddr1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtAddr1.Size = New System.Drawing.Size(169, 19)
        Me.txtAddr1.TabIndex = 11
        '
        'txtCustomer
        '
        Me.txtCustomer.AcceptsReturn = True
        Me.txtCustomer.BackColor = System.Drawing.SystemColors.Window
        Me.txtCustomer.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCustomer.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustomer.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCustomer.Location = New System.Drawing.Point(96, 64)
        Me.txtCustomer.MaxLength = 0
        Me.txtCustomer.Name = "txtCustomer"
        Me.txtCustomer.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCustomer.Size = New System.Drawing.Size(233, 19)
        Me.txtCustomer.TabIndex = 10
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.SystemColors.Control
        Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label8.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label8.Location = New System.Drawing.Point(24, 256)
        Me.Label8.Name = "Label8"
        Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label8.Size = New System.Drawing.Size(57, 25)
        Me.Label8.TabIndex = 9
        Me.Label8.Text = "Postal Code"
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(8, 88)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(81, 17)
        Me.Label7.TabIndex = 8
        Me.Label7.Text = "Billing Address"
        '
        'lblStateProvince
        '
        Me.lblStateProvince.BackColor = System.Drawing.SystemColors.Control
        Me.lblStateProvince.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblStateProvince.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStateProvince.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblStateProvince.Location = New System.Drawing.Point(24, 232)
        Me.lblStateProvince.Name = "lblStateProvince"
        Me.lblStateProvince.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblStateProvince.Size = New System.Drawing.Size(49, 17)
        Me.lblStateProvince.TabIndex = 7
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(24, 208)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(33, 17)
        Me.Label6.TabIndex = 6
        Me.Label6.Text = "City"
        '
        'lblAddr4
        '
        Me.lblAddr4.BackColor = System.Drawing.SystemColors.Control
        Me.lblAddr4.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblAddr4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAddr4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblAddr4.Location = New System.Drawing.Point(24, 184)
        Me.lblAddr4.Name = "lblAddr4"
        Me.lblAddr4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblAddr4.Size = New System.Drawing.Size(33, 17)
        Me.lblAddr4.TabIndex = 5
        Me.lblAddr4.Text = "Addr4"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(24, 160)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(33, 17)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "Addr3"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(24, 136)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(33, 17)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Addr2"
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
        Me.Label2.Size = New System.Drawing.Size(33, 17)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Addr1"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(8, 64)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(81, 25)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Customer Name"
        '
        'lblNationality
        '
        Me.lblNationality.BackColor = System.Drawing.SystemColors.Control
        Me.lblNationality.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblNationality.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNationality.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblNationality.Location = New System.Drawing.Point(8, 16)
        Me.lblNationality.Name = "lblNationality"
        Me.lblNationality.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblNationality.Size = New System.Drawing.Size(353, 17)
        Me.lblNationality.TabIndex = 0
        Me.lblNationality.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'frmUS_CDN_CustomerAdd
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(370, 359)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.cmdClearForm)
        Me.Controls.Add(Me.cmdAddCustomer)
        Me.Controls.Add(Me.txtPostalCode)
        Me.Controls.Add(Me.txtStateProvince)
        Me.Controls.Add(Me.txtCity)
        Me.Controls.Add(Me.txtAddr4)
        Me.Controls.Add(Me.txtAddr3)
        Me.Controls.Add(Me.txtAddr2)
        Me.Controls.Add(Me.txtAddr1)
        Me.Controls.Add(Me.txtCustomer)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.lblStateProvince)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.lblAddr4)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lblNationality)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(4, 23)
        Me.Name = "frmUS_CDN_CustomerAdd"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "US / Canadian CustomerAdd"
        Me.ResumeLayout(False)

    End Sub
#End Region
End Class