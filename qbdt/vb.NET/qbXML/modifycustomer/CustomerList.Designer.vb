<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class CustomerList
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
	Public WithEvents Combo_CustList As System.Windows.Forms.ComboBox
	Public WithEvents Comm_Exit As System.Windows.Forms.Button
	Public WithEvents Comm_View_res As System.Windows.Forms.Button
	Public WithEvents Comm_View_Req As System.Windows.Forms.Button
	Public WithEvents Comm_Submit As System.Windows.Forms.Button
	Public WithEvents Label7 As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents Image_QBBANNER As System.Windows.Forms.PictureBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(CustomerList))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Combo_CustList = New System.Windows.Forms.ComboBox()
        Me.Comm_Exit = New System.Windows.Forms.Button()
        Me.Comm_View_res = New System.Windows.Forms.Button()
        Me.Comm_View_Req = New System.Windows.Forms.Button()
        Me.Comm_Submit = New System.Windows.Forms.Button()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Image_QBBANNER = New System.Windows.Forms.PictureBox()
        Me.Frame1.SuspendLayout()
        CType(Me.Image_QBBANNER, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Combo_CustList
        '
        Me.Combo_CustList.BackColor = System.Drawing.SystemColors.Window
        Me.Combo_CustList.Cursor = System.Windows.Forms.Cursors.Default
        Me.Combo_CustList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Combo_CustList.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Combo_CustList.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Combo_CustList.Location = New System.Drawing.Point(152, 55)
        Me.Combo_CustList.Name = "Combo_CustList"
        Me.Combo_CustList.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Combo_CustList.Size = New System.Drawing.Size(169, 24)
        Me.Combo_CustList.TabIndex = 4
        '
        'Comm_Exit
        '
        Me.Comm_Exit.BackColor = System.Drawing.SystemColors.Control
        Me.Comm_Exit.Cursor = System.Windows.Forms.Cursors.Default
        Me.Comm_Exit.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Comm_Exit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Comm_Exit.Location = New System.Drawing.Point(344, 64)
        Me.Comm_Exit.Name = "Comm_Exit"
        Me.Comm_Exit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Comm_Exit.Size = New System.Drawing.Size(121, 25)
        Me.Comm_Exit.TabIndex = 3
        Me.Comm_Exit.Text = "&Exit"
        Me.Comm_Exit.UseVisualStyleBackColor = False
        '
        'Comm_View_res
        '
        Me.Comm_View_res.BackColor = System.Drawing.SystemColors.Control
        Me.Comm_View_res.Cursor = System.Windows.Forms.Cursors.Default
        Me.Comm_View_res.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Comm_View_res.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Comm_View_res.Location = New System.Drawing.Point(344, 152)
        Me.Comm_View_res.Name = "Comm_View_res"
        Me.Comm_View_res.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Comm_View_res.Size = New System.Drawing.Size(121, 32)
        Me.Comm_View_res.TabIndex = 2
        Me.Comm_View_res.Text = "View qbXML Re&sponse"
        Me.Comm_View_res.UseVisualStyleBackColor = False
        '
        'Comm_View_Req
        '
        Me.Comm_View_Req.BackColor = System.Drawing.SystemColors.Control
        Me.Comm_View_Req.Cursor = System.Windows.Forms.Cursors.Default
        Me.Comm_View_Req.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Comm_View_Req.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Comm_View_Req.Location = New System.Drawing.Point(344, 120)
        Me.Comm_View_Req.Name = "Comm_View_Req"
        Me.Comm_View_Req.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Comm_View_Req.Size = New System.Drawing.Size(121, 25)
        Me.Comm_View_Req.TabIndex = 1
        Me.Comm_View_Req.Text = "View qbXML  Re&quest"
        Me.Comm_View_Req.UseVisualStyleBackColor = False
        '
        'Comm_Submit
        '
        Me.Comm_Submit.BackColor = System.Drawing.SystemColors.Control
        Me.Comm_Submit.Cursor = System.Windows.Forms.Cursors.Default
        Me.Comm_Submit.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Comm_Submit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Comm_Submit.Location = New System.Drawing.Point(344, 32)
        Me.Comm_Submit.Name = "Comm_Submit"
        Me.Comm_Submit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Comm_Submit.Size = New System.Drawing.Size(121, 25)
        Me.Comm_Submit.TabIndex = 0
        Me.Comm_Submit.Text = "&Modify"
        Me.Comm_Submit.UseVisualStyleBackColor = False
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.Label7)
        Me.Frame1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(16, 32)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(313, 129)
        Me.Frame1.TabIndex = 5
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Please select the customer whose data you want to modify: "
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(8, 27)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(121, 25)
        Me.Label7.TabIndex = 6
        Me.Label7.Text = "Customer Full Name List:"
        '
        'Image_QBBANNER
        '
        Me.Image_QBBANNER.Cursor = System.Windows.Forms.Cursors.Default
        Me.Image_QBBANNER.Image = CType(resources.GetObject("Image_QBBANNER.Image"), System.Drawing.Image)
        Me.Image_QBBANNER.Location = New System.Drawing.Point(16, 200)
        Me.Image_QBBANNER.Name = "Image_QBBANNER"
        Me.Image_QBBANNER.Size = New System.Drawing.Size(429, 25)
        Me.Image_QBBANNER.TabIndex = 6
        Me.Image_QBBANNER.TabStop = False
        '
        'CustomerList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(477, 246)
        Me.Controls.Add(Me.Combo_CustList)
        Me.Controls.Add(Me.Comm_Exit)
        Me.Controls.Add(Me.Comm_View_res)
        Me.Controls.Add(Me.Comm_View_Req)
        Me.Controls.Add(Me.Comm_Submit)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.Image_QBBANNER)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(3, 22)
        Me.Name = "CustomerList"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "qbXML Sample: Select Customer"
        Me.Frame1.ResumeLayout(False)
        CType(Me.Image_QBBANNER, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
#End Region
End Class