<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class Main_Renamed
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
	Public WithEvents Comm_Browse As System.Windows.Forms.Button
	Public CommDlg_BrowseOpen As System.Windows.Forms.OpenFileDialog
	Public WithEvents CompanyFileNameInput As System.Windows.Forms.TextBox
	Public WithEvents Comm_Exit As System.Windows.Forms.Button
	Public WithEvents Comm_Submit As System.Windows.Forms.Button
	Public WithEvents Image_QBBANNER As System.Windows.Forms.PictureBox
	Public WithEvents Label2 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Main_Renamed))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Comm_Browse = New System.Windows.Forms.Button()
        Me.CommDlg_BrowseOpen = New System.Windows.Forms.OpenFileDialog()
        Me.CompanyFileNameInput = New System.Windows.Forms.TextBox()
        Me.Comm_Exit = New System.Windows.Forms.Button()
        Me.Comm_Submit = New System.Windows.Forms.Button()
        Me.Image_QBBANNER = New System.Windows.Forms.PictureBox()
        Me.Label2 = New System.Windows.Forms.Label()
        CType(Me.Image_QBBANNER, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Comm_Browse
        '
        Me.Comm_Browse.BackColor = System.Drawing.SystemColors.Control
        Me.Comm_Browse.Cursor = System.Windows.Forms.Cursors.Default
        Me.Comm_Browse.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Comm_Browse.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Comm_Browse.Location = New System.Drawing.Point(376, 72)
        Me.Comm_Browse.Name = "Comm_Browse"
        Me.Comm_Browse.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Comm_Browse.Size = New System.Drawing.Size(73, 25)
        Me.Comm_Browse.TabIndex = 4
        Me.Comm_Browse.Text = "&Browse..."
        Me.Comm_Browse.UseVisualStyleBackColor = False
        '
        'CompanyFileNameInput
        '
        Me.CompanyFileNameInput.AcceptsReturn = True
        Me.CompanyFileNameInput.BackColor = System.Drawing.SystemColors.Window
        Me.CompanyFileNameInput.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.CompanyFileNameInput.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CompanyFileNameInput.ForeColor = System.Drawing.SystemColors.WindowText
        Me.CompanyFileNameInput.Location = New System.Drawing.Point(32, 72)
        Me.CompanyFileNameInput.MaxLength = 0
        Me.CompanyFileNameInput.Name = "CompanyFileNameInput"
        Me.CompanyFileNameInput.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.CompanyFileNameInput.Size = New System.Drawing.Size(329, 23)
        Me.CompanyFileNameInput.TabIndex = 3
        Me.CompanyFileNameInput.Text = "c:\program files\intuit\quickbooks pro\sample_product-based business.qbw"
        '
        'Comm_Exit
        '
        Me.Comm_Exit.BackColor = System.Drawing.SystemColors.Control
        Me.Comm_Exit.Cursor = System.Windows.Forms.Cursors.Default
        Me.Comm_Exit.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Comm_Exit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Comm_Exit.Location = New System.Drawing.Point(272, 120)
        Me.Comm_Exit.Name = "Comm_Exit"
        Me.Comm_Exit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Comm_Exit.Size = New System.Drawing.Size(89, 25)
        Me.Comm_Exit.TabIndex = 2
        Me.Comm_Exit.Text = "&Exit"
        Me.Comm_Exit.UseVisualStyleBackColor = False
        '
        'Comm_Submit
        '
        Me.Comm_Submit.BackColor = System.Drawing.SystemColors.Control
        Me.Comm_Submit.Cursor = System.Windows.Forms.Cursors.Default
        Me.Comm_Submit.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Comm_Submit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Comm_Submit.Location = New System.Drawing.Point(88, 120)
        Me.Comm_Submit.Name = "Comm_Submit"
        Me.Comm_Submit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Comm_Submit.Size = New System.Drawing.Size(89, 25)
        Me.Comm_Submit.TabIndex = 1
        Me.Comm_Submit.Text = "&Next"
        Me.Comm_Submit.UseVisualStyleBackColor = False
        '
        'Image_QBBANNER
        '
        Me.Image_QBBANNER.Cursor = System.Windows.Forms.Cursors.Default
        Me.Image_QBBANNER.Image = CType(resources.GetObject("Image_QBBANNER.Image"), System.Drawing.Image)
        Me.Image_QBBANNER.Location = New System.Drawing.Point(24, 176)
        Me.Image_QBBANNER.Name = "Image_QBBANNER"
        Me.Image_QBBANNER.Size = New System.Drawing.Size(425, 28)
        Me.Image_QBBANNER.TabIndex = 5
        Me.Image_QBBANNER.TabStop = False
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(32, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(265, 17)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Please enter the QB file name in full path:"
        '
        'Main_Renamed
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(479, 229)
        Me.Controls.Add(Me.Comm_Browse)
        Me.Controls.Add(Me.CompanyFileNameInput)
        Me.Controls.Add(Me.Comm_Exit)
        Me.Controls.Add(Me.Comm_Submit)
        Me.Controls.Add(Me.Image_QBBANNER)
        Me.Controls.Add(Me.Label2)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(3, 22)
        Me.Name = "Main_Renamed"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "qbXML Sample: Modify Customer"
        CType(Me.Image_QBBANNER, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
End Class