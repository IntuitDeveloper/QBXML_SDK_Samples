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
	Public WithEvents showResButton As System.Windows.Forms.Button
	Public WithEvents showReqButton As System.Windows.Forms.Button
	Public WithEvents exitButton As System.Windows.Forms.Button
	Public WithEvents goButton As System.Windows.Forms.Button
	Public WithEvents qbFile As System.Windows.Forms.TextBox
	Public WithEvents browseButton As System.Windows.Forms.Button
	Public WithEvents htmlFile As System.Windows.Forms.TextBox
    Public browseDialogOpen As System.Windows.Forms.OpenFileDialog
    Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.showResButton = New System.Windows.Forms.Button()
        Me.showReqButton = New System.Windows.Forms.Button()
        Me.exitButton = New System.Windows.Forms.Button()
        Me.goButton = New System.Windows.Forms.Button()
        Me.qbFile = New System.Windows.Forms.TextBox()
        Me.browseButton = New System.Windows.Forms.Button()
        Me.htmlFile = New System.Windows.Forms.TextBox()
        Me.browseDialogOpen = New System.Windows.Forms.OpenFileDialog()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'showResButton
        '
        Me.showResButton.BackColor = System.Drawing.SystemColors.Control
        Me.showResButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.showResButton.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.showResButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.showResButton.Location = New System.Drawing.Point(420, 84)
        Me.showResButton.Name = "showResButton"
        Me.showResButton.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.showResButton.Size = New System.Drawing.Size(145, 25)
        Me.showResButton.TabIndex = 10
        Me.showResButton.Text = "Show Response qbXML"
        Me.showResButton.UseVisualStyleBackColor = False
        '
        'showReqButton
        '
        Me.showReqButton.BackColor = System.Drawing.SystemColors.Control
        Me.showReqButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.showReqButton.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.showReqButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.showReqButton.Location = New System.Drawing.Point(248, 84)
        Me.showReqButton.Name = "showReqButton"
        Me.showReqButton.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.showReqButton.Size = New System.Drawing.Size(145, 25)
        Me.showReqButton.TabIndex = 9
        Me.showReqButton.Text = "Show Request qbXML"
        Me.showReqButton.UseVisualStyleBackColor = False
        '
        'exitButton
        '
        Me.exitButton.BackColor = System.Drawing.SystemColors.Control
        Me.exitButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.exitButton.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.exitButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.exitButton.Location = New System.Drawing.Point(592, 84)
        Me.exitButton.Name = "exitButton"
        Me.exitButton.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.exitButton.Size = New System.Drawing.Size(145, 25)
        Me.exitButton.TabIndex = 8
        Me.exitButton.Text = "&Exit Program"
        Me.exitButton.UseVisualStyleBackColor = False
        '
        'goButton
        '
        Me.goButton.BackColor = System.Drawing.SystemColors.Control
        Me.goButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.goButton.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.goButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.goButton.Location = New System.Drawing.Point(76, 84)
        Me.goButton.Name = "goButton"
        Me.goButton.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.goButton.Size = New System.Drawing.Size(145, 25)
        Me.goButton.TabIndex = 4
        Me.goButton.Text = "&Generate Report"
        Me.goButton.UseVisualStyleBackColor = False
        '
        'qbFile
        '
        Me.qbFile.AcceptsReturn = True
        Me.qbFile.BackColor = System.Drawing.SystemColors.Window
        Me.qbFile.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.qbFile.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.qbFile.ForeColor = System.Drawing.SystemColors.WindowText
        Me.qbFile.Location = New System.Drawing.Point(292, 16)
        Me.qbFile.MaxLength = 0
        Me.qbFile.Name = "qbFile"
        Me.qbFile.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.qbFile.Size = New System.Drawing.Size(241, 25)
        Me.qbFile.TabIndex = 3
        '
        'browseButton
        '
        Me.browseButton.BackColor = System.Drawing.SystemColors.Control
        Me.browseButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.browseButton.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.browseButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.browseButton.Location = New System.Drawing.Point(536, 16)
        Me.browseButton.Name = "browseButton"
        Me.browseButton.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.browseButton.Size = New System.Drawing.Size(105, 25)
        Me.browseButton.TabIndex = 2
        Me.browseButton.Text = "&Browse"
        Me.browseButton.UseVisualStyleBackColor = False
        '
        'htmlFile
        '
        Me.htmlFile.AcceptsReturn = True
        Me.htmlFile.BackColor = System.Drawing.SystemColors.Window
        Me.htmlFile.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.htmlFile.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.htmlFile.ForeColor = System.Drawing.SystemColors.WindowText
        Me.htmlFile.Location = New System.Drawing.Point(292, 48)
        Me.htmlFile.MaxLength = 0
        Me.htmlFile.Name = "htmlFile"
        Me.htmlFile.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.htmlFile.Size = New System.Drawing.Size(349, 25)
        Me.htmlFile.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(124, 28)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(165, 13)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "(blank to use currently open file)"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(124, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(165, 21)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Select QuickBooks Company File:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(124, 52)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(165, 17)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Select Output Html File:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'UIForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(812, 159)
        Me.Controls.Add(Me.showResButton)
        Me.Controls.Add(Me.showReqButton)
        Me.Controls.Add(Me.exitButton)
        Me.Controls.Add(Me.goButton)
        Me.Controls.Add(Me.qbFile)
        Me.Controls.Add(Me.browseButton)
        Me.Controls.Add(Me.htmlFile)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label2)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(4, 23)
        Me.Name = "UIForm"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "QBSDK 2.0 --- Expense By Vendor Summary Report"
        Me.ResumeLayout(False)

    End Sub
#End Region
End Class