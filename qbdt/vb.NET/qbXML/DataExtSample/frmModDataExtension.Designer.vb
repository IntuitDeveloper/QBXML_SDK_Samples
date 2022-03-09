<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmModDataExtension
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
	Public WithEvents cmdModResponse As System.Windows.Forms.Button
	Public WithEvents cmdModRequest As System.Windows.Forms.Button
	Public WithEvents Text1 As System.Windows.Forms.TextBox
	Public WithEvents lstCustomers As System.Windows.Forms.ListBox
	Public WithEvents lstUsedDataExts As System.Windows.Forms.ListBox
	Public WithEvents txtDataExtValue As System.Windows.Forms.TextBox
	Public WithEvents cmdModValue As System.Windows.Forms.Button
	Public WithEvents cmdCloseWindow As System.Windows.Forms.Button
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmModDataExtension))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdModResponse = New System.Windows.Forms.Button()
        Me.cmdModRequest = New System.Windows.Forms.Button()
        Me.Text1 = New System.Windows.Forms.TextBox()
        Me.lstCustomers = New System.Windows.Forms.ListBox()
        Me.lstUsedDataExts = New System.Windows.Forms.ListBox()
        Me.txtDataExtValue = New System.Windows.Forms.TextBox()
        Me.cmdModValue = New System.Windows.Forms.Button()
        Me.cmdCloseWindow = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'cmdModResponse
        '
        Me.cmdModResponse.BackColor = System.Drawing.SystemColors.Control
        Me.cmdModResponse.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdModResponse.Enabled = False
        Me.cmdModResponse.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdModResponse.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdModResponse.Location = New System.Drawing.Point(136, 472)
        Me.cmdModResponse.Name = "cmdModResponse"
        Me.cmdModResponse.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdModResponse.Size = New System.Drawing.Size(121, 41)
        Me.cmdModResponse.TabIndex = 10
        Me.cmdModResponse.Text = "Show Modify Response"
        Me.cmdModResponse.UseVisualStyleBackColor = False
        '
        'cmdModRequest
        '
        Me.cmdModRequest.BackColor = System.Drawing.SystemColors.Control
        Me.cmdModRequest.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdModRequest.Enabled = False
        Me.cmdModRequest.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdModRequest.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdModRequest.Location = New System.Drawing.Point(8, 472)
        Me.cmdModRequest.Name = "cmdModRequest"
        Me.cmdModRequest.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdModRequest.Size = New System.Drawing.Size(121, 41)
        Me.cmdModRequest.TabIndex = 9
        Me.cmdModRequest.Text = "Show Modify Request"
        Me.cmdModRequest.UseVisualStyleBackColor = False
        '
        'Text1
        '
        Me.Text1.AcceptsReturn = True
        Me.Text1.BackColor = System.Drawing.SystemColors.Menu
        Me.Text1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Text1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text1.Location = New System.Drawing.Point(8, 8)
        Me.Text1.MaxLength = 0
        Me.Text1.Multiline = True
        Me.Text1.Name = "Text1"
        Me.Text1.ReadOnly = True
        Me.Text1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text1.Size = New System.Drawing.Size(377, 73)
        Me.Text1.TabIndex = 8
        Me.Text1.TabStop = False
        Me.Text1.Text = resources.GetString("Text1.Text")
        '
        'lstCustomers
        '
        Me.lstCustomers.BackColor = System.Drawing.SystemColors.Window
        Me.lstCustomers.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstCustomers.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstCustomers.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstCustomers.ItemHeight = 14
        Me.lstCustomers.Location = New System.Drawing.Point(8, 104)
        Me.lstCustomers.Name = "lstCustomers"
        Me.lstCustomers.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstCustomers.Size = New System.Drawing.Size(377, 130)
        Me.lstCustomers.TabIndex = 4
        Me.lstCustomers.TabStop = False
        '
        'lstUsedDataExts
        '
        Me.lstUsedDataExts.BackColor = System.Drawing.SystemColors.Window
        Me.lstUsedDataExts.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstUsedDataExts.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstUsedDataExts.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstUsedDataExts.ItemHeight = 14
        Me.lstUsedDataExts.Location = New System.Drawing.Point(8, 272)
        Me.lstUsedDataExts.Name = "lstUsedDataExts"
        Me.lstUsedDataExts.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstUsedDataExts.Size = New System.Drawing.Size(377, 74)
        Me.lstUsedDataExts.TabIndex = 3
        Me.lstUsedDataExts.TabStop = False
        '
        'txtDataExtValue
        '
        Me.txtDataExtValue.AcceptsReturn = True
        Me.txtDataExtValue.BackColor = System.Drawing.SystemColors.Window
        Me.txtDataExtValue.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDataExtValue.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDataExtValue.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDataExtValue.Location = New System.Drawing.Point(144, 368)
        Me.txtDataExtValue.MaxLength = 0
        Me.txtDataExtValue.Name = "txtDataExtValue"
        Me.txtDataExtValue.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDataExtValue.Size = New System.Drawing.Size(177, 19)
        Me.txtDataExtValue.TabIndex = 2
        '
        'cmdModValue
        '
        Me.cmdModValue.BackColor = System.Drawing.SystemColors.Control
        Me.cmdModValue.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdModValue.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdModValue.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdModValue.Location = New System.Drawing.Point(120, 408)
        Me.cmdModValue.Name = "cmdModValue"
        Me.cmdModValue.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdModValue.Size = New System.Drawing.Size(161, 41)
        Me.cmdModValue.TabIndex = 1
        Me.cmdModValue.Text = "Modify Data Extension Value"
        Me.cmdModValue.UseVisualStyleBackColor = False
        '
        'cmdCloseWindow
        '
        Me.cmdCloseWindow.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCloseWindow.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCloseWindow.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCloseWindow.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCloseWindow.Location = New System.Drawing.Point(264, 472)
        Me.cmdCloseWindow.Name = "cmdCloseWindow"
        Me.cmdCloseWindow.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCloseWindow.Size = New System.Drawing.Size(121, 41)
        Me.cmdCloseWindow.TabIndex = 0
        Me.cmdCloseWindow.Text = "Close Window"
        Me.cmdCloseWindow.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(8, 88)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(57, 17)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Customers"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(8, 256)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(177, 17)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Data Extension Values"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(8, 368)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(129, 23)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "New Data Extension Value"
        '
        'frmModDataExtension
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(398, 519)
        Me.Controls.Add(Me.cmdModResponse)
        Me.Controls.Add(Me.cmdModRequest)
        Me.Controls.Add(Me.Text1)
        Me.Controls.Add(Me.lstCustomers)
        Me.Controls.Add(Me.lstUsedDataExts)
        Me.Controls.Add(Me.txtDataExtValue)
        Me.Controls.Add(Me.cmdModValue)
        Me.Controls.Add(Me.cmdCloseWindow)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label3)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(292, 121)
        Me.Name = "frmModDataExtension"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Modify Data Extension Value for Customer"
        Me.ResumeLayout(False)

    End Sub
#End Region
End Class