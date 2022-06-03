<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmAddDataExtension
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
	Public WithEvents cmdShowResponse As System.Windows.Forms.Button
	Public WithEvents cmdShowRequest As System.Windows.Forms.Button
	Public WithEvents Text1 As System.Windows.Forms.TextBox
	Public WithEvents lstUsedDataExts As System.Windows.Forms.ListBox
	Public WithEvents lstAvailableDataExts As System.Windows.Forms.ListBox
	Public WithEvents cmdCloseWindow As System.Windows.Forms.Button
	Public WithEvents cmdAddValue As System.Windows.Forms.Button
	Public WithEvents txtDataExtValue As System.Windows.Forms.TextBox
	Public WithEvents lstUnusedDataExtDefs As System.Windows.Forms.ListBox
	Public WithEvents lstCustomers As System.Windows.Forms.ListBox
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAddDataExtension))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdShowResponse = New System.Windows.Forms.Button()
        Me.cmdShowRequest = New System.Windows.Forms.Button()
        Me.Text1 = New System.Windows.Forms.TextBox()
        Me.lstUsedDataExts = New System.Windows.Forms.ListBox()
        Me.lstAvailableDataExts = New System.Windows.Forms.ListBox()
        Me.cmdCloseWindow = New System.Windows.Forms.Button()
        Me.cmdAddValue = New System.Windows.Forms.Button()
        Me.txtDataExtValue = New System.Windows.Forms.TextBox()
        Me.lstUnusedDataExtDefs = New System.Windows.Forms.ListBox()
        Me.lstCustomers = New System.Windows.Forms.ListBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'cmdShowResponse
        '
        Me.cmdShowResponse.BackColor = System.Drawing.SystemColors.Control
        Me.cmdShowResponse.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdShowResponse.Enabled = False
        Me.cmdShowResponse.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdShowResponse.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdShowResponse.Location = New System.Drawing.Point(136, 472)
        Me.cmdShowResponse.Name = "cmdShowResponse"
        Me.cmdShowResponse.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdShowResponse.Size = New System.Drawing.Size(113, 41)
        Me.cmdShowResponse.TabIndex = 12
        Me.cmdShowResponse.Text = "Show Add Response"
        Me.cmdShowResponse.UseVisualStyleBackColor = False
        '
        'cmdShowRequest
        '
        Me.cmdShowRequest.BackColor = System.Drawing.SystemColors.Control
        Me.cmdShowRequest.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdShowRequest.Enabled = False
        Me.cmdShowRequest.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdShowRequest.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdShowRequest.Location = New System.Drawing.Point(8, 472)
        Me.cmdShowRequest.Name = "cmdShowRequest"
        Me.cmdShowRequest.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdShowRequest.Size = New System.Drawing.Size(113, 41)
        Me.cmdShowRequest.TabIndex = 11
        Me.cmdShowRequest.Text = "Show Add Request"
        Me.cmdShowRequest.UseVisualStyleBackColor = False
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
        Me.Text1.TabIndex = 10
        Me.Text1.Text = resources.GetString("Text1.Text")
        '
        'lstUsedDataExts
        '
        Me.lstUsedDataExts.BackColor = System.Drawing.SystemColors.Window
        Me.lstUsedDataExts.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstUsedDataExts.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstUsedDataExts.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstUsedDataExts.ItemHeight = 14
        Me.lstUsedDataExts.Items.AddRange(New Object() {"lstUsedDataExts"})
        Me.lstUsedDataExts.Location = New System.Drawing.Point(328, 368)
        Me.lstUsedDataExts.Name = "lstUsedDataExts"
        Me.lstUsedDataExts.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstUsedDataExts.Size = New System.Drawing.Size(9, 18)
        Me.lstUsedDataExts.Sorted = True
        Me.lstUsedDataExts.TabIndex = 9
        Me.lstUsedDataExts.Visible = False
        '
        'lstAvailableDataExts
        '
        Me.lstAvailableDataExts.BackColor = System.Drawing.SystemColors.Window
        Me.lstAvailableDataExts.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstAvailableDataExts.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstAvailableDataExts.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstAvailableDataExts.ItemHeight = 14
        Me.lstAvailableDataExts.Location = New System.Drawing.Point(312, 368)
        Me.lstAvailableDataExts.Name = "lstAvailableDataExts"
        Me.lstAvailableDataExts.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstAvailableDataExts.Size = New System.Drawing.Size(9, 18)
        Me.lstAvailableDataExts.Sorted = True
        Me.lstAvailableDataExts.TabIndex = 8
        Me.lstAvailableDataExts.Visible = False
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
        Me.cmdCloseWindow.TabIndex = 7
        Me.cmdCloseWindow.Text = "Close Window"
        Me.cmdCloseWindow.UseVisualStyleBackColor = False
        '
        'cmdAddValue
        '
        Me.cmdAddValue.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAddValue.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAddValue.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAddValue.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAddValue.Location = New System.Drawing.Point(120, 408)
        Me.cmdAddValue.Name = "cmdAddValue"
        Me.cmdAddValue.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAddValue.Size = New System.Drawing.Size(145, 41)
        Me.cmdAddValue.TabIndex = 6
        Me.cmdAddValue.Text = "Add Data Extension Value"
        Me.cmdAddValue.UseVisualStyleBackColor = False
        '
        'txtDataExtValue
        '
        Me.txtDataExtValue.AcceptsReturn = True
        Me.txtDataExtValue.BackColor = System.Drawing.SystemColors.Window
        Me.txtDataExtValue.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDataExtValue.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDataExtValue.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDataExtValue.Location = New System.Drawing.Point(120, 368)
        Me.txtDataExtValue.MaxLength = 0
        Me.txtDataExtValue.Name = "txtDataExtValue"
        Me.txtDataExtValue.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDataExtValue.Size = New System.Drawing.Size(177, 19)
        Me.txtDataExtValue.TabIndex = 5
        '
        'lstUnusedDataExtDefs
        '
        Me.lstUnusedDataExtDefs.BackColor = System.Drawing.SystemColors.Window
        Me.lstUnusedDataExtDefs.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstUnusedDataExtDefs.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstUnusedDataExtDefs.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstUnusedDataExtDefs.ItemHeight = 14
        Me.lstUnusedDataExtDefs.Location = New System.Drawing.Point(8, 272)
        Me.lstUnusedDataExtDefs.Name = "lstUnusedDataExtDefs"
        Me.lstUnusedDataExtDefs.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstUnusedDataExtDefs.Size = New System.Drawing.Size(377, 74)
        Me.lstUnusedDataExtDefs.TabIndex = 2
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
        Me.lstCustomers.TabIndex = 0
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
        Me.Label3.Size = New System.Drawing.Size(105, 25)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Data Extension Value"
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
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Unused Data Extension Definitions"
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
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Customers"
        '
        'frmAddDataExtension
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(397, 522)
        Me.Controls.Add(Me.cmdShowResponse)
        Me.Controls.Add(Me.cmdShowRequest)
        Me.Controls.Add(Me.Text1)
        Me.Controls.Add(Me.lstUsedDataExts)
        Me.Controls.Add(Me.lstAvailableDataExts)
        Me.Controls.Add(Me.cmdCloseWindow)
        Me.Controls.Add(Me.cmdAddValue)
        Me.Controls.Add(Me.txtDataExtValue)
        Me.Controls.Add(Me.lstUnusedDataExtDefs)
        Me.Controls.Add(Me.lstCustomers)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(276, 202)
        Me.Name = "frmAddDataExtension"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Add Data Extension To Customer"
        Me.ResumeLayout(False)

    End Sub
#End Region
End Class