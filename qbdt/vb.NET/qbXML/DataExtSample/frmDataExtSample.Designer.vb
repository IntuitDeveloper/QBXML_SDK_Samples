<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmDataExtSample
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
	Public WithEvents cmdShowCustomResponse As System.Windows.Forms.Button
	Public WithEvents cmdShowCustomRequest As System.Windows.Forms.Button
	Public WithEvents cmdShowQueryResponse As System.Windows.Forms.Button
	Public WithEvents cmdShowQueryRequest As System.Windows.Forms.Button
	Public WithEvents Text1 As System.Windows.Forms.TextBox
	Public WithEvents txtUtility As System.Windows.Forms.TextBox
	Public WithEvents txtDataExts As System.Windows.Forms.TextBox
	Public WithEvents cmdExit As System.Windows.Forms.Button
	Public WithEvents cmdModDataExt As System.Windows.Forms.Button
	Public WithEvents cmdAddDataExtension As System.Windows.Forms.Button
	Public WithEvents cmdDefineDataExt As System.Windows.Forms.Button
	Public WithEvents txtCustomFields As System.Windows.Forms.TextBox
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdShowCustomResponse = New System.Windows.Forms.Button()
        Me.cmdShowCustomRequest = New System.Windows.Forms.Button()
        Me.cmdShowQueryResponse = New System.Windows.Forms.Button()
        Me.cmdShowQueryRequest = New System.Windows.Forms.Button()
        Me.Text1 = New System.Windows.Forms.TextBox()
        Me.txtUtility = New System.Windows.Forms.TextBox()
        Me.txtDataExts = New System.Windows.Forms.TextBox()
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.cmdModDataExt = New System.Windows.Forms.Button()
        Me.cmdAddDataExtension = New System.Windows.Forms.Button()
        Me.cmdDefineDataExt = New System.Windows.Forms.Button()
        Me.txtCustomFields = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'cmdShowCustomResponse
        '
        Me.cmdShowCustomResponse.BackColor = System.Drawing.SystemColors.Control
        Me.cmdShowCustomResponse.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdShowCustomResponse.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdShowCustomResponse.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdShowCustomResponse.Location = New System.Drawing.Point(320, 512)
        Me.cmdShowCustomResponse.Name = "cmdShowCustomResponse"
        Me.cmdShowCustomResponse.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdShowCustomResponse.Size = New System.Drawing.Size(89, 55)
        Me.cmdShowCustomResponse.TabIndex = 13
        Me.cmdShowCustomResponse.Text = "Show Custom Field Query Response"
        Me.cmdShowCustomResponse.UseVisualStyleBackColor = False
        '
        'cmdShowCustomRequest
        '
        Me.cmdShowCustomRequest.BackColor = System.Drawing.SystemColors.Control
        Me.cmdShowCustomRequest.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdShowCustomRequest.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdShowCustomRequest.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdShowCustomRequest.Location = New System.Drawing.Point(216, 512)
        Me.cmdShowCustomRequest.Name = "cmdShowCustomRequest"
        Me.cmdShowCustomRequest.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdShowCustomRequest.Size = New System.Drawing.Size(89, 55)
        Me.cmdShowCustomRequest.TabIndex = 12
        Me.cmdShowCustomRequest.Text = "Show Custom Field Query Request"
        Me.cmdShowCustomRequest.UseVisualStyleBackColor = False
        '
        'cmdShowQueryResponse
        '
        Me.cmdShowQueryResponse.BackColor = System.Drawing.SystemColors.Control
        Me.cmdShowQueryResponse.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdShowQueryResponse.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdShowQueryResponse.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdShowQueryResponse.Location = New System.Drawing.Point(112, 512)
        Me.cmdShowQueryResponse.Name = "cmdShowQueryResponse"
        Me.cmdShowQueryResponse.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdShowQueryResponse.Size = New System.Drawing.Size(89, 55)
        Me.cmdShowQueryResponse.TabIndex = 11
        Me.cmdShowQueryResponse.Text = "Show Data Extension Query Response"
        Me.cmdShowQueryResponse.UseVisualStyleBackColor = False
        '
        'cmdShowQueryRequest
        '
        Me.cmdShowQueryRequest.BackColor = System.Drawing.SystemColors.Control
        Me.cmdShowQueryRequest.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdShowQueryRequest.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdShowQueryRequest.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdShowQueryRequest.Location = New System.Drawing.Point(8, 512)
        Me.cmdShowQueryRequest.Name = "cmdShowQueryRequest"
        Me.cmdShowQueryRequest.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdShowQueryRequest.Size = New System.Drawing.Size(89, 55)
        Me.cmdShowQueryRequest.TabIndex = 10
        Me.cmdShowQueryRequest.Text = "Show Data Extension Query Request"
        Me.cmdShowQueryRequest.UseVisualStyleBackColor = False
        '
        'Text1
        '
        Me.Text1.AcceptsReturn = True
        Me.Text1.BackColor = System.Drawing.SystemColors.Menu
        Me.Text1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Text1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text1.Location = New System.Drawing.Point(8, 32)
        Me.Text1.MaxLength = 0
        Me.Text1.Multiline = True
        Me.Text1.Name = "Text1"
        Me.Text1.ReadOnly = True
        Me.Text1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text1.Size = New System.Drawing.Size(385, 49)
        Me.Text1.TabIndex = 9
        Me.Text1.Text = "Select the buttons below to add data extension definitions to a company file, add" &
    " data extension values to a customer record or modify the value of a data extens" &
    "ion already added to a customer" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'txtUtility
        '
        Me.txtUtility.AcceptsReturn = True
        Me.txtUtility.BackColor = System.Drawing.SystemColors.Window
        Me.txtUtility.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtUtility.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtUtility.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtUtility.Location = New System.Drawing.Point(400, 264)
        Me.txtUtility.MaxLength = 0
        Me.txtUtility.Multiline = True
        Me.txtUtility.Name = "txtUtility"
        Me.txtUtility.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtUtility.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtUtility.Size = New System.Drawing.Size(10, 13)
        Me.txtUtility.TabIndex = 8
        Me.txtUtility.TabStop = False
        Me.txtUtility.Text = "T" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "e" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "x" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "t" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "1" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.txtUtility.Visible = False
        Me.txtUtility.WordWrap = False
        '
        'txtDataExts
        '
        Me.txtDataExts.AcceptsReturn = True
        Me.txtDataExts.BackColor = System.Drawing.SystemColors.Window
        Me.txtDataExts.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDataExts.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDataExts.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDataExts.Location = New System.Drawing.Point(8, 88)
        Me.txtDataExts.MaxLength = 0
        Me.txtDataExts.Multiline = True
        Me.txtDataExts.Name = "txtDataExts"
        Me.txtDataExts.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDataExts.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtDataExts.Size = New System.Drawing.Size(401, 161)
        Me.txtDataExts.TabIndex = 7
        Me.txtDataExts.WordWrap = False
        '
        'cmdExit
        '
        Me.cmdExit.BackColor = System.Drawing.SystemColors.Control
        Me.cmdExit.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdExit.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdExit.Location = New System.Drawing.Point(328, 440)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdExit.Size = New System.Drawing.Size(81, 49)
        Me.cmdExit.TabIndex = 6
        Me.cmdExit.Text = "Exit"
        Me.cmdExit.UseVisualStyleBackColor = False
        '
        'cmdModDataExt
        '
        Me.cmdModDataExt.BackColor = System.Drawing.SystemColors.Control
        Me.cmdModDataExt.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdModDataExt.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdModDataExt.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdModDataExt.Location = New System.Drawing.Point(216, 440)
        Me.cmdModDataExt.Name = "cmdModDataExt"
        Me.cmdModDataExt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdModDataExt.Size = New System.Drawing.Size(97, 49)
        Me.cmdModDataExt.TabIndex = 5
        Me.cmdModDataExt.Text = "Modify Customer Data Extension"
        Me.cmdModDataExt.UseVisualStyleBackColor = False
        '
        'cmdAddDataExtension
        '
        Me.cmdAddDataExtension.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAddDataExtension.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAddDataExtension.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAddDataExtension.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAddDataExtension.Location = New System.Drawing.Point(104, 440)
        Me.cmdAddDataExtension.Name = "cmdAddDataExtension"
        Me.cmdAddDataExtension.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAddDataExtension.Size = New System.Drawing.Size(97, 49)
        Me.cmdAddDataExtension.TabIndex = 4
        Me.cmdAddDataExtension.Text = "Add Customer Data Extension"
        Me.cmdAddDataExtension.UseVisualStyleBackColor = False
        '
        'cmdDefineDataExt
        '
        Me.cmdDefineDataExt.BackColor = System.Drawing.SystemColors.Control
        Me.cmdDefineDataExt.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdDefineDataExt.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDefineDataExt.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDefineDataExt.Location = New System.Drawing.Point(8, 440)
        Me.cmdDefineDataExt.Name = "cmdDefineDataExt"
        Me.cmdDefineDataExt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdDefineDataExt.Size = New System.Drawing.Size(81, 49)
        Me.cmdDefineDataExt.TabIndex = 3
        Me.cmdDefineDataExt.Text = "Define Data Extension"
        Me.cmdDefineDataExt.UseVisualStyleBackColor = False
        '
        'txtCustomFields
        '
        Me.txtCustomFields.AcceptsReturn = True
        Me.txtCustomFields.BackColor = System.Drawing.SystemColors.Window
        Me.txtCustomFields.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCustomFields.Enabled = False
        Me.txtCustomFields.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustomFields.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCustomFields.Location = New System.Drawing.Point(8, 288)
        Me.txtCustomFields.MaxLength = 0
        Me.txtCustomFields.Multiline = True
        Me.txtCustomFields.Name = "txtCustomFields"
        Me.txtCustomFields.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCustomFields.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtCustomFields.Size = New System.Drawing.Size(401, 145)
        Me.txtCustomFields.TabIndex = 2
        Me.txtCustomFields.WordWrap = False
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(8, 264)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(385, 17)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Defined Custom Fields (read only, to show query for Custom Field Defs)"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(393, 23)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Data Extensions for OwnerID {E09C86CF-9D6E-4EF2-BCBE-4D66B6B0F754}"
        '
        'frmDataExtSample
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(418, 568)
        Me.Controls.Add(Me.cmdShowCustomResponse)
        Me.Controls.Add(Me.cmdShowCustomRequest)
        Me.Controls.Add(Me.cmdShowQueryResponse)
        Me.Controls.Add(Me.cmdShowQueryRequest)
        Me.Controls.Add(Me.Text1)
        Me.Controls.Add(Me.txtUtility)
        Me.Controls.Add(Me.txtDataExts)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.cmdModDataExt)
        Me.Controls.Add(Me.cmdAddDataExtension)
        Me.Controls.Add(Me.cmdDefineDataExt)
        Me.Controls.Add(Me.txtCustomFields)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(276, 164)
        Me.Name = "frmDataExtSample"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Data Extension Sample Main Window"
        Me.ResumeLayout(False)

    End Sub
#End Region
End Class