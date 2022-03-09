<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class QBFCEventsSubscriber
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
	Public WithEvents Subscribe As System.Windows.Forms.Button
	Public WithEvents Unsubscribe As System.Windows.Forms.Button
	Public WithEvents ErrorMsg As System.Windows.Forms.TextBox
	Public WithEvents ResponseXML As System.Windows.Forms.TextBox
	Public WithEvents RequestXML As System.Windows.Forms.TextBox
	Public WithEvents Errors As System.Windows.Forms.Label
	Public WithEvents Response As System.Windows.Forms.Label
	Public WithEvents Request As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Subscribe = New System.Windows.Forms.Button()
        Me.Unsubscribe = New System.Windows.Forms.Button()
        Me.ErrorMsg = New System.Windows.Forms.TextBox()
        Me.ResponseXML = New System.Windows.Forms.TextBox()
        Me.RequestXML = New System.Windows.Forms.TextBox()
        Me.Errors = New System.Windows.Forms.Label()
        Me.Response = New System.Windows.Forms.Label()
        Me.Request = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Subscribe
        '
        Me.Subscribe.BackColor = System.Drawing.SystemColors.Control
        Me.Subscribe.Cursor = System.Windows.Forms.Cursors.Default
        Me.Subscribe.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Subscribe.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Subscribe.Location = New System.Drawing.Point(64, 464)
        Me.Subscribe.Name = "Subscribe"
        Me.Subscribe.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Subscribe.Size = New System.Drawing.Size(89, 33)
        Me.Subscribe.TabIndex = 8
        Me.Subscribe.Text = "Subscribe"
        Me.Subscribe.UseVisualStyleBackColor = False
        '
        'Unsubscribe
        '
        Me.Unsubscribe.BackColor = System.Drawing.SystemColors.Control
        Me.Unsubscribe.Cursor = System.Windows.Forms.Cursors.Default
        Me.Unsubscribe.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Unsubscribe.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Unsubscribe.Location = New System.Drawing.Point(184, 464)
        Me.Unsubscribe.Name = "Unsubscribe"
        Me.Unsubscribe.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Unsubscribe.Size = New System.Drawing.Size(89, 33)
        Me.Unsubscribe.TabIndex = 7
        Me.Unsubscribe.Text = "Unsubscribe"
        Me.Unsubscribe.UseVisualStyleBackColor = False
        '
        'ErrorMsg
        '
        Me.ErrorMsg.AcceptsReturn = True
        Me.ErrorMsg.BackColor = System.Drawing.SystemColors.Window
        Me.ErrorMsg.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.ErrorMsg.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ErrorMsg.ForeColor = System.Drawing.SystemColors.WindowText
        Me.ErrorMsg.Location = New System.Drawing.Point(24, 352)
        Me.ErrorMsg.MaxLength = 0
        Me.ErrorMsg.Multiline = True
        Me.ErrorMsg.Name = "ErrorMsg"
        Me.ErrorMsg.ReadOnly = True
        Me.ErrorMsg.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ErrorMsg.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.ErrorMsg.Size = New System.Drawing.Size(345, 97)
        Me.ErrorMsg.TabIndex = 5
        Me.ErrorMsg.WordWrap = False
        '
        'ResponseXML
        '
        Me.ResponseXML.AcceptsReturn = True
        Me.ResponseXML.BackColor = System.Drawing.SystemColors.Window
        Me.ResponseXML.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.ResponseXML.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ResponseXML.ForeColor = System.Drawing.SystemColors.WindowText
        Me.ResponseXML.Location = New System.Drawing.Point(24, 216)
        Me.ResponseXML.MaxLength = 0
        Me.ResponseXML.Multiline = True
        Me.ResponseXML.Name = "ResponseXML"
        Me.ResponseXML.ReadOnly = True
        Me.ResponseXML.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ResponseXML.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.ResponseXML.Size = New System.Drawing.Size(345, 97)
        Me.ResponseXML.TabIndex = 2
        Me.ResponseXML.WordWrap = False
        '
        'RequestXML
        '
        Me.RequestXML.AcceptsReturn = True
        Me.RequestXML.BackColor = System.Drawing.SystemColors.Window
        Me.RequestXML.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.RequestXML.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RequestXML.ForeColor = System.Drawing.SystemColors.WindowText
        Me.RequestXML.Location = New System.Drawing.Point(24, 80)
        Me.RequestXML.MaxLength = 0
        Me.RequestXML.Multiline = True
        Me.RequestXML.Name = "RequestXML"
        Me.RequestXML.ReadOnly = True
        Me.RequestXML.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.RequestXML.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.RequestXML.Size = New System.Drawing.Size(345, 97)
        Me.RequestXML.TabIndex = 1
        Me.RequestXML.WordWrap = False
        '
        'Errors
        '
        Me.Errors.BackColor = System.Drawing.SystemColors.Control
        Me.Errors.Cursor = System.Windows.Forms.Cursors.Default
        Me.Errors.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Errors.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Errors.Location = New System.Drawing.Point(24, 328)
        Me.Errors.Name = "Errors"
        Me.Errors.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Errors.Size = New System.Drawing.Size(73, 17)
        Me.Errors.TabIndex = 6
        Me.Errors.Text = "Errors"
        '
        'Response
        '
        Me.Response.BackColor = System.Drawing.SystemColors.Control
        Me.Response.Cursor = System.Windows.Forms.Cursors.Default
        Me.Response.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Response.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Response.Location = New System.Drawing.Point(24, 192)
        Me.Response.Name = "Response"
        Me.Response.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Response.Size = New System.Drawing.Size(89, 17)
        Me.Response.TabIndex = 4
        Me.Response.Text = "Response"
        '
        'Request
        '
        Me.Request.BackColor = System.Drawing.SystemColors.Control
        Me.Request.Cursor = System.Windows.Forms.Cursors.Default
        Me.Request.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Request.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Request.Location = New System.Drawing.Point(24, 56)
        Me.Request.Name = "Request"
        Me.Request.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Request.Size = New System.Drawing.Size(89, 17)
        Me.Request.TabIndex = 3
        Me.Request.Text = "Request"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(24, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(345, 40)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "This sample shows how to subscribe and unsubscribe to Data Event:CustomerAdd and " &
    "UI Event:Company File Close Events."
        '
        'QBFCEventsSubscriber
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(410, 525)
        Me.Controls.Add(Me.Subscribe)
        Me.Controls.Add(Me.Unsubscribe)
        Me.Controls.Add(Me.ErrorMsg)
        Me.Controls.Add(Me.ResponseXML)
        Me.Controls.Add(Me.RequestXML)
        Me.Controls.Add(Me.Errors)
        Me.Controls.Add(Me.Response)
        Me.Controls.Add(Me.Request)
        Me.Controls.Add(Me.Label1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(4, 30)
        Me.Name = "QBFCEventsSubscriber"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "QBFC Events Subscriber"
        Me.ResumeLayout(False)

    End Sub
#End Region
End Class