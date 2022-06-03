<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class DisplayXML
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
	Public WithEvents Comm_Exit As System.Windows.Forms.Button
	Public WithEvents Text_Content As System.Windows.Forms.TextBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(DisplayXML))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Comm_Exit = New System.Windows.Forms.Button
		Me.Text_Content = New System.Windows.Forms.TextBox
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "Request XML"
		Me.ClientSize = New System.Drawing.Size(544, 276)
		Me.Location = New System.Drawing.Point(3, 22)
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
		Me.Name = "DisplayXML"
		Me.Comm_Exit.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Comm_Exit.Text = "&Done"
		Me.Comm_Exit.Size = New System.Drawing.Size(97, 25)
		Me.Comm_Exit.Location = New System.Drawing.Point(224, 232)
		Me.Comm_Exit.TabIndex = 1
		Me.Comm_Exit.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Comm_Exit.BackColor = System.Drawing.SystemColors.Control
		Me.Comm_Exit.CausesValidation = True
		Me.Comm_Exit.Enabled = True
		Me.Comm_Exit.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Comm_Exit.Cursor = System.Windows.Forms.Cursors.Default
		Me.Comm_Exit.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Comm_Exit.TabStop = True
		Me.Comm_Exit.Name = "Comm_Exit"
		Me.Text_Content.AutoSize = False
		Me.Text_Content.Size = New System.Drawing.Size(529, 209)
		Me.Text_Content.Location = New System.Drawing.Point(8, 8)
		Me.Text_Content.MultiLine = True
		Me.Text_Content.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
		Me.Text_Content.TabIndex = 0
		Me.Text_Content.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text_Content.AcceptsReturn = True
		Me.Text_Content.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text_Content.BackColor = System.Drawing.SystemColors.Window
		Me.Text_Content.CausesValidation = True
		Me.Text_Content.Enabled = True
		Me.Text_Content.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text_Content.HideSelection = True
		Me.Text_Content.ReadOnly = False
		Me.Text_Content.Maxlength = 0
		Me.Text_Content.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text_Content.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text_Content.TabStop = True
		Me.Text_Content.Visible = True
		Me.Text_Content.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text_Content.Name = "Text_Content"
		Me.Controls.Add(Comm_Exit)
		Me.Controls.Add(Text_Content)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class