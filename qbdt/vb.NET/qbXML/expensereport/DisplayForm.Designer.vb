<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class DisplayForm
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
	Public WithEvents closeButton As System.Windows.Forms.Button
	Public WithEvents displayText As System.Windows.Forms.TextBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(DisplayForm))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.closeButton = New System.Windows.Forms.Button
		Me.displayText = New System.Windows.Forms.TextBox
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.ClientSize = New System.Drawing.Size(556, 416)
		Me.Location = New System.Drawing.Point(4, 23)
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
		Me.Name = "DisplayForm"
		Me.closeButton.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.closeButton.Text = "Close"
		Me.closeButton.Size = New System.Drawing.Size(157, 17)
		Me.closeButton.Location = New System.Drawing.Point(200, 384)
		Me.closeButton.TabIndex = 1
		Me.closeButton.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.closeButton.BackColor = System.Drawing.SystemColors.Control
		Me.closeButton.CausesValidation = True
		Me.closeButton.Enabled = True
		Me.closeButton.ForeColor = System.Drawing.SystemColors.ControlText
		Me.closeButton.Cursor = System.Windows.Forms.Cursors.Default
		Me.closeButton.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.closeButton.TabStop = True
		Me.closeButton.Name = "closeButton"
		Me.displayText.AutoSize = False
		Me.displayText.Size = New System.Drawing.Size(549, 365)
		Me.displayText.Location = New System.Drawing.Point(4, 4)
		Me.displayText.MultiLine = True
		Me.displayText.ScrollBars = System.Windows.Forms.ScrollBars.Both
		Me.displayText.WordWrap = False
		Me.displayText.TabIndex = 0
		Me.displayText.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.displayText.AcceptsReturn = True
		Me.displayText.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.displayText.BackColor = System.Drawing.SystemColors.Window
		Me.displayText.CausesValidation = True
		Me.displayText.Enabled = True
		Me.displayText.ForeColor = System.Drawing.SystemColors.WindowText
		Me.displayText.HideSelection = True
		Me.displayText.ReadOnly = False
		Me.displayText.Maxlength = 0
		Me.displayText.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.displayText.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.displayText.TabStop = True
		Me.displayText.Visible = True
		Me.displayText.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.displayText.Name = "displayText"
		Me.Controls.Add(closeButton)
		Me.Controls.Add(displayText)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class