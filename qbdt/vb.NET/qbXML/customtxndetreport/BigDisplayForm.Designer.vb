<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class BigDisplayForm
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
	Public WithEvents browser As System.Windows.Forms.WebBrowser
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(BigDisplayForm))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.closeButton = New System.Windows.Forms.Button
		Me.browser = New System.Windows.Forms.WebBrowser
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "Your Report"
		Me.ClientSize = New System.Drawing.Size(806, 504)
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
		Me.Name = "BigDisplayForm"
		Me.closeButton.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.closeButton.Text = "Close Window"
		Me.closeButton.Size = New System.Drawing.Size(161, 29)
		Me.closeButton.Location = New System.Drawing.Point(324, 468)
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
		Me.browser.Size = New System.Drawing.Size(781, 441)
		Me.browser.Location = New System.Drawing.Point(12, 16)
		Me.browser.TabIndex = 0
		Me.browser.AllowWebBrowserDrop = True
		Me.browser.Name = "browser"
		Me.Controls.Add(closeButton)
		Me.Controls.Add(browser)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class