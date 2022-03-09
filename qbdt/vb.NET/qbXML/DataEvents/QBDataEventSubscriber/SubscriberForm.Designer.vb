<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class SubscriberForm
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
	Public WithEvents Unsubscribe As System.Windows.Forms.Button
	Public WithEvents SubscribeBtn As System.Windows.Forms.Button
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(SubscriberForm))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Unsubscribe = New System.Windows.Forms.Button
		Me.SubscribeBtn = New System.Windows.Forms.Button
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "QBDataEventSubscriber"
		Me.ClientSize = New System.Drawing.Size(332, 143)
		Me.Location = New System.Drawing.Point(4, 30)
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
		Me.Name = "SubscriberForm"
		Me.Unsubscribe.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Unsubscribe.Text = "UNSubscribe From Events!"
		Me.Unsubscribe.Size = New System.Drawing.Size(281, 33)
		Me.Unsubscribe.Location = New System.Drawing.Point(24, 72)
		Me.Unsubscribe.TabIndex = 1
		Me.Unsubscribe.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Unsubscribe.BackColor = System.Drawing.SystemColors.Control
		Me.Unsubscribe.CausesValidation = True
		Me.Unsubscribe.Enabled = True
		Me.Unsubscribe.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Unsubscribe.Cursor = System.Windows.Forms.Cursors.Default
		Me.Unsubscribe.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Unsubscribe.TabStop = True
		Me.Unsubscribe.Name = "Unsubscribe"
		Me.SubscribeBtn.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.SubscribeBtn.Text = "Subscribe To Events!"
		Me.SubscribeBtn.Size = New System.Drawing.Size(281, 33)
		Me.SubscribeBtn.Location = New System.Drawing.Point(24, 16)
		Me.SubscribeBtn.TabIndex = 0
		Me.SubscribeBtn.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.SubscribeBtn.BackColor = System.Drawing.SystemColors.Control
		Me.SubscribeBtn.CausesValidation = True
		Me.SubscribeBtn.Enabled = True
		Me.SubscribeBtn.ForeColor = System.Drawing.SystemColors.ControlText
		Me.SubscribeBtn.Cursor = System.Windows.Forms.Cursors.Default
		Me.SubscribeBtn.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.SubscribeBtn.TabStop = True
		Me.SubscribeBtn.Name = "SubscribeBtn"
		Me.Controls.Add(Unsubscribe)
		Me.Controls.Add(SubscribeBtn)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class