<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmShowRequest
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
	Public WithEvents cmdDone As System.Windows.Forms.Button
	Public WithEvents txtRequest As System.Windows.Forms.TextBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmShowRequest))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdDone = New System.Windows.Forms.Button
		Me.txtRequest = New System.Windows.Forms.TextBox
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "Form1"
		Me.ClientSize = New System.Drawing.Size(546, 590)
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
		Me.Name = "frmShowRequest"
		Me.cmdDone.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdDone.Text = "Done"
		Me.cmdDone.Size = New System.Drawing.Size(113, 41)
		Me.cmdDone.Location = New System.Drawing.Point(216, 544)
		Me.cmdDone.TabIndex = 1
		Me.cmdDone.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdDone.BackColor = System.Drawing.SystemColors.Control
		Me.cmdDone.CausesValidation = True
		Me.cmdDone.Enabled = True
		Me.cmdDone.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdDone.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdDone.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdDone.TabStop = True
		Me.cmdDone.Name = "cmdDone"
		Me.txtRequest.AutoSize = False
		Me.txtRequest.Size = New System.Drawing.Size(529, 529)
		Me.txtRequest.Location = New System.Drawing.Point(8, 8)
		Me.txtRequest.ReadOnly = True
		Me.txtRequest.MultiLine = True
		Me.txtRequest.ScrollBars = System.Windows.Forms.ScrollBars.Both
		Me.txtRequest.WordWrap = False
		Me.txtRequest.TabIndex = 0
		Me.txtRequest.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtRequest.AcceptsReturn = True
		Me.txtRequest.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtRequest.BackColor = System.Drawing.SystemColors.Window
		Me.txtRequest.CausesValidation = True
		Me.txtRequest.Enabled = True
		Me.txtRequest.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtRequest.HideSelection = True
		Me.txtRequest.Maxlength = 0
		Me.txtRequest.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtRequest.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtRequest.TabStop = True
		Me.txtRequest.Visible = True
		Me.txtRequest.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtRequest.Name = "txtRequest"
		Me.Controls.Add(cmdDone)
		Me.Controls.Add(txtRequest)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class