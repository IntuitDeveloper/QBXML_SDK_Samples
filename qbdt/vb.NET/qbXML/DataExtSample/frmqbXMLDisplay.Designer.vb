<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmqbXMLDisplay
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
	Public WithEvents cmdCloseWindow As System.Windows.Forms.Button
	Public WithEvents txtqbXML As System.Windows.Forms.TextBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmqbXMLDisplay))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdCloseWindow = New System.Windows.Forms.Button
		Me.txtqbXML = New System.Windows.Forms.TextBox
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.Text = "qbXML Display"
		Me.ClientSize = New System.Drawing.Size(442, 458)
		Me.Location = New System.Drawing.Point(292, 141)
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
		Me.Name = "frmqbXMLDisplay"
		Me.cmdCloseWindow.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdCloseWindow.Text = "Close Window"
		Me.cmdCloseWindow.Size = New System.Drawing.Size(129, 49)
		Me.cmdCloseWindow.Location = New System.Drawing.Point(152, 408)
		Me.cmdCloseWindow.TabIndex = 1
		Me.cmdCloseWindow.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCloseWindow.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCloseWindow.CausesValidation = True
		Me.cmdCloseWindow.Enabled = True
		Me.cmdCloseWindow.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCloseWindow.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCloseWindow.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCloseWindow.TabStop = True
		Me.cmdCloseWindow.Name = "cmdCloseWindow"
		Me.txtqbXML.AutoSize = False
		Me.txtqbXML.Size = New System.Drawing.Size(425, 393)
		Me.txtqbXML.Location = New System.Drawing.Point(8, 8)
		Me.txtqbXML.MultiLine = True
		Me.txtqbXML.ScrollBars = System.Windows.Forms.ScrollBars.Both
		Me.txtqbXML.WordWrap = False
		Me.txtqbXML.TabIndex = 0
		Me.txtqbXML.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtqbXML.AcceptsReturn = True
		Me.txtqbXML.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtqbXML.BackColor = System.Drawing.SystemColors.Window
		Me.txtqbXML.CausesValidation = True
		Me.txtqbXML.Enabled = True
		Me.txtqbXML.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtqbXML.HideSelection = True
		Me.txtqbXML.ReadOnly = False
		Me.txtqbXML.Maxlength = 0
		Me.txtqbXML.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtqbXML.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtqbXML.TabStop = True
		Me.txtqbXML.Visible = True
		Me.txtqbXML.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtqbXML.Name = "txtqbXML"
		Me.Controls.Add(cmdCloseWindow)
		Me.Controls.Add(txtqbXML)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class