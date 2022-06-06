<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class ShowItems
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
	Public WithEvents Done As System.Windows.Forms.Button
	Public WithEvents ItemsList As System.Windows.Forms.TextBox
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(ShowItems))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Done = New System.Windows.Forms.Button
		Me.ItemsList = New System.Windows.Forms.TextBox
		Me.Label1 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "Items List"
		Me.ClientSize = New System.Drawing.Size(471, 302)
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
		Me.Name = "ShowItems"
		Me.Done.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Done.Text = "Done"
		Me.Done.Size = New System.Drawing.Size(73, 25)
		Me.Done.Location = New System.Drawing.Point(192, 264)
		Me.Done.TabIndex = 1
		Me.Done.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Done.BackColor = System.Drawing.SystemColors.Control
		Me.Done.CausesValidation = True
		Me.Done.Enabled = True
		Me.Done.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Done.Cursor = System.Windows.Forms.Cursors.Default
		Me.Done.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Done.TabStop = True
		Me.Done.Name = "Done"
		Me.ItemsList.AutoSize = False
		Me.ItemsList.Size = New System.Drawing.Size(425, 201)
		Me.ItemsList.Location = New System.Drawing.Point(16, 48)
		Me.ItemsList.MultiLine = True
		Me.ItemsList.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
		Me.ItemsList.TabIndex = 0
		Me.ItemsList.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ItemsList.AcceptsReturn = True
		Me.ItemsList.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.ItemsList.BackColor = System.Drawing.SystemColors.Window
		Me.ItemsList.CausesValidation = True
		Me.ItemsList.Enabled = True
		Me.ItemsList.ForeColor = System.Drawing.SystemColors.WindowText
		Me.ItemsList.HideSelection = True
		Me.ItemsList.ReadOnly = False
		Me.ItemsList.Maxlength = 0
		Me.ItemsList.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.ItemsList.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ItemsList.TabStop = True
		Me.ItemsList.Visible = True
		Me.ItemsList.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.ItemsList.Name = "ItemsList"
		Me.Label1.Text = "Query result for the first 30 items in the current company file."
		Me.Label1.Size = New System.Drawing.Size(393, 17)
		Me.Label1.Location = New System.Drawing.Point(16, 16)
		Me.Label1.TabIndex = 2
		Me.Label1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.BackColor = System.Drawing.SystemColors.Control
		Me.Label1.Enabled = True
		Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Controls.Add(Done)
		Me.Controls.Add(ItemsList)
		Me.Controls.Add(Label1)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class