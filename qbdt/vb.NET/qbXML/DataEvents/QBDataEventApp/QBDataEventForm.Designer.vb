<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class QBDataEventForm
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
	Public WithEvents CheckEventTimer As System.Windows.Forms.Timer
	Public WithEvents CustomerList As System.Windows.Forms.ListView
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(QBDataEventForm))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.CheckEventTimer = New System.Windows.Forms.Timer(components)
		Me.CustomerList = New System.Windows.Forms.ListView
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "QB Customer Change Tracker"
		Me.ClientSize = New System.Drawing.Size(563, 335)
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
		Me.Name = "QBDataEventForm"
		Me.CheckEventTimer.Interval = 2000
		Me.CheckEventTimer.Enabled = True
		Me.CustomerList.Size = New System.Drawing.Size(545, 265)
		Me.CustomerList.Location = New System.Drawing.Point(8, 56)
		Me.CustomerList.TabIndex = 0
		Me.CustomerList.LabelWrap = True
		Me.CustomerList.HideSelection = True
		Me.CustomerList.ForeColor = System.Drawing.SystemColors.WindowText
		Me.CustomerList.BackColor = System.Drawing.SystemColors.Window
		Me.CustomerList.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.CustomerList.LabelEdit = True
		Me.CustomerList.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.CustomerList.Name = "CustomerList"
		Me.Controls.Add(CustomerList)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class