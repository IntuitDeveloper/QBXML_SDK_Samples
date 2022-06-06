<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class MenuEventSample
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
	Public WithEvents EventData As System.Windows.Forms.TextBox
	Public WithEvents RemoveMenu As System.Windows.Forms.Button
	Public WithEvents AddMenu As System.Windows.Forms.Button
	Public WithEvents Label1 As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(MenuEventSample))
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
        Me.EventData = New System.Windows.Forms.TextBox
        Me.RemoveMenu = New System.Windows.Forms.Button
        Me.AddMenu = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        Me.ToolTip1.Active = True
        Me.Text = "Menu Event Sample"
        Me.ClientSize = New System.Drawing.Size(427, 276)
        Me.Location = New System.Drawing.Point(4, 30)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
        Me.Name = "MenuEventSample"
        Me.EventData.AutoSize = False
        Me.EventData.Size = New System.Drawing.Size(369, 129)
        Me.EventData.Location = New System.Drawing.Point(8, 120)
        Me.EventData.MultiLine = True
        Me.EventData.TabIndex = 2
        Me.EventData.Text = "No event received"
        Me.EventData.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.EventData.AcceptsReturn = True
        Me.EventData.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
        Me.EventData.BackColor = System.Drawing.SystemColors.Window
        Me.EventData.CausesValidation = True
        Me.EventData.Enabled = True
        Me.EventData.ForeColor = System.Drawing.SystemColors.WindowText
        Me.EventData.HideSelection = True
        Me.EventData.ReadOnly = False
        Me.EventData.Maxlength = 0
        Me.EventData.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.EventData.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.EventData.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.EventData.TabStop = True
        Me.EventData.Visible = True
        Me.EventData.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.EventData.Name = "EventData"
        Me.RemoveMenu.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.RemoveMenu.Text = "Remove Menu From QuickBooks"
        Me.RemoveMenu.Size = New System.Drawing.Size(185, 33)
        Me.RemoveMenu.Location = New System.Drawing.Point(8, 56)
        Me.RemoveMenu.TabIndex = 1
        Me.RemoveMenu.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RemoveMenu.BackColor = System.Drawing.SystemColors.Control
        Me.RemoveMenu.CausesValidation = True
        Me.RemoveMenu.Enabled = True
        Me.RemoveMenu.ForeColor = System.Drawing.SystemColors.ControlText
        Me.RemoveMenu.Cursor = System.Windows.Forms.Cursors.Default
        Me.RemoveMenu.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.RemoveMenu.TabStop = True
        Me.RemoveMenu.Name = "RemoveMenu"
        Me.AddMenu.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.AddMenu.Text = "Add Menu To QuickBooks"
        Me.AddMenu.Size = New System.Drawing.Size(185, 33)
        Me.AddMenu.Location = New System.Drawing.Point(8, 8)
        Me.AddMenu.TabIndex = 0
        Me.AddMenu.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AddMenu.BackColor = System.Drawing.SystemColors.Control
        Me.AddMenu.CausesValidation = True
        Me.AddMenu.Enabled = True
        Me.AddMenu.ForeColor = System.Drawing.SystemColors.ControlText
        Me.AddMenu.Cursor = System.Windows.Forms.Cursors.Default
        Me.AddMenu.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.AddMenu.TabStop = True
        Me.AddMenu.Name = "AddMenu"
        Me.Label1.Text = "Menu Event Data"
        Me.Label1.Size = New System.Drawing.Size(201, 17)
        Me.Label1.Location = New System.Drawing.Point(8, 104)
        Me.Label1.TabIndex = 3
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
        Me.Controls.Add(EventData)
        Me.Controls.Add(RemoveMenu)
        Me.Controls.Add(AddMenu)
        Me.Controls.Add(Label1)
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub
    Private Sub frmProgramma_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Dim process As System.Diagnostics.Process = Process.GetCurrentProcess
        process.Kill()
    End Sub
#End Region
End Class