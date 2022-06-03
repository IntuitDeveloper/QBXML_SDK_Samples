<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmDepositAdd
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
	Public WithEvents cmdExit As System.Windows.Forms.Button
	Public WithEvents cmdDepositFunds As System.Windows.Forms.Button
	Public WithEvents lstFundsForDeposit As System.Windows.Forms.ListBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmDepositAdd))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdExit = New System.Windows.Forms.Button
		Me.cmdDepositFunds = New System.Windows.Forms.Button
		Me.lstFundsForDeposit = New System.Windows.Forms.ListBox
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "Deposit Add"
		Me.ClientSize = New System.Drawing.Size(501, 431)
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
		Me.Name = "frmDepositAdd"
		Me.cmdExit.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdExit.Text = "Exit"
		Me.cmdExit.Size = New System.Drawing.Size(153, 57)
		Me.cmdExit.Location = New System.Drawing.Point(280, 360)
		Me.cmdExit.TabIndex = 2
		Me.cmdExit.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdExit.BackColor = System.Drawing.SystemColors.Control
		Me.cmdExit.CausesValidation = True
		Me.cmdExit.Enabled = True
		Me.cmdExit.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdExit.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdExit.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdExit.TabStop = True
		Me.cmdExit.Name = "cmdExit"
		Me.cmdDepositFunds.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdDepositFunds.Text = "Deposit Selected Funds"
		Me.cmdDepositFunds.Size = New System.Drawing.Size(153, 57)
		Me.cmdDepositFunds.Location = New System.Drawing.Point(64, 360)
		Me.cmdDepositFunds.TabIndex = 1
		Me.cmdDepositFunds.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdDepositFunds.BackColor = System.Drawing.SystemColors.Control
		Me.cmdDepositFunds.CausesValidation = True
		Me.cmdDepositFunds.Enabled = True
		Me.cmdDepositFunds.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdDepositFunds.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdDepositFunds.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdDepositFunds.TabStop = True
		Me.cmdDepositFunds.Name = "cmdDepositFunds"
		Me.lstFundsForDeposit.Size = New System.Drawing.Size(465, 332)
		Me.lstFundsForDeposit.Location = New System.Drawing.Point(16, 16)
		Me.lstFundsForDeposit.TabIndex = 0
		Me.lstFundsForDeposit.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lstFundsForDeposit.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lstFundsForDeposit.BackColor = System.Drawing.SystemColors.Window
		Me.lstFundsForDeposit.CausesValidation = True
		Me.lstFundsForDeposit.Enabled = True
		Me.lstFundsForDeposit.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lstFundsForDeposit.IntegralHeight = True
		Me.lstFundsForDeposit.Cursor = System.Windows.Forms.Cursors.Default
		Me.lstFundsForDeposit.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.lstFundsForDeposit.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lstFundsForDeposit.Sorted = False
		Me.lstFundsForDeposit.TabStop = True
		Me.lstFundsForDeposit.Visible = True
		Me.lstFundsForDeposit.MultiColumn = False
		Me.lstFundsForDeposit.Name = "lstFundsForDeposit"
		Me.Controls.Add(cmdExit)
		Me.Controls.Add(cmdDepositFunds)
		Me.Controls.Add(lstFundsForDeposit)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class