<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmInvoiceModMain
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
	Public WithEvents cmdQuit As System.Windows.Forms.Button
	Public WithEvents cmdQBFC As System.Windows.Forms.Button
	Public WithEvents cmdQBXMLRP As System.Windows.Forms.Button
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmInvoiceModMain))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdQuit = New System.Windows.Forms.Button
		Me.cmdQBFC = New System.Windows.Forms.Button
		Me.cmdQBXMLRP = New System.Windows.Forms.Button
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.Text = "Invoice Modify Sample "
		Me.ClientSize = New System.Drawing.Size(314, 242)
		Me.Location = New System.Drawing.Point(382, 269)
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
		Me.Name = "frmInvoiceModMain"
		Me.cmdQuit.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdQuit.Text = "Quit"
		Me.cmdQuit.Size = New System.Drawing.Size(97, 49)
		Me.cmdQuit.Location = New System.Drawing.Point(208, 184)
		Me.cmdQuit.TabIndex = 4
		Me.cmdQuit.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdQuit.BackColor = System.Drawing.SystemColors.Control
		Me.cmdQuit.CausesValidation = True
		Me.cmdQuit.Enabled = True
		Me.cmdQuit.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdQuit.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdQuit.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdQuit.TabStop = True
		Me.cmdQuit.Name = "cmdQuit"
		Me.cmdQBFC.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdQBFC.Text = "Use QBFC"
		Me.cmdQBFC.Size = New System.Drawing.Size(145, 65)
		Me.cmdQBFC.Location = New System.Drawing.Point(160, 104)
		Me.cmdQBFC.TabIndex = 3
		Me.cmdQBFC.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdQBFC.BackColor = System.Drawing.SystemColors.Control
		Me.cmdQBFC.CausesValidation = True
		Me.cmdQBFC.Enabled = True
		Me.cmdQBFC.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdQBFC.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdQBFC.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdQBFC.TabStop = True
		Me.cmdQBFC.Name = "cmdQBFC"
		Me.cmdQBXMLRP.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdQBXMLRP.Text = "Use qbXML Request Processor"
		Me.cmdQBXMLRP.Size = New System.Drawing.Size(145, 65)
		Me.cmdQBXMLRP.Location = New System.Drawing.Point(8, 104)
		Me.cmdQBXMLRP.TabIndex = 2
		Me.cmdQBXMLRP.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdQBXMLRP.BackColor = System.Drawing.SystemColors.Control
		Me.cmdQBXMLRP.CausesValidation = True
		Me.cmdQBXMLRP.Enabled = True
		Me.cmdQBXMLRP.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdQBXMLRP.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdQBXMLRP.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdQBXMLRP.TabStop = True
		Me.cmdQBXMLRP.Name = "cmdQBXMLRP"
		Me.Label2.Text = "The sample program has been implemented to use either the qbXML request processor directly or use QBFC.  You choose the method to be used here."
		Me.Label2.Size = New System.Drawing.Size(297, 41)
		Me.Label2.Location = New System.Drawing.Point(8, 48)
		Me.Label2.TabIndex = 1
		Me.Label2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label2.BackColor = System.Drawing.SystemColors.Control
		Me.Label2.Enabled = True
		Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label2.UseMnemonic = True
		Me.Label2.Visible = True
		Me.Label2.AutoSize = False
		Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label2.Name = "Label2"
		Me.Label1.Text = "This sample program demonstrates the use of the SDK version 2.1 invoice modify functionality."
		Me.Label1.Size = New System.Drawing.Size(297, 25)
		Me.Label1.Location = New System.Drawing.Point(8, 8)
		Me.Label1.TabIndex = 0
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
		Me.Controls.Add(cmdQuit)
		Me.Controls.Add(cmdQBFC)
		Me.Controls.Add(cmdQBXMLRP)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label1)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class