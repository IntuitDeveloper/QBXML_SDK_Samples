<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmPurchaseOrderModifyMain
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
	Public WithEvents txtToDate As System.Windows.Forms.TextBox
	Public WithEvents txtFromDate As System.Windows.Forms.TextBox
	Public WithEvents cmdQBXMLRP As System.Windows.Forms.Button
	Public WithEvents cmdQBFC As System.Windows.Forms.Button
	Public WithEvents cmdQuit As System.Windows.Forms.Button
	Public WithEvents Label5 As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmPurchaseOrderModifyMain))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.txtToDate = New System.Windows.Forms.TextBox
		Me.txtFromDate = New System.Windows.Forms.TextBox
		Me.cmdQBXMLRP = New System.Windows.Forms.Button
		Me.cmdQBFC = New System.Windows.Forms.Button
		Me.cmdQuit = New System.Windows.Forms.Button
		Me.Label5 = New System.Windows.Forms.Label
		Me.Label4 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.Text = "Purchase Order Modify"
		Me.ClientSize = New System.Drawing.Size(313, 347)
		Me.Location = New System.Drawing.Point(356, 246)
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
		Me.Name = "frmPurchaseOrderModifyMain"
		Me.txtToDate.AutoSize = False
		Me.txtToDate.Size = New System.Drawing.Size(65, 19)
		Me.txtToDate.Location = New System.Drawing.Point(144, 184)
		Me.txtToDate.TabIndex = 9
		Me.txtToDate.Text = "2038-01-18"
		Me.txtToDate.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtToDate.AcceptsReturn = True
		Me.txtToDate.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtToDate.BackColor = System.Drawing.SystemColors.Window
		Me.txtToDate.CausesValidation = True
		Me.txtToDate.Enabled = True
		Me.txtToDate.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtToDate.HideSelection = True
		Me.txtToDate.ReadOnly = False
		Me.txtToDate.Maxlength = 0
		Me.txtToDate.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtToDate.MultiLine = False
		Me.txtToDate.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtToDate.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtToDate.TabStop = True
		Me.txtToDate.Visible = True
		Me.txtToDate.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtToDate.Name = "txtToDate"
		Me.txtFromDate.AutoSize = False
		Me.txtFromDate.Size = New System.Drawing.Size(65, 19)
		Me.txtFromDate.Location = New System.Drawing.Point(40, 184)
		Me.txtFromDate.TabIndex = 7
		Me.txtFromDate.Text = "1970-01-01"
		Me.txtFromDate.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtFromDate.AcceptsReturn = True
		Me.txtFromDate.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtFromDate.BackColor = System.Drawing.SystemColors.Window
		Me.txtFromDate.CausesValidation = True
		Me.txtFromDate.Enabled = True
		Me.txtFromDate.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtFromDate.HideSelection = True
		Me.txtFromDate.ReadOnly = False
		Me.txtFromDate.Maxlength = 0
		Me.txtFromDate.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtFromDate.MultiLine = False
		Me.txtFromDate.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtFromDate.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtFromDate.TabStop = True
		Me.txtFromDate.Visible = True
		Me.txtFromDate.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtFromDate.Name = "txtFromDate"
		Me.cmdQBXMLRP.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdQBXMLRP.Text = "Use qbXML Request Processor"
		Me.cmdQBXMLRP.Size = New System.Drawing.Size(145, 65)
		Me.cmdQBXMLRP.Location = New System.Drawing.Point(8, 216)
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
		Me.cmdQBFC.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdQBFC.Text = "Use QBFC"
		Me.cmdQBFC.Size = New System.Drawing.Size(145, 65)
		Me.cmdQBFC.Location = New System.Drawing.Point(160, 216)
		Me.cmdQBFC.TabIndex = 1
		Me.cmdQBFC.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdQBFC.BackColor = System.Drawing.SystemColors.Control
		Me.cmdQBFC.CausesValidation = True
		Me.cmdQBFC.Enabled = True
		Me.cmdQBFC.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdQBFC.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdQBFC.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdQBFC.TabStop = True
		Me.cmdQBFC.Name = "cmdQBFC"
		Me.cmdQuit.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdQuit.Text = "Quit"
		Me.cmdQuit.Size = New System.Drawing.Size(97, 49)
		Me.cmdQuit.Location = New System.Drawing.Point(208, 288)
		Me.cmdQuit.TabIndex = 0
		Me.cmdQuit.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdQuit.BackColor = System.Drawing.SystemColors.Control
		Me.cmdQuit.CausesValidation = True
		Me.cmdQuit.Enabled = True
		Me.cmdQuit.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdQuit.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdQuit.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdQuit.TabStop = True
		Me.cmdQuit.Name = "cmdQuit"
		Me.Label5.Text = "To"
		Me.Label5.Size = New System.Drawing.Size(17, 17)
		Me.Label5.Location = New System.Drawing.Point(120, 184)
		Me.Label5.TabIndex = 8
		Me.Label5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label5.BackColor = System.Drawing.SystemColors.Control
		Me.Label5.Enabled = True
		Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label5.UseMnemonic = True
		Me.Label5.Visible = True
		Me.Label5.AutoSize = False
		Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label5.Name = "Label5"
		Me.Label4.Text = "From"
		Me.Label4.Size = New System.Drawing.Size(25, 17)
		Me.Label4.Location = New System.Drawing.Point(8, 184)
		Me.Label4.TabIndex = 6
		Me.Label4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label4.BackColor = System.Drawing.SystemColors.Control
		Me.Label4.Enabled = True
		Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label4.UseMnemonic = True
		Me.Label4.Visible = True
		Me.Label4.AutoSize = False
		Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label4.Name = "Label4"
		Me.Label3.Text = "Open Purchase Order Query Date(s)"
		Me.Label3.Size = New System.Drawing.Size(177, 17)
		Me.Label3.Location = New System.Drawing.Point(8, 160)
		Me.Label3.TabIndex = 5
		Me.Label3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label3.BackColor = System.Drawing.SystemColors.Control
		Me.Label3.Enabled = True
		Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label3.UseMnemonic = True
		Me.Label3.Visible = True
		Me.Label3.AutoSize = False
		Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label3.Name = "Label3"
		Me.Label1.Text = "This sample program demonstrates the use of the SDK version 2.1 purchase order modify functionality."
		Me.Label1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.Size = New System.Drawing.Size(297, 49)
		Me.Label1.Location = New System.Drawing.Point(8, 8)
		Me.Label1.TabIndex = 4
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
		Me.Label2.Text = "The sample program has been implemented to use either the qbXML request processor directly or use QBFC.  You choose the method to be used here."
		Me.Label2.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Size = New System.Drawing.Size(297, 65)
		Me.Label2.Location = New System.Drawing.Point(8, 72)
		Me.Label2.TabIndex = 3
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
		Me.Controls.Add(txtToDate)
		Me.Controls.Add(txtFromDate)
		Me.Controls.Add(cmdQBXMLRP)
		Me.Controls.Add(cmdQBFC)
		Me.Controls.Add(cmdQuit)
		Me.Controls.Add(Label5)
		Me.Controls.Add(Label4)
		Me.Controls.Add(Label3)
		Me.Controls.Add(Label1)
		Me.Controls.Add(Label2)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class