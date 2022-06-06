<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class Macro
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
	Public WithEvents _SendTwoRequests_0 As System.Windows.Forms.Button
	Public WithEvents Frame2 As System.Windows.Forms.GroupBox
	Public WithEvents _SendOneRequest_1 As System.Windows.Forms.Button
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents estimatesList As System.Windows.Forms.ListBox
	Public WithEvents CloseWindow As System.Windows.Forms.Button
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents SendOneRequest As Microsoft.VisualBasic.Compatibility.VB6.ButtonArray
	Public WithEvents SendTwoRequests As Microsoft.VisualBasic.Compatibility.VB6.ButtonArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Macro))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Frame2 = New System.Windows.Forms.GroupBox
		Me._SendTwoRequests_0 = New System.Windows.Forms.Button
		Me._SendOneRequest_1 = New System.Windows.Forms.Button
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me.estimatesList = New System.Windows.Forms.ListBox
		Me.CloseWindow = New System.Windows.Forms.Button
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.SendOneRequest = New Microsoft.VisualBasic.Compatibility.VB6.ButtonArray(components)
		Me.SendTwoRequests = New Microsoft.VisualBasic.Compatibility.VB6.ButtonArray(components)
		Me.Frame2.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.SendOneRequest, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.SendTwoRequests, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.Text = "UseMacro and DefMacro Sample Program"
		Me.ClientSize = New System.Drawing.Size(463, 473)
		Me.Location = New System.Drawing.Point(4, 23)
		Me.ControlBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MaximizeBox = True
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "Macro"
		Me.Frame2.Text = "Press this button to send the InvoiceAdd and ReceivePayment in separate requests."
		Me.Frame2.Size = New System.Drawing.Size(425, 57)
		Me.Frame2.Location = New System.Drawing.Point(8, 360)
		Me.Frame2.TabIndex = 6
		Me.Frame2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame2.BackColor = System.Drawing.SystemColors.Control
		Me.Frame2.Enabled = True
		Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame2.Visible = True
		Me.Frame2.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame2.Name = "Frame2"
		Me._SendTwoRequests_0.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._SendTwoRequests_0.Text = "Send Two Requests"
		Me._SendTwoRequests_0.Size = New System.Drawing.Size(113, 33)
		Me._SendTwoRequests_0.Location = New System.Drawing.Point(152, 16)
		Me._SendTwoRequests_0.TabIndex = 7
		Me._SendTwoRequests_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._SendTwoRequests_0.BackColor = System.Drawing.SystemColors.Control
		Me._SendTwoRequests_0.CausesValidation = True
		Me._SendTwoRequests_0.Enabled = True
		Me._SendTwoRequests_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._SendTwoRequests_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._SendTwoRequests_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._SendTwoRequests_0.TabStop = True
		Me._SendTwoRequests_0.Name = "_SendTwoRequests_0"
		Me._SendOneRequest_1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._SendOneRequest_1.Text = "Send One Request"
		Me._SendOneRequest_1.Size = New System.Drawing.Size(113, 33)
		Me._SendOneRequest_1.Location = New System.Drawing.Point(160, 304)
		Me._SendOneRequest_1.TabIndex = 5
		Me._SendOneRequest_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._SendOneRequest_1.BackColor = System.Drawing.SystemColors.Control
		Me._SendOneRequest_1.CausesValidation = True
		Me._SendOneRequest_1.Enabled = True
		Me._SendOneRequest_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._SendOneRequest_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._SendOneRequest_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._SendOneRequest_1.TabStop = True
		Me._SendOneRequest_1.Name = "_SendOneRequest_1"
		Me.Frame1.Text = "Press this button to send the InvoiceAdd and ReceivePayment in one request."
		Me.Frame1.Size = New System.Drawing.Size(425, 57)
		Me.Frame1.Location = New System.Drawing.Point(8, 288)
		Me.Frame1.TabIndex = 4
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.BackColor = System.Drawing.SystemColors.Control
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame1.Name = "Frame1"
		Me.estimatesList.Size = New System.Drawing.Size(321, 137)
		Me.estimatesList.Location = New System.Drawing.Point(24, 112)
		Me.estimatesList.TabIndex = 1
		Me.estimatesList.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.estimatesList.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.estimatesList.BackColor = System.Drawing.SystemColors.Window
		Me.estimatesList.CausesValidation = True
		Me.estimatesList.Enabled = True
		Me.estimatesList.ForeColor = System.Drawing.SystemColors.WindowText
		Me.estimatesList.IntegralHeight = True
		Me.estimatesList.Cursor = System.Windows.Forms.Cursors.Default
		Me.estimatesList.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.estimatesList.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.estimatesList.Sorted = False
		Me.estimatesList.TabStop = True
		Me.estimatesList.Visible = True
		Me.estimatesList.MultiColumn = False
		Me.estimatesList.Name = "estimatesList"
		Me.CloseWindow.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.CloseWindow
		Me.CloseWindow.Text = "Close Window"
		Me.CloseWindow.Size = New System.Drawing.Size(113, 33)
		Me.CloseWindow.Location = New System.Drawing.Point(160, 432)
		Me.CloseWindow.TabIndex = 0
		Me.CloseWindow.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.CloseWindow.BackColor = System.Drawing.SystemColors.Control
		Me.CloseWindow.CausesValidation = True
		Me.CloseWindow.Enabled = True
		Me.CloseWindow.ForeColor = System.Drawing.SystemColors.ControlText
		Me.CloseWindow.Cursor = System.Windows.Forms.Cursors.Default
		Me.CloseWindow.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.CloseWindow.TabStop = True
		Me.CloseWindow.Name = "CloseWindow"
		Me.Label3.Text = "This sample program will show how to use the ""UseMacro"" and ""DefMacro"".  The ""Send One Request"" Button will build the InvoiceAdd and ReceivePayment transactions in one request.  The ReceivePayment uses the UseMacro to refer to the InvoiceAdd transaction.  The ""Send Two Requests"" button will show that the DefMacro value is persistent over separate requests."
		Me.Label3.Size = New System.Drawing.Size(425, 65)
		Me.Label3.Location = New System.Drawing.Point(16, 8)
		Me.Label3.TabIndex = 8
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
		Me.Label2.Text = "An InvoiceAdd and ReceivePayment for that invoice will be sent to QuickBooks."
		Me.Label2.Size = New System.Drawing.Size(417, 17)
		Me.Label2.Location = New System.Drawing.Point(16, 256)
		Me.Label2.TabIndex = 3
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
		Me.Label1.Text = "Select an Estimate below.  The information from the estimate will be used in the InvoiceAdd."
		Me.Label1.Size = New System.Drawing.Size(545, 17)
		Me.Label1.Location = New System.Drawing.Point(8, 88)
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
		Me.Controls.Add(Frame2)
		Me.Controls.Add(_SendOneRequest_1)
		Me.Controls.Add(Frame1)
		Me.Controls.Add(estimatesList)
		Me.Controls.Add(CloseWindow)
		Me.Controls.Add(Label3)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label1)
		Me.Frame2.Controls.Add(_SendTwoRequests_0)
		Me.SendOneRequest.SetIndex(_SendOneRequest_1, CType(1, Short))
		Me.SendTwoRequests.SetIndex(_SendTwoRequests_0, CType(0, Short))
		CType(Me.SendTwoRequests, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.SendOneRequest, System.ComponentModel.ISupportInitialize).EndInit()
		Me.Frame2.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class