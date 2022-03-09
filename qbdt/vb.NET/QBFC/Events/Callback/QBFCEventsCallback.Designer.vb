<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class QBFCEventsCallbackForm
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
	Public WithEvents Query As System.Windows.Forms.Button
	Public WithEvents QueryResult As System.Windows.Forms.TextBox
	Public WithEvents ErrorMsg As System.Windows.Forms.TextBox
	Public WithEvents UIEvent As System.Windows.Forms.TextBox
	Public WithEvents DataEvent As System.Windows.Forms.TextBox
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(QBFCEventsCallbackForm))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Query = New System.Windows.Forms.Button
		Me.QueryResult = New System.Windows.Forms.TextBox
		Me.ErrorMsg = New System.Windows.Forms.TextBox
		Me.UIEvent = New System.Windows.Forms.TextBox
		Me.DataEvent = New System.Windows.Forms.TextBox
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me.Label4 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "QBFCEventsCallback"
		Me.ClientSize = New System.Drawing.Size(614, 399)
		Me.Location = New System.Drawing.Point(4, 30)
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
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
		Me.Name = "QBFCEventsCallbackForm"
		Me.Query.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Query.Text = "Query"
		Me.Query.Size = New System.Drawing.Size(105, 41)
		Me.Query.Location = New System.Drawing.Point(72, 312)
		Me.Query.TabIndex = 8
		Me.Query.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Query.BackColor = System.Drawing.SystemColors.Control
		Me.Query.CausesValidation = True
		Me.Query.Enabled = True
		Me.Query.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Query.Cursor = System.Windows.Forms.Cursors.Default
		Me.Query.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Query.TabStop = True
		Me.Query.Name = "Query"
		Me.QueryResult.AutoSize = False
		Me.QueryResult.Size = New System.Drawing.Size(329, 97)
		Me.QueryResult.Location = New System.Drawing.Point(208, 280)
		Me.QueryResult.ReadOnly = True
		Me.QueryResult.MultiLine = True
		Me.QueryResult.ScrollBars = System.Windows.Forms.ScrollBars.Both
		Me.QueryResult.WordWrap = False
		Me.QueryResult.TabIndex = 7
		Me.QueryResult.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.QueryResult.AcceptsReturn = True
		Me.QueryResult.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.QueryResult.BackColor = System.Drawing.SystemColors.Window
		Me.QueryResult.CausesValidation = True
		Me.QueryResult.Enabled = True
		Me.QueryResult.ForeColor = System.Drawing.SystemColors.WindowText
		Me.QueryResult.HideSelection = True
		Me.QueryResult.Maxlength = 0
		Me.QueryResult.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.QueryResult.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.QueryResult.TabStop = True
		Me.QueryResult.Visible = True
		Me.QueryResult.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.QueryResult.Name = "QueryResult"
		Me.ErrorMsg.AutoSize = False
		Me.ErrorMsg.Size = New System.Drawing.Size(361, 73)
		Me.ErrorMsg.Location = New System.Drawing.Point(96, 168)
		Me.ErrorMsg.ReadOnly = True
		Me.ErrorMsg.MultiLine = True
		Me.ErrorMsg.ScrollBars = System.Windows.Forms.ScrollBars.Both
		Me.ErrorMsg.WordWrap = False
		Me.ErrorMsg.TabIndex = 5
		Me.ErrorMsg.Text = "" & Chr(13) & Chr(10) & "  " & Chr(13) & Chr(10)
		Me.ErrorMsg.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ErrorMsg.AcceptsReturn = True
		Me.ErrorMsg.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.ErrorMsg.BackColor = System.Drawing.SystemColors.Window
		Me.ErrorMsg.CausesValidation = True
		Me.ErrorMsg.Enabled = True
		Me.ErrorMsg.ForeColor = System.Drawing.SystemColors.WindowText
		Me.ErrorMsg.HideSelection = True
		Me.ErrorMsg.Maxlength = 0
		Me.ErrorMsg.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.ErrorMsg.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ErrorMsg.TabStop = True
		Me.ErrorMsg.Visible = True
		Me.ErrorMsg.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.ErrorMsg.Name = "ErrorMsg"
		Me.UIEvent.AutoSize = False
		Me.UIEvent.Size = New System.Drawing.Size(265, 73)
		Me.UIEvent.Location = New System.Drawing.Point(312, 72)
		Me.UIEvent.ReadOnly = True
		Me.UIEvent.MultiLine = True
		Me.UIEvent.ScrollBars = System.Windows.Forms.ScrollBars.Both
		Me.UIEvent.WordWrap = False
		Me.UIEvent.TabIndex = 1
		Me.UIEvent.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.UIEvent.AcceptsReturn = True
		Me.UIEvent.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.UIEvent.BackColor = System.Drawing.SystemColors.Window
		Me.UIEvent.CausesValidation = True
		Me.UIEvent.Enabled = True
		Me.UIEvent.ForeColor = System.Drawing.SystemColors.WindowText
		Me.UIEvent.HideSelection = True
		Me.UIEvent.Maxlength = 0
		Me.UIEvent.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.UIEvent.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.UIEvent.TabStop = True
		Me.UIEvent.Visible = True
		Me.UIEvent.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.UIEvent.Name = "UIEvent"
		Me.DataEvent.AutoSize = False
		Me.DataEvent.Size = New System.Drawing.Size(265, 73)
		Me.DataEvent.Location = New System.Drawing.Point(40, 72)
		Me.DataEvent.ReadOnly = True
		Me.DataEvent.MultiLine = True
		Me.DataEvent.ScrollBars = System.Windows.Forms.ScrollBars.Both
		Me.DataEvent.WordWrap = False
		Me.DataEvent.TabIndex = 0
		Me.DataEvent.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.DataEvent.AcceptsReturn = True
		Me.DataEvent.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.DataEvent.BackColor = System.Drawing.SystemColors.Window
		Me.DataEvent.CausesValidation = True
		Me.DataEvent.Enabled = True
		Me.DataEvent.ForeColor = System.Drawing.SystemColors.WindowText
		Me.DataEvent.HideSelection = True
		Me.DataEvent.Maxlength = 0
		Me.DataEvent.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.DataEvent.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.DataEvent.TabStop = True
		Me.DataEvent.Visible = True
		Me.DataEvent.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.DataEvent.Name = "DataEvent"
		Me.Frame1.Text = "Customer Query"
		Me.Frame1.Size = New System.Drawing.Size(537, 121)
		Me.Frame1.Location = New System.Drawing.Point(40, 264)
		Me.Frame1.TabIndex = 9
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.BackColor = System.Drawing.SystemColors.Control
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame1.Name = "Frame1"
		Me.Label4.Text = "Error Msg"
		Me.Label4.Size = New System.Drawing.Size(169, 25)
		Me.Label4.Location = New System.Drawing.Point(104, 152)
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
		Me.Label3.Text = "Company File Close Event"
		Me.Label3.Size = New System.Drawing.Size(185, 25)
		Me.Label3.Location = New System.Drawing.Point(312, 48)
		Me.Label3.TabIndex = 4
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
		Me.Label2.Text = "Customer Add Event"
		Me.Label2.Size = New System.Drawing.Size(161, 17)
		Me.Label2.Location = New System.Drawing.Point(40, 48)
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
		Me.Label1.Text = "This sample shows how to receive events.  This sample works with the QBFCEventsSubscriber sample."
		Me.Label1.Size = New System.Drawing.Size(497, 33)
		Me.Label1.Location = New System.Drawing.Point(56, 8)
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
		Me.Controls.Add(Query)
		Me.Controls.Add(QueryResult)
		Me.Controls.Add(ErrorMsg)
		Me.Controls.Add(UIEvent)
		Me.Controls.Add(DataEvent)
		Me.Controls.Add(Frame1)
		Me.Controls.Add(Label4)
		Me.Controls.Add(Label3)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label1)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class