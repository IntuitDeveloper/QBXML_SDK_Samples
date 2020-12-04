Option Strict Off
Option Explicit On 
Imports Interop.QBFC13

Friend Class UIForm
	Inherits System.Windows.Forms.Form
#Region "Windows Form Designer generated code "
	Public Sub New()
		MyBase.New()
		If m_vb6FormDefInstance Is Nothing Then
			If m_InitializingDefInstance Then
				m_vb6FormDefInstance = Me
			Else
				Try 
					'For the start-up form, the first instance created is the default instance.
					If System.Reflection.Assembly.GetExecutingAssembly.EntryPoint.DeclaringType Is Me.GetType Then
						m_vb6FormDefInstance = Me
					End If
				Catch
				End Try
			End If
		End If
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
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
	Public WithEvents showResButton As System.Windows.Forms.Button
	Public WithEvents showReqButton As System.Windows.Forms.Button
	Public WithEvents exitButton As System.Windows.Forms.Button
	Public WithEvents goButton As System.Windows.Forms.Button
	Public WithEvents qbFile As System.Windows.Forms.TextBox
	Public WithEvents browseButton As System.Windows.Forms.Button
	Public WithEvents htmlFile As System.Windows.Forms.TextBox
	Public WithEvents embedBrowser As AxSHDocVw.AxWebBrowser
	Public WithEvents browseDialog As AxMSComDlg.AxCommonDialog
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(UIForm))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me.showResButton = New System.Windows.Forms.Button
		Me.showReqButton = New System.Windows.Forms.Button
		Me.exitButton = New System.Windows.Forms.Button
		Me.goButton = New System.Windows.Forms.Button
		Me.qbFile = New System.Windows.Forms.TextBox
		Me.browseButton = New System.Windows.Forms.Button
		Me.htmlFile = New System.Windows.Forms.TextBox
		Me.embedBrowser = New AxSHDocVw.AxWebBrowser
		Me.browseDialog = New AxMSComDlg.AxCommonDialog
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		CType(Me.embedBrowser, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.browseDialog, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.Text = "QBSDK 2.0 --- Expense By Vendor Summary Report"
		Me.ClientSize = New System.Drawing.Size(828, 539)
		Me.Location = New System.Drawing.Point(4, 23)
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
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
		Me.Name = "UIForm"
		Me.showResButton.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.showResButton.Text = "Show Response qbXML"
		Me.showResButton.Size = New System.Drawing.Size(145, 25)
		Me.showResButton.Location = New System.Drawing.Point(420, 84)
		Me.showResButton.TabIndex = 10
		Me.showResButton.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.showResButton.BackColor = System.Drawing.SystemColors.Control
		Me.showResButton.CausesValidation = True
		Me.showResButton.Enabled = True
		Me.showResButton.ForeColor = System.Drawing.SystemColors.ControlText
		Me.showResButton.Cursor = System.Windows.Forms.Cursors.Default
		Me.showResButton.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.showResButton.TabStop = True
		Me.showResButton.Name = "showResButton"
		Me.showReqButton.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.showReqButton.Text = "Show Request qbXML"
		Me.showReqButton.Size = New System.Drawing.Size(145, 25)
		Me.showReqButton.Location = New System.Drawing.Point(248, 84)
		Me.showReqButton.TabIndex = 9
		Me.showReqButton.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.showReqButton.BackColor = System.Drawing.SystemColors.Control
		Me.showReqButton.CausesValidation = True
		Me.showReqButton.Enabled = True
		Me.showReqButton.ForeColor = System.Drawing.SystemColors.ControlText
		Me.showReqButton.Cursor = System.Windows.Forms.Cursors.Default
		Me.showReqButton.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.showReqButton.TabStop = True
		Me.showReqButton.Name = "showReqButton"
		Me.exitButton.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.exitButton.Text = "&Exit Program"
		Me.exitButton.Size = New System.Drawing.Size(145, 25)
		Me.exitButton.Location = New System.Drawing.Point(592, 84)
		Me.exitButton.TabIndex = 8
		Me.exitButton.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.exitButton.BackColor = System.Drawing.SystemColors.Control
		Me.exitButton.CausesValidation = True
		Me.exitButton.Enabled = True
		Me.exitButton.ForeColor = System.Drawing.SystemColors.ControlText
		Me.exitButton.Cursor = System.Windows.Forms.Cursors.Default
		Me.exitButton.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.exitButton.TabStop = True
		Me.exitButton.Name = "exitButton"
		Me.goButton.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.goButton.Text = "&Generate Report"
		Me.goButton.Size = New System.Drawing.Size(145, 25)
		Me.goButton.Location = New System.Drawing.Point(76, 84)
		Me.goButton.TabIndex = 4
		Me.goButton.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.goButton.BackColor = System.Drawing.SystemColors.Control
		Me.goButton.CausesValidation = True
		Me.goButton.Enabled = True
		Me.goButton.ForeColor = System.Drawing.SystemColors.ControlText
		Me.goButton.Cursor = System.Windows.Forms.Cursors.Default
		Me.goButton.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.goButton.TabStop = True
		Me.goButton.Name = "goButton"
		Me.qbFile.AutoSize = False
		Me.qbFile.Size = New System.Drawing.Size(241, 25)
		Me.qbFile.Location = New System.Drawing.Point(292, 16)
		Me.qbFile.TabIndex = 3
		Me.qbFile.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.qbFile.AcceptsReturn = True
		Me.qbFile.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.qbFile.BackColor = System.Drawing.SystemColors.Window
		Me.qbFile.CausesValidation = True
		Me.qbFile.Enabled = True
		Me.qbFile.ForeColor = System.Drawing.SystemColors.WindowText
		Me.qbFile.HideSelection = True
		Me.qbFile.ReadOnly = False
		Me.qbFile.Maxlength = 0
		Me.qbFile.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.qbFile.MultiLine = False
		Me.qbFile.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.qbFile.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.qbFile.TabStop = True
		Me.qbFile.Visible = True
		Me.qbFile.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.qbFile.Name = "qbFile"
		Me.browseButton.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.browseButton.Text = "&Browse"
		Me.browseButton.Size = New System.Drawing.Size(105, 25)
		Me.browseButton.Location = New System.Drawing.Point(536, 16)
		Me.browseButton.TabIndex = 2
		Me.browseButton.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.browseButton.BackColor = System.Drawing.SystemColors.Control
		Me.browseButton.CausesValidation = True
		Me.browseButton.Enabled = True
		Me.browseButton.ForeColor = System.Drawing.SystemColors.ControlText
		Me.browseButton.Cursor = System.Windows.Forms.Cursors.Default
		Me.browseButton.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.browseButton.TabStop = True
		Me.browseButton.Name = "browseButton"
		Me.htmlFile.AutoSize = False
		Me.htmlFile.Size = New System.Drawing.Size(349, 25)
		Me.htmlFile.Location = New System.Drawing.Point(292, 48)
		Me.htmlFile.TabIndex = 1
		Me.htmlFile.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.htmlFile.AcceptsReturn = True
		Me.htmlFile.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.htmlFile.BackColor = System.Drawing.SystemColors.Window
		Me.htmlFile.CausesValidation = True
		Me.htmlFile.Enabled = True
		Me.htmlFile.ForeColor = System.Drawing.SystemColors.WindowText
		Me.htmlFile.HideSelection = True
		Me.htmlFile.ReadOnly = False
		Me.htmlFile.Maxlength = 0
		Me.htmlFile.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.htmlFile.MultiLine = False
		Me.htmlFile.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.htmlFile.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.htmlFile.TabStop = True
		Me.htmlFile.Visible = True
		Me.htmlFile.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.htmlFile.Name = "htmlFile"
		embedBrowser.OcxState = CType(resources.GetObject("embedBrowser.OcxState"), System.Windows.Forms.AxHost.State)
		Me.embedBrowser.Size = New System.Drawing.Size(789, 409)
		Me.embedBrowser.Location = New System.Drawing.Point(20, 116)
		Me.embedBrowser.TabIndex = 0
		browseDialog.OcxState = CType(resources.GetObject("browseDialog.OcxState"), System.Windows.Forms.AxHost.State)
		Me.browseDialog.Location = New System.Drawing.Point(24, 56)
		Me.browseDialog.Name = "browseDialog"
		Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.Label3.Text = "(blank to use currently open file)"
		Me.Label3.Size = New System.Drawing.Size(165, 13)
		Me.Label3.Location = New System.Drawing.Point(124, 28)
		Me.Label3.TabIndex = 7
		Me.Label3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label1.Text = "Select QuickBooks Company File:"
		Me.Label1.Size = New System.Drawing.Size(165, 21)
		Me.Label1.Location = New System.Drawing.Point(124, 16)
		Me.Label1.TabIndex = 6
		Me.Label1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label2.Text = "Select Output Html File:"
		Me.Label2.Size = New System.Drawing.Size(165, 17)
		Me.Label2.Location = New System.Drawing.Point(124, 52)
		Me.Label2.TabIndex = 5
		Me.Label2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
		Me.Controls.Add(showResButton)
		Me.Controls.Add(showReqButton)
		Me.Controls.Add(exitButton)
		Me.Controls.Add(goButton)
		Me.Controls.Add(qbFile)
		Me.Controls.Add(browseButton)
		Me.Controls.Add(htmlFile)
		Me.Controls.Add(embedBrowser)
		Me.Controls.Add(browseDialog)
		Me.Controls.Add(Label3)
		Me.Controls.Add(Label1)
		Me.Controls.Add(Label2)
		CType(Me.browseDialog, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.embedBrowser, System.ComponentModel.ISupportInitialize).EndInit()
	End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As UIForm
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As UIForm
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New UIForm()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	' UIForm.frm
    ' Created Sept, 2002
	'
	' This file contains the code for the User Interface form of
	' the project.  The UI allows the user to first select a
	' QuickBooks company file and a location to store the output
	' .html file.  When the GenerateReport button is pressed,
	' it kicks off all of the main action.
	'
	' Copyright © 2002-2013 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	'
	'

    ' Variables needed for displaying qbXML request and response
	' in DisplayForm forms.
	Private requestDisplay As New DisplayForm
	Private responseDisplay As New DisplayForm
	Private displayInit As Boolean

    '
	' Set default values when the form is loaded.
	'
	Private Sub UIForm_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		' Load a blank page on the embedded browser to avoid
		' file not found pages
		embedBrowser.Navigate2(VB6.GetPath & "\blank.html")
		
		' Default location for html file
		htmlFile.Text = VB6.GetPath & "\output.html"
		
	End Sub

    '
	' Browse for a company file using the MS Common Dialog
	' Control
	'
	Private Sub browseButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles browseButton.Click
		On Error GoTo ErrorHandler
		
		' Set CancelError to True
		browseDialog.CancelError = True
		
		' Set flags
		browseDialog.Flags = MSComDlg.FileOpenConstants.cdlOFNHideReadOnly
		
		' Set filters
		browseDialog.Filter = "All Files (*.*)|*.*|QB Company Files" & "(*.qbw)|*.qbw"
		
		' Specify default filter
		browseDialog.FilterIndex = 2
		
		' Display the Open dialog box
		browseDialog.ShowOpen()
		
		' Set up the selected file in textBox
		qbFile.Text = browseDialog.fileName
		Exit Sub
		
ErrorHandler: 
		' If an error occurs, do nothing
		Exit Sub
	End Sub

    '
	' Generate and display the report when the go button is
	' clicked if a company file has already been selected.
	'
	Private Sub goButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles goButton.Click
        On Error GoTo ErrorHandler

		' Get company file name, use open company file if blank
		Dim companyFile As String
		companyFile = qbFile.Text
		
		' Make sure we have an output file path
		Dim url As String
		
		If ("" = htmlFile.Text) Then
			MsgBox("Please fill in a path for the html output file.")
			Exit Sub
		Else
			url = htmlFile.Text
		End If
		
		UIForm.DefInstance.Cursor = System.Windows.Forms.Cursors.WaitCursor
		
        Dim xmlResponse As IMsgSetResponse
        Dim xmlRequest As IMsgSetRequest

        ' create the request and send the request to QB
        ' receive the response back from QB
        Dim bError As Boolean
        bError = sendReqToQB(companyFile, xmlRequest, xmlResponse)

        ' Exit sub if error from QuickBooks was already
		' reported in sendReqToQB
        If (bError) Then
            GoTo ErrorHandler2
        End If

        ' Set the two DisplayForm forms so that the user will be
        ' able to view the request and response by clicking on the
        ' Show Request and Show Response buttons.
        requestDisplay.setXML(xmlRequest.ToXMLString())
        responseDisplay.setXML(xmlResponse.ToXMLString())
        displayInit = True

        ' process the response and create the html page to display
        Dim bSuccess As Boolean
        bSuccess = parseResponse(xmlResponse, url)

        ' Exit sub if the xml was not parsed correctly
        If (Not bSuccess) Then
            GoTo ErrorHandler2
        End If

        ' Finally, display the result html in the embedded browser
        embedBrowser.Navigate2(url)

        UIForm.DefInstance.Cursor = System.Windows.Forms.Cursors.Default
        Exit Sub

ErrorHandler:
        UIForm.DefInstance.Cursor = System.Windows.Forms.Cursors.Default
        MsgBox(Err.Description, MsgBoxStyle.Critical)
        Exit Sub

        ' Used if another function has already displayed
        ' and error message
ErrorHandler2:
        UIForm.DefInstance.Cursor = System.Windows.Forms.Cursors.Default
        Exit Sub
	End Sub

    '
	' Display XML Request if one has been generated already.
	'
	Private Sub showReqButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles showReqButton.Click
		On Error GoTo showReqErrorHandler
		
		If displayInit Then
			requestDisplay.showXML("qbXML Request:")
		Else
			MsgBox("Please click Generate Report before trying to " & "view the qbXML request.")
		End If
		
		Exit Sub
		
showReqErrorHandler: 
		MsgBox(Err.Description)
		Exit Sub
	End Sub

    '
	' Display XML Response if one has been generated already.
	'
	Private Sub showResButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles showResButton.Click
		On Error GoTo showResErrorHandler
		
		If displayInit Then
			responseDisplay.showXML("qbXML Response:")
		Else
			MsgBox("Please click Generate Report before trying to " & "view the qbXML response.")
		End If
		
		Exit Sub
		
showResErrorHandler: 
		MsgBox(Err.Description)
		Exit Sub
	End Sub

    '
	' Quit the program.
	'
	Private Sub exitButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles exitButton.Click
		Me.Close()
	End Sub
End Class
