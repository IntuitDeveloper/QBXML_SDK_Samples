Option Strict Off
Option Explicit On
Friend Class UIForm
	Inherits System.Windows.Forms.Form
    ' UIForm.frm
    ' Created July, 2002
    '
    ' This file contains the code for the User Interface form of
    ' the project.  The UI allows the user to first select a
    ' QuickBooks company file and a location to store the output
    ' .html file.  When the GenerateReport button is pressed,
    ' it kicks off all of the main action.
    '
    ' Copyright © 2021-2022 Intuit Inc. All rights reserved.
    ' Use is subject to the terms specified at:
    '      http://developer.intuit.com/legal/devsite_tos.html
    '
    '
    '



    ' Variables needed for displaying qbXML request and response
    ' in DisplayForm forms.
    Private requestDisplayStr As String
    Private responseDisplayStr As String
    Private displayInit As Boolean



    '
    ' Set default values when the form is loaded.
    '
    Private Sub UIForm_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        ' Default location for html file
        htmlFile.Text = My.Application.Info.DirectoryPath & "\output.html"

    End Sub



    '
    ' Browse for a company file using the MS Common Dialog
    ' Control
    '
    Private Sub browseButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles browseButton.Click
		On Error GoTo ErrorHandler

        ' Set CancelError to True
        
        'browseDialog.CancelError = True

        ' Set flags
        
        
        browseDialogOpen.ShowReadOnly = False
		
		' Set filters
		
		browseDialogOpen.Filter = "All Files (*.*)|*.*|QB Company Files" & "(*.qbw)|*.qbw"
		
		' Specify default filter
		browseDialogOpen.FilterIndex = 2
		
		' Display the Open dialog box
		browseDialogOpen.ShowDialog()
		
		' Set up the selected file in textBox
		qbFile.Text = browseDialogOpen.FileName
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
		
		Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
		
		Dim xmlRequest As String
		xmlRequest = generateXMLRequest()
		
		' Exit sub if the xml was not generated correctly
		If (ERRORCODE = xmlRequest) Then
			GoTo ErrorHandler2
		End If
		
		Dim xmlResponse As String
		xmlResponse = sendReqToQB(xmlRequest, companyFile)
		
		' Exit sub if xml error from QuickBooks was already
		' reported in sendReqToQB
		If (ERRORCODE = xmlResponse) Then
			GoTo ErrorHandler2
		End If

        ' Set the two DisplayForm forms so that the user will be
        ' able to view the request and response by clicking on the
        ' Show Request and Show Response buttons.
        requestDisplayStr = PrettyXMLString(xmlRequest)
        responseDisplayStr = PrettyXMLString(xmlResponse)
        displayInit = True
		
		Dim newStatus As String
		newStatus = parseXMLResponse(xmlResponse, url)
		
		' Exit sub if the xml was not parsed correctly
		If (ERRORCODE = newStatus) Then
			GoTo ErrorHandler2
		End If

        ' Finally, display the result html in the embedded browser
        
        Process.Start(url)

        Me.Cursor = System.Windows.Forms.Cursors.Default
		Exit Sub
		
ErrorHandler: 
		Me.Cursor = System.Windows.Forms.Cursors.Default
		MsgBox(Err.Description, MsgBoxStyle.Critical)
		Exit Sub
		
		' Used if another function has already displayed
		' and error message
ErrorHandler2: 
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Exit Sub
	End Sub
	
	
	
	'
	' Display XML Request if one has been generated already.
	'
	Private Sub showReqButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles showReqButton.Click
		On Error GoTo showReqErrorHandler
		
		If displayInit Then
            Dim requestDisplay As New DisplayForm
            requestDisplay.setXML(requestDisplayStr)
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
            Dim responseDisplay As New DisplayForm
            responseDisplay.setXML(responseDisplayStr)
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