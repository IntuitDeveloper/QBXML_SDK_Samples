Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
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
	Private requestDisplay As New DisplayForm
	Private responseDisplay As New DisplayForm
	Private displayInit As Boolean
	
	
	
	Private Sub cmdClearFilters_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClearFilters.Click
		AccountFilterList.SelectedIndex = -1
		AccountFilterList.Refresh()
		
		EntityFilterList.SelectedIndex = -1
		EntityFilterList.Refresh()
		
		ItemTypeFilterList.SelectedIndex = -1
		ItemTypeFilterList.Refresh()
		
		TxnTypeFilterList.SelectedIndex = -1
		TxnTypeFilterList.Refresh()
		
		RowsList.SelectedIndex = -1
		RowsList.Refresh()
	End Sub
	
	'
	' Set the default output path when the form is first loaded.
	'
	Private Sub UIForm_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		' Default location for html file
		htmlFile.Text = My.Application.Info.DirectoryPath & "\output.html"
		displayInit = False
		
	End Sub
	
	
	
	'
	' Browse for a company file using the MS Common Dialog
	' Control
	'
	Private Sub browseButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles browseButton.Click
		On Error GoTo ErrorHandler


        ' Set flags
        browseDialogOpen.ShowReadOnly = False

        ' Set filters
        browseDialogOpen.Filter = "All Files (*.*)|*.*|QB Company Files" & "(*.qbw)|*.qbw"

        ' Specify default filter
        browseDialogOpen.FilterIndex = 2
        Dim dlgResult As DialogResult


        ' Display the Open dialog box
        dlgResult = browseDialogOpen.ShowDialog()
        If dlgResult = DialogResult.OK Then
            ' set up the selected file in Text_Schema textBox
            qbFile.Text = browseDialogOpen.FileName
        End If
        Exit Sub

ErrorHandler: 
		' error occurs, do nothing
		Exit Sub
		
	End Sub
	
	
	
	
	
	
	'
	' Generate and display the report when the go button is
	' clicked if a company file has already been selected.
	'
	Private Sub goButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles goButton.Click
		' On Error GoTo ErrorHandler
		
		' Get company file name, use open file if blank
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
		
		' Make sure at least one Column is selected
		If (ColumnsList.SelectedItems.Count = 0) Then
			MsgBox("You must select at least one column to display.")
			Exit Sub
		End If
		
		' Make sure Summarize Rows By is selected
		If (RowsList.SelectedIndex < 0) Then
			MsgBox("You must select summary data.")
			Exit Sub
		End If
		
		' Make sure dates are in a format that can be understood
		' by QuickBooks, i.e. yyyy-mm-dd
		If Not "" = fromDateText.Text Then
			If (Len(fromDateText.Text) <> 10) Or Not Mid(fromDateText.Text, 5, 1) = "-" Or Not Mid(fromDateText.Text, 8, 1) = "-" Then
				MsgBox("Dates must be in the format yyyy-mm-dd")
				Exit Sub
			End If
		End If
		
		If Not "" = toDateText.Text Then
			If (Len(toDateText.Text) <> 10) Or Not Mid(toDateText.Text, 5, 1) = "-" Or Not Mid(toDateText.Text, 8, 1) = "-" Then
				MsgBox("Dates must be in the format yyyy-mm-dd")
				Exit Sub
			End If
		End If
		
		Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
		
		Dim NumColumns As Short
		NumColumns = ColumnsList.SelectedItems.Count
		
		Dim ColumnsArray() As String
		ReDim ColumnsArray(NumColumns)
		
		' Loop through the columns list and add all of the
		' selected columns to the array
		Dim counter As Short
		Dim arrayIndex As Short
		arrayIndex = 1
		For counter = 0 To ColumnsList.Items.Count - 1
			If ColumnsList.GetSelected(counter) Then
				ColumnsArray(arrayIndex) = VB6.GetItemString(ColumnsList, counter)
				arrayIndex = arrayIndex + 1
			End If
		Next 
		
		Dim AccountFilter As String
		Dim EntityFilter As String
		Dim ItemFilter As String
		Dim TxnFilter As String
		Dim FromDate As String
		Dim ToDate As String
		Dim SummarizeRowsBy As String
		
		If (AccountFilterList.SelectedItems.Count = 1) Then
			AccountFilter = AccountFilterList.Text
		Else
			AccountFilter = ""
		End If
		
		If (EntityFilterList.SelectedItems.Count = 1) Then
			EntityFilter = EntityFilterList.Text
		Else
			EntityFilter = ""
		End If
		
		If (ItemTypeFilterList.SelectedItems.Count = 1) Then
			ItemFilter = ItemTypeFilterList.Text
		Else
			ItemFilter = ""
		End If
		
		If (TxnTypeFilterList.SelectedItems.Count = 1) Then
			TxnFilter = TxnTypeFilterList.Text
		Else
			TxnFilter = ""
		End If
		
		FromDate = fromDateText.Text
		ToDate = toDateText.Text
		SummarizeRowsBy = VB6.GetItemString(RowsList, RowsList.SelectedIndex)
		
		Dim xmlRequest As String
		xmlRequest = generateXMLRequest(AccountFilter, EntityFilter, ItemFilter, TxnFilter, FromDate, ToDate, ColumnsArray, SummarizeRowsBy)
		
		' Exit sub if the xml was not generated correctly
		If (ErrorCode = xmlRequest) Then
			GoTo ErrorHandler2
		End If
		
		Dim xmlResponse As String
		xmlResponse = sendReqToQB(xmlRequest, companyFile)
		
		' Exit sub if xml error from QuickBooks was already
		' reported in sendReqToQB
		If (ErrorCode = xmlResponse) Then
			GoTo ErrorHandler2
		End If
		
		' Set the two DisplayForm forms so that the user will be
		' able to view the request and response by clicking on the
		' Show Request and Show Response buttons.
		requestDisplay.setXML(PrettyXMLString(xmlRequest))
		responseDisplay.setXML(PrettyXMLString(xmlResponse))
		displayInit = True
		
		Dim newStatus As String
		newStatus = parseXMLResponse(xmlResponse, url)
		
		' Exit sub if the xml was not parsed correctly
		If (ErrorCode = newStatus) Then
			GoTo ErrorHandler2
		End If
		
		BigDisplayForm.displayResults(url)
		
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Exit Sub
		
ErrorHandler: 
		MsgBox(Err.Description)
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
	
	'----------------------------------------------------------
	'
	' Routine: PrettyPrintXMLToFile
	'
	' Description
	'
	' Takes an XML message set string and a file name as input.
	' Creates a new copy of the file and pretty prints the XML
	' message set into the file.
	'
	'----------------------------------------------------------
	Public Sub PrettyPrintXMLToFile(ByRef XMLString As String, ByRef XMLFile As String)
		Dim FileNum As Object
		
		Dim SplitXMLString() As String
		Dim IndentString As String
		Dim XMLStringLength As Integer
		Dim SplitIndex As Object
		
		IndentString = CStr(Nothing)

        FileNum = FreeFile()
        FileOpen(FileNum, XMLFile, OpenMode.Output)

        'Remove the linefeeds from the XML output string
        XMLString = Replace(XMLString, vbLf, vbNullString)
		
		SplitXMLString = Split(XMLString, "<")

        'We're expecting the first character of the XML output to be "<"
        'which result in an empty first array element, so skip it.
        SplitIndex = LBound(SplitXMLString) + 1

        Do
            If VB.Left(SplitXMLString(SplitIndex), 1) = "/" Then
                IndentString = VB.Left(IndentString, Len(IndentString) - 3)
                PrintLine(FileNum, IndentString & "<" & SplitXMLString(SplitIndex))
                SplitIndex = SplitIndex + 1
            ElseIf VB.Left(SplitXMLString(SplitIndex + 1), 1) = "/" Then
                If InStr(1, VB.Left(SplitXMLString(SplitIndex), InStr(1, SplitXMLString(SplitIndex), ">")), " ") > 0 Then
                    PrintLine(FileNum, IndentString & "<" & SplitXMLString(SplitIndex))
                    SplitIndex = SplitIndex + 1
                Else
                    PrintLine(FileNum, IndentString & "<" & SplitXMLString(SplitIndex) & "<" & SplitXMLString(SplitIndex + 1))
                    SplitIndex = SplitIndex + 2
                End If
            Else
                PrintLine(FileNum, IndentString & "<" & SplitXMLString(SplitIndex))
                If SplitXMLString(SplitIndex) <> "?xml version=""1.0"" ?>" And SplitXMLString(SplitIndex) <> "!DOCTYPE QBXML PUBLIC '-//INTUIT//DTD QBXML QBD 1.0//EN' >" And VB.Right(SplitXMLString(SplitIndex), 2) <> "/>" Then
                    IndentString = IndentString & "   "
                End If
                SplitIndex = SplitIndex + 1
            End If
        Loop Until SplitIndex >= UBound(SplitXMLString)

        If VB.Left(SplitXMLString(UBound(SplitXMLString)), 1) = "/" Then
			IndentString = VB.Left(IndentString, Len(IndentString) - 3)
		End If

        PrintLine(FileNum, IndentString & "<" & SplitXMLString(UBound(SplitXMLString)))

        FileClose(FileNum)
    End Sub
End Class