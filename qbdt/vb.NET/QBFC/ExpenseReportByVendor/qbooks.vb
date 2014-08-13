Option Strict Off
Option Explicit On 
Imports Interop.QBFC13

Module qbooks
	' QBooks.bas
    ' Created Sept, 2002
	'
	' This module contains everything needed to interact with QuickBooks.
	' It has a function to generate the qbxml request for the report,
	' a function to call ParserHelper to parse the response, and
	' most importantly, the function sendReqToQB which connects to
	' QuickBooks to send the request and receive a response.
	'
	' Copyright © 2002-2013 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	'
	
    '
    ' Parses the response using the class module ParserHelper,
	' and generate the .html report based on the response.
	'
    Public Function parseResponse(ByRef responseSet As IMsgSetResponse, _
            ByRef url As String) As Boolean
        On Error GoTo ErrorHandler

        Dim responseParser As New ParserHelper()
        Dim rpResult As Boolean
        rpResult = responseParser.ParseReportResponseSet(responseSet)

        If False = rpResult Then
            parseResponse = False
            Exit Function
        End If

        Dim outputString As String
        outputString = responseParser.OutputStr

        CreateFile(url, outputString)

        parseResponse = True

        Exit Function

ErrorHandler:
        MsgBox("Error parsing response: " & Err.Description, MsgBoxStyle.Critical)
        parseResponse = False
        Exit Function
    End Function

    '
    ' Send the request to QuickBooks using the specified file or the
    ' open file if none is specified.  Return the response set.
    '
    Public Function sendReqToQB(ByRef qbFile As String, _
            ByRef requestSet As IMsgSetRequest, _
            ByRef responseSet As IMsgSetResponse) As Boolean
        On Error GoTo ErrorHandler

        'Create the session manager object
        Dim sessionManager As QBSessionManager = New QBSessionManager()

        ' Open a connection with QuickBooks
        sessionManager.OpenConnection("Desktop VB QBD Expense Report Sample App", "IDN Desktop VB QBD Expense Report Sample App")

        ' Begin a session which obtains a ticket
        sessionManager.BeginSession(qbFile, ENOpenMode.omDontCare)

        'Check to make sure the QuickBooks we're working with supports version 2 of the SDK
        Dim strXMLVersions() As String
        strXMLVersions = sessionManager.QBXMLVersionsForSession

        Dim booSupports2dot0 As Boolean
        booSupports2dot0 = False
        Dim i As Long
        For i = LBound(strXMLVersions) To UBound(strXMLVersions)
            If (strXMLVersions(i) = "2.0") Then
                booSupports2dot0 = True
                requestSet = sessionManager.CreateMsgSetRequest("US", 2, 0)
                Exit For
            End If
        Next

        If Not booSupports2dot0 Then
            MsgBox("This program only runs against QuickBooks installations which support the 2.0 qbXML spec.  Your version of QuickBooks does not support qbXML 2.0")
            End
        End If

        'Initialize the message set request's attributes
        requestSet.Attributes.OnError = ENRqOnError.roeContinue

        'Append the request to the message set request object
        Dim GeneralSummaryReportQ As IGeneralSummaryReportQuery
        GeneralSummaryReportQ = requestSet.AppendGeneralSummaryReportQueryRq

        'Set the values of the report we want
        '	<GeneralSummaryReportType>ExpenseByVendorSummary</GeneralSummaryReportType>
        GeneralSummaryReportQ.GeneralSummaryReportType.SetValue(ENGeneralSummaryReportType.gsrtExpenseByVendorSummary)
        '	<ReportDateMacro>ThisYearToDate</ReportDateMacro>
        GeneralSummaryReportQ.ORReportPeriod.ReportDateMacro.SetValue(ENReportDateMacro.rdmThisYearToDate)
        '	<SummarizeColumnsBy>Month</SummarizeColumnsBy>
        GeneralSummaryReportQ.SummarizeColumnsBy.SetValue(ENSummarizeColumnsBy.scbMonth)
        '	<ReportCalendar>FiscalYear</ReportCalendar>
        GeneralSummaryReportQ.ReportCalendar.SetValue(ENReportCalendar.rcFiscalYear)

        'Perform the request
        responseSet = sessionManager.DoRequests(requestSet)

        ' Uncomment the following to see the request and response XML for debugging
        'MsgBox requestSet.ToXMLString, vbOKOnly, "RequestXML"
        'MsgBox responseSet.ToXMLString, vbOKOnly, "ResponseXML"

        ' End the session and close the connection
        sessionManager.EndSession()
        sessionManager.CloseConnection()

        ' Return the xml error status
        sendReqToQB = False

        Exit Function

ErrorHandler:
        MsgBox("Error:" & CStr(Hex(Err.Number)) & ", " & Err.Description, MsgBoxStyle.Critical, "Send Request to QB: ")
        sendReqToQB = True
        Exit Function
    End Function

    '
    ' Save the output html string to a file.
    '
    Public Function CreateFile(ByRef fileName As String, ByRef outPut As String) As Boolean
        On Error GoTo ErrorHandler

        If outPut = "" Then
            CreateFile = False
            Exit Function
        End If

        Dim fs, fileObj As Object
        fs = CreateObject("Scripting.FileSystemObject")

        fileObj = fs.CreateTextFile(fileName, True)

        fileObj.Write(outPut)
        fileObj.Close()

        CreateFile = True
        Exit Function

ErrorHandler:
        MsgBox("Create File Error: " & Err.Description, MsgBoxStyle.Critical)
        CreateFile = False
    End Function
End Module
