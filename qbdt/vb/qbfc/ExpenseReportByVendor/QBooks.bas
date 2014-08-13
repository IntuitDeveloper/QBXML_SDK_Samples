Attribute VB_Name = "qbooks"
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
    ' Updated to QBFC 12.0 and fixed setting max QBXML version: 09/2012
    
    '
    ' Parses the response using the class module ParserHelper,
    ' and generate the .html report based on the response.
    '
    Public Function parseResponse(ByRef responseSet As IMsgSetResponse, _
            ByRef url As String) As Boolean
        On Error GoTo ErrorHandler

        Dim responseParser As New ParserHelper
        Dim rpResult As Boolean
        rpResult = responseParser.ParseReportResponseSet(responseSet)

        If False = rpResult Then
            parseResponse = False
            Exit Function
        End If

        Dim outputString As String
        outputString = responseParser.OutputStr

        CreateFile url, outputString

        parseResponse = True

        Exit Function

ErrorHandler:
        MsgBox "Error parsing response: " & Err.Description
        parseResponse = False
        Exit Function
    End Function

    '
    ' Send the request to QuickBooks using the specified file or the
    ' open file if none is specified.  Return the response set.
    '
    Public Function sendReqToQB(qbFile As String, _
            requestSet As IMsgSetRequest, _
            responseSet As IMsgSetResponse) As Boolean
        On Error GoTo ErrorHandler

        'Create the session manager object
        Dim SessionManager As QBSessionManager
        Set SessionManager = New QBSessionManager

        ' Open a connection with QuickBooks
        SessionManager.OpenConnection "", "IDN Desktop VB QBD Expense Report Sample App"

        ' Begin a session which obtains a ticket
        SessionManager.BeginSession qbFile, omDontCare

        'Check to make sure the QuickBooks we're working with supports version 2 of the SDK
        Dim supportedVersion As Double
        supportedVersion = Val(QBFCLatestVersion(SessionManager))
        If (supportedVersion >= 6#) Then
            booSupports2dot0 = True
            Set requestSet = SessionManager.CreateMsgSetRequest("US", 6, 0)
        ElseIf (supportedVersion >= 5#) Then
            booSupports2dot0 = True
            Set requestSet = SessionManager.CreateMsgSetRequest("US", 5, 0)
        ElseIf (supportedVersion >= 4#) Then
            booSupports2dot0 = True
            Set requestSet = SessionManager.CreateMsgSetRequest("US", 4, 0)
        ElseIf (supportedVersion >= 3#) Then
            booSupports2dot0 = True
            Set requestSet = SessionManager.CreateMsgSetRequest("US", 3, 0)
        ElseIf (supportedVersion >= 2#) Then
            booSupports2dot0 = True
            Set requestSet = SessionManager.CreateMsgSetRequest("US", 2, 0)
        End If
        
        If Not booSupports2dot0 Then
            MsgBox "This program only runs against QuickBooks installations which support the 2.0 qbXML spec.  Your version of QuickBooks does not support qbXML 2.0"
            End
        End If
        
        'Initialize the message set request's attributes
        requestSet.Attributes.OnError = ENRqOnError.roeContinue

        'Append the request to the message set request object
        Dim GeneralSummaryReportQ As IGeneralSummaryReportQuery
        Set GeneralSummaryReportQ = requestSet.AppendGeneralSummaryReportQueryRq

        'Set the values of the report we want
        '   <GeneralSummaryReportType>ExpenseByVendorSummary</GeneralSummaryReportType>
        GeneralSummaryReportQ.GeneralSummaryReportType.SetValue (ENGeneralSummaryReportType.gsrtExpenseByVendorSummary)
        '   <ReportDateMacro>ThisYearToDate</ReportDateMacro>
        GeneralSummaryReportQ.ORReportPeriod.ReportDateMacro.SetValue (ENReportDateMacro.rdmThisYearToDate)
        '   <SummarizeColumnsBy>Month</SummarizeColumnsBy>
        GeneralSummaryReportQ.SummarizeColumnsBy.SetValue (ENSummarizeColumnsBy.scbMonth)
        '   <ReportCalendar>FiscalYear</ReportCalendar>
        GeneralSummaryReportQ.ReportCalendar.SetValue (ENReportCalendar.rcFiscalYear)

        'Perform the request
        Set responseSet = SessionManager.DoRequests(requestSet)

        ' Uncomment the following to see the request and response XML for debugging
        'MsgBox requestSet.ToXMLString, vbOKOnly, "RequestXML"
        'MsgBox responseSet.ToXMLString, vbOKOnly, "ResponseXML"

        ' End the session and close the connection
        SessionManager.EndSession
        SessionManager.CloseConnection

        ' Return the xml error status
        sendReqToQB = False

        Exit Function

ErrorHandler:
        MsgBox "Error:" & CStr(Hex(Err.Number)) & ", " & Err.Description, vbOKOnly, "Send Request to QB: "
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
        Set fs = CreateObject("Scripting.FileSystemObject")

        Set fileObj = fs.CreateTextFile(fileName, True)

        fileObj.Write (outPut)
        fileObj.Close

        CreateFile = True
        Exit Function

ErrorHandler:
        MsgBox "Create File Error: " & Err.Description
        CreateFile = False
    End Function
Function QBFCLatestVersion(SessionManager As QBSessionManager) As String
    Dim strXMLVersions() As String
    'Should be able to use this, but there appears to be a bug that may cause 2.0 to be returned
    'when it should not.
    'strXMLVersions = SessionManager.QBXMLVersionsForSession
    
    Dim msgset As IMsgSetRequest
    'Use oldest version to ensure that we work with any QuickBooks (US)
    Set msgset = SessionManager.CreateMsgSetRequest("US", 1, 0)
    msgset.AppendHostQueryRq
    Dim QueryResponse As IMsgSetResponse
    Set QueryResponse = SessionManager.DoRequests(msgset)
    Dim response As IResponse
    
    ' The response list contains only one response,
    ' which corresponds to our single HostQuery request
    Set response = QueryResponse.responseList.GetAt(0)
    Dim HostResponse As IHostRet
    Set HostResponse = response.Detail
    Dim supportedVersions As IBSTRList
    Set supportedVersions = HostResponse.SupportedQBXMLVersionList
    
    Dim i As Long
    Dim vers As Double
    Dim LastVers As Double
    LastVers = 0
    For i = 0 To supportedVersions.count - 1
        vers = Val(supportedVersions.GetAt(i))
        If (vers > LastVers) Then
            LastVers = vers
            QBFCLatestVersion = supportedVersions.GetAt(i)
        End If
    Next i
End Function
