Option Strict Off
Option Explicit On
Module QBooks
    ' QBooks.bas
    ' Created July, 2002
    '
    ' This module handles all of the communication with QuickBooks
    ' in the function sendReqToQB.  It also has functions to generate
    ' the QBXML Request and to kick off the parsing of the QBXML Response.
    '
    ' Copyright © 2021-2022 Intuit Inc. All rights reserved.
    ' Use is subject to the terms specified at:
    '      http://developer.intuit.com/legal/devsite_tos.html
    '
    '


    Public Const ErrorCode As String = "ERROR"
	Public Const SuccessCode As String = "SUCCESS"
	
	
	'
	' Generate the necessary request qbxml for a CustomTxnDetail
	' report with the given parameters.  Uses MSXML version 4.
	'
	Public Function generateXMLRequest(ByRef AccountFilterStr As String, ByRef EntityFilterStr As String, ByRef ItemFilterStr As String, ByRef TxnFilterStr As String, ByRef FromDateStr As String, ByRef ToDateStr As String, ByRef ColumnsArray() As String, ByRef SummarizeRowsBy As String) As String
        '    On Error GoTo ErrorHandler

        Dim doc As New MSXML2.DOMDocument60
        doc.async = False
		doc.validateOnParse = True
		
		Dim root As MSXML2.IXMLDOMNode
		root = doc.createElement("QBXML")
		doc.appendChild(root)
		
		Dim msgsRq As MSXML2.IXMLDOMNode
		msgsRq = doc.createElement("QBXMLMsgsRq")
		root.appendChild(msgsRq)
		
		Dim onErrorAttr As MSXML2.IXMLDOMAttribute
		onErrorAttr = doc.createAttribute("onError")
		onErrorAttr.Text = "stopOnError"
        msgsRq.attributes.setNamedItem(onErrorAttr)

        Dim custDetReportQuery As MSXML2.IXMLDOMNode
		custDetReportQuery = doc.createElement("CustomDetailReportQueryRq")
		msgsRq.appendChild(custDetReportQuery)
		
		' Add the request ID for the message set
		Dim rqIdAttr As MSXML2.IXMLDOMAttribute
		rqIdAttr = doc.createAttribute("requestID")
		rqIdAttr.Text = "1"
        custDetReportQuery.attributes.setNamedItem(rqIdAttr)

        ' Add the type of report to the query
        Dim custDetRepType As MSXML2.IXMLDOMNode
		custDetRepType = doc.createElement("CustomDetailReportType")
		custDetRepType.Text = "CustomTxnDetail"
		custDetReportQuery.appendChild(custDetRepType)
		
		' Add the date filter, if one exists
		Dim repPeriod As MSXML2.IXMLDOMNode
		Dim fromRepDate As MSXML2.IXMLDOMNode
		Dim toRepDate As MSXML2.IXMLDOMNode
		If Not (FromDateStr = "") Or Not (ToDateStr = "") Then
			repPeriod = doc.createElement("ReportPeriod")
			custDetReportQuery.appendChild(repPeriod)
			If Not (FromDateStr = "") Then
				fromRepDate = doc.createElement("FromReportDate")
				fromRepDate.Text = FromDateStr
				repPeriod.appendChild(fromRepDate)
			End If
			If Not (ToDateStr = "") Then
				toRepDate = doc.createElement("ToReportDate")
				toRepDate.Text = ToDateStr
				repPeriod.appendChild(toRepDate)
			End If
		End If
		
		' Add account filter if one exists
		Dim repAccFilter As MSXML2.IXMLDOMNode
		Dim accTypeFilter As MSXML2.IXMLDOMNode
		If Not (AccountFilterStr = "") Then
			repAccFilter = doc.createElement("ReportAccountFilter")
			custDetReportQuery.appendChild(repAccFilter)
			accTypeFilter = doc.createElement("AccountTypeFilter")
			accTypeFilter.Text = AccountFilterStr
			repAccFilter.appendChild(accTypeFilter)
		End If
		
		' Add entity filter if one exists
		Dim repEntFilter As MSXML2.IXMLDOMNode
		Dim entityTypeFilter As MSXML2.IXMLDOMNode
		If Not (EntityFilterStr = "") Then
			repEntFilter = doc.createElement("ReportEntityFilter")
			custDetReportQuery.appendChild(repEntFilter)
			entityTypeFilter = doc.createElement("EntityTypeFilter")
			entityTypeFilter.Text = EntityFilterStr
			repEntFilter.appendChild(entityTypeFilter)
		End If
		
		' Add item filter if one exists
		Dim repItemFilter As MSXML2.IXMLDOMNode
		Dim itemTypeFilter As MSXML2.IXMLDOMNode
		If Not (ItemFilterStr = "") Then
			repItemFilter = doc.createElement("ReportItemFilter")
			custDetReportQuery.appendChild(repItemFilter)
			itemTypeFilter = doc.createElement("ItemTypeFilter")
			itemTypeFilter.Text = ItemFilterStr
			repItemFilter.appendChild(itemTypeFilter)
		End If
		
		' Add txn type filter if one exists
		Dim repTxnTypeFilter As MSXML2.IXMLDOMNode
		Dim txnTypeFilter As MSXML2.IXMLDOMNode
		If Not (TxnFilterStr = "") Then
			repTxnTypeFilter = doc.createElement("ReportTxnTypeFilter")
			custDetReportQuery.appendChild(repTxnTypeFilter)
			txnTypeFilter = doc.createElement("TxnTypeFilter")
			txnTypeFilter.Text = TxnFilterStr
			repTxnTypeFilter.appendChild(txnTypeFilter)
		End If
		
		'Add SummarizeRowsBy
		Dim summarizeRowsByElement As MSXML2.IXMLDOMNode
		summarizeRowsByElement = doc.createElement("SummarizeRowsBy")
		summarizeRowsByElement.Text = SummarizeRowsBy
		custDetReportQuery.appendChild(summarizeRowsByElement)
		
		' Finally add all of the elements for columns
		Dim length As Short
		length = UBound(ColumnsArray)
		Dim includeCols() As MSXML2.IXMLDOMNode
		ReDim includeCols(length)
		Dim arrayIndex As Short
		For arrayIndex = 1 To length
			includeCols(arrayIndex) = doc.createElement("IncludeColumn")
			includeCols(arrayIndex).Text = ColumnsArray(arrayIndex)
			custDetReportQuery.appendChild(includeCols(arrayIndex))
		Next 
		
		' Add typical header lines to the beginning of the string
		' of xml now stored in doc.xml.  Notice the version number
		' of QBXML is now 2.0.  This is the DOCTYPE which should be
		' used with the SDK 2.0.
		Dim requestXML As String
		requestXML = "<?xml version=""1.0"" ?>" & "<?qbxml version=""2.0"" ?>" & doc.xml
		
		generateXMLRequest = requestXML
		Exit Function
		
ErrorHandler: 
		MsgBox("Error generating xml request: " & Err.Description)
		generateXMLRequest = ErrorCode
		Exit Function
	End Function
	
	
	
	
	'
	' Parses the response using the class module ParserHelper,
	' and generate the .html report based on the response.
	'
	Public Function parseXMLResponse(ByRef responseXML As String, ByRef url As String) As String
		On Error GoTo ErrorHandler
		
		Dim ResponseParser As New ParserHelper
		Dim RPResult As Boolean
		RPResult = ResponseParser.ParseReportResponseXML(responseXML)
		
		If RPResult = False Then
			parseXMLResponse = ErrorCode
			Exit Function
		End If
		
		Dim outputString As String
		outputString = ResponseParser.OutputStr
		
		CreateFile(url, outputString)
		Exit Function
		
ErrorHandler: 
		MsgBox("Error parsing response: " & Err.Description)
		parseXMLResponse = ErrorCode
		Exit Function
	End Function
	
	
	
	'
	' Send the request to QuickBooks using the specified file or the
	' open file if none is specified.  Return the qbxml response.
	'
	Public Function sendReqToQB(ByRef requestXML As String, ByRef qbFile As String) As String
		On Error GoTo ErrorHandler
		
		Dim requestProcessor As New QBXMLRP2Lib.RequestProcessor2
		
		' Open a connection with QuickBooks
		requestProcessor.OpenConnection("IDN Desktop VB QBD CustomTxnDetail Reporting Sample App", "Desktop VB QBD CustomTxnDetail Reporting Sample App")
		
		' Begin a session and obtain a ticket
		Dim ticket As String
		ticket = requestProcessor.BeginSession(qbFile, QBXMLRP2Lib.QBFileMode.qbFileOpenDoNotCare)
		
		'Check that the right version of qbXML is supported and warn if not
		Dim latestVersion As String
        latestVersion = qbXMLLatestVersion(requestProcessor, ticket)
        Dim vers As Double

        vers = CDbl(latestVersion)
        If (vers < 2.0) Then
            MsgBox("This sample requires a version of QuickBooks which supports qbXML 2.0," & " the latest version the currently running QuickBooks supports is " & latestVersion, MsgBoxStyle.Exclamation)
        End If

        ' Send the request xml to QuickBooks for processing
        Dim responseXML As String
		responseXML = requestProcessor.ProcessRequest(ticket, requestXML)
		
		' End the session and close the connection
		requestProcessor.EndSession((ticket))
		requestProcessor.CloseConnection()
		
		' Return the xml response from QuickBooks
		sendReqToQB = responseXML
		Exit Function
		
ErrorHandler: 
		MsgBox(Err.Description, MsgBoxStyle.Critical)
		sendReqToQB = ErrorCode
		Exit Function
	End Function
	
	
	
	
	'
	' Save the output to a file.
	'
	Public Function CreateFile(ByRef fileName As String, ByRef output As String) As Boolean
		On Error GoTo ErrorHandler
		
		If output = "" Then
			CreateFile = False
			Exit Function
		End If
		
		Dim fs, fileObj As Object
		fs = CreateObject("Scripting.FileSystemObject")

        fileObj = fs.CreateTextFile(fileName, True)

        fileObj.Write(output)
        fileObj.Close()

        CreateFile = True
		Exit Function
		
ErrorHandler: 
		MsgBox("Create File Error: " & Err.Description, MsgBoxStyle.Critical)
		CreateFile = False
	End Function
	
	
	
	
	
	Public Function PrettyXMLString(ByRef InXMLString As String) As String
		
		Dim SplitXMLString() As String
		Dim IndentString As String
		Dim OutString As String
		Dim XMLString As String
		Dim XMLStringLength As Integer
		Dim SplitIndex As Object
		
		' Preserve the passed in string
		XMLString = InXMLString
		
		IndentString = CStr(Nothing)
		OutString = CStr(Nothing)
		
		'Remove the linefeeds from the XML string
		XMLString = Replace(XMLString, vbLf, vbNullString)
		
		SplitXMLString = Split(XMLString, "<")

        'We're expecting the first character of the XML string to be "<"
        'which result in an empty first array element, so skip it.
        SplitIndex = LBound(SplitXMLString) + 1

        Do
            If Left(SplitXMLString(SplitIndex), 1) = "/" Then
                IndentString = Left(IndentString, Len(IndentString) - 3)
                OutString = OutString & IndentString & "<" & SplitXMLString(SplitIndex) & vbCrLf
                SplitIndex = SplitIndex + 1
            ElseIf Left(SplitXMLString(SplitIndex + 1), 1) = "/" Then
                If InStr(1, Left(SplitXMLString(SplitIndex), InStr(1, SplitXMLString(SplitIndex), ">")), " ") > 0 Then
                    OutString = OutString & IndentString & "<" & SplitXMLString(SplitIndex) & vbCrLf
                    SplitIndex = SplitIndex + 1
                Else
                    OutString = OutString & IndentString & "<" & SplitXMLString(SplitIndex) & "<" & SplitXMLString(SplitIndex + 1) & vbCrLf
                    SplitIndex = SplitIndex + 2
                End If
            Else
                OutString = OutString & IndentString & "<" & SplitXMLString(SplitIndex) & vbCrLf
                If SplitXMLString(SplitIndex) <> "?xml version=""1.0"" ?>" And SplitXMLString(SplitIndex) <> "!DOCTYPE QBXML PUBLIC '-//INTUIT//DTD QBXML QBD 1.0//EN' >" And SplitXMLString(SplitIndex) <> "!DOCTYPE QBXML PUBLIC '-//INTUIT//DTD QBXML QBD 1.1//EN' >" And SplitXMLString(SplitIndex) <> "!DOCTYPE QBXML PUBLIC '-//INTUIT//DTD QBXML QBD 2.0//EN' >" And SplitXMLString(SplitIndex) <> "?qbxml version=""2.0"" ?>" And Right(SplitXMLString(SplitIndex), 2) <> "/>" Then
                    IndentString = IndentString & "   "
                End If
                SplitIndex = SplitIndex + 1
            End If
        Loop Until SplitIndex >= UBound(SplitXMLString)

        If Left(SplitXMLString(UBound(SplitXMLString)), 1) = "/" Then
			IndentString = Left(IndentString, Len(IndentString) - 3)
		End If
		
		OutString = OutString & IndentString & "<" & SplitXMLString(UBound(SplitXMLString))
		
		PrettyXMLString = OutString
		
	End Function
	
	Function qbXMLLatestVersion(ByRef rp As QBXMLRP2Lib.RequestProcessor2, ByRef ticket As String) As String
		Dim strXMLVersions() As String
        'Should be able to use this, but there appears to be a bug that may cause 2.0 to be returned
        'when it should not.
        'strXMLVersions = rp.QBXMLVersionsForSession(ticket)

        'Instead we use HostQuery
        'Create a DOM document object for creating our request.
        Dim xml As New MSXML2.DOMDocument60

        'Create the QBXML aggregate
        Dim rootElement As MSXML2.IXMLDOMNode
		rootElement = xml.createElement("QBXML")
		xml.appendChild(rootElement)
		
		'Add the QBXMLMsgsRq aggregate to the QBXML aggregate
		Dim QBXMLMsgsRqNode As MSXML2.IXMLDOMNode
		QBXMLMsgsRqNode = xml.createElement("QBXMLMsgsRq")
		rootElement.appendChild(QBXMLMsgsRqNode)
		
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'If we were writing a real application this is where we would add
		'a newMessageSetID so we could perform error recovery.  Any time a
		'request contains an add, delete, modify or void request developers
		'should use the error recovery mechanisms.
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		
		'Set the QBXMLMsgsRq onError attribute to continueOnError
		Dim onErrorAttr As MSXML2.IXMLDOMAttribute
		onErrorAttr = xml.createAttribute("onError")
		onErrorAttr.Text = "stopOnError"
        QBXMLMsgsRqNode.attributes.setNamedItem(onErrorAttr)

        'Add the InvoiceAddRq aggregate to QBXMLMsgsRq aggregate
        Dim HostQuery As MSXML2.IXMLDOMNode
		HostQuery = xml.createElement("HostQueryRq")
		QBXMLMsgsRqNode.appendChild(HostQuery)
		
		Dim strXMLRequest As String
		strXMLRequest = "<?xml version=""1.0"" ?>" & "<!DOCTYPE QBXML PUBLIC '-//INTUIT//DTD QBXML QBD 1.0//EN' 'http://developer.intuit.com'>" & rootElement.xml
		
		Dim strXMLResponse As String
		strXMLResponse = rp.ProcessRequest(ticket, strXMLRequest)
        Dim QueryResponse As New MSXML2.DOMDocument60

        'Parse the response XML
        QueryResponse.async = False
		QueryResponse.loadXML(strXMLResponse)
		
		Dim supportedVersions As MSXML2.IXMLDOMNodeList
		supportedVersions = QueryResponse.getElementsByTagName("SupportedQBXMLVersion")
		
		Dim VersNode As MSXML2.IXMLDOMNode
		
		Dim i As Integer
		Dim vers As Double
		Dim LastVers As Double
		LastVers = 0
		For i = 0 To supportedVersions.length - 1
			VersNode = supportedVersions.Item(i)
			vers = CDbl(VersNode.firstChild.Text)
			If (vers > LastVers) Then
				LastVers = vers
				qbXMLLatestVersion = VersNode.firstChild.Text
			End If
		Next i
	End Function
End Module