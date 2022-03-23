Option Strict Off
Option Explicit On
Module qbooks
	' QBooks.bas
	' Created July, 2002
	'
	' This module contains everything needed to interact with QuickBooks.
	' It has a function to generate the qbxml request for the report,
	' a function to call ParserHelper to parse the response, and
	' most importantly, the function sendReqToQB which connects to
	' QuickBooks to send the request and receive a response.
	'
	' Copyright © 2021-2022 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	'
	
	Public Const ERRORCODE As String = "ERROR"
	Public Const SUCCESSCODE As String = "SUCCESS"
	
	
	
	'
	' Generate request QBXML asking for an ExpenseByVendorSummary
	' report.  MSXML v4.0 is needed for this function.  The
	' xml request will look something like:
	'
	' <?xml version="1.0" ?>
	' <!DOCTYPE QBXML PUBLIC '-//INTUIT//DTD QBXML QBD 2.0//EN' >
	' <QBXML>
	'     <QBXMLMsgsRq onError="stopOnError" >
	'         <GeneralSummaryReportQueryRq requestID="1" >
	'             <GeneralSummaryReportType>ExpenseByVendorSummary</GeneralSummaryReportType>
	'             <ReportDateMacro>ThisYearToDate</ReportDateMacro>
	'             <SummarizeColumnsBy>Month</SummarizeColumnsBy>
	'             <ReportCalendar>FiscalYear</ReportCalendar>
	'         </GeneralSummaryReportQueryRq>
	'     </QBXMLMsgsRq>
	' </QBXML>
	'
	'
	Public Function generateXMLRequest() As String
		On Error GoTo ErrorHandler

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
		
		msgsRq.Attributes.setNamedItem(onErrorAttr)
		
		Dim genSumReportQuery As MSXML2.IXMLDOMNode
		genSumReportQuery = doc.createElement("GeneralSummaryReportQueryRq")
		msgsRq.appendChild(genSumReportQuery)
		
		Dim rqIdAttr As MSXML2.IXMLDOMAttribute
		rqIdAttr = doc.createAttribute("requestID")
		rqIdAttr.Text = "1"
		
		genSumReportQuery.Attributes.setNamedItem(rqIdAttr)
		
		Dim genSumRepType As MSXML2.IXMLDOMNode
		genSumRepType = doc.createElement("GeneralSummaryReportType")
		genSumRepType.Text = "ExpenseByVendorSummary"
		genSumReportQuery.appendChild(genSumRepType)
		
		Dim repDateMacro As MSXML2.IXMLDOMNode
		repDateMacro = doc.createElement("ReportDateMacro")
		repDateMacro.Text = "ThisYearToDate"
		genSumReportQuery.appendChild(repDateMacro)
		
		Dim sumColsBy As MSXML2.IXMLDOMNode
		sumColsBy = doc.createElement("SummarizeColumnsBy")
		sumColsBy.Text = "Month"
		genSumReportQuery.appendChild(sumColsBy)
		
		Dim reportCalendar As MSXML2.IXMLDOMNode
		reportCalendar = doc.createElement("ReportCalendar")
		reportCalendar.Text = "FiscalYear"
		genSumReportQuery.appendChild(reportCalendar)
		
		' Add typical header lines to the beginning of the string
		' of xml now stored in doc.xml.  Notice the version number
		' of QBXML is now 2.0.  This is the DOCTYPE which should be
		' used with the SDK 2.0.
		Dim requestXML As String
		requestXML = "<?xml version=""1.0"" ?>" & "<?qbxml version=""2.0"" ?>" & doc.xml
		
		generateXMLRequest = requestXML
		Exit Function
		
ErrorHandler: 
		MsgBox("Error generating xml request: " & Err.Description, MsgBoxStyle.Critical)
		generateXMLRequest = ERRORCODE
		Exit Function
	End Function
	
	
	
	'
	' Parses the response using the class module ParserHelper,
	' and generate the .html report based on the response.
	'
	Public Function parseXMLResponse(ByRef responseXML As String, ByRef url As String) As String
		On Error GoTo ErrorHandler
		
		Dim responseParser As New ParserHelper
		Dim rpResult As Boolean
		rpResult = responseParser.ParseReportResponseXML(responseXML)
		
		If rpResult = False Then
			parseXMLResponse = ERRORCODE
			Exit Function
		End If
		
		Dim outputString As String
		outputString = responseParser.OutputStr
		
		CreateFile(url, outputString)
		Exit Function
		
ErrorHandler: 
		MsgBox("Error parsing response: " & Err.Description, MsgBoxStyle.Critical)
		parseXMLResponse = ERRORCODE
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
		requestProcessor.OpenConnection("", "IDN Desktop VB QBD Expense Report Sample App")
		
		' Begin a session and obtain a ticket
		Dim ticket As String
		ticket = requestProcessor.BeginSession(qbFile, QBXMLRP2Lib.QBFileMode.qbFileOpenDoNotCare)

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
		sendReqToQB = ERRORCODE
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
		
		QBXMLMsgsRqNode.Attributes.setNamedItem(onErrorAttr)
		
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
	
	Function qbXMLAddProlog(ByRef supportedVersion As String, ByRef xml As String) As String
		Dim qbXMLVersionSpec As String
		If (Val(supportedVersion) >= 2) Then
			qbXMLVersionSpec = "<?qbxml version=""" & supportedVersion & """?>"
		ElseIf (supportedVersion = "1.1") Then 
			qbXMLVersionSpec = "<!DOCTYPE QBXML PUBLIC '-//INTUIT//DTD QBXML QBD " & supportedVersion & "//EN' 'http://developer.intuit.com'>"
		Else
			MsgBox("You are apparently running QuickBooks 2002 Release 1, we strongly recommend that you use QuickBooks' online update feature to obtain the latest fixes and enhancements", MsgBoxStyle.Exclamation)
			qbXMLVersionSpec = "<!DOCTYPE QBXML PUBLIC '-//INTUIT//DTD QBXML QBD " & supportedVersion & "//EN' 'http://developer.intuit.com'>"
		End If
		qbXMLAddProlog = "<?xml version=""1.0""?>" & vbCrLf & qbXMLVersionSpec & xml
	End Function
End Module