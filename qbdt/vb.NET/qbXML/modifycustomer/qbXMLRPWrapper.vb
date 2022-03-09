Option Strict Off
Option Explicit On
Friend Class qbXMLRPWrapper
	'-----------------------------------------------------------
	' Class Module: qbXMLRPWrapper
	'
	' Description:  Encapsulates the calls to the Request Processor which
	'               is the interface to communicate with QuickBooks.
	'               The proper calling sequence is
	'                   Start
	'                   DoRequest 1
	'                   DoRequest 2
	'                   ...
	'                   DoRequest n
	'                   Finish
	'               Each method call returns a status that must be checked
	'               to determine whether the call was successful or not.
	'
	' Created On: 11/08/2001
	' Updated to SDK 2.0: 08/05/2002
	'
	' Copyright © 2021-2022 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	'----------------------------------------------------------
	
	
	Private m_ErrNumber As Integer
	Private m_ErrMsg As String

    Private m_RP As QBXMLRP2Lib.RequestProcessor2
    Private m_Ticket As String
	
	' Initiate connection to QuickBooks (qbXMLRP)
	' Return true if successful, false otherwise
	Public Function Start(ByRef appID As String, ByRef appName As String, ByRef companyFile As String) As Boolean
		
		On Error GoTo ErrHandler
		
		m_RP = New QBXMLRP2Lib.RequestProcessor2
		
		' Open connection to qbXMLRp COM
		m_RP.OpenConnection(appID, appName)
		
		' Begin Session
		m_Ticket = m_RP.BeginSession(companyFile, QBXMLRP2Lib.QBFileMode.qbFileOpenDoNotCare)
		
		Start = True
		Exit Function
		
ErrHandler: 
		Start = False
		m_ErrNumber = Err.Number
		m_ErrMsg = Err.Description
		Exit Function
		
	End Function
	
	' Send qbXML request.  The response is returned in responseXML.
	' Return true if successful, false otherwise
	Public Function DoRequest(ByRef requestXML As String, ByRef responseXML As String) As Boolean
		
		On Error GoTo ErrHandler
		
		'
		' Determine the maximum SDK version supported and use it
		'
		Dim supportedVersion As String
		supportedVersion = qbXMLLatestVersion(m_RP, m_Ticket)
		requestXML = qbXMLAddProlog(supportedVersion, requestXML)
		
		' Get the responseXML.  Will throw error if ticket is invalid
		responseXML = m_RP.ProcessRequest(m_Ticket, requestXML)
		
		DoRequest = True
		Exit Function
		
ErrHandler: 
		DoRequest = False
		m_ErrNumber = Err.Number
		m_ErrMsg = Err.Description
		Exit Function
		
	End Function
	
	
	' Close connection to QuickBooks
	' Return true if successful, false otherwise
	Public Function Finish() As Boolean
		
		On Error GoTo ErrHandler
		
		If m_Ticket <> "" Then
			' End the session
			m_RP.EndSession(m_Ticket)
			
			' Close the connection
			m_RP.CloseConnection()
		End If
		
		m_Ticket = ""
        m_RP = Nothing

        Finish = True
		Exit Function
		
ErrHandler: 
		Finish = False
		m_ErrNumber = Err.Number
		m_ErrMsg = Err.Description
		Exit Function
		
	End Function
	
	' Return detail error information
	Public Sub GetErrorInfo(ByRef errNumber As Integer, ByRef errMsg As String)
		errNumber = m_ErrNumber
		errMsg = m_ErrMsg
	End Sub
	
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
End Class