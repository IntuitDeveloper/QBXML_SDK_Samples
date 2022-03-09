Option Strict Off
Option Explicit On
Friend Class AddCust
	Inherits System.Windows.Forms.Form
    '-----------------------------------------------------------
    ' Form Module: AddCust
    '
    ' Description:  This sample application demonstrates how to
    '               construct a qbXML request, send it to QuickBooks,
    '               and parse the qbXML response.  It prompts the
    '               user for customer information and adds a new
    '               customer to QuickBooks.
    '
    '               QuickBooks must be running with a data file open.
    '               The current data file is used.
    '
    ' Created On: 11/05/2001
    ' Updated to SDK 2.0 On: 07/30/2002
    '
    ' Copyright © 2021-2022 Intuit Inc. All rights reserved.
    ' Use is subject to the terms specified at:
    '      http://developer.intuit.com/legal/devsite_tos.html
    '
    '----------------------------------------------------------


    ' AppID and AppName sent to QuickBooks
    Const cAppID As String = "123"
	Const cAppName As String = "IDN VB AddCustomer Sample"
	
	' We only have one request, thus we simply use "1".  For multiple
	' requests, it is recommended you use unique requestID for each
	' request
	Const cRequestID As String = "1"
	
	' Request and response strings
	Dim requestXML As String
	Dim responseXML As String
	
	' Customer info
	Dim firstName As String
	Dim lastName As String
	Dim phoneNumber As String
	Dim custName As String
	
	' Response info
	Dim resCustName As String
	Dim resCustFullName As String
	Dim resListID As String
	
	
	' Submit button
	Private Sub Comm_Submit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Comm_Submit.Click
		
		On Error GoTo ErrHandler
		
		' Initialize
		requestXML = ""
		responseXML = ""
		
		' Get input data
		If Not CollectFormData Then
			Exit Sub
		End If
		
		'
		' Set up to communicate with QuickBooks
		'
		' Create qbXMLRP2 COM object
		'
		Dim qbXMLRP As New QBXMLRP2Lib.RequestProcessor2
		
		Dim ticket As String
		
		'
		' Open connection to QuickBooks
		'
		qbXMLRP.OpenConnection(cAppID, cAppName)
		'
		' Begin Session
		' Pass empty string for the data file name to use the currently
		' open data file.
		'
		ticket = qbXMLRP.BeginSession("", QBXMLRP2Lib.QBFileMode.qbFileOpenDoNotCare)

        '
        ' Build the request XML
        '
        ' XMLBuilder
        Dim builder As New MSXML2.DOMDocument60

        Dim QBXML As MSXML2.IXMLDOMNode
		QBXML = builder.createElement("QBXML")
		builder.appendChild(QBXML)
		Dim msgsRq As MSXML2.IXMLDOMElement
        msgsRq = QBXML.appendChild(builder.createElement("QBXMLMsgsRq"))
        msgsRq.setAttribute("onError", "continueOnError")
		Dim CustomerAddRq As MSXML2.IXMLDOMElement
		Dim CustomerAdd As MSXML2.IXMLDOMElement
        CustomerAddRq = msgsRq.appendChild(builder.createElement("CustomerAddRq"))
        CustomerAdd = CustomerAddRq.appendChild(builder.createElement("CustomerAdd"))
        Dim dataElement As MSXML2.IXMLDOMElement
        dataElement = CustomerAdd.appendChild(builder.createElement("Name"))
        dataElement.appendChild(builder.createTextNode(custName))
        If firstName <> "" Then
            dataElement = CustomerAdd.appendChild(builder.createElement("FirstName"))
            dataElement.appendChild(builder.createTextNode(firstName))
        End If
		If lastName <> "" Then
            dataElement = CustomerAdd.appendChild(builder.createElement("LastName"))
            dataElement.appendChild(builder.createTextNode(lastName))
        End If
		If phoneNumber <> "" Then
            dataElement = CustomerAdd.appendChild(builder.createElement("Phone"))
            dataElement.appendChild(builder.createTextNode(phoneNumber))
        End If
		
		Dim supportedVersion As String
		supportedVersion = qbXMLLatestVersion(qbXMLRP, ticket)
		requestXML = qbXMLAddProlog(supportedVersion, (builder.xml))
		
		'
		' Send request to QuickBooks
		'
		responseXML = qbXMLRP.ProcessRequest(ticket, requestXML)
		
		
		'
		' Done communicating with QuickBooks, so close up the connection
		'
		' End the session
		'
		qbXMLRP.EndSession(ticket)
		'
		' Close the connection
		'
		qbXMLRP.CloseConnection()
		
		'
		' Now we'll parse the response from QuickBooks
		'
		If Not ParseResponseXML Then
			Exit Sub
		End If
		
		'
		' Display the result
		'
		MsgBox("Customer " & custName & " has been sucessfully created" & vbCr & "ListID = " & resListID & vbCr & "FullName = " & resCustFullName, MsgBoxStyle.OKOnly, "Success")
		
		Exit Sub
		
ErrHandler: 
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
		Exit Sub
		
	End Sub
	
	
	' View qbXML Request button
	Private Sub Comm_View_Req_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Comm_View_Req.Click
		
		On Error GoTo ErrHandler
		
		Dim reqFrm As DisplayXML
		If requestXML <> "" Then
			
			reqFrm = New DisplayXML
			
			reqFrm.Text_Content.Text = requestXML
			reqFrm.Text = "Request XML"
			reqFrm.Show()
			
		Else
			MsgBox("Request is empty.  Please add a customer first", MsgBoxStyle.Information, "Request XML")
		End If
		Exit Sub
		
		
ErrHandler: 
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
		Exit Sub
		
	End Sub
	
	' View Response Button
	'
	Private Sub Comm_View_Res_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Comm_View_Res.Click
		
		On Error GoTo ErrHandler
		
		' Instantiate a RequestXML Obj
		'
		
		Dim resFrm As New DisplayXML
		Dim tmpResponseXML As String
		If responseXML <> "" Then
			
			
			' replace Lf to CrLf, this is for display only
			tmpResponseXML = Replace(responseXML, vbLf, vbCrLf)
			resFrm.Text_Content.Text = tmpResponseXML
			resFrm.Text = "Response XML"
			resFrm.Show()
			
		Else
			MsgBox("Response is empty.  Please add a customer first", MsgBoxStyle.Information, "Response XML")
		End If
		
		Exit Sub
		
ErrHandler: 
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
		Exit Sub
		
	End Sub
	
	' Exit button
	'
	Private Sub Comm_Exit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Comm_Exit.Click
		Me.Close() ' close the window
	End Sub
	
	
	
	' Parse response XML from QuickBooks using MSXML parser
	Private Function ParseResponseXML() As Boolean
		
		On Error GoTo ErrHandler
		
		Dim retStatusCode As String
		Dim retStatusMessage As String
		Dim retStatusSeverity As String

        ' Create xmlDoc Obj

        ' DOM Document Object
        Dim xmlDoc As New MSXML2.DOMDocument60

        ' DOM Node list object for looping through
        Dim objNodeList As MSXML2.IXMLDOMNodeList
		
		' Node objects
		Dim objChild As MSXML2.IXMLDOMNode
		Dim custChildNode As MSXML2.IXMLDOMNode
		
		' Attributes Name Mapping
		Dim attrNamedNodeMap As MSXML2.IXMLDOMNamedNodeMap
		
		Dim i As Short
		Dim ret As Boolean
		Dim errorMsg As String
		
		errorMsg = ""
		
		' Load xml doc
		ret = xmlDoc.loadXML(responseXML)
		If Not ret Then
			errorMsg = "loadXML failed, reason: " & xmlDoc.parseError.reason
			GoTo ErrHandler
		End If
		
		' Get CustomerAddRs nodes list
		objNodeList = xmlDoc.getElementsByTagName("CustomerAddRs")
		
		' Loop through each CustomerAddRs node
		' Since we have only one request, we should only have one
		' CustomerAddRs.  The loop is actually unnecessary, but it
		' is a good programming practice
		For i = 0 To (objNodeList.length - 1)
			
			' Get the CustomerRetRs
			attrNamedNodeMap = objNodeList.Item(i).attributes

            ' Get the status Code, info and Severity
            '
            retStatusCode = attrNamedNodeMap.getNamedItem("statusCode").nodeValue
            retStatusSeverity = attrNamedNodeMap.getNamedItem("statusSeverity").nodeValue
            retStatusMessage = attrNamedNodeMap.getNamedItem("statusMessage").nodeValue

            ' Check status code to see if there is error or warning
            If retStatusCode <> "0" Then
				' Checking for Warning is a good practice, although unlikely to happen
				' on an add request.
				If retStatusSeverity = "Warning" Then
					' Show the warning, then continue normal processing
					MsgBox(retStatusMessage, MsgBoxStyle.Exclamation, "Warning from QuickBooks")
				ElseIf retStatusSeverity = "Error" Then 
					MsgBox(retStatusMessage, MsgBoxStyle.Exclamation, "Error from QuickBooks")
					' We only have one response thus we will exit.  If we have multiple
					' responses, then we may want to continue with the loop.
					ParseResponseXML = False
					Exit Function
				End If
			End If
			
			' Walk through the child nodes of CustomerAddRs node, to get
			' the detail info
			' This loop is not necessary actually, since we only have one
			' CustomerRet.
			For	Each objChild In objNodeList.Item(i).childNodes
				
				' Get the CustomerRet block
				If objChild.nodeName = "CustomerRet" Then
					
					' Get the elements in this block
					For	Each custChildNode In objChild.childNodes
						If custChildNode.nodeName = "ListID" Then
							resListID = custChildNode.Text
						ElseIf custChildNode.nodeName = "Name" Then 
							resCustName = custChildNode.Text
						ElseIf custChildNode.nodeName = "FullName" Then 
							resCustFullName = custChildNode.Text
						End If
					Next custChildNode
					
				End If ' End of customerRet
				
			Next objChild ' End of customerAddret
		Next 
		
		ParseResponseXML = True
		Exit Function
		
ErrHandler: 
		If errorMsg <> "" Then
			MsgBox(errorMsg, MsgBoxStyle.Exclamation, "Error")
		Else
			MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
		End If
		ParseResponseXML = False
		Exit Function
		
	End Function
	
	' Form Data Collection
	'
	Private Function CollectFormData() As Boolean
		
		On Error GoTo ErrHandler
		
		' Get data from the form
		firstName = Text_FirstName.Text
		lastName = Text_LastName.Text
		phoneNumber = Text_Phone.Text
		custName = Text_CustomerName.Text
		
		' Customer Name is required
		If custName = "" Then
			MsgBox("Customer Name is empty", MsgBoxStyle.OKOnly, "Error")
			CollectFormData = False
			GoTo ExitProc
		End If
		
		CollectFormData = True
		
ExitProc: 
		Exit Function
		
ErrHandler: 
		CollectFormData = False
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
		Exit Function
		
	End Function
	
	Function qbXMLLatestVersion(ByRef rp As QBXMLRP2Lib.RequestProcessor2, ByRef ticket As String) As String
        Dim strXMLVersions() As String
        qbXMLLatestVersion = ""
        'Should be able to use this, but there appears to be a bug that may cause 2.0 to be returned
        'when it should not.
        strXMLVersions = rp.QBXMLVersionsForSession(ticket)

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