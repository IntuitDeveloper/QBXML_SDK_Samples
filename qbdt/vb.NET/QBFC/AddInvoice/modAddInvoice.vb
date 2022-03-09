Option Strict Off
Option Explicit On
Module modAddInvoice
	'----------------------------------------------------------
	' Module: modAddInvoice
	'
	' Description:  This module implements two routines which add an
	'               invoice to QuickBooks.  One uses the DOM parser and
	'               raw qbXML to build the add request and parse the
	'               response.  The second uses QBFC objects to do the
	'               building and parsing.  The user determines which
	'               routine is called by selecting a radio button in
	'               the main form.
	'
	' Created On: 09/10/2002
	'
	' Copyright © 2021-2022 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	'----------------------------------------------------------
	
	Public Sub qbXML_AddInvoice()
		
		'Declare various utility variables
		Dim booSessionBegun As Boolean
		'We want to know if we've begun a session so we can end it if an
		'error sends us to the exception handler
		booSessionBegun = False
		
		Dim strTicket As String
		Dim strXMLRequest As String
		Dim strXMLResponse As String
		Dim qbXMLVersionSpec As String
		
		'Set up an error handler to catch exceptions
		On Error GoTo Errs
		
		'Create the RequestProcessor object using QBXMLRPLib
		
		Dim qbXMLCOM As QBXMLRP2Lib.RequestProcessor2
		
		'Set up the RequestProcessor object and connect to QuickBooks
		qbXMLCOM = New QBXMLRP2Lib.RequestProcessor2
		
		'Our App will be named Add Invoice Sample
		qbXMLCOM.OpenConnection("", "IDN Add Invoice Sample")
		
		'Use the currently open company file
		strTicket = qbXMLCOM.BeginSession("", QBXMLRP2Lib.QBFileMode.qbFileOpenDoNotCare)
		booSessionBegun = True
		
		'Check which version of qbXML to use
		Dim supportedVersion As String
		supportedVersion = qbXMLLatestVersion(qbXMLCOM, strTicket)
		
		Dim addr4supported As Boolean
		addr4supported = False
		
		Dim requestMsgSet As QBFC15Lib.IMsgSetRequest
		If (Val(supportedVersion) >= 2) Then
			qbXMLVersionSpec = "<?qbxml version=""" & supportedVersion & """?>"
			addr4supported = True
		ElseIf (supportedVersion = CStr(1.1)) Then 
			qbXMLVersionSpec = "<!DOCTYPE QBXML PUBLIC '-//INTUIT//DTD QBXML QBD " & supportedVersion & "//EN' 'http://developer.intuit.com'>"
		Else
			MsgBox("You are apparently running QuickBooks 2002 Release 1, we strongly recommend that you use QuickBooks' online update feature to obtain the latest fixes and enhancements", MsgBoxStyle.Exclamation)
			qbXMLVersionSpec = "<!DOCTYPE QBXML PUBLIC '-//INTUIT//DTD QBXML QBD " & supportedVersion & "//EN' 'http://developer.intuit.com'>"
		End If

        'Create a DOM document object for creating our request.
        Dim xmlInvoiceAdd As New MSXML2.DOMDocument60
        xmlInvoiceAdd = CreateObject("MSXML2.DOMDocument")
		
		'Create the QBXML aggregate
		Dim rootElement As MSXML2.IXMLDOMNode
		rootElement = xmlInvoiceAdd.createElement("QBXML")
		xmlInvoiceAdd.appendChild(rootElement)
		
		'Add the QBXMLMsgsRq aggregate to the QBXML aggregate
		Dim QBXMLMsgsRqNode As MSXML2.IXMLDOMNode
		QBXMLMsgsRqNode = xmlInvoiceAdd.createElement("QBXMLMsgsRq")
		rootElement.appendChild(QBXMLMsgsRqNode)
		
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'If we were writing a real application this is where we would add
		'a newMessageSetID so we could perform error recovery.  Any time a
		'request contains an add, delete, modify or void request developers
		'should use the error recovery mechanisms.
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		
		'Set the QBXMLMsgsRq onError attribute to continueOnError
		Dim onErrorAttr As MSXML2.IXMLDOMAttribute
		onErrorAttr = xmlInvoiceAdd.createAttribute("onError")
		onErrorAttr.Text = "stopOnError"
		
		QBXMLMsgsRqNode.Attributes.setNamedItem(onErrorAttr)
		
		'Add the InvoiceAddRq aggregate to QBXMLMsgsRq aggregate
		Dim InvoiceAddRqNode As MSXML2.IXMLDOMNode
		InvoiceAddRqNode = xmlInvoiceAdd.createElement("InvoiceAddRq")
		QBXMLMsgsRqNode.appendChild(InvoiceAddRqNode)
		
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'Note - Starting requestID from 0 and incrementing it for each
		'request allows a developer to load the response into a QBFC and
		'parse it with QBFC
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		
		'Set the requestID attribute to 0
		Dim requestIDAttr As MSXML2.IXMLDOMAttribute
		requestIDAttr = xmlInvoiceAdd.createAttribute("requestID")
		requestIDAttr.Text = "0"
		
		InvoiceAddRqNode.Attributes.setNamedItem(requestIDAttr)
		
		'Add the InvoiceAdd aggregate to InvoiceAddRq aggregate
		Dim InvoiceAddNode As MSXML2.IXMLDOMNode
		InvoiceAddNode = xmlInvoiceAdd.createElement("InvoiceAdd")
		InvoiceAddRqNode.appendChild(InvoiceAddNode)
		
		'Add the CustomerRef aggregate to InvoiceAdd aggregate
		Dim CustomerRefNode As MSXML2.IXMLDOMNode
		CustomerRefNode = xmlInvoiceAdd.createElement("CustomerRef")
		InvoiceAddNode.appendChild(CustomerRefNode)
		
		'Add the FullName element to CustomerRef aggregate
		Dim FullNameElement As MSXML2.IXMLDOMElement
		FullNameElement = xmlInvoiceAdd.createElement("FullName")
		FullNameElement.Text = "Abercrombie, Kristy"
		
		CustomerRefNode.appendChild(FullNameElement)
		
		'Add the RefNumber element to InvoiceAdd aggregate
		Dim RefNumberElement As MSXML2.IXMLDOMElement
		RefNumberElement = xmlInvoiceAdd.createElement("RefNumber")
		RefNumberElement.Text = "qbXMLAdd"
		
		InvoiceAddNode.appendChild(RefNumberElement)
		
		'Add the ShipAddress aggregate to InvoiceAdd aggregate
		Dim ShipAddressNode As MSXML2.IXMLDOMNode
		ShipAddressNode = xmlInvoiceAdd.createElement("ShipAddress")
		InvoiceAddNode.appendChild(ShipAddressNode)
		
		'Add the Addr1 element to ShipAddress aggregate
		Dim Addr1Element As MSXML2.IXMLDOMElement
		Addr1Element = xmlInvoiceAdd.createElement("Addr1")
		Addr1Element.Text = "Kristy Abercrombie"
		
		ShipAddressNode.appendChild(Addr1Element)
		
		'Add the Addr2 element to ShipAddress aggregate
		Dim Addr2Element As MSXML2.IXMLDOMElement
		Addr2Element = xmlInvoiceAdd.createElement("Addr2")
		Addr2Element.Text = "c/o Enviromental Designs"
		
		ShipAddressNode.appendChild(Addr2Element)
		
		'Add the Addr3 element to ShipAddress aggregate
		Dim Addr3Element As MSXML2.IXMLDOMElement
		Addr3Element = xmlInvoiceAdd.createElement("Addr3")
		Addr3Element.Text = "1521 Main Street"
		
		ShipAddressNode.appendChild(Addr3Element)
		
		'Add the Addr4 element to ShipAddress aggregate if supported
		Dim Addr4Element As MSXML2.IXMLDOMElement
		If (addr4supported) Then
			Addr4Element = xmlInvoiceAdd.createElement("Addr4")
			Addr4Element.Text = "Suite 204"
			
			ShipAddressNode.appendChild(Addr4Element)
		End If
		
		'Add the City element to ShipAddress aggregate
		Dim CityElement As MSXML2.IXMLDOMElement
		CityElement = xmlInvoiceAdd.createElement("City")
		CityElement.Text = "Culver City"
		
		ShipAddressNode.appendChild(CityElement)
		
		'Add the State element to ShipAddress aggregate
		Dim StateElement As MSXML2.IXMLDOMElement
		StateElement = xmlInvoiceAdd.createElement("State")
		StateElement.Text = "CA"
		
		ShipAddressNode.appendChild(StateElement)
		
		'Add the PostalCode element to ShipAddress aggregate
		Dim PostalCodeElement As MSXML2.IXMLDOMElement
		PostalCodeElement = xmlInvoiceAdd.createElement("PostalCode")
		PostalCodeElement.Text = "90139"
		
		ShipAddressNode.appendChild(PostalCodeElement)
		
		'We're going to add three line items, so use the same code and just
		'change the values on each iteration
		'Declare all of the DOM objects we'll be using here
		Dim InvoiceLineNode As MSXML2.IXMLDOMNode
		Dim ItemRefNode As MSXML2.IXMLDOMNode
		Dim DescElement As MSXML2.IXMLDOMElement
		Dim QuantityElement As MSXML2.IXMLDOMElement
		Dim RateElement As MSXML2.IXMLDOMElement
		
		'Declare the other variables we'll be using in the loop
		Dim strDesc As String
		Dim strItemRefName As String
		Dim strItemFullName As String
		Dim strQuantity As String
		Dim strRate As String
		
		Dim i As Short
		For i = 1 To 3
			'We'll add standard items for the first two lines and a group item
			'for the third
			If i <> 3 Then
				'Add the InvoiceLine aggregate to InvoiceAdd aggregate
				InvoiceLineNode = xmlInvoiceAdd.createElement("InvoiceLineAdd")
				InvoiceAddNode.appendChild(InvoiceLineNode)
				strItemRefName = "ItemRef"
				
				If i = 1 Then
					strItemFullName = "Installation"
					strDesc = ""
					strQuantity = "2"
					strRate = ""
				Else
					strItemFullName = "Lumber:Trim"
					strDesc = "Door trim priced by the foot"
					strQuantity = "26"
					strRate = "1.37"
				End If
			Else
				'Add the InvoiceLineGroupAdd aggregate to InvoiceAdd aggregate
				InvoiceLineNode = xmlInvoiceAdd.createElement("InvoiceLineGroupAdd")
				InvoiceAddNode.appendChild(InvoiceLineNode)
				strItemRefName = "ItemGroupRef"
				
				strItemFullName = "Door Set"
				strQuantity = ""
				strRate = ""
				strDesc = ""
			End If
			
			'Add the ItemRef aggregate to InvoiceAdd aggregate
			ItemRefNode = xmlInvoiceAdd.createElement(strItemRefName)
			InvoiceLineNode.appendChild(ItemRefNode)
			
			'Add the FullName element to ItemRef aggregate
			FullNameElement = xmlInvoiceAdd.createElement("FullName")
			FullNameElement.Text = strItemFullName
			
			ItemRefNode.appendChild(FullNameElement)
			
			'We don't want to cause QuickBooks to do more work than it needs
			'to so only include elements where we want to set the value
			If strDesc <> "" Then
				DescElement = xmlInvoiceAdd.createElement("Desc")
				DescElement.Text = strDesc
				
				InvoiceLineNode.appendChild(DescElement)
			End If
			
			If strQuantity <> "" Then
				QuantityElement = xmlInvoiceAdd.createElement("Quantity")
				QuantityElement.Text = strQuantity
				
				InvoiceLineNode.appendChild(QuantityElement)
			End If
			
			If strRate <> "" Then
				RateElement = xmlInvoiceAdd.createElement("Rate")
				RateElement.Text = strRate
				
				InvoiceLineNode.appendChild(RateElement)
			End If
			
		Next 
		
		'We're adding the prolog using text strings
		strXMLRequest = "<?xml version=""1.0"" ?>" & qbXMLVersionSpec & rootElement.xml
		
		strXMLResponse = qbXMLCOM.ProcessRequest(strTicket, strXMLRequest)
		
		
		' Uncomment the following to see the request and response XML for debugging
		' MsgBox strXMLRequest, vbOKOnly, "RequestXML"
		' MsgBox strXMLResponse, vbOKOnly, "ResponseXML"
		
		qbXMLCOM.EndSession(strTicket)
		booSessionBegun = False
		qbXMLCOM.CloseConnection()

        'Set up a DOM document object to load the response into
        Dim xmlInvoiceAddResponse As New MSXML2.DOMDocument60
        xmlInvoiceAddResponse = CreateObject("MSXML2.DOMDocument")
		
		'Parse the response XML
		xmlInvoiceAddResponse.async = False
		xmlInvoiceAddResponse.loadXML(strXMLResponse)
		
		'Get the status for our add request
		Dim nodeInvoiceAddRs As MSXML2.IXMLDOMNodeList
		nodeInvoiceAddRs = xmlInvoiceAddResponse.getElementsByTagName("InvoiceAddRs")
		
		Dim rsStatusAttr As MSXML2.IXMLDOMNamedNodeMap
		rsStatusAttr = nodeInvoiceAddRs.Item(0).Attributes
		Dim strAddStatus As String
		
		strAddStatus = rsStatusAttr.getNamedItem("statusCode").nodeValue
		
		'If the status is bad, report it to the user
		If strAddStatus <> "0" Then
			MsgBox("qbXML_AddInvoice unexpexcted Error - " & rsStatusAttr.getNamedItem("statusMessage").nodeValue)
			Exit Sub
		End If
		
		'Parse the Add response and add the InvoiceAdds to the InvoiceAdd text box
		Dim InvoiceAddRsNode As MSXML2.IXMLDOMNode
		InvoiceAddRsNode = nodeInvoiceAddRs.Item(0)
		
		'We only expect one InvoiceRet aggregate in an InvoiceAddRs
		Dim InvoiceRetNode As MSXML2.IXMLDOMNode
		InvoiceRetNode = InvoiceAddRsNode.selectSingleNode("InvoiceRet")
		
		'Get the data for the InvoiceDisplay form and set it
		frmInvoiceDisplay.txtInvoiceNumber.Text = InvoiceRetNode.selectSingleNode("RefNumber").Text
		frmInvoiceDisplay.txtInvoiceDate.Text = InvoiceRetNode.selectSingleNode("TxnDate").Text
		frmInvoiceDisplay.txtDueDate.Text = InvoiceRetNode.selectSingleNode("DueDate").Text
		frmInvoiceDisplay.txtSubtotal.Text = InvoiceRetNode.selectSingleNode("Subtotal").Text
		frmInvoiceDisplay.txtSalesTax.Text = InvoiceRetNode.selectSingleNode("SalesTaxTotal").Text
		frmInvoiceDisplay.txtInvoiceTotal.Text = Str(CDbl(frmInvoiceDisplay.txtSubtotal.Text) + CDbl(frmInvoiceDisplay.txtSalesTax.Text))
		
		'Get the invoice lines and display information about them
		
		'We know there are three invoice lines, but lets pretend we didn't
		'know that.  We'll start at the last child node and move back
		'until we first get an InvoiceLineRet or InvoiceLineGroupRet, then
		'stop when we stop finding them and move forward again until we
		'don't.
		
		'Get the child nodes of the InvoiceAddRs
		Dim InvoiceNodeList As MSXML2.IXMLDOMNodeList
		InvoiceNodeList = InvoiceRetNode.childNodes
		
		Dim l As Integer
		l = InvoiceNodeList.length - 1
		
		Do While InvoiceNodeList.Item(l).nodeName <> "InvoiceLineRet" And InvoiceNodeList.Item(l).nodeName <> "InvoiceLineGroupRet"
			l = l - 1
		Loop 
		
		Do While InvoiceNodeList.Item(l).nodeName = "InvoiceLineRet" Or InvoiceNodeList.Item(l).nodeName = "InvoiceLineGroupRet"
			l = l - 1
		Loop 
		l = l + 1
		
		Do While l < InvoiceNodeList.length
			If InvoiceNodeList.Item(l).nodeName = "InvoiceLineRet" Then
				frmInvoiceDisplay.txtInvoiceLines.Text = frmInvoiceDisplay.txtInvoiceLines.Text & InvoiceNodeList.Item(l).selectSingleNode("ItemRef").selectSingleNode("FullName").Text & vbCrLf
				l = l + 1
			ElseIf InvoiceNodeList.Item(l).nodeName = "InvoiceLineGroupRet" Then 
				frmInvoiceDisplay.txtInvoiceLines.Text = frmInvoiceDisplay.txtInvoiceLines.Text & InvoiceNodeList.Item(l).selectSingleNode("ItemGroupRef").selectSingleNode("FullName").Text & vbCrLf
				l = l + 1
			Else
				l = InvoiceNodeList.length
			End If
		Loop

        
        Dim Frm As New frmInvoiceDisplay
        Frm.ShowDialog()

        Exit Sub
		
Errs: 
		If Err.Number = &H80040416 Then
			MsgBox("You must have QuickBooks running with the company" & vbCrLf & "file open to use this program.")
			qbXMLCOM.CloseConnection()
			End
		ElseIf Err.Number = &H80040422 Then 
			MsgBox("This QuickBooks company file is open in single user mode and" & vbCrLf & "another application is already accessing it.  Please exit the" & vbCrLf & "other application and run this application again.")
			qbXMLCOM.CloseConnection()
			End
		Else
			MsgBox("HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)
			
			If booSessionBegun Then
				qbXMLCOM.EndSession(strTicket)
				qbXMLCOM.CloseConnection()
			End If
			End
		End If
	End Sub
	
	
	Public Sub QBFC_AddInvoice()
		Dim msg As Object
		
		'Declare various utility variables
		Dim booSessionBegun As Boolean
		booSessionBegun = False
		
		On Error GoTo ErrHandler
		
		'We want to know if we've begun a session so we can end it if an
		'error sends us to the exception handler
		booSessionBegun = False
		
		' Create the session manager object using QBFC, and use this
		' object to open a connection and begin a session with QuickBooks.
		Dim SessionManager As New QBFC15Lib.QBSessionManager
		SessionManager.OpenConnection("", "IDN Add Invoice Sample")
		SessionManager.BeginSession("", QBFC15Lib.ENOpenMode.omDontCare)
		booSessionBegun = True
		
		Dim supportedVersion As Double
		supportedVersion = Val(QBFCLatestVersion(SessionManager))
		
		Dim addr4supported As Boolean
		addr4supported = False
		' Create the message set request object
		Dim requestMsgSet As QBFC15Lib.IMsgSetRequest
		If (supportedVersion >= 6#) Then
			requestMsgSet = SessionManager.CreateMsgSetRequest("US", 6, 0)
			addr4supported = True
		ElseIf (supportedVersion >= 5#) Then 
			requestMsgSet = SessionManager.CreateMsgSetRequest("US", 5, 0)
			addr4supported = True
		ElseIf (supportedVersion >= 4#) Then 
			requestMsgSet = SessionManager.CreateMsgSetRequest("US", 4, 0)
			addr4supported = True
		ElseIf (supportedVersion >= 3#) Then 
			requestMsgSet = SessionManager.CreateMsgSetRequest("US", 3, 0)
			addr4supported = True
		ElseIf (supportedVersion >= 2#) Then 
			requestMsgSet = SessionManager.CreateMsgSetRequest("US", 2, 0)
			addr4supported = True
		ElseIf (supportedVersion = 1.1) Then 
			requestMsgSet = SessionManager.CreateMsgSetRequest("US", 1, 1)
		Else
			MsgBox("You are apparently running QuickBooks 2002 Release 1, we strongly recommend that you use QuickBooks' online update feature to obtain the latest fixes and enhancements", MsgBoxStyle.Exclamation)
			requestMsgSet = SessionManager.CreateMsgSetRequest("US", 1, 0)
		End If
		
		' Initialize the message set request's attributes
		requestMsgSet.Attributes.OnError = QBFC15Lib.ENRqOnError.roeStop
		
		' Add the request to the message set request object
		Dim invoiceAdd As QBFC15Lib.IInvoiceAdd
		invoiceAdd = requestMsgSet.AppendInvoiceAddRq
		
		' Set the IinvoiceAdd field values
		invoiceAdd.CustomerRef.FullName.setValue("Abercrombie, Kristy")
		invoiceAdd.RefNumber.setValue("QBFCAdd")
		
		' Set up the shipping address for the invoice
		invoiceAdd.ShipAddress.Addr1.setValue("Kristy Abercrombie")
		invoiceAdd.ShipAddress.Addr2.setValue("c/o Enviromental Designs")
		invoiceAdd.ShipAddress.Addr3.setValue("1521 Main Street")
		If (addr4supported) Then invoiceAdd.ShipAddress.Addr4.setValue("Suite 204")
		invoiceAdd.ShipAddress.City.setValue("Culver City")
		invoiceAdd.ShipAddress.State.setValue("CA")
		invoiceAdd.ShipAddress.PostalCode.setValue("90139")
		
		' Create the first line item for the invoice
		Dim invoiceLineAdd As QBFC15Lib.IInvoiceLineAdd
		invoiceLineAdd = invoiceAdd.ORInvoiceLineAddList.Append.invoiceLineAdd
		
		' Set the values for the invoice line
		invoiceLineAdd.ItemRef.FullName.setValue("Installation")
		invoiceLineAdd.Quantity.setValue(2)
		
		'Create the second line item for the invoice
		invoiceLineAdd = invoiceAdd.ORInvoiceLineAddList.Append.invoiceLineAdd
		
		' Set the values for the invoice line
		invoiceLineAdd.ItemRef.FullName.setValue("Lumber:Trim")
		invoiceLineAdd.Desc.setValue("Door trim priced by the foot")
		invoiceLineAdd.Quantity.setValue(26)
		invoiceLineAdd.ORRatePriceLevel.Rate.setValue(1.37)
		
		'Create the third line item with is a item group line
		Dim invoiceGroupLineAdd As QBFC15Lib.IInvoiceLineGroupAdd
		invoiceGroupLineAdd = invoiceAdd.ORInvoiceLineAddList.Append.InvoiceLineGroupAdd
		
		' Set the values for the invoice line
		invoiceGroupLineAdd.ItemGroupRef.FullName.setValue("Door Set")
		
		' Perform the request and obtain a response from QuickBooks
		Dim responseMsgSet As QBFC15Lib.IMsgSetResponse
		responseMsgSet = SessionManager.DoRequests(requestMsgSet)
		
		' Close the session and connection with QuickBooks.
		SessionManager.EndSession()
		booSessionBegun = False
		SessionManager.CloseConnection()
		
		' Uncomment the following to see the request and response XML for debugging
		' MsgBox requestMsgSet.ToXMLString, vbOKOnly, "RequestXML"
		' MsgBox responseMsgSet.ToXMLString, vbOKOnly, "ResponseXML"
		
		' Interpret the response
		Dim response As QBFC15Lib.IResponse
		
		' The response list contains only one response,
		' which corresponds to our single request
		response = responseMsgSet.ResponseList.GetAt(0)
		
		
		msg = "Status: Code = " & CStr(response.StatusCode) & ", Message = " & response.StatusMessage & ", Severity = " & response.StatusSeverity & vbCrLf
		
		' The Detail property of the IResponse object
		' returns a Ret object for Add and Mod requests.
		' In this case, the Ret object is IInvoiceRet.
		
		'For help finding out the Detail's type, uncomment the following
		'line:
		'MsgBox response.Detail.Type.GetAsString
		
		Dim invoiceRet As QBFC15Lib.IInvoiceRet
		invoiceRet = response.Detail
		
		If (invoiceRet Is Nothing) Then
			MsgBox(msg)
			Exit Sub
		End If
		'Declare variables for subtotal and tax so we can compute the
		'invoice total.  Although we could use BalanceRemaining in this
		'case, lets compute it anyway.
		Dim dblSubtotal As Double
        Dim dblTax As Double
        Dim frm As New frmInvoiceDisplay

        'Fill in the text fields of the Invoice Display form
        frm.txtInvoiceNumber.Text = invoiceRet.RefNumber.GetValue()
        Frm.txtInvoiceDate.Text = CStr(invoiceRet.TxnDate.GetValue())
        Frm.txtDueDate.Text = CStr(invoiceRet.DueDate.GetValue())
        dblSubtotal = invoiceRet.Subtotal.GetValue()
        Frm.txtSubtotal.Text = invoiceRet.Subtotal.GetAsString
        dblTax = invoiceRet.SalesTaxTotal.GetValue()
        Frm.txtSalesTax.Text = invoiceRet.SalesTaxTotal.GetAsString
        Frm.txtInvoiceTotal.Text = Str(dblSubtotal + dblTax)

        'Put the lines in the display
        Dim orInvoiceLineRetList As QBFC15Lib.IORInvoiceLineRetList
		orInvoiceLineRetList = invoiceRet.orInvoiceLineRetList
		
		Dim i As Short
		For i = 0 To orInvoiceLineRetList.Count - 1
			If invoiceRet.orInvoiceLineRetList.GetAt(i).ortype = QBFC15Lib.ENORInvoiceLineRet.orilrInvoiceLineRet Then
				frmInvoiceDisplay.txtInvoiceLines.Text = frmInvoiceDisplay.txtInvoiceLines.Text & invoiceRet.orInvoiceLineRetList.GetAt(i).InvoiceLineRet.ItemRef.FullName.getValue & vbCrLf
			Else
				frmInvoiceDisplay.txtInvoiceLines.Text = frmInvoiceDisplay.txtInvoiceLines.Text & invoiceRet.orInvoiceLineRetList.GetAt(i).InvoiceLineGroupRet.ItemGroupRef.FullName.getValue & vbCrLf
			End If
		Next

        'Show the Invoice Display form
        

        frm.ShowDialog()
        Exit Sub
		
ErrHandler: 
		If Err.Number = &H80040416 Then
			MsgBox("You must have QuickBooks running with the company" & vbCrLf & "file open to use this program.")
			SessionManager.CloseConnection()
			End
		ElseIf Err.Number = &H80040422 Then 
			MsgBox("This QuickBooks company file is open in single user mode and" & vbCrLf & "another application is already accessing it.  Please exit the" & vbCrLf & "other application and run this application again.")
			SessionManager.CloseConnection()
			End
		Else
			MsgBox("HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)
			
			If booSessionBegun Then
				SessionManager.EndSession()
				SessionManager.CloseConnection()
			End If
			
			End
		End If
	End Sub
	
	Function QBFCLatestVersion(ByRef SessionManager As QBFC15Lib.QBSessionManager) As String
		Dim strXMLVersions() As String
		'Should be able to use this, but there appears to be a bug that may cause 2.0 to be returned
		'when it should not.
		'strXMLVersions = SessionManager.QBXMLVersionsForSession
		
		Dim msgset As QBFC15Lib.IMsgSetRequest
		'Use oldest version to ensure that we work with any QuickBooks (US)
		msgset = SessionManager.CreateMsgSetRequest("US", 1, 0)
		msgset.AppendHostQueryRq()
		Dim QueryResponse As QBFC15Lib.IMsgSetResponse
		QueryResponse = SessionManager.DoRequests(msgset)
		Dim response As QBFC15Lib.IResponse
		
		' The response list contains only one response,
		' which corresponds to our single HostQuery request
		response = QueryResponse.ResponseList.GetAt(0)
		Dim HostResponse As QBFC15Lib.IHostRet
		HostResponse = response.Detail
		Dim supportedVersions As QBFC15Lib.IBSTRList
		supportedVersions = HostResponse.SupportedQBXMLVersionList
		
		Dim i As Integer
		Dim vers As Double
		Dim LastVers As Double
		LastVers = 0
		For i = 0 To supportedVersions.Count - 1
			vers = Val(supportedVersions.GetAt(i))
			If (vers > LastVers) Then
				LastVers = vers
				QBFCLatestVersion = supportedVersions.GetAt(i)
			End If
		Next i
	End Function
	
	Function qbXMLLatestVersion(ByRef rp As QBXMLRP2Lib.RequestProcessor2, ByRef ticket As String) As String
		Dim strXMLResponse As Object
		Dim strXMLRequest As Object
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
		
		
		strXMLRequest = "<?xml version=""1.0"" ?>" & "<!DOCTYPE QBXML PUBLIC '-//INTUIT//DTD QBXML QBD 1.0//EN' 'http://developer.intuit.com'>" & rootElement.xml
		
		
		
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