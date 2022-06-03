Option Strict Off
Option Explicit On
Module modDepositAdd
    '-----------------------------------------------------------
    ' Form Module: modDepositAdd
    '
    ' Description: this module contains the code which creates qbXML
    '              messages, exchanges them with QuickBooks, interprets
    '              the responses and loads information into form objects.
    '
    ' Copyright © 2021-2022 Intuit Inc. All rights reserved.
    ' Use is subject to the terms specified at:
    '      http://developer.intuit.com/legal/devsite_tos.html
    '
    '----------------------------------------------------------

    Dim booConnected As Boolean
	
	Dim strTicket As String

    'Module objects
    Dim qbXMLCOM As QBXMLRP2Lib.RequestProcessor2

    Public Sub Connect()
		
		On Error GoTo Errs
		qbXMLCOM = New QBXMLRP2Lib.RequestProcessor2
		
		qbXMLCOM.OpenConnection("", "IDN Deposit Add")
		
		strTicket = qbXMLCOM.BeginSession("", QBXMLRP2Lib.QBFileMode.qbFileOpenDoNotCare)
		booConnected = True
		
		'Check to make sure the QuickBooks we're working with supports version 2 of the SDK
		Dim supportedVersion As String
		supportedVersion = qbXMLLatestVersion(qbXMLCOM, strTicket)
		
		Dim booSupports2dot0 As Boolean
		booSupports2dot0 = False
        If supportedVersion >= "15.0" Then booSupports2dot0 = True

        If Not booSupports2dot0 Then
            MsgBox("This program only runs against QuickBooks installations which support the 2.0 qbXML spec.  Your version of QuickBooks does not support qbXML 2.0")
            End
        End If

        Exit Sub

Errs: 
		If Err.Number = -2147220458 Then
			MsgBox("You must have QuickBooks running to run this program.")
			End
		ElseIf Err.Number = -2147220446 Then 
			MsgBox("The company file is open in single user mode and there is already another application accessing it.  You need to have the other application end its session with the company file before this program can be run successfully.")
			End
		Else
			MsgBox(Str(Err.Number) & "     " & Err.Description)
		End If
	End Sub
	
	Public Function CreateReceivePaymentForDepositQuery() As String

        'Create a DOM document object for creating our request.
        Dim xmlDepositFundsQuery As Object
        xmlDepositFundsQuery = CreateObject("MSXML2.DOMDocument")
		
		'Add the QBXML aggregate
		Dim rootElement As MSXML2.IXMLDOMNode
		rootElement = xmlDepositFundsQuery.createElement("QBXML")
		xmlDepositFundsQuery.appendChild(rootElement)
		
		'Add the QBXMLMsgsRq aggregate
		Dim QBXMLMsgsRqNode As MSXML2.IXMLDOMNode
		QBXMLMsgsRqNode = xmlDepositFundsQuery.createElement("QBXMLMsgsRq")
		rootElement.appendChild(QBXMLMsgsRqNode)
		
		'Set the QBXMLMsgsRq onError attribute to continueOnError
		Dim onErrorAttr As MSXML2.IXMLDOMAttribute
		onErrorAttr = xmlDepositFundsQuery.createAttribute("onError")
		onErrorAttr.Text = "stopOnError"
        QBXMLMsgsRqNode.attributes.setNamedItem(onErrorAttr)

        'Add the DepositFundsQueryRq aggregate
        Dim DepositFundsQueryRqNode As MSXML2.IXMLDOMNode
		DepositFundsQueryRqNode = xmlDepositFundsQuery.createElement("ReceivePaymentToDepositQueryRq")
		QBXMLMsgsRqNode.appendChild(DepositFundsQueryRqNode)
		
		'Set the requestID attribute to 0
		Dim requestIDAttr As MSXML2.IXMLDOMAttribute
		requestIDAttr = xmlDepositFundsQuery.createAttribute("requestID")
		requestIDAttr.Text = "0"
        DepositFundsQueryRqNode.attributes.setNamedItem(requestIDAttr)

        'We're adding the prolog using text strings
        CreateReceivePaymentForDepositQuery = "<?xml version=""1.0"" ?>" & "<!DOCTYPE QBXML PUBLIC '-//INTUIT//DTD QBXML QBD 2.0//EN' >" & rootElement.xml
	End Function
	Public Sub GetFundsForDeposit(ByRef lstFundsForDeposit As System.Windows.Forms.ListBox)
		
		Dim strXMLRequest As String
		Dim strXMLResponse As String
		
		On Error GoTo Errs
		
		strXMLRequest = CreateReceivePaymentForDepositQuery()
		
		strXMLResponse = qbXMLCOM.ProcessRequest(strTicket, strXMLRequest)
		
		FillFundsList(lstFundsForDeposit, strXMLResponse)
		
		Exit Sub
Errs: 
		If Err.Number = -2147220458 Then
			MsgBox("You must have QuickBooks running to run this program.")
			End
		Else
			MsgBox(Str(Err.Number) & "     " & Err.Description)
		End If
	End Sub
	
	Public Sub EndSessionCloseConnection()
		If booConnected Then
			qbXMLCOM.EndSession((strTicket))
			qbXMLCOM.CloseConnection()
		End If
	End Sub
	
	Public Sub FillFundsList(ByRef lstListBox As System.Windows.Forms.ListBox, ByRef strXMLResponse As String)

        'Set up a DOM document object to load the response into
        Dim xmlFundsResponse As Object
        xmlFundsResponse = CreateObject("MSXML2.DOMDocument")
		
		'Clear the list box
		lstListBox.Items.Clear()
		
		'Parse the response XML
		xmlFundsResponse.async = False
		xmlFundsResponse.loadXML(strXMLResponse)
		
		'Get the status for our query request
		Dim nodeFundsQueryRs As MSXML2.IXMLDOMNodeList
		nodeFundsQueryRs = xmlFundsResponse.getElementsByTagName("ReceivePaymentToDepositQueryRs")
		
		Dim rsStatusAttr As MSXML2.IXMLDOMNamedNodeMap
		rsStatusAttr = nodeFundsQueryRs.Item(0).Attributes
		Dim strQueryStatus As String
        strQueryStatus = rsStatusAttr.getNamedItem("statusCode").nodeValue

        'If the status is bad, report it to the user
        If strQueryStatus <> "0" Then
			MsgBox(rsStatusAttr.getNamedItem("statusMessage").nodeValue)
			Exit Sub
		End If
		
		'Parse the query response and add the Fundss to the Funds list box
		Dim FundsQueryRsNode As MSXML2.IXMLDOMNode
		FundsQueryRsNode = nodeFundsQueryRs.Item(0)
		
		Dim FundsNodeList As MSXML2.IXMLDOMNodeList
		FundsNodeList = xmlFundsResponse.getElementsByTagName("ReceivePaymentToDepositRet")
		
		Dim numItems As Integer
		numItems = FundsNodeList.length
		
		'Declare the XML objects outside of the loop
		Dim ReturnElement As MSXML2.IXMLDOMElement
		Dim itemNode As MSXML2.IXMLDOMNode
		Dim i As Short
		Dim strDisplayLine As String
		For i = 0 To numItems - 1
			itemNode = FundsNodeList.Item(i)
			strDisplayLine = ""
			ReturnElement = itemNode.selectSingleNode("TxnDate")
			strDisplayLine = ReturnElement.Text
			ReturnElement = itemNode.selectSingleNode("Amount")
			strDisplayLine = strDisplayLine & "    " & ReturnElement.Text
			ReturnElement = itemNode.selectSingleNode("TxnType")
			strDisplayLine = strDisplayLine & "    " & ReturnElement.Text
			ReturnElement = itemNode.selectSingleNode("TxnID")
			strDisplayLine = strDisplayLine & "    " & ReturnElement.Text
			lstListBox.Items.Add(strDisplayLine)
		Next 
	End Sub
	
	Public Sub DepositFunds(ByRef strFundsInfo As String)
		
		Dim strTxnID As String
		Dim strXMLRequest As String
		Dim strXMLResponse As String
		
		strTxnID = Right(strFundsInfo, Len(strFundsInfo) - InStrRev(strFundsInfo, " "))

        'Create a DOM document object for creating our request.
        Dim xmlDepositAdd As Object
        xmlDepositAdd = CreateObject("MSXML2.DOMDocument")
		
		'Add the QBXML aggregate
		Dim rootElement As MSXML2.IXMLDOMNode
		rootElement = xmlDepositAdd.createElement("QBXML")
		xmlDepositAdd.appendChild(rootElement)
		
		'Add the QBXMLMsgsRq aggregate
		Dim QBXMLMsgsRqNode As MSXML2.IXMLDOMNode
		QBXMLMsgsRqNode = xmlDepositAdd.createElement("QBXMLMsgsRq")
		rootElement.appendChild(QBXMLMsgsRqNode)
		
		'Set the QBXMLMsgsRq onError attribute to continueOnError
		Dim onErrorAttr As MSXML2.IXMLDOMAttribute
		onErrorAttr = xmlDepositAdd.createAttribute("onError")
		onErrorAttr.Text = "stopOnError"
        QBXMLMsgsRqNode.attributes.setNamedItem(onErrorAttr)

        'Add the DepositAddRq aggregate
        Dim DepositAddRqNode As MSXML2.IXMLDOMNode
		DepositAddRqNode = xmlDepositAdd.createElement("DepositAddRq")
		QBXMLMsgsRqNode.appendChild(DepositAddRqNode)
		
		'Add the DepositAdd aggregate
		Dim DepositAddNode As MSXML2.IXMLDOMNode
		DepositAddNode = xmlDepositAdd.createElement("DepositAdd")
		DepositAddRqNode.appendChild(DepositAddNode)
		
		'Add the  aggregate
		Dim DepositToAccountRefNode As MSXML2.IXMLDOMNode
		DepositToAccountRefNode = xmlDepositAdd.createElement("DepositToAccountRef")
		DepositAddNode.appendChild(DepositToAccountRefNode)
		
		'Add the FullName element
		Dim FullNameElement As MSXML2.IXMLDOMElement
		FullNameElement = xmlDepositAdd.createElement("FullName")
		FullNameElement.Text = "Checking"
        DepositToAccountRefNode.appendChild(FullNameElement)

        'Add the DepositLineAdd aggregate
        Dim DepositLineAddNode As MSXML2.IXMLDOMNode
		DepositLineAddNode = xmlDepositAdd.createElement("DepositLineAdd")
		DepositAddNode.appendChild(DepositLineAddNode)
		
		'Add the PaymentTxnID element
		Dim PaymentTxnIDElement As MSXML2.IXMLDOMElement
		PaymentTxnIDElement = xmlDepositAdd.createElement("PaymentTxnID")
		PaymentTxnIDElement.Text = strTxnID
        DepositLineAddNode.appendChild(PaymentTxnIDElement)

        'Set the requestID attribute to 0
        Dim requestIDAttr As MSXML2.IXMLDOMAttribute
		requestIDAttr = xmlDepositAdd.createAttribute("requestID")
		requestIDAttr.Text = "0"
        DepositAddRqNode.attributes.setNamedItem(requestIDAttr)

        'We're adding the prolog using text strings
        strXMLRequest = "<?xml version=""1.0"" ?>" & "<!DOCTYPE QBXML PUBLIC '-//INTUIT//DTD QBXML QBD 2.0//EN' >" & rootElement.xml
		
		strXMLResponse = qbXMLCOM.ProcessRequest(strTicket, strXMLRequest)

        'Set up a DOM document object to load the response into
        Dim xmlDepositAddRqResponse As Object
        xmlDepositAddRqResponse = CreateObject("MSXML2.DOMDocument")
		
		'Parse the response XML
		xmlDepositAddRqResponse.async = False
		xmlDepositAddRqResponse.loadXML(strXMLResponse)
		
		'Get the status for our query request
		Dim nodeDepositAddRqQueryRs As MSXML2.IXMLDOMNodeList
		nodeDepositAddRqQueryRs = xmlDepositAddRqResponse.getElementsByTagName("DepositAddRs")
		
		Dim rsStatusAttr As MSXML2.IXMLDOMNamedNodeMap
		rsStatusAttr = nodeDepositAddRqQueryRs.Item(0).Attributes
		Dim strQueryStatus As String
        strQueryStatus = rsStatusAttr.getNamedItem("statusCode").nodeValue

        'If the status is bad, report it to the user
        If strQueryStatus <> "0" Then
			MsgBox(rsStatusAttr.getNamedItem("statusMessage").nodeValue)
			Exit Sub
		Else
			MsgBox("The funds were successfully deposited in Checking")
		End If
		
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
End Module