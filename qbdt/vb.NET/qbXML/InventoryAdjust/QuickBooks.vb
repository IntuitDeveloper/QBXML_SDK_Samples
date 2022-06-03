Option Strict Off
Option Explicit On
Module QuickBooks
	Dim SessionTicket As String
    Dim qbXMLRP As QBXMLRP2Lib.RequestProcessor2

    Function qbConnect() As Boolean
		On Error GoTo errHandler
		qbXMLRP = New QBXMLRP2Lib.RequestProcessor2
        qbXMLRP.OpenConnection("", My.Application.Info.Title)
        SessionTicket = qbXMLRP.BeginSession("", QBXMLRP2Lib.QBFileMode.qbFileOpenSingleUser)
		qbConnect = True

        'Check that we have at least qbXML 2.0 support
        Dim supportedVersion As Decimal
        supportedVersion = qbXMLLatestVersion(qbXMLRP, SessionTicket)
        If (supportedVersion < 2.0) Then
            MsgBox("This sample requires support for qbXML 2.0 or later (QuickBooks 2003 or later) " & "Expect a parsing error when attempting to send requests to QuickBooks", MsgBoxStyle.Exclamation)
        End If
        Exit Function
errHandler: 
		Dim errInfo As ErrObject
		MsgBox("Got error trying to connect with QuickBooks: " & Err.Description)
		qbConnect = False
	End Function
	
	Function ClassesEnabled() As Boolean
        Dim query As MSXML2.DOMDocument60
        query = BuildSimpleQuery("PreferencesQueryRq")
        Dim resp As MSXML2.DOMDocument60
        resp = processRq((query.xml))
		Dim nodes As MSXML2.IXMLDOMNodeList
		nodes = resp.getElementsByTagName("IsUsingClassTracking")
        Dim val_Renamed As String
        val_Renamed = nodes.item(0).text
		If "true" = val_Renamed Then
			ClassesEnabled = True
		Else
			ClassesEnabled = False
		End If
	End Function
	Sub fillAccountList(ByRef AccountList As System.Windows.Forms.ComboBox)
		If ("" = SessionTicket) Then
			Exit Sub
		End If
        Dim query As MSXML2.DOMDocument60
        query = BuildSimpleQuery("AccountQueryRq")
		fillComboBox(AccountList, query, "AccountRet", "FullName")
	End Sub
	Sub fillItemList(ByRef ItemList As System.Windows.Forms.ComboBox)
		If ("" = SessionTicket) Then
			Exit Sub
		End If
        Dim query As MSXML2.DOMDocument60
        query = BuildSimpleQuery("ItemInventoryQueryRq")
		fillComboBox(ItemList, query, "ItemInventoryRet", "FullName")
	End Sub
	Sub fillCustomerList(ByRef CustList As System.Windows.Forms.ComboBox)
		If "" = SessionTicket Then
			Exit Sub
		End If
        Dim query As MSXML2.DOMDocument60
        query = BuildSimpleQuery("CustomerQueryRq")
		fillComboBox(CustList, query, "CustomerRet", "FullName")
	End Sub
	
	Sub fillClassList(ByRef ClassList As System.Windows.Forms.ComboBox)
		If "" = SessionTicket Then
			Exit Sub
		End If
        Dim query As MSXML2.DOMDocument60
        query = BuildSimpleQuery("ClassQueryRq")
		fillComboBox(ClassList, query, "ClassRet", "FullName")
	End Sub
    Sub fillComboBox(ByRef ItemList As System.Windows.Forms.ComboBox, ByRef query As MSXML2.DOMDocument60, ByRef retElem As String, ByRef displayElem As String)
        Dim resp As MSXML2.DOMDocument60
        resp = processRq((query.xml))
        Dim items As MSXML2.IXMLDOMNodeList
        Dim item As MSXML2.IXMLDOMElement
        items = resp.getElementsByTagName(retElem)
        item = items.nextNode
        Dim name As MSXML2.IXMLDOMElement
        Do While (Not item Is Nothing)
            name = item.getElementsByTagName(displayElem).item(0)
            ItemList.Items.Add((name.text))
            item = items.nextNode
        Loop
    End Sub
    Function AdjustInventory(ByRef acct As String, ByRef cust As String, ByRef class_Renamed As String, ByRef Memo As String, ByRef LineItems As System.Windows.Forms.ListView) As String
        Dim i As Object
        Dim doc As New MSXML2.DOMDocument60
        Dim pi As MSXML2.IXMLDOMProcessingInstruction
        pi = doc.createProcessingInstruction("xml", "version=""1.0""")
        doc.appendChild(pi)
        pi = doc.createProcessingInstruction("qbxml", "version=""2.0""")
        doc.appendChild(pi)
        Dim qbxml As MSXML2.IXMLDOMElement
        qbxml = doc.createElement("QBXML")
        doc.appendChild(qbxml)
        Dim msgsrq As MSXML2.IXMLDOMElement
        msgsrq = doc.createElement("QBXMLMsgsRq")
        qbxml.appendChild(msgsrq)
        msgsrq.setAttribute("onError", "continueOnError")
        Dim elem As MSXML2.IXMLDOMElement
        Dim inventoryAdjRq As MSXML2.IXMLDOMElement
        inventoryAdjRq = doc.createElement("InventoryAdjustmentAddRq")
        msgsrq.appendChild(inventoryAdjRq)
        Dim inventoryAdj As MSXML2.IXMLDOMElement
        inventoryAdj = doc.createElement("InventoryAdjustmentAdd")
        inventoryAdjRq.appendChild(inventoryAdj)
        Dim ref As MSXML2.IXMLDOMElement
        ref = doc.createElement("AccountRef")
        inventoryAdj.appendChild(ref)
        ref.appendChild(createSimpleElement(doc, "FullName", acct))
        If Not cust = "" Then
            ref = doc.createElement("CustomerRef")
            inventoryAdj.appendChild(ref)
            ref.appendChild(createSimpleElement(doc, "FullName", cust))
        End If
        If Not class_Renamed = "" Then
            ref = doc.createElement("ClassRef")
            inventoryAdj.appendChild(ref)
            ref.appendChild(createSimpleElement(doc, "FullName", class_Renamed))
        End If
        inventoryAdj.appendChild(createSimpleElement(doc, "Memo", Memo))
        Dim item As String
        Dim what As String
        Dim diff As String
        Dim quant As String
        Dim value As String
        Dim adjline As MSXML2.IXMLDOMElement
        Dim quantAdj As MSXML2.IXMLDOMElement
        Dim itemref As MSXML2.IXMLDOMElement
        Dim j As Int32
        For j = 1 To LineItems.Items.Count
            i = j - 1
            adjline = doc.createElement("InventoryAdjustmentLineAdd")
            inventoryAdj.appendChild(adjline)
            itemref = doc.createElement("ItemRef")
            adjline.appendChild(itemref)
            item = LineItems.Items.Item(i).Text
            itemref.appendChild(createSimpleElement(doc, "FullName", item))
            what = LineItems.Items.Item(i).SubItems(1).Text
            If "Quantity" = what Then
                quantAdj = doc.createElement("QuantityAdjustment")
                adjline.appendChild(quantAdj)
                diff = LineItems.Items.Item(i).SubItems(2).Text
                quant = LineItems.Items.Item(i).SubItems(3).Text
                If "Relative" = diff Then
                    quantAdj.appendChild(createSimpleElement(doc, "QuantityDifference", quant))
                Else
                    quantAdj.appendChild(createSimpleElement(doc, "NewQuantity", quant))
                End If
            Else
                value = LineItems.Items.Item(i).SubItems(2).Text
                quant = LineItems.Items.Item(i).SubItems(3).Text
                quantAdj = doc.createElement("ValueAdjustment")
                adjline.appendChild(quantAdj)
                If Not "0" = quant Then
                    quantAdj.appendChild(createSimpleElement(doc, "NewQuantity", quant))
                End If
                quantAdj.appendChild(createSimpleElement(doc, "NewValue", value))
            End If
        Next j
        Dim resp As MSXML2.DOMDocument60
        resp = processRq((doc.xml))
        Dim rsList As MSXML2.IXMLDOMNodeList
        rsList = resp.getElementsByTagName("InventoryAdjustmentAddRs")
        Dim status As String
        Dim res As MSXML2.IXMLDOMElement
        res = rsList.item(0)
        status = res.getAttribute("statusMessage")
        AdjustInventory = status & vbCrLf & resp.xml
    End Function

    Function createSimpleElement(ByRef doc As MSXML2.DOMDocument60, ByRef name As String, ByRef value As String) As MSXML2.IXMLDOMElement
        Dim elem As MSXML2.IXMLDOMElement
        elem = doc.createElement(name)
        Dim text As MSXML2.IXMLDOMText
        text = doc.createTextNode(value)
        elem.appendChild(text)
        createSimpleElement = elem
    End Function
    Function processRq(ByRef req As String) As MSXML2.DOMDocument60
        Dim doc As New MSXML2.DOMDocument60
        If "" = SessionTicket Then
            Exit Function
        End If
        Dim resp As String
        resp = qbXMLRP.ProcessRequest(SessionTicket, req)
        doc.loadXML(resp)
        processRq = doc
    End Function

    Function BuildSimpleQuery(ByRef query As String) As MSXML2.DOMDocument60
        Dim doc As New MSXML2.DOMDocument60
        Dim pi As MSXML2.IXMLDOMProcessingInstruction
        pi = doc.createProcessingInstruction("xml", "version=""1.0""")
        doc.appendChild(pi)
        pi = doc.createProcessingInstruction("qbxml", "version=""2.0""")
        doc.appendChild(pi)
        Dim qbxml As MSXML2.IXMLDOMElement
        qbxml = doc.createElement("QBXML")
        doc.appendChild(qbxml)
        Dim msgsrq As MSXML2.IXMLDOMElement
        msgsrq = doc.createElement("QBXMLMsgsRq")
        qbxml.appendChild(msgsrq)
        msgsrq.setAttribute("onError", "continueOnError")
        Dim queryElem As MSXML2.IXMLDOMElement
        queryElem = doc.createElement(query)
        msgsrq.appendChild(queryElem)
        BuildSimpleQuery = doc
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
		onErrorAttr.text = "stopOnError"
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
			VersNode = supportedVersions.item(i)
			vers = CDbl(VersNode.firstChild.text)
			If (vers > LastVers) Then
				LastVers = vers
				qbXMLLatestVersion = VersNode.firstChild.text
			End If
		Next i
	End Function
End Module