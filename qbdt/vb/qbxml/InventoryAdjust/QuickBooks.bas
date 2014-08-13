Attribute VB_Name = "QuickBooks"
Dim SessionTicket As String
Dim qbXMLRP As QBXMLRP2Lib.RequestProcessor2

Function qbConnect() As Boolean
    On Error GoTo errHandler
    Set qbXMLRP = New QBXMLRP2Lib.RequestProcessor2
    qbXMLRP.OpenConnection "", App.Title
    SessionTicket = qbXMLRP.BeginSession("", qbFileOpenSingleUser)
    qbConnect = True
    
    'Check that we have at least qbXML 2.0 support
            
    Dim supportedVersion As String
    supportedVersion = qbXMLLatestVersion(qbXMLRP, SessionTicket)
    If (supportedVersion < "2.0") Then
        MsgBox "This sample requires support for qbXML 2.0 or later (QuickBooks 2003 or later) " & _
               "Expect a parsing error when attempting to send requests to QuickBooks", _
                vbExclamation
    End If
    Exit Function
errHandler:
    Dim errInfo As ErrObject
    MsgBox ("Got error trying to connect with QuickBooks: " & Err.Description)
    qbConnect = False
End Function

Function ClassesEnabled() As Boolean
    Dim query As DOMDocument40
    Set query = BuildSimpleQuery("PreferencesQueryRq")
    Dim resp As DOMDocument40
    Set resp = processRq(query.xml)
    Dim nodes As IXMLDOMNodeList
    Set nodes = resp.getElementsByTagName("IsUsingClassTracking")
    Dim val As String
    val = nodes.item(0).text
    If "true" = val Then
        ClassesEnabled = True
    Else
        ClassesEnabled = False
    End If
End Function
Sub fillAccountList(AccountList As ComboBox)
    If ("" = SessionTicket) Then
        Exit Sub
    End If
    Dim query As DOMDocument40
    Set query = BuildSimpleQuery("AccountQueryRq")
    fillComboBox AccountList, query, "AccountRet", "FullName"
End Sub
Sub fillItemList(ItemList As ComboBox)
    If ("" = SessionTicket) Then
        Exit Sub
    End If
    Dim query As DOMDocument40
    Set query = BuildSimpleQuery("ItemInventoryQueryRq")
    fillComboBox ItemList, query, "ItemInventoryRet", "FullName"
End Sub
Sub fillCustomerList(CustList As ComboBox)
    If "" = SessionTicket Then
        Exit Sub
    End If
    Dim query As DOMDocument40
    Set query = BuildSimpleQuery("CustomerQueryRq")
    fillComboBox CustList, query, "CustomerRet", "FullName"
End Sub

Sub fillClassList(ClassList As ComboBox)
    If "" = SessionTicket Then
        Exit Sub
    End If
    Dim query As DOMDocument40
    Set query = BuildSimpleQuery("ClassQueryRq")
    fillComboBox ClassList, query, "ClassRet", "FullName"
End Sub
Sub fillComboBox(ItemList As ComboBox, query As DOMDocument40, retElem As String, displayElem As String)
    Dim resp As DOMDocument40
    Set resp = processRq(query.xml)
    Dim items As IXMLDOMNodeList
    Dim item As IXMLDOMElement
    Set items = resp.getElementsByTagName(retElem)
    Set item = items.nextNode
    Do While (Not item Is Nothing)
        Dim name As IXMLDOMElement
        Set name = item.getElementsByTagName(displayElem).item(0)
        ItemList.AddItem (name.text)
        Set item = items.nextNode
    Loop
End Sub
Function AdjustInventory(acct As String, cust As String, class As String, Memo As String, LineItems As ListView) As String
    Dim doc As New DOMDocument40
    Dim pi As IXMLDOMProcessingInstruction
    Set pi = doc.createProcessingInstruction("xml", "version=""1.0""")
    doc.appendChild pi
    Set pi = doc.createProcessingInstruction("qbxml", "version=""2.0""")
    doc.appendChild pi
    Dim qbxml As IXMLDOMElement
    Set qbxml = doc.createElement("QBXML")
    doc.appendChild qbxml
    Dim msgsrq As IXMLDOMElement
    Set msgsrq = doc.createElement("QBXMLMsgsRq")
    qbxml.appendChild msgsrq
    msgsrq.setAttribute "onError", "continueOnError"
    Dim elem As IXMLDOMElement
    Dim inventoryAdjRq As IXMLDOMElement
    Set inventoryAdjRq = doc.createElement("InventoryAdjustmentAddRq")
    msgsrq.appendChild inventoryAdjRq
    Dim inventoryAdj As IXMLDOMElement
    Set inventoryAdj = doc.createElement("InventoryAdjustmentAdd")
    inventoryAdjRq.appendChild inventoryAdj
    Dim ref As IXMLDOMElement
    Set ref = doc.createElement("AccountRef")
    inventoryAdj.appendChild ref
    ref.appendChild createSimpleElement(doc, "FullName", acct)
    If Not cust = "" Then
        Set ref = doc.createElement("CustomerRef")
        inventoryAdj.appendChild ref
        ref.appendChild createSimpleElement(doc, "FullName", cust)
    End If
    If Not class = "" Then
        Set ref = doc.createElement("ClassRef")
        inventoryAdj.appendChild ref
        ref.appendChild createSimpleElement(doc, "FullName", class)
    End If
    inventoryAdj.appendChild createSimpleElement(doc, "Memo", Memo)
    For i = 1 To LineItems.ListItems.Count
        Dim item As String
        Dim what As String
        Dim diff As String
        Dim quant As String
        Dim value As String
        Dim adjline As IXMLDOMElement
        Dim quantAdj As IXMLDOMElement
        Set adjline = doc.createElement("InventoryAdjustmentLineAdd")
        inventoryAdj.appendChild adjline
        Dim itemref As IXMLDOMElement
        Set itemref = doc.createElement("ItemRef")
        adjline.appendChild itemref
        item = LineItems.ListItems.item(i).text
        itemref.appendChild createSimpleElement(doc, "FullName", item)
        what = LineItems.ListItems.item(i).SubItems(1)
        If "Quantity" = what Then
            Set quantAdj = doc.createElement("QuantityAdjustment")
            adjline.appendChild quantAdj
            diff = LineItems.ListItems.item(i).SubItems(2)
            quant = LineItems.ListItems.item(i).SubItems(3)
            If "Relative" = diff Then
                quantAdj.appendChild createSimpleElement(doc, "QuantityDifference", quant)
            Else
                quantAdj.appendChild createSimpleElement(doc, "NewQuantity", quant)
            End If
        Else
            value = LineItems.ListItems.item(i).SubItems(2)
            quant = LineItems.ListItems.item(i).SubItems(3)
            Set quantAdj = doc.createElement("ValueAdjustment")
            adjline.appendChild quantAdj
            If Not "0" = quant Then
                quantAdj.appendChild createSimpleElement(doc, "NewQuantity", quant)
            End If
            quantAdj.appendChild createSimpleElement(doc, "NewValue", value)
        End If
    Next i
    Dim resp As DOMDocument40
    Set resp = processRq(doc.xml)
    Dim rsList As IXMLDOMNodeList
    Set rsList = resp.getElementsByTagName("InventoryAdjustmentAddRs")
    Dim status As String
    Dim res As IXMLDOMElement
    Set res = rsList.item(0)
    status = res.getAttribute("statusMessage")
    AdjustInventory = status & vbCrLf & resp.xml
End Function

Function createSimpleElement(doc As DOMDocument40, name As String, value As String) As IXMLDOMElement
    Dim elem As IXMLDOMElement
    Set elem = doc.createElement(name)
    Dim text As IXMLDOMText
    Set text = doc.createTextNode(value)
    elem.appendChild text
    Set createSimpleElement = elem
End Function
Function processRq(req As String) As DOMDocument40
    Dim doc As New DOMDocument40
    If "" = SessionTicket Then
        Exit Function
    End If
    Dim resp As String
    resp = qbXMLRP.ProcessRequest(SessionTicket, req)
    doc.loadXML (resp)
    Set processRq = doc
End Function

Function BuildSimpleQuery(query As String) As DOMDocument40
    Dim doc As New DOMDocument40
    Dim pi As IXMLDOMProcessingInstruction
    Set pi = doc.createProcessingInstruction("xml", "version=""1.0""")
    doc.appendChild pi
    Set pi = doc.createProcessingInstruction("qbxml", "version=""2.0""")
    doc.appendChild pi
    Dim qbxml As IXMLDOMElement
    Set qbxml = doc.createElement("QBXML")
    doc.appendChild qbxml
    Dim msgsrq As IXMLDOMElement
    Set msgsrq = doc.createElement("QBXMLMsgsRq")
    qbxml.appendChild msgsrq
    msgsrq.setAttribute "onError", "continueOnError"
    Dim queryElem As IXMLDOMElement
    Set queryElem = doc.createElement(query)
    msgsrq.appendChild queryElem
    Set BuildSimpleQuery = doc
End Function

Function qbXMLLatestVersion(rp As RequestProcessor2, ticket As String) As String
    Dim strXMLVersions() As String
    'Should be able to use this, but there appears to be a bug that may cause 2.0 to be returned
    'when it should not.
    'strXMLVersions = rp.QBXMLVersionsForSession(ticket)
    
    'Instead we use HostQuery
    'Create a DOM document object for creating our request.
    Dim xml As New DOMDocument

    'Create the QBXML aggregate
    Dim rootElement As IXMLDOMNode
    Set rootElement = xml.createElement("QBXML")
    xml.appendChild rootElement
  
    'Add the QBXMLMsgsRq aggregate to the QBXML aggregate
    Dim QBXMLMsgsRqNode As IXMLDOMNode
    Set QBXMLMsgsRqNode = xml.createElement("QBXMLMsgsRq")
    rootElement.appendChild QBXMLMsgsRqNode

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'If we were writing a real application this is where we would add
    'a newMessageSetID so we could perform error recovery.  Any time a
    'request contains an add, delete, modify or void request developers
    'should use the error recovery mechanisms.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Set the QBXMLMsgsRq onError attribute to continueOnError
    Dim onErrorAttr As IXMLDOMAttribute
    Set onErrorAttr = xml.createAttribute("onError")
    onErrorAttr.text = "stopOnError"
    QBXMLMsgsRqNode.Attributes.setNamedItem onErrorAttr
  
    'Add the InvoiceAddRq aggregate to QBXMLMsgsRq aggregate
    Dim HostQuery As IXMLDOMNode
    Set HostQuery = xml.createElement("HostQueryRq")
    QBXMLMsgsRqNode.appendChild HostQuery
    
    Dim strXMLRequest As String
    strXMLRequest = _
        "<?xml version=""1.0"" ?>" & _
        "<!DOCTYPE QBXML PUBLIC '-//INTUIT//DTD QBXML QBD 1.0//EN' 'http://developer.intuit.com'>" _
        & rootElement.xml

    Dim strXMLResponse As String
    strXMLResponse = rp.ProcessRequest(ticket, strXMLRequest)
    Dim QueryResponse As New DOMDocument

    'Parse the response XML
    QueryResponse.async = False
    QueryResponse.loadXML (strXMLResponse)

    Dim supportedVersions As IXMLDOMNodeList
    Set supportedVersions = QueryResponse.getElementsByTagName("SupportedQBXMLVersion")
    
    Dim VersNode As IXMLDOMNode
    
    Dim i As Long
    Dim vers As Double
    Dim LastVers As Double
    LastVers = 0
    For i = 0 To supportedVersions.length - 1
        Set VersNode = supportedVersions.item(i)
        vers = VersNode.firstChild.text
        If (vers > LastVers) Then
            LastVers = vers
            qbXMLLatestVersion = VersNode.firstChild.text
        End If
    Next i
End Function



