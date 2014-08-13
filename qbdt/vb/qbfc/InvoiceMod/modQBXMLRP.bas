Attribute VB_Name = "modQBXMLRP"
'----------------------------------------------------------
' Copyright © 2003-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/rdmgr/?ID=100
'
'----------------------------------------------------------

Option Explicit

Dim objRequestProcessor As RequestProcessor2

Dim strTicket As String
Dim booSessionBegun As Boolean

Dim strqbXMLVersionLine As String


Dim objSavedDOMDocument As New DOMDocument40
Dim objSavedInvoiceRet As IXMLDOMNode

Dim strSavedRequest As String

Public Sub QBXMLRP_OpenConnectionBeginSession()
  
  On Error GoTo Errs
  
  Set objRequestProcessor = New RequestProcessor2
  
  objRequestProcessor.OpenConnection "", "IDN Invoice Modify Sample Application"
  strTicket = objRequestProcessor.BeginSession("", qbFileOpenDoNotCare)
  booSessionBegun = True
  Exit Sub
  
Errs:
  If Err.Number = &H80040416 Then
    MsgBox "You must have QuickBooks running with a company" & vbCrLf & _
           "file open to use this program."
    objRequestProcessor.CloseConnection
    End
  ElseIf Err.Number = &H80040422 Then
    MsgBox "This QuickBooks company file is open in single user mode and" & vbCrLf & _
           "another application is already accessing it.  Please exit the" & vbCrLf & _
           "other application and run this application again."
    objRequestProcessor.CloseConnection
    End
  ElseIf Err.Number = &H1AD Then
    MsgBox _
      "It appears that the qbXML Request Processor has not" & vbCrLf & _
      "been installed, indicating QuickBooks 2002 or later" & vbCrLf & _
      "may not have been installed.  Please run this sample" & vbCrLf & _
      "after installing QuickBooks 2003 and running the Upgrade."
    End
  ElseIf Err.Number = &H1AE Then
    MsgBox _
      "It appears that you have QuickBooks 2002 R1P installed." & vbCrLf & _
      "This program requires QuickBooks 2003 to work."
    End
  Else
    MsgBox Err.Number & vbCrLf & Hex(Err.Number) & vbCrLf & _
           Err.Description
    End
  End If
End Sub


Public Sub QBXMLRP_EndSessionCloseConnection()
  If booSessionBegun Then
    objRequestProcessor.EndSession strTicket
    objRequestProcessor.CloseConnection
  End If
End Sub


Function QBXMLRP_MaxVersionSupported() As String

'This section contained obsolete code that has been stubbed out

  QBXMLRP_MaxVersionSupported = "3.0"
End Function


Public Sub QBXMLRP_FillComboBox(cmbComboBox As ComboBox, _
                                strQueryType As String, _
                                strNameElement As String, _
                                strFilter As String, _
                                booMarkGroupItems As Boolean)
  
  'Clear the combo box
  cmbComboBox.Clear
  
  Dim strSplits() As String
  strSplits = Split(strQueryType, ",")
  
  Dim strNameElementSplits() As String
  strNameElementSplits = Split(strNameElement, ",")
  
  Dim strXMLRequest As String
  Dim strXMLResponse As String
  
  Dim i As Integer
  For i = 0 To UBound(strSplits)
    strXMLRequest = "<?xml version=""1.0""?>" & "<?qbxml version=""3.0""?>" & _
      "<QBXML><QBXMLMsgsRq onError=""stopOnError""><" & strSplits(i) & _
      "QueryRq>" & strFilter & "</" & strSplits(i) & "QueryRq>" & _
      "</QBXMLMsgsRq></QBXML>"

    strXMLResponse = objRequestProcessor.ProcessRequest(strTicket, strXMLRequest)
  
    'Set up a DOM document object to load the response into
    Dim objDOMResponseDoc As New DOMDocument40

    'Parse the response XML
    objDOMResponseDoc.async = False
    objDOMResponseDoc.loadXML (strXMLResponse)

    'Get the status for our query request
    Dim objQueryRs As IXMLDOMNodeList
    Set objQueryRs = objDOMResponseDoc.getElementsByTagName(strSplits(i) & "QueryRs")
  
    Dim rsStatusAttr        As IXMLDOMNamedNodeMap
    Set rsStatusAttr = objQueryRs.Item(0).Attributes
    Dim strQueryStatus As String
    strQueryStatus = rsStatusAttr.getNamedItem("statusCode").nodeValue

    'If the status is bad, report it to the user
    If strQueryStatus <> "0" Then
      If strQueryStatus = "1" Then
        Exit Sub
      Else
        MsgBox rsStatusAttr.getNamedItem("statusMessage").nodeValue
        Exit Sub
      End If
    End If
  
    'Parse the query response and add the query values to the combo box
    Dim QueryRsNode As IXMLDOMNode
    Set QueryRsNode = objQueryRs.Item(0)
  
    Dim NodeList    As IXMLDOMNodeList
    Set NodeList = objDOMResponseDoc.getElementsByTagName(strSplits(i) & "Ret")
  
    Dim numItems            As Long
    numItems = NodeList.length
  
    If UBound(strNameElementSplits) > 0 Then
      strNameElement = strNameElementSplits(i)
    End If
    
    'Declare the XML objects outside of the loop
    Dim itemNode          As IXMLDOMNode
    Dim j                 As Integer
    For j = 0 To numItems - 1
      Set itemNode = NodeList.Item(j)
      If strSplits(i) = "ItemGroup" And booMarkGroupItems Then
        cmbComboBox.AddItem itemNode.selectSingleNode(strNameElement).Text & _
          " - Group Item"
      Else
        cmbComboBox.AddItem itemNode.selectSingleNode(strNameElement).Text
      End If
    Next j
  Next i
End Sub


Public Sub QBXMLRP_FillInvoiceList(lstInvoices As ListBox, _
                                   strRefNumber As String, _
                                   strFromDateTime As String, _
                                   strToDateTime As String, _
                                   strDateQueryType As String, _
                                   strDateMacro As String, _
                                   strCustomerJob As String, _
                                   booCustomerWithChildren As Boolean, _
                                   strAccount As String, _
                                   booAccountWithChildren As Boolean, _
                                   strFromRefNumberRange As String, _
                                   strToRefNumberRange As String, _
                                   strRefNumberPiece As String, _
                                   strRefNumberCriteria As String, _
                                   strPaidStatus As String)

  Dim strXMLRequest As String
  
  If strRefNumber <> "" Then
    'We only need to query for the ref number, don't bother building
    'the XML with MSXML
    strXMLRequest = _
      "<?xml version=""1.0""?>" & "<?qbxml version=""3.0""?>" & _
      "<QBXML><QBXMLMsgsRq onError=""stopOnError""><InvoiceQueryRq>" & _
      "<RefNumberCaseSensitive>" & strRefNumber & "</RefNumberCaseSensitive>" & _
      "<IncludeLineItems>true</IncludeLineItems></InvoiceQueryRq>" & _
      "</QBXMLMsgsRq></QBXML>"
  Else 'strRefNumber <> ""
    Dim objDOMDocument As New DOMDocument40
    
    Dim objRootNode As IXMLDOMNode
    Dim objRequestNode As IXMLDOMNode
    Dim objElement As IXMLDOMElement
    Dim objAttribute As IXMLDOMAttribute
    
    CreateStandardRequestNode _
      False, "continueOnError", objDOMDocument, objRootNode, objRequestNode, objAttribute
    
    Dim objInvoiceQueryNode As IXMLDOMNode
    AddMSXMLNode "InvoiceQueryRq", objDOMDocument, objRequestNode, objInvoiceQueryNode
    
    'Limit our response to 30 invoices
    AddMSXMLElement "MaxReturned", "30", objDOMDocument, objInvoiceQueryNode, objElement
    
    If strFromDateTime <> Empty Or strToDateTime <> Empty Then
      Dim objDateTimeFilter As IXMLDOMNode
      AddMSXMLNode strDateQueryType, objDOMDocument, objInvoiceQueryNode, objDateTimeFilter
      If strDateQueryType = "ModifiedDateRangeFilter" Then
        If strFromDateTime <> Empty Then
          AddMSXMLElement _
            "FromModifiedDate", strFromDateTime, objDOMDocument, objDateTimeFilter, objElement
        End If
        If strToDateTime <> Empty Then
          AddMSXMLElement _
            "ToModifiedDate", strToDateTime, objDOMDocument, objDateTimeFilter, objElement
        End If
      Else 'strDateQueryType = "ModifiedDateRangeFilter"
        If strFromDateTime <> Empty Then
          AddMSXMLElement _
            "FromTxnDate", strFromDateTime, objDOMDocument, objDateTimeFilter, objElement
        End If
        If strToDateTime <> Empty Then
          AddMSXMLElement _
            "ToTxnDate", strToDateTime, objDOMDocument, objDateTimeFilter, objElement
        End If
      End If 'strDateQueryType = "ModifiedDateRangeFilter"
    End If 'strFromDateTime <> Empty Or strToDateTime <> Empty
    
    If strDateMacro <> Empty Then
      Dim objDateFilter As IXMLDOMNode
      AddMSXMLNode "TxnDateRangeFilter", objDOMDocument, objInvoiceQueryNode, objDateFilter
      AddMSXMLElement "DateMacro", strDateMacro, objDOMDocument, objDateFilter, objElement
    End If
    
    If strCustomerJob <> Empty Then
      Dim objEntityFilter As IXMLDOMNode
      AddMSXMLNode _
        "EntityFilter", objDOMDocument, objInvoiceQueryNode, objEntityFilter
      
      If booCustomerWithChildren Then
        AddMSXMLElement _
          "FullNameWithChildren", strCustomerJob, objDOMDocument, objEntityFilter, objElement
      Else
        AddMSXMLElement _
          "FullName", strCustomerJob, objDOMDocument, objEntityFilter, objElement
      End If
    End If
    
    If strAccount <> Empty Then
      Dim objAccountFilter As IXMLDOMNode
      AddMSXMLNode _
        "AccountFilter", objDOMDocument, objInvoiceQueryNode, objAccountFilter
      
      If booAccountWithChildren Then
        AddMSXMLElement _
          "FullNameWithChildren", strAccount, objDOMDocument, objAccountFilter, objElement
      Else
        AddMSXMLElement _
          "FullName", strAccount, objDOMDocument, objAccountFilter, objElement
      End If
    End If
    
    If strFromRefNumberRange <> Empty Or strToRefNumberRange <> Empty Then
      Dim objRefNumberRangeFilter As IXMLDOMNode
      AddMSXMLNode "RefNumberRangeFilter", objDOMDocument, objInvoiceQueryNode, objRefNumberRangeFilter
      If strFromRefNumberRange <> Empty Then
        AddMSXMLElement _
          "FromRefNumber", strFromRefNumberRange, objDOMDocument, objRefNumberRangeFilter, objElement
      End If
      If strToRefNumberRange <> Empty Then
        AddMSXMLElement _
        "ToRefNumber", strToRefNumberRange, objDOMDocument, objRefNumberRangeFilter, objElement
      End If
    End If 'strFromRefNumberRange <> Empty Or strToRefNumberRange <> Empty

    If strRefNumberPiece <> Empty Then
      Dim objRefNumberFilter As IXMLDOMNode
      AddMSXMLNode "RefNumberFilter", objDOMDocument, objInvoiceQueryNode, objRefNumberFilter
      AddMSXMLElement "MatchCriterion", strRefNumberCriteria, objDOMDocument, objRefNumberFilter, objElement
      AddMSXMLElement "RefNumber", strRefNumberPiece, objDOMDocument, objRefNumberFilter, objElement
    End If

    If strPaidStatus <> Empty Then
      AddMSXMLElement "PaidStatus", strPaidStatus, objDOMDocument, objInvoiceQueryNode, objElement
    End If
  
    AddMSXMLElement "IncludeLineItems", "true", objDOMDocument, objInvoiceQueryNode, objElement
    
    strXMLRequest = "<?xml version=""1.0""?>" & "<?qbxml version=""3.0""?>" & objRootNode.xml
  End If 'strRefNumber <> ""
  PrettyPrintXMLToFile strXMLRequest, "C:\debugrq.xml"
  strSavedRequest = PrettyPrintXMLToString(strXMLRequest)
  
  Dim strXMLResponse As String
  strXMLResponse = objRequestProcessor.ProcessRequest(strTicket, strXMLRequest)
  PrettyPrintXMLToFile strXMLResponse, "C:\debugrs.xml"
  
  objDOMDocument.async = False
  objDOMDocument.loadXML (strXMLResponse)
  
  Dim objInvoiceQueryNodeList As IXMLDOMNodeList
  Set objInvoiceQueryNodeList = objDOMDocument.getElementsByTagName("InvoiceQueryRs")
  
  Set objInvoiceQueryNode = objInvoiceQueryNodeList.Item(0)
  
  Dim objInvoiceQueryAttributes As IXMLDOMNamedNodeMap
  Set objInvoiceQueryAttributes = objInvoiceQueryNode.Attributes
  
  If objInvoiceQueryAttributes.getNamedItem("statusCode").nodeValue <> "0" Then
    MsgBox "Error getting Invoices" & vbCrLf & vbCrLf & _
      "Error = " & _
      objInvoiceQueryAttributes.getNamedItem("statusCode").nodeValue & _
      vbCrLf & vbCrLf & "Message = " & _
      objInvoiceQueryAttributes.getNamedItem("statusMessage").nodeValue
      
    lstInvoices.AddItem "No invoices match the query filter used"
    Exit Sub
  End If
  
  Dim objInvoiceRetNodeList As IXMLDOMNodeList
  Set objInvoiceRetNodeList = objDOMDocument.getElementsByTagName("InvoiceRet")
  
  Dim objInvoiceRet As IXMLDOMNode
  Dim objNodeList As IXMLDOMNodeList
  Dim intItems As Integer
  Dim strReturnedRefNumber As String
  Dim strItems As String
  
  Dim i As Integer
  For i = 0 To objInvoiceRetNodeList.length - 1
    Set objInvoiceRet = objInvoiceRetNodeList.Item(i)
    
    If Not objInvoiceRet.selectSingleNode("RefNumberCaseSensitive") Is Nothing Then
      strReturnedRefNumber = _
        "Invoice " & objInvoiceRet.selectSingleNode("RefNumberCaseSensitive").Text
    Else
      strReturnedRefNumber = "Un-numbered "
    End If
    
    Set objNodeList = objInvoiceRet.selectNodes("InvoiceLineRet")
    If objNodeList Is Nothing Then
      intItems = 0
    Else
      intItems = objNodeList.length
    End If
    Set objNodeList = objInvoiceRet.selectNodes("InvoiceLineGroupRet")
    If Not objNodeList Is Nothing Then
      intItems = intItems + objNodeList.length
    End If
    
    strItems = Str(intItems)
    If Len(strItems) = 1 Then strItems = "  " & strItems
    If Len(strItems) = 2 Then strItems = " " & strItems
    lstInvoices.AddItem _
      strReturnedRefNumber & "     " & _
      objInvoiceRet.selectSingleNode("TxnDate").Text & _
      "     " & strItems & " items     " & _
      objInvoiceRet.selectSingleNode("CustomerRef").selectSingleNode("FullName").Text & _
      "     Balance " & objInvoiceRet.selectSingleNode("BalanceRemaining").Text & _
      "     " & objInvoiceRet.selectSingleNode("TxnID").Text
  Next i
End Sub


Public Sub QBXMLRP_GetInvoice(TxnID As String)

  Dim strXMLRequest As String
  
  'We only need to query for the TxnID, don't bother building
  'the XML with MSXML
  strXMLRequest = _
    "<?xml version=""1.0""?>" & "<?qbxml version=""3.0""?>" & _
    "<QBXML><QBXMLMsgsRq onError=""stopOnError""><InvoiceQueryRq>" & _
    "<TxnID>" & TxnID & "</TxnID><IncludeLineItems>1</IncludeLineItems>" & _
    "</InvoiceQueryRq></QBXMLMsgsRq></QBXML>"

  Dim strXMLResponse As String
  strXMLResponse = objRequestProcessor.ProcessRequest(strTicket, strXMLRequest)

  objSavedDOMDocument.async = False
  objSavedDOMDocument.loadXML (strXMLResponse)
  
  Dim objInvoiceQueryNodeList As IXMLDOMNodeList
  Set objInvoiceQueryNodeList = objSavedDOMDocument.getElementsByTagName("InvoiceQueryRs")
  
  Dim objInvoiceQueryNode As IXMLDOMNode
  Set objInvoiceQueryNode = objInvoiceQueryNodeList.Item(0)
  
  Dim objInvoiceQueryAttributes As IXMLDOMNamedNodeMap
  Set objInvoiceQueryAttributes = objInvoiceQueryNode.Attributes
  
  If objInvoiceQueryAttributes.getNamedItem("statusCode").nodeValue <> "0" Then
    MsgBox "Error getting Invoices" & vbCrLf & vbCrLf & _
      "Error = " & _
      objInvoiceQueryAttributes.getNamedItem("statusCode").nodeValue & _
      vbCrLf & vbCrLf & "Message = " & _
      objInvoiceQueryAttributes.getNamedItem("statusMessage").nodeValue
    Exit Sub
  End If

  Dim objInvoiceRetNodeList As IXMLDOMNodeList
  Set objInvoiceRetNodeList = objSavedDOMDocument.getElementsByTagName("InvoiceRet")
  
  Set objSavedInvoiceRet = objInvoiceRetNodeList.Item(0)
End Sub


Public Sub QBXMLRP_LoadInvoiceModifyForm()
  frmInvoiceModify.txtTxnID.Text = _
    objSavedInvoiceRet.selectSingleNode("TxnID").Text
  
  frmInvoiceModify.txtEditSequence.Text = _
    objSavedInvoiceRet.selectSingleNode("EditSequence").Text
  
  If Not objSavedInvoiceRet.selectSingleNode("RefNumberCaseSensitive") Is Nothing Then
    frmInvoiceModify.txtRefNumber.Text = _
      objSavedInvoiceRet.selectSingleNode("RefNumberCaseSensitive").Text
  End If
  
  frmInvoiceModify.txtTxnDate.Text = _
    objSavedInvoiceRet.selectSingleNode("TxnDate").Text
  
  If Not objSavedInvoiceRet.selectSingleNode("IsPending") Is Nothing Then
    If objSavedInvoiceRet.selectSingleNode("IsPending").Text = "true" Then
      frmInvoiceModify.chkPending.Value = 1 'Checked
    End If
  End If
  
  If Not objSavedInvoiceRet.selectSingleNode("IsToBePrinted") Is Nothing Then
    If objSavedInvoiceRet.selectSingleNode("IsToBePrinted").Text = "true" Then
      frmInvoiceModify.chkToBePrinted.Value = 1 'Checked
    End If
  End If
  
  frmInvoiceModify.cmbCustomer.Text = _
    objSavedInvoiceRet.selectSingleNode("CustomerRef").selectSingleNode("FullName").Text
  
  If Not objSavedInvoiceRet.selectSingleNode("ClassRef") Is Nothing Then
    frmInvoiceModify.cmbClass.Text = _
      objSavedInvoiceRet.selectSingleNode("ClassRef").selectSingleNode("FullName").Text
  End If
  
  If Not objSavedInvoiceRet.selectSingleNode("BillAddress") Is Nothing Then
    If Not objSavedInvoiceRet.selectSingleNode("BillAddress").selectSingleNode("Addr1") Is Nothing Then
      frmInvoiceModify.txtBillAddr1.Text = objSavedInvoiceRet.selectSingleNode("BillAddress").selectSingleNode("Addr1").Text
    End If
    
    If Not objSavedInvoiceRet.selectSingleNode("BillAddress").selectSingleNode("Addr2") Is Nothing Then
      frmInvoiceModify.txtBillAddr2.Text = objSavedInvoiceRet.selectSingleNode("BillAddress").selectSingleNode("Addr2").Text
    End If
    
    If Not objSavedInvoiceRet.selectSingleNode("BillAddress").selectSingleNode("Addr3") Is Nothing Then
      frmInvoiceModify.txtBillAddr3.Text = objSavedInvoiceRet.selectSingleNode("BillAddress").selectSingleNode("Addr3").Text
    End If
    
    If Not objSavedInvoiceRet.selectSingleNode("BillAddress").selectSingleNode("Addr4") Is Nothing Then
      frmInvoiceModify.txtBillAddr4.Text = objSavedInvoiceRet.selectSingleNode("BillAddress").selectSingleNode("Addr4").Text
    End If
    
    If Not objSavedInvoiceRet.selectSingleNode("BillAddress").selectSingleNode("City") Is Nothing Then
      frmInvoiceModify.txtBillCity.Text = objSavedInvoiceRet.selectSingleNode("BillAddress").selectSingleNode("City").Text
    End If
    
    If Not objSavedInvoiceRet.selectSingleNode("BillAddress").selectSingleNode("State") Is Nothing Then
      frmInvoiceModify.txtBillState.Text = objSavedInvoiceRet.selectSingleNode("BillAddress").selectSingleNode("State").Text
    End If
    
    If Not objSavedInvoiceRet.selectSingleNode("BillAddress").selectSingleNode("PostalCode") Is Nothing Then
      frmInvoiceModify.txtBillPostalCode.Text = objSavedInvoiceRet.selectSingleNode("BillAddress").selectSingleNode("PostalCode").Text
    End If
    
    If Not objSavedInvoiceRet.selectSingleNode("BillAddress").selectSingleNode("Country") Is Nothing Then
      frmInvoiceModify.txtBillCountry.Text = objSavedInvoiceRet.selectSingleNode("BillAddress").selectSingleNode("Country").Text
    End If
  End If
  
  If Not objSavedInvoiceRet.selectSingleNode("ShipAddress") Is Nothing Then
    If Not objSavedInvoiceRet.selectSingleNode("ShipAddress").selectSingleNode("Addr1") Is Nothing Then
      frmInvoiceModify.txtShipAddr1.Text = objSavedInvoiceRet.selectSingleNode("ShipAddress").selectSingleNode("Addr1").Text
    End If
    
    If Not objSavedInvoiceRet.selectSingleNode("ShipAddress").selectSingleNode("Addr2") Is Nothing Then
      frmInvoiceModify.txtShipAddr2.Text = objSavedInvoiceRet.selectSingleNode("ShipAddress").selectSingleNode("Addr2").Text
    End If
    
    If Not objSavedInvoiceRet.selectSingleNode("ShipAddress").selectSingleNode("Addr3") Is Nothing Then
      frmInvoiceModify.txtShipAddr3.Text = objSavedInvoiceRet.selectSingleNode("ShipAddress").selectSingleNode("Addr3").Text
    End If
    
    If Not objSavedInvoiceRet.selectSingleNode("ShipAddress").selectSingleNode("Addr4") Is Nothing Then
      frmInvoiceModify.txtShipAddr4.Text = objSavedInvoiceRet.selectSingleNode("ShipAddress").selectSingleNode("Addr4").Text
    End If
    
    If Not objSavedInvoiceRet.selectSingleNode("ShipAddress").selectSingleNode("City") Is Nothing Then
      frmInvoiceModify.txtShipCity.Text = objSavedInvoiceRet.selectSingleNode("ShipAddress").selectSingleNode("City").Text
    End If
    
    If Not objSavedInvoiceRet.selectSingleNode("ShipAddress").selectSingleNode("State") Is Nothing Then
      frmInvoiceModify.txtShipState.Text = objSavedInvoiceRet.selectSingleNode("ShipAddress").selectSingleNode("State").Text
    End If
    
    If Not objSavedInvoiceRet.selectSingleNode("ShipAddress").selectSingleNode("PostalCode") Is Nothing Then
      frmInvoiceModify.txtShipPostalCode.Text = objSavedInvoiceRet.selectSingleNode("ShipAddress").selectSingleNode("PostalCode").Text
    End If
    
    If Not objSavedInvoiceRet.selectSingleNode("ShipAddress").selectSingleNode("Country") Is Nothing Then
      frmInvoiceModify.txtShipCountry.Text = objSavedInvoiceRet.selectSingleNode("ShipAddress").selectSingleNode("Country").Text
    End If
  End If
  
  If Not objSavedInvoiceRet.selectSingleNode("ARAccountRef") Is Nothing Then
    frmInvoiceModify.cmbARAccount.Text = _
      objSavedInvoiceRet.selectSingleNode("ARAccountRef").selectSingleNode("FullName").Text
  End If
  
  If Not objSavedInvoiceRet.selectSingleNode("TermsRef") Is Nothing Then
    frmInvoiceModify.cmbTerms.Text = _
      objSavedInvoiceRet.selectSingleNode("TermsRef").selectSingleNode("FullName").Text
  End If
  
  If Not objSavedInvoiceRet.selectSingleNode("PONumber") Is Nothing Then
    frmInvoiceModify.txtPONumber.Text = _
      objSavedInvoiceRet.selectSingleNode("PONumber").Text
  End If
  
  If Not objSavedInvoiceRet.selectSingleNode("DueDate") Is Nothing Then
    frmInvoiceModify.txtDueDate.Text = _
      objSavedInvoiceRet.selectSingleNode("DueDate").Text
  End If
  
  If Not objSavedInvoiceRet.selectSingleNode("ShipDate") Is Nothing Then
    frmInvoiceModify.txtShipDate.Text = _
      objSavedInvoiceRet.selectSingleNode("ShipDate").Text
  End If
  
  If Not objSavedInvoiceRet.selectSingleNode("FOB") Is Nothing Then
    frmInvoiceModify.txtFOB.Text = _
      objSavedInvoiceRet.selectSingleNode("FOB").Text
  End If
  
  If Not objSavedInvoiceRet.selectSingleNode("SalesRepRef") Is Nothing Then
    frmInvoiceModify.cmbSalesRep.Text = _
      objSavedInvoiceRet.selectSingleNode("SalesRepRef").selectSingleNode("FullName").Text
  End If
  
  If Not objSavedInvoiceRet.selectSingleNode("ShipMethodRef") Is Nothing Then
    frmInvoiceModify.cmbShipMethod.Text = _
      objSavedInvoiceRet.selectSingleNode("ShipMethodRef").selectSingleNode("FullName").Text
  End If
  
  If Not objSavedInvoiceRet.selectSingleNode("ItemSalesTaxRef") Is Nothing Then
    frmInvoiceModify.cmbItemSalesTax.Text = _
      objSavedInvoiceRet.selectSingleNode("ItemSalesTaxRef").selectSingleNode("FullName").Text
  End If
  
  If Not objSavedInvoiceRet.selectSingleNode("CustomerSalesTaxCodeRef") Is Nothing Then
    frmInvoiceModify.cmbCustTaxCode.Text = _
      objSavedInvoiceRet.selectSingleNode("CustomerSalesTaxCodeRef").selectSingleNode("FullName").Text
  End If
  
  If Not objSavedInvoiceRet.selectSingleNode("CustomerMsgRef") Is Nothing Then
    frmInvoiceModify.cmbCustomerMsg.Text = _
      objSavedInvoiceRet.selectSingleNode("CustomerMsgRef").selectSingleNode("FullName").Text
  End If
  
  If Not objSavedInvoiceRet.selectSingleNode("Memo") Is Nothing Then
    frmInvoiceModify.txtMemo.Text = _
      objSavedInvoiceRet.selectSingleNode("Memo").Text
  End If
End Sub


Public Sub QBXMLRP_LoadInvoiceLineArray(strLineArray() As String)
  
  Dim objNodeList As IXMLDOMNodeList
  Set objNodeList = objSavedInvoiceRet.selectNodes("InvoiceLineRet")
  
  Dim objNode As IXMLDOMNode
  Dim strFoo As String
  If objNodeList.length > 0 Then
    Set objNode = objNodeList.Item(0)
    strFoo = objNode.previousSibling.nodeName
    If strFoo = "InvoiceLineGroupRet" Then
      Set objNodeList = objSavedInvoiceRet.selectNodes("InvoiceLineGroupRet")
    End If
  Else
    Set objNodeList = objSavedInvoiceRet.selectNodes("InvoiceLineGroupRet")
  End If
  
  Set objNode = objNodeList.Item(0)
  Dim objGroupLines As IXMLDOMNodeList
  Dim objLine As IXMLDOMNode
  Dim i As Integer
  Dim j As Integer
  i = 1
  Do While objNode.nodeName = "InvoiceLineRet" Or _
           objNode.nodeName = "InvoiceLineGroupRet"
    If objNode.nodeName = "InvoiceLineRet" Then
      strLineArray(i) = LineInfo(objNode) & "Item<spliter>Original"
    Else
      strLineArray(i) = LineInfo(objNode) & "Group<spliter>Original"
      
      Set objGroupLines = objNode.selectNodes("InvoiceLineRet")
      If objGroupLines.length > 0 Then
        For j = 0 To objGroupLines.length - 1
          i = i + 1
          Set objLine = objGroupLines.Item(j)
          strLineArray(i) = LineInfo(objLine) & "SubItem<spliter>Original"
        Next j
      End If 'objGroupLines.length > 0
    End If 'objNode.nodeName = "InvoiceLineRet"
    
    If objNode.nextSibling Is Nothing Then
      Set objNode = objSavedInvoiceRet.selectSingleNode("TxnID")
    Else
      Set objNode = objNode.nextSibling
      i = i + 1
    End If
  Loop
End Sub


Private Function LineInfo(objNode As IXMLDOMNode) As String

  Dim strLineInfo As String
  Dim strRateOrPercent As String

  strLineInfo = objNode.selectSingleNode("TxnLineID").Text & "<spliter>"
  
  If Not objNode.selectSingleNode("Quantity") Is Nothing Then
    strLineInfo = strLineInfo & objNode.selectSingleNode("Quantity").Text
  End If
  strLineInfo = strLineInfo & "<spliter>"
  
  If Not objNode.selectSingleNode("ItemRef") Is Nothing Then
    strLineInfo = strLineInfo & _
      objNode.selectSingleNode("ItemRef").selectSingleNode("FullName").Text
  ElseIf Not objNode.selectSingleNode("ItemGroupRef") Is Nothing Then
    strLineInfo = strLineInfo & _
      objNode.selectSingleNode("ItemGroupRef").selectSingleNode("FullName").Text
  End If
  strLineInfo = strLineInfo & "<spliter>"
  
  If Not objNode.selectSingleNode("Desc") Is Nothing Then
    strLineInfo = strLineInfo & objNode.selectSingleNode("Desc").Text
  End If
  strLineInfo = strLineInfo & "<spliter>"

  If Not objNode.selectSingleNode("Rate") Is Nothing Then
    strLineInfo = strLineInfo & objNode.selectSingleNode("Rate").Text
    strRateOrPercent = "Rate"
  ElseIf Not objNode.selectSingleNode("RatePercent") Is Nothing Then
    strLineInfo = strLineInfo & objNode.selectSingleNode("RatePercent").Text
    strRateOrPercent = "RatePercent"
  Else
    strRateOrPercent = "Neither"
  End If
  strLineInfo = strLineInfo & "<spliter>"

  If Not objNode.selectSingleNode("Amount") Is Nothing Then
    strLineInfo = strLineInfo & objNode.selectSingleNode("Amount").Text
  End If
  strLineInfo = strLineInfo & "<spliter>"
  
  If Not objNode.selectSingleNode("ClassRef") Is Nothing Then
    strLineInfo = strLineInfo & _
      objNode.selectSingleNode("ClassRef").selectSingleNode("FullName").Text
  End If
  strLineInfo = strLineInfo & "<spliter>"
  
  If Not objNode.selectSingleNode("ServiceDate") Is Nothing Then
    strLineInfo = strLineInfo & objNode.selectSingleNode("ServiceDate").Text
  End If
  strLineInfo = strLineInfo & "<spliter>"
    
  If Not objNode.selectSingleNode("SalesTaxCodeRef") Is Nothing Then
    strLineInfo = strLineInfo & _
      objNode.selectSingleNode("SalesTaxCodeRef").selectSingleNode("FullName").Text
  End If
  strLineInfo = strLineInfo & "<spliter>" & strRateOrPercent & _
    "<spliter><spliter>"
  
  
  LineInfo = strLineInfo
End Function


Public Sub QBXMLRP_ModifyInvoice(strInvoiceChangeString As String)

  Dim strTemp As String
  strTemp = strInvoiceChangeString
  
  Dim objDOMDoc As New DOMDocument40
  Dim objRootNode As IXMLDOMNode
  Dim objRequestNode As IXMLDOMNode
  Dim objElement As IXMLDOMElement
  Dim objAttribute As IXMLDOMAttribute
  
  CreateStandardRequestNode _
    True, "continueOnError", objDOMDoc, objRootNode, objRequestNode, objAttribute

  Dim objInvoiceModRqNode As IXMLDOMNode
  AddMSXMLNode "InvoiceModRq", objDOMDoc, objRequestNode, objInvoiceModRqNode
  
  Dim objInvoiceModNode As IXMLDOMNode
  AddMSXMLNode "InvoiceMod", objDOMDoc, objInvoiceModRqNode, objInvoiceModNode
  
  'We know the invoice change string starts out as <TxnID>value</TxnID>
  'then <EditSequence>value</EditSequence> so use that knowledge to pull
  'out the values and put them in the request
  
  strTemp = Right(strTemp, Len(strTemp) - 7)
  AddMSXMLElement "TxnID", Left(strTemp, InStr(1, strTemp, "</") - 1), _
    objDOMDoc, objInvoiceModNode, objElement

  strTemp = Right(strTemp, Len(strTemp) - (InStr(1, strTemp, "<EditS") + 13))
  AddMSXMLElement "EditSequence", Left(strTemp, InStr(1, strTemp, "</") - 1), _
    objDOMDoc, objInvoiceModNode, objElement

  strTemp = Right(strTemp, Len(strTemp) - (InStr(1, strTemp, "</Edit") + 14))
  
  'The rest of our invoice change string is either going to be elements to
  'add to the InvoiceModRq aggregate or aggregates to add to the aggregate.
  'We'll pull each one out of the invoice change string and treat them the
  'same letting the recursive procedure we call figure it out.

  Dim strStartTag As String
  Dim strEndTag As String
  Dim strSegment As String
  Dim intTagLength As Integer
  
  Dim objNode As IXMLDOMNode
  
  Do
    GetTags strTemp, strStartTag, strEndTag, intTagLength
    If intTagLength = 0 Then
      MsgBox "Error processing invoice change string, exiting"
      End
    End If
    
    strSegment = Left(strTemp, InStr(1, strTemp, strEndTag) + intTagLength)
    strTemp = Right(strTemp, Len(strTemp) - Len(strSegment))
    
    AddElementOrAggregate _
      strSegment, strStartTag, intTagLength, objDOMDoc, _
      objInvoiceModNode, objNode, objElement

  Loop Until strTemp = Empty
  
  Dim strXMLRequest As String
  strXMLRequest = "<?xml version=""1.0""?>" & "<?qbxml version=""3.0""?>" & objRootNode.xml
  PrettyPrintXMLToFile strXMLRequest, "C:\InvoiceModRq.xml"
  strSavedRequest = PrettyPrintXMLToString(strXMLRequest)
  
  Dim strXMLResponse As String
  strXMLResponse = objRequestProcessor.ProcessRequest(strTicket, strXMLRequest)
  PrettyPrintXMLToFile strXMLResponse, "C:\InvoiceModRs.xml"
  
  objDOMDoc.async = False
  objDOMDoc.loadXML (strXMLResponse)
  
  Dim objInvoiceModNodeList As IXMLDOMNodeList
  Set objInvoiceModNodeList = objDOMDoc.getElementsByTagName("InvoiceModRs")
  
  Set objInvoiceModNode = objInvoiceModNodeList.Item(0)
  
  Dim objInvoiceModAttributes As IXMLDOMNamedNodeMap
  Set objInvoiceModAttributes = objInvoiceModNode.Attributes
  
  If objInvoiceModAttributes.getNamedItem("statusCode").nodeValue <> "0" Then
    MsgBox "Error modifying invoice" & vbCrLf & vbCrLf & _
      "Error = " & _
      objInvoiceModAttributes.getNamedItem("statusCode").nodeValue & _
      vbCrLf & vbCrLf & "Message = " & _
      objInvoiceModAttributes.getNamedItem("statusMessage").nodeValue
      
    Exit Sub
  Else
    MsgBox "Successfully modified invoice"
  End If
  
End Sub


Private Sub AddElementOrAggregate( _
              strElOrAggString As String, _
              strStartTag As String, _
              intTagLength As Integer, _
              objDOMDoc As DOMDocument40, _
              objParentNode As IXMLDOMNode, _
              objNode As IXMLDOMNode, _
              objElement As IXMLDOMElement)

  Dim strTemp As String
  strTemp = strElOrAggString
  Dim strTagName As String
  strTagName = Left(strStartTag, Len(strStartTag) - 1)
  strTagName = Right(strTagName, Len(strTagName) - 1)
  
  Dim strValue As String
  strValue = Left(strTemp, Len(strTemp) - (intTagLength + 1))
  strValue = Right(strValue, Len(strValue) - intTagLength)
  
  Dim strInnerStartTag As String
  Dim strInnerEndTag As String
  Dim intInnerTagLength As Integer
  
  GetTags strValue, strInnerStartTag, strInnerEndTag, intInnerTagLength
  
  If strInnerStartTag = Empty Then
    AddMSXMLElement strTagName, strValue, objDOMDoc, objParentNode, objElement
  Else
    Dim objChildNode As IXMLDOMNode
    AddMSXMLNode strTagName, objDOMDoc, objParentNode, objChildNode
    
    Dim strSegment As String
    Do
      strSegment = Left(strValue, InStr(1, strValue, strInnerEndTag) + intInnerTagLength)
      strValue = Right(strValue, Len(strValue) - Len(strSegment))

      AddElementOrAggregate _
        strSegment, strInnerStartTag, intInnerTagLength, objDOMDoc, _
        objChildNode, objNode, objElement
    
      If strValue <> Empty Then GetTags strValue, strInnerStartTag, strInnerEndTag, intInnerTagLength
    Loop Until strValue = Empty
  End If
End Sub
             

Public Sub QBXMLRP_GetItemInfo(strItemFullName As String, _
                               strDesc As String, _
                               strRate As String, _
                               strRateOrPercent As String, _
                               strSalesTaxCode As String)

  Dim objDOMDoc As New DOMDocument40
  Dim objRootNode As IXMLDOMNode
  Dim objRequestNode As IXMLDOMNode
  Dim objElement As IXMLDOMElement
  Dim objAttribute As IXMLDOMAttribute
  
  CreateStandardRequestNode _
    False, "continueOnError", objDOMDoc, objRootNode, objRequestNode, objAttribute

  Dim objQueryNode As IXMLDOMNode
  AddMSXMLNode "ItemQueryRq", objDOMDoc, objRequestNode, objQueryNode
  
  AddMSXMLElement "FullName", strItemFullName, objDOMDoc, objQueryNode, objElement
  
  Dim strXMLRequest As String
  strXMLRequest = "<?xml version=""1.0""?>" & "<?qbxml version=""3.0""?>" & objRootNode.xml
  PrettyPrintXMLToFile strXMLRequest, "C:\debugrq.xml"
  
  Dim strXMLResponse As String
  strXMLResponse = objRequestProcessor.ProcessRequest(strTicket, strXMLRequest)
  PrettyPrintXMLToFile strXMLResponse, "C:\debugrs.xml"
  
  objDOMDoc.async = False
  objDOMDoc.loadXML (strXMLResponse)
  
  Dim objItemQueryNodeList As IXMLDOMNodeList
  Set objItemQueryNodeList = objDOMDoc.getElementsByTagName("ItemQueryRs")
  
  Dim objItemQueryNode As IXMLDOMNode
  Set objItemQueryNode = objItemQueryNodeList.Item(0)
  
  Dim objItemQueryAttributes As IXMLDOMNamedNodeMap
  Set objItemQueryAttributes = objItemQueryNode.Attributes
  
  If objItemQueryAttributes.getNamedItem("statusCode").nodeValue <> "0" Then
    MsgBox "Error getting Item information" & vbCrLf & vbCrLf & _
      "Error = " & _
      objItemQueryAttributes.getNamedItem("statusCode").nodeValue & _
      vbCrLf & vbCrLf & "Message = " & _
      objItemQueryAttributes.getNamedItem("statusMessage").nodeValue
      
    Exit Sub
  End If
  
  Dim objItemRetNode As IXMLDOMNode
  Set objItemRetNode = objItemQueryNode.firstChild
    
  Dim strItemType As String
  strItemType = objItemRetNode.nodeName
  
  If strItemType = "ItemInventoryRet" Or _
     strItemType = "ItemInventoryAssemblyRet" Then
    
    If Not objItemRetNode.selectSingleNode("SalesDesc") Is Nothing Then
      strDesc = objItemRetNode.selectSingleNode("SalesDesc").Text
    End If
  
    If Not objItemRetNode.selectSingleNode("SalesPrice") Is Nothing Then
      strRate = objItemRetNode.selectSingleNode("SalesPrice").Text
    End If
  
    If Not objItemRetNode.selectSingleNode("SalesTaxCodeRef") Is Nothing Then
      strSalesTaxCode = _
        objItemRetNode.selectSingleNode("SalesTaxCodeRef").selectSingleNode("FullName").Text
    End If
  
  ElseIf strItemType = "ItemSubtotalRet" Or _
         strItemType = "ItemPaymentRet" Or _
         strItemType = "ItemSalesTaxRet" Or _
         strItemType = "ItemGroupRet" Then
    
    If Not objItemRetNode.selectSingleNode("ItemDesc") Is Nothing Then
      strDesc = objItemRetNode.selectSingleNode("ItemDesc").Text
    End If
  
  ElseIf strItemType = "ItemDiscountRet" Then
    
    If Not objItemRetNode.selectSingleNode("ItemDesc") Is Nothing Then
      strDesc = objItemRetNode.selectSingleNode("ItemDesc").Text
    End If
  
    If Not objItemRetNode.selectSingleNode("SalesTaxCodeRef") Is Nothing Then
      strSalesTaxCode = _
        objItemRetNode.selectSingleNode("SalesTaxCodeRef").selectSingleNode("FullName").Text
    End If
  
    If Not objItemRetNode.selectSingleNode("DiscountRate") Is Nothing Then
      strRate = objItemRetNode.selectSingleNode("DiscountRate").Text
    End If
  
    If Not objItemRetNode.selectSingleNode("DiscountRatePercent") Is Nothing Then
      strRate = objItemRetNode.selectSingleNode("DiscountRatePercent").Text
      strRateOrPercent = "RatePercent"
    End If
  
  Else
    
    If Not objItemRetNode.selectSingleNode("SalesTaxCodeRef") Is Nothing Then
      strSalesTaxCode = _
        objItemRetNode.selectSingleNode("SalesTaxCodeRef").selectSingleNode("FullName").Text
    End If
  
    If Not objItemRetNode.selectSingleNode("SalesOrPurchase") Is Nothing Then
      If Not objItemRetNode.selectSingleNode("SalesOrPurchase").selectSingleNode("Desc") Is Nothing Then
        strDesc = objItemRetNode.selectSingleNode("SalesOrPurchase").selectSingleNode("Desc").Text
      End If
  
      If Not objItemRetNode.selectSingleNode("SalesOrPurchase").selectSingleNode("Price") Is Nothing Then
        strRate = objItemRetNode.selectSingleNode("SalesOrPurchase").selectSingleNode("Price").Text
      End If
  
      If Not objItemRetNode.selectSingleNode("SalesOrPurchase").selectSingleNode("PricePercent") Is Nothing Then
        strRate = objItemRetNode.selectSingleNode("SalesOrPurchase").selectSingleNode("PricePercent").Text
        strRateOrPercent = "RatePercent"
      End If
    Else
      
      If Not objItemRetNode.selectSingleNode("SalesAndPurchase").selectSingleNode("SalesDesc") Is Nothing Then
        strDesc = objItemRetNode.selectSingleNode("SalesAndPurchase").selectSingleNode("SalesDesc").Text
      End If
  
      If Not objItemRetNode.selectSingleNode("SalesAndPurchase").selectSingleNode("SalesPrice") Is Nothing Then
        strRate = objItemRetNode.selectSingleNode("SalesAndPurchase").selectSingleNode("SalesPrice").Text
        strRateOrPercent = "Rate"
      End If
    End If
  End If
End Sub


Public Sub QBXMLRP_LoadRequest(strRequestText As String)
  strRequestText = strSavedRequest
End Sub
