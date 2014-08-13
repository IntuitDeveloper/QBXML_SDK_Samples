Attribute VB_Name = "modQBXMLRP"
'----------------------------------------------------------
' Copyright © 2003-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'----------------------------------------------------------

Option Explicit

Dim objRequestProcessor As QBXMLRP2Lib.RequestProcessor2

Dim strTicket As String
Dim booSessionBegun As Boolean




Public Sub QBXMLRP_OpenConnectionBeginSession()
  
  On Error GoTo Errs
  
  Set objRequestProcessor = New QBXMLRP2Lib.RequestProcessor2
  
  objRequestProcessor.OpenConnection "", "IDN PO Modify Sample Application"
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

  QBXMLRP_MaxVersionSupported = "3.0"
End Function


Public Sub QBXMLRP_LoadPOInfoArray(strPOInfo() As String, _
                                   dateFrom As Date, _
                                   dateTo As Date, _
                                   booGiveWarning As Boolean)
  
  Dim strXMLRequest As String
  
  strXMLRequest = _
    "<?xml version=""1.0"" ?>" & "<?qbxml version=""3.0""?>" & _
    "<QBXML><QBXMLMsgsRq onError = ""stopOnError"">" & _
    "<GeneralDetailReportQueryRq>" & _
    "<GeneralDetailReportType>OpenPOs</GeneralDetailReportType>" & _
    "<ReportPeriod>" & _
    "<FromReportDate>" & Format(dateFrom, "yyyy-mm-dd") & "</FromReportDate>" & _
    "<ToReportDate>" & Format(dateTo, "yyyy-mm-dd") & "</ToReportDate>" & _
    "</ReportPeriod>" & _
    "<IncludeColumn>RefNumber</IncludeColumn>" & _
    "<IncludeColumn>Date</IncludeColumn>" & _
    "<IncludeColumn>Name</IncludeColumn>" & _
    "<IncludeColumn>DeliveryDate</IncludeColumn>" & _
    "<IncludeColumn>Amount</IncludeColumn>" & _
    "<IncludeColumn>TxnID</IncludeColumn>" & _
    "</GeneralDetailReportQueryRq></QBXMLMsgsRq></QBXML>"

  Dim strXMLResponse As String
  strXMLResponse = objRequestProcessor.ProcessRequest(strTicket, strXMLRequest)

  Dim objXMLDoc As New DOMDocument40
  
  objXMLDoc.async = False
  objXMLDoc.loadXML (strXMLResponse)
  
  Dim objReportResponseNodeList As IXMLDOMNodeList
  Set objReportResponseNodeList = objXMLDoc.getElementsByTagName("GeneralDetailReportQueryRs")
  
  Dim objReportResponseNode As IXMLDOMNode
  Set objReportResponseNode = objReportResponseNodeList.Item(0)
  
  Dim objReportResponseAttributes As IXMLDOMNamedNodeMap
  Set objReportResponseAttributes = objReportResponseNode.Attributes
  
  If objReportResponseAttributes.getNamedItem("statusCode").nodeValue <> "0" Then
    MsgBox "Error getting PO list" & vbCrLf & vbCrLf & _
     "Error = " & _
     objReportResponseAttributes.getNamedItem("statusCode").nodeValue & _
     vbCrLf & vbCrLf & "Message = " & _
     objReportResponseAttributes.getNamedItem("statusMessage").nodeValue
     
    strPOInfo(0) = "<spliter><spliter>There were no open purchase orders returned<spliter><spliter>"
    Exit Sub
  End If
  
  Dim objDataRowList As IXMLDOMNodeList
  Set objDataRowList = objXMLDoc.getElementsByTagName("DataRow")
  
  Dim intDataRows As Integer
  intDataRows = objDataRowList.length
  
  Dim objColDataList As IXMLDOMNodeList
  Dim objColData As IXMLDOMNode
  
  Dim i As Integer
  Dim j As Integer
  If intDataRows > UBound(strPOInfo) Then
    If booGiveWarning Then
      MsgBox "This program limits the number of selected open purchase orders to " & _
        UBound(strPOInfo) & vbCrLf & vbCrLf & _
        "This program will display the first " & UBound(strPOInfo) & _
        " open purchase orders returned from your query period"
    End If
    j = UBound(strPOInfo)
  Else
    j = intDataRows
  End If
  For i = 0 To j - 1
    Set objColDataList = objDataRowList.Item(i).selectNodes("ColData")
      If objColDataList.length = 6 Then
        strPOInfo(i + 1) = _
        objColDataList.Item(0).Attributes.getNamedItem("value").nodeValue & _
        "<spliter>" & _
        objColDataList.Item(1).Attributes.getNamedItem("value").nodeValue & _
        "<spliter>" & _
        objColDataList.Item(2).Attributes.getNamedItem("value").nodeValue & _
        "<spliter>" & _
        objColDataList.Item(3).Attributes.getNamedItem("value").nodeValue & _
        "<spliter>" & _
        objColDataList.Item(4).Attributes.getNamedItem("value").nodeValue & _
        "<spliter>" & _
        objColDataList.Item(5).Attributes.getNamedItem("value").nodeValue
      Else
        strPOInfo(i + 1) = _
        objColDataList.Item(0).Attributes.getNamedItem("value").nodeValue & _
        "<spliter>" & _
        objColDataList.Item(1).Attributes.getNamedItem("value").nodeValue & _
        "<spliter><spliter>" & _
        objColDataList.Item(2).Attributes.getNamedItem("value").nodeValue & _
        "<spliter>" & _
        objColDataList.Item(3).Attributes.getNamedItem("value").nodeValue & _
        "<spliter>" & _
        objColDataList.Item(4).Attributes.getNamedItem("value").nodeValue
    End If
  Next
End Sub


Public Sub QBXMLRP_GetPOInfo(strTxnID As String, _
                             strEditSequence As String, _
                             strRefNumber As String, _
                             strTxnDate As String, _
                             strVendor As String, _
                             strPOLines() As String, _
                             strSelectedPOInfo As String)

  Dim strPOInfoSplit() As String
  
  strPOInfoSplit = Split(strSelectedPOInfo, "<spliter>")

  Dim strXMLRequest As String
  strXMLRequest = _
    "<?xml version=""1.0"" ?>" & "<?qbxml version=""3.0""?>" & _
    "<QBXML><QBXMLMsgsRq onError = ""stopOnError"">" & _
    "<PurchaseOrderQueryRq>" & _
    "<TxnID>" & strPOInfoSplit(5) & "</TxnID>" & _
    "<IncludeLineItems>true</IncludeLineItems>" & _
    "</PurchaseOrderQueryRq>" & _
    "</QBXMLMsgsRq></QBXML>"

  Dim strXMLResponse As String
  strXMLResponse = objRequestProcessor.ProcessRequest(strTicket, strXMLRequest)
  
  Dim objDOMDoc As New DOMDocument40
  
  objDOMDoc.async = False
  objDOMDoc.loadXML strXMLResponse
  
  Dim objPORetList As IXMLDOMNodeList
  Set objPORetList = objDOMDoc.getElementsByTagName("PurchaseOrderRet")
  
  Dim objPORet As IXMLDOMNode
  Set objPORet = objPORetList.Item(0)
  
  strTxnID = objPORet.selectSingleNode("TxnID").Text
  strEditSequence = objPORet.selectSingleNode("EditSequence").Text
  strTxnDate = objPORet.selectSingleNode("TxnDate").Text
  
  If Not objPORet.selectSingleNode("RefNumber") Is Nothing Then
    strRefNumber = objPORet.selectSingleNode("RefNumber").Text
  End If
  
  If Not objPORet.selectSingleNode("VendorRef") Is Nothing Then
    strVendor = _
      objPORet.selectSingleNode("VendorRef").selectSingleNode("FullName").Text
  End If
  
  Dim objChildNode As IXMLDOMNode
  Set objChildNode = objPORet.firstChild
  
  Do
    Set objChildNode = objChildNode.nextSibling
  Loop Until objChildNode.nodeName = "PurchaseOrderLineRet" Or _
             objChildNode.nodeName = "PurchaseOrderLineGroupRet"
  
  Dim objGroupLineList As IXMLDOMNodeList
  Dim booDone As Boolean
  Dim i As Integer
  Dim j As Integer
  booDone = False
  i = 1
  
  Do
    strPOLines(i) = POLineString(objChildNode, "")
    
    If objChildNode.nodeName = "PurchaseOrderLineGroupRet" Then
      Set objGroupLineList = objChildNode.selectNodes("PurchaseOrderLineRet")
      For j = 0 To objGroupLineList.length - 1
        i = i + 1
        strPOLines(i) = POLineString(objGroupLineList.Item(j), "GroupSubItem")
      Next j
    End If
    
    Set objChildNode = objChildNode.nextSibling
    i = i + 1
    booDone = (objChildNode Is Nothing)
  Loop Until booDone
  
End Sub


Private Function POLineString(objLineNode As IXMLDOMNode, _
                              strLineType As String)

  Dim strOrdered As String
  Dim strReceived As String

  Dim strOutput As String
  
  If objLineNode.nodeName = "PurchaseOrderLineRet" Then
    If Not objLineNode.selectSingleNode("ItemRef") Is Nothing Then
      strOutput = _
        objLineNode.selectSingleNode("ItemRef").selectSingleNode("FullName").Text
    End If
  Else
    strOutput = _
      objLineNode.selectSingleNode("ItemGroupRef").selectSingleNode("FullName").Text
  End If
  strOutput = strOutput & "<spliter>"
  
  If Not objLineNode.selectSingleNode("Desc") Is Nothing Then
    strOutput = strOutput & _
      objLineNode.selectSingleNode("Desc").Text
  End If
  strOutput = strOutput & "<spliter>"
  
  If Not objLineNode.selectSingleNode("Quantity") Is Nothing Then
    strOrdered = objLineNode.selectSingleNode("Quantity").Text
    strOutput = strOutput & strOrdered
  Else
    strOrdered = Empty
  End If
  strOutput = strOutput & "<spliter>"
  
  If Not objLineNode.selectSingleNode("Rate") Is Nothing Then
    strOutput = strOutput & _
      objLineNode.selectSingleNode("Rate").Text
  End If
  strOutput = strOutput & "<spliter>"
  
  If Not objLineNode.selectSingleNode("CustomerRef") Is Nothing Then
    strOutput = strOutput & _
      objLineNode.selectSingleNode("CustomerRef").selectSingleNode("FullName").Text
  End If
  strOutput = strOutput & "<spliter>"
  
  If Not objLineNode.selectSingleNode("Amount") Is Nothing Then
    strOutput = strOutput & _
      objLineNode.selectSingleNode("Amount").Text
  End If
  strOutput = strOutput & "<spliter>"
  
  If Not objLineNode.selectSingleNode("ReceivedQuantity") Is Nothing Then
    strReceived = objLineNode.selectSingleNode("ReceivedQuantity").Text
    strOutput = strOutput & strReceived
  Else
    strReceived = Empty
  End If
  strOutput = strOutput & "<spliter>"
  
  If Not objLineNode.selectSingleNode("IsManuallyClosed") Is Nothing Then
    If objLineNode.selectSingleNode("IsManuallyClosed").Text = "true" Or _
       (strOrdered <> Empty And strOrdered = strReceived) Or _
       (strOrdered = "0") Then
      strOutput = strOutput & "X"
    End If
  End If
  strOutput = strOutput & "<spliter>"
  
  strOutput = strOutput & _
    objLineNode.selectSingleNode("TxnLineID").Text & "<spliter>"
  
  If strLineType = Empty Then
    If objLineNode.nodeName = "PurchaseOrderLineRet" Then
      strOutput = strOutput & "Item"
    Else
      strOutput = strOutput & "GroupItem"
    End If
  Else
    strOutput = strOutput & strLineType
  End If
  strOutput = strOutput & "<spliter>"
    
  If Not objLineNode.selectSingleNode("IsManuallyClosed") Is Nothing Then
    strOutput = strOutput & _
      objLineNode.selectSingleNode("IsManuallyClosed").Text
  End If
  strOutput = strOutput & "<spliter>"
  
  POLineString = strOutput
End Function

Public Sub QBXMLRP_ClosePO(strTxnID As String, _
                           strEditSequence As String)

  Dim strXMLRequest As String
  strXMLRequest = _
    "<?xml version=""1.0"" ?>" & "<?qbxml version=""3.0""?>" & _
    "<QBXML><QBXMLMsgsRq onError = ""stopOnError"">" & _
    "<PurchaseOrderModRq><PurchaseOrderMod>" & _
    "<TxnID>" & strTxnID & "</TxnID>" & _
    "<EditSequence>" & strEditSequence & "</EditSequence>" & _
    "<IsManuallyClosed>true</IsManuallyClosed>" & _
    "</PurchaseOrderMod></PurchaseOrderModRq>" & _
    "</QBXMLMsgsRq></QBXML>"

  Dim strXMLResponse As String
  strXMLResponse = objRequestProcessor.ProcessRequest(strTicket, strXMLRequest)
  
  Dim objXMLDoc As New DOMDocument40
  
  objXMLDoc.async = False
  objXMLDoc.loadXML (strXMLResponse)
  
  Dim objReportResponseNodeList As IXMLDOMNodeList
  Set objReportResponseNodeList = objXMLDoc.getElementsByTagName("PurchaseOrderModRs")
  
  Dim objReportResponseNode As IXMLDOMNode
  Set objReportResponseNode = objReportResponseNodeList.Item(0)
  
  Dim objReportResponseAttributes As IXMLDOMNamedNodeMap
  Set objReportResponseAttributes = objReportResponseNode.Attributes
  
  If objReportResponseAttributes.getNamedItem("statusCode").nodeValue <> "0" Then
    MsgBox "Error closing purchase order" & vbCrLf & vbCrLf & _
      "Error = " & _
      objReportResponseAttributes.getNamedItem("statusCode").nodeValue & _
      vbCrLf & vbCrLf & "Message = " & _
      objReportResponseAttributes.getNamedItem("statusMessage").nodeValue
  Else
    MsgBox "Purchase order successfully closed"
  End If
End Sub


Public Sub QBXMLRP_ChangePOLine(strAction As String, _
                                strTxnID As String, _
                                strEditSequence As String, _
                                strPOLines() As String, _
                                intSelectedPOLine As Integer, _
                                intGroupLine As Integer)
  
  Dim strXMLRequest As String
  strXMLRequest = _
    "<?xml version=""1.0"" ?>" & "<?qbxml version=""3.0""?>" & _
    "<QBXML><QBXMLMsgsRq onError = ""stopOnError"">" & _
    "<PurchaseOrderModRq><PurchaseOrderMod>" & _
    "<TxnID>" & strTxnID & "</TxnID>" & _
    "<EditSequence>" & strEditSequence & "</EditSequence>"

  Dim Splits() As String
  
  Dim i As Integer
  Dim booDone As Boolean
  Dim booInGroup As Boolean
  Dim booInSelectedGroup As Boolean
  booInGroup = False
  booInSelectedGroup = False
  booDone = False
  i = 1
  Do
    Splits = Split(strPOLines(i), "<spliter>")
      
    If booInGroup And Splits(9) <> "GroupSubItem" Then
      strXMLRequest = strXMLRequest & "</PurchaseOrderLineGroupMod>"
      booInGroup = False
    End If
    
    If Splits(9) = "GroupItem" Then
      strXMLRequest = strXMLRequest & "<PurchaseOrderLineGroupMod>" & _
        "<TxnLineID>" & Splits(8) & "</TxnLineID>"
      booInGroup = True
    ElseIf (Splits(9) = "GroupSubItem" And _
           (booInSelectedGroup Or strAction = "ReceiveAll")) Or _
           Splits(9) = "Item" Then
      strXMLRequest = strXMLRequest & "<PurchaseOrderLineMod>" & _
        "<TxnLineID>" & Splits(8) & "</TxnLineID>"
    End If
    
    If Not booInSelectedGroup Then
      If i = intGroupLine Then
        booInSelectedGroup = True
      End If
    ElseIf Splits(9) <> "GroupSubItem" Then
      booInSelectedGroup = False
    End If
    
    If i = intSelectedPOLine Or strAction = "ReceiveAll" Then
      If strAction = "Close" Then
        strXMLRequest = strXMLRequest & _
          "<IsManuallyClosed>true</IsManuallyClosed>"
      ElseIf Splits(2) <> Empty Then
        strXMLRequest = strXMLRequest & _
          "<>" & Splits(2) & "</>"
      End If
    End If
    
    If (Splits(9) = "GroupSubItem" And booInSelectedGroup) Or _
       Splits(9) = "Item" Then
      strXMLRequest = strXMLRequest & "</PurchaseOrderLineMod>"
    End If
    
    If i = UBound(strPOLines) Then
      booDone = True
    ElseIf strPOLines(i + 1) = Empty Then
      booDone = True
      If booInGroup Then
        strXMLRequest = strXMLRequest & "</PurchaseOrderLineGroupMod>"
      End If
    End If
    
    i = i + 1
  Loop Until booDone
  
  strXMLRequest = strXMLRequest & "</PurchaseOrderMod></PurchaseOrderModRq>" & _
    "</QBXMLMsgsRq></QBXML>"

'  PrettyPrintXMLToFile strXMLRequest, "C:\Temp\PurchaseOrderModify.xml"
'  EndSessionCloseConnection
'  End
  
  Dim strXMLResponse As String
  strXMLResponse = objRequestProcessor.ProcessRequest(strTicket, strXMLRequest)
  
  Dim objXMLDoc As New DOMDocument40
  
  objXMLDoc.async = False
  objXMLDoc.loadXML (strXMLResponse)
  
  Dim objReportResponseNodeList As IXMLDOMNodeList
  Set objReportResponseNodeList = objXMLDoc.getElementsByTagName("PurchaseOrderModRs")
  
  Dim objReportResponseNode As IXMLDOMNode
  Set objReportResponseNode = objReportResponseNodeList.Item(0)
  
  Dim objReportResponseAttributes As IXMLDOMNamedNodeMap
  Set objReportResponseAttributes = objReportResponseNode.Attributes
  
  If objReportResponseAttributes.getNamedItem("statusCode").nodeValue <> "0" Then
    MsgBox "Error getting PO list" & vbCrLf & vbCrLf & _
      "Error = " & _
      objReportResponseAttributes.getNamedItem("statusCode").nodeValue & _
      vbCrLf & vbCrLf & "Message = " & _
      objReportResponseAttributes.getNamedItem("statusMessage").nodeValue
  Else
    MsgBox "Purchase Order Line successfully closed"
  End If
End Sub


Public Sub QBXMLRP_SetQuantitiesAndBillForRemainingItems( _
             strTxnID As String, _
             strEditSequence As String, _
             strVendor As String, _
             strRefNumber As String, _
             strTxnDate As String, _
             strPOLines() As String, _
             intPOLine As Integer)

  Dim objDOMDoc As New DOMDocument40
  Dim objRootNode As IXMLDOMNode
  Dim objRequestNode As IXMLDOMNode
  Dim objModNode As IXMLDOMNode
  Dim objNode As IXMLDOMNode
  Dim objElement As IXMLDOMElement
  Dim objAttribute As IXMLDOMAttribute
  
  'We're creating a standard request message set node setting the
  'newMessageSetID because this includes a modify and an add request and
  'setting the onError attribute to stopOnError so we don't create a bill
  'if our purchase order modify fails
  
  CreateStandardRequestNode True, "stopOnError", objDOMDoc, objRootNode, objRequestNode, objAttribute
  
  'Add the PurchaseOrderModRq and PurchaseOrderMod nodes
  AddMSXMLNode "PurchaseOrderModRq", objDOMDoc, objRequestNode, objNode
  AddMSXMLAttribute "requestID", "0", objDOMDoc, objNode, objAttribute
  
  AddMSXMLNode "PurchaseOrderMod", objDOMDoc, objNode, objModNode
  
  'Add the TxnID and EditSequence of the purchase order we're modifying
  AddMSXMLElement "TxnID", strTxnID, objDOMDoc, objModNode, objElement
  AddMSXMLElement "EditSequence", strEditSequence, objDOMDoc, objModNode, objElement
  
  'Loop through the PO lines and change the ordered quantity to match the
  'received quantity
  Dim i As Integer
  Dim booDone As Boolean
  Dim booPOModified As Boolean
  Dim objParentNode As IXMLDOMNode
  Dim strSplits() As String
  
  booPOModified = False
  Set objParentNode = objModNode
  i = 1
  Do
    strSplits = Split(strPOLines(i), "<spliter>")
    
    'The quantity ordered or received can be empty, set then to zero if they're empty
    If strSplits(6) = "" Then strSplits(6) = "0"
    If strSplits(2) = "" Then strSplits(2) = "0"
    
    If strSplits(9) <> "GroupSubItem" Then
      Set objParentNode = objModNode
    End If
    
    If strSplits(9) = "GroupItem" Then
      AddMSXMLNode "PurchaseOrderLineGroupMod", objDOMDoc, objModNode, objNode
      AddMSXMLElement "TxnLineID", strSplits(8), objDOMDoc, objNode, objElement
      Set objParentNode = objNode
    Else
      AddMSXMLNode "PurchaseOrderLineMod", objDOMDoc, objParentNode, objNode
      AddMSXMLElement "TxnLineID", strSplits(8), objDOMDoc, objNode, objElement
      
      'If the PO line hasn't been manually closed and the number of items
      'received is less than the number ordered change the number ordered
      'to the number received
      If strSplits(7) <> "X" And Int(strSplits(6)) < Int(strSplits(2)) And _
         ((intPOLine = 0) Or (intPOLine = i)) Then
        AddMSXMLElement "Quantity", strSplits(6), objDOMDoc, objNode, objElement
        booPOModified = True
      End If
    End If
    
    If i = UBound(strPOLines) Then
      booDone = True
    ElseIf strPOLines(i + 1) = Empty Then
      booDone = True
    Else
      i = i + 1
    End If
  Loop Until booDone
  
  If Not booPOModified Then
    MsgBox "The purchase order selected has either has all lines " & _
      "manually closed or all lines received.  Ending processing " & _
      "on the existing purchase order and not creating a new bill " & _
      "for missing items."
    Exit Sub
  End If
  
  'Now add the bill for the items we modified in the purchase order
  AddMSXMLNode "BillAddRq", objDOMDoc, objRequestNode, objNode
  AddMSXMLAttribute "requestID", "1", objDOMDoc, objNode, objAttribute
  
  Dim objBillAddNode As IXMLDOMNode
  AddMSXMLNode "BillAdd", objDOMDoc, objNode, objBillAddNode
  
  AddMSXMLNode "VendorRef", objDOMDoc, objBillAddNode, objNode
  AddMSXMLElement "FullName", strVendor, objDOMDoc, objNode, objElement
  
  AddMSXMLElement "Memo", _
    "Created by Purchase Order Modify Sample for modified PO " & strRefNumber & _
    "; TxnDate " & strTxnDate & "; TxnID " & strTxnID & _
    " .  Quantity ordered " & _
    "reduced to quantity recieved for one or all lines, differences " & _
    "reflected in this bill.", _
    objDOMDoc, objBillAddNode, objElement

  Dim objItemLineNode As IXMLDOMNode
  Dim j As Integer
  For j = 1 To i
    strSplits = Split(strPOLines(j), "<spliter>")
    
    'The quantity ordered or received can be empty, set then to zero if they're empty
    If strSplits(6) = "" Then strSplits(6) = "0"
    If strSplits(2) = "" Then strSplits(2) = "0"
    
    If strSplits(9) <> "GroupItem" And strSplits(7) <> "X" And _
       Int(strSplits(6)) < Int(strSplits(2)) And _
       ((intPOLine = 0) Or (intPOLine = j)) Then

      AddMSXMLNode "ItemLineAdd", objDOMDoc, objBillAddNode, objItemLineNode
      
      AddMSXMLNode "ItemRef", objDOMDoc, objItemLineNode, objNode
      AddMSXMLElement "FullName", strSplits(0), objDOMDoc, objNode, objElement

      AddMSXMLElement "Quantity", Str(Int(strSplits(2)) - Int(strSplits(6))), _
        objDOMDoc, objItemLineNode, objElement
    
      If strSplits(3) <> Empty Then
        AddMSXMLElement "Cost", strSplits(3), objDOMDoc, objItemLineNode, objElement
      ElseIf strSplits(5) <> Empty Then
        AddMSXMLElement "Amount", strSplits(5), objDOMDoc, objItemLineNode, objElement
      End If
      
      If strSplits(4) <> Empty Then
        AddMSXMLNode "CustomerRef", objDOMDoc, objItemLineNode, objNode
        AddMSXMLElement "FullName", strSplits(4), objDOMDoc, objNode, objElement
      End If
    End If
  Next j
  
  Dim strXMLRequest As String
  strXMLRequest = "<?xml version=""1.0""?>" & "<?qbxml version=""3.0""?>" & objRootNode.xml
  Dim fnum As Integer
  fnum = FreeFile()
  
  Dim strXMLResponse As String
  strXMLResponse = objRequestProcessor.ProcessRequest(strTicket, strXMLRequest)
  
  Dim objXMLDoc As New DOMDocument40
  
  objXMLDoc.async = False
  objXMLDoc.loadXML (strXMLResponse)
  
  Dim objResponseNodeList As IXMLDOMNodeList
  Set objResponseNodeList = objXMLDoc.getElementsByTagName("PurchaseOrderModRs")
  
  Dim objResponseNode As IXMLDOMNode
  Set objResponseNode = objResponseNodeList.Item(0)
  
  If objResponseNode Is Nothing Then
      MsgBox "Error from QuickBooks processing request: " & strXMLResponse
  End If
  
  
  Dim objResponseAttributes As IXMLDOMNamedNodeMap
  Set objResponseAttributes = objResponseNode.Attributes
  
  If objResponseAttributes.getNamedItem("statusCode").nodeValue <> "0" Then
    MsgBox "Error modifying PO" & vbCrLf & vbCrLf & _
      "Error = " & _
      objResponseAttributes.getNamedItem("statusCode").nodeValue & _
      vbCrLf & vbCrLf & "Message = " & _
      objResponseAttributes.getNamedItem("statusMessage").nodeValue
    Exit Sub
  End If
  
  'Get the new edit sequence and Memo for updating the Memo field of the PO
  Dim strNewEditSequence As String
  strNewEditSequence = objResponseNode.selectSingleNode("PurchaseOrderRet").selectSingleNode("EditSequence").Text
  Dim strMemo As String
  If Not objResponseNode.selectSingleNode("PurchaseOrderRet").selectSingleNode("Memo") Is Nothing Then
    strMemo = objResponseNode.selectSingleNode("PurchaseOrderRet").selectSingleNode("Memo").Text
  Else
    strMemo = ""
  End If
  
  Set objResponseNodeList = objXMLDoc.getElementsByTagName("BillAddRs")
  Set objResponseNode = objResponseNodeList.Item(0)
  Set objResponseAttributes = objResponseNode.Attributes
  
  If objResponseAttributes.getNamedItem("statusCode").nodeValue <> "0" Then
    MsgBox "Error adding Bill" & vbCrLf & vbCrLf & _
      "Error = " & _
      objResponseAttributes.getNamedItem("statusCode").nodeValue & _
      vbCrLf & vbCrLf & "Message = " & _
      objResponseAttributes.getNamedItem("statusMessage").nodeValue
    Exit Sub
  Else
    If intPOLine = 0 Then
      MsgBox "Successfully set PO lines and billed for items"
    Else
      MsgBox "Successfully set PO line and billed for items"
    End If
  End If
  
  'Now let's update the PO Memo to include the manipulation we performed
  Dim strBillTxnID As String
  Dim strBillRefNumber As String
  Dim strBillTxnDate As String
  
  Dim objBillRet As IXMLDOMNode
  
  Set objBillRet = objResponseNode.selectSingleNode("BillRet")
  
  strBillTxnID = objBillRet.selectSingleNode("TxnID").Text
  strBillTxnDate = objBillRet.selectSingleNode("TxnDate").Text
  
  If Not objBillRet.selectSingleNode("RefNumber") Is Nothing Then
    strBillRefNumber = objBillRet.selectSingleNode("RefNumber").Text
  End If
  
  strMemo = strMemo & _
    " >>> Purchase order modified by ""Purchase Order Modify Sample"": " & _
    "Quantity ordered changed to quantity received value "
  If intPOLine = 0 Then
    strMemo = strMemo & "on all lines not closed or fully recieved.  "
  Else
    strMemo = strMemo & "on one line.  "
  End If
  strMemo = strMemo & "Created bill " & strBillRefNumber & _
    " with TxnDate " & strBillTxnDate & "; TxnID " & strBillTxnID & _
    " for these items."

  strXMLRequest = _
    "<?xml version=""1.0"" ?>" & "<?qbxml version=""3.0""?>" & _
    "<QBXML><QBXMLMsgsRq onError = ""stopOnError"">" & _
    "<PurchaseOrderModRq><PurchaseOrderMod>" & _
    "<TxnID>" & strTxnID & "</TxnID>" & _
    "<EditSequence>" & strNewEditSequence & "</EditSequence>" & _
    "<Memo>" & strMemo & "</Memo>" & _
    "</PurchaseOrderMod></PurchaseOrderModRq>" & _
    "</QBXMLMsgsRq></QBXML>"

  strXMLResponse = objRequestProcessor.ProcessRequest(strTicket, strXMLRequest)

  objXMLDoc.async = False
  objXMLDoc.loadXML (strXMLResponse)
  
  Set objResponseNodeList = objXMLDoc.getElementsByTagName("PurchaseOrderModRs")
  Set objResponseNode = objResponseNodeList.Item(0)
  Set objResponseAttributes = objResponseNode.Attributes
  
  If objResponseAttributes.getNamedItem("statusCode").nodeValue <> "0" Then
    MsgBox "Error modifying PO Memo field" & vbCrLf & vbCrLf & _
      "Error = " & _
      objResponseAttributes.getNamedItem("statusCode").nodeValue & _
      vbCrLf & vbCrLf & "Message = " & _
      objResponseAttributes.getNamedItem("statusMessage").nodeValue
  End If
End Sub




