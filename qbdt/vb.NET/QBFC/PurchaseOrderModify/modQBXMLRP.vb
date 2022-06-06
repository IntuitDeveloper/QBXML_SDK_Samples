Option Strict Off
Option Explicit On
Module modQBXMLRP
	'----------------------------------------------------------
	' Copyright © 2003-2013 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	'----------------------------------------------------------
	
	
	'UPGRADE_WARNING: Arrays in structure objRequestProcessor may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Dim objRequestProcessor As QBXMLRP2Lib.RequestProcessor2
	
	Dim strTicket As String
	Dim booSessionBegun As Boolean
	
	
	
	
	Public Sub QBXMLRP_OpenConnectionBeginSession()
		
		On Error GoTo Errs
		
		objRequestProcessor = New QBXMLRP2Lib.RequestProcessor2
		
		objRequestProcessor.OpenConnection("", "IDN PO Modify Sample Application")
		strTicket = objRequestProcessor.BeginSession("", QBXMLRP2Lib.QBFileMode.qbFileOpenDoNotCare)
		booSessionBegun = True
		Exit Sub
		
Errs: 
		If Err.Number = &H80040416 Then
			MsgBox("You must have QuickBooks running with a company" & vbCrLf & "file open to use this program.")
			objRequestProcessor.CloseConnection()
			End
		ElseIf Err.Number = &H80040422 Then 
			MsgBox("This QuickBooks company file is open in single user mode and" & vbCrLf & "another application is already accessing it.  Please exit the" & vbCrLf & "other application and run this application again.")
			objRequestProcessor.CloseConnection()
			End
		ElseIf Err.Number = &H1AD Then 
			MsgBox("It appears that the qbXML Request Processor has not" & vbCrLf & "been installed, indicating QuickBooks 2002 or later" & vbCrLf & "may not have been installed.  Please run this sample" & vbCrLf & "after installing QuickBooks 2003 and running the Upgrade.")
			End
		ElseIf Err.Number = &H1AE Then 
			MsgBox("It appears that you have QuickBooks 2002 R1P installed." & vbCrLf & "This program requires QuickBooks 2003 to work.")
			End
		Else
			MsgBox(Err.Number & vbCrLf & Hex(Err.Number) & vbCrLf & Err.Description)
			End
		End If
	End Sub
	
	
	Public Sub QBXMLRP_EndSessionCloseConnection()
		If booSessionBegun Then
			objRequestProcessor.EndSession(strTicket)
			objRequestProcessor.CloseConnection()
		End If
	End Sub
	
	
	Function QBXMLRP_MaxVersionSupported() As String
		
		QBXMLRP_MaxVersionSupported = "3.0"
	End Function
	
	
	Public Sub QBXMLRP_LoadPOInfoArray(ByRef strPOInfo() As String, ByRef dateFrom As Date, ByRef dateTo As Date, ByRef booGiveWarning As Boolean)
		
		Dim strXMLRequest As String
		
		strXMLRequest = "<?xml version=""1.0"" ?>" & "<?qbxml version=""3.0""?>" & "<QBXML><QBXMLMsgsRq onError = ""stopOnError"">" & "<GeneralDetailReportQueryRq>" & "<GeneralDetailReportType>OpenPOs</GeneralDetailReportType>" & "<ReportPeriod>" & "<FromReportDate>" & VB6.Format(dateFrom, "yyyy-mm-dd") & "</FromReportDate>" & "<ToReportDate>" & VB6.Format(dateTo, "yyyy-mm-dd") & "</ToReportDate>" & "</ReportPeriod>" & "<IncludeColumn>RefNumber</IncludeColumn>" & "<IncludeColumn>Date</IncludeColumn>" & "<IncludeColumn>Name</IncludeColumn>" & "<IncludeColumn>DeliveryDate</IncludeColumn>" & "<IncludeColumn>Amount</IncludeColumn>" & "<IncludeColumn>TxnID</IncludeColumn>" & "</GeneralDetailReportQueryRq></QBXMLMsgsRq></QBXML>"
		
		Dim strXMLResponse As String
		strXMLResponse = objRequestProcessor.ProcessRequest(strTicket, strXMLRequest)

        Dim objXMLDoc As New MSXML2.DOMDocument60

        objXMLDoc.async = False
		objXMLDoc.loadXML(strXMLResponse)
		
		Dim objReportResponseNodeList As MSXML2.IXMLDOMNodeList
		objReportResponseNodeList = objXMLDoc.getElementsByTagName("GeneralDetailReportQueryRs")
		
		Dim objReportResponseNode As MSXML2.IXMLDOMNode
		objReportResponseNode = objReportResponseNodeList.Item(0)
		
		Dim objReportResponseAttributes As MSXML2.IXMLDOMNamedNodeMap
		objReportResponseAttributes = objReportResponseNode.Attributes
		
		If objReportResponseAttributes.getNamedItem("statusCode").nodeValue <> "0" Then
			MsgBox("Error getting PO list" & vbCrLf & vbCrLf & "Error = " & objReportResponseAttributes.getNamedItem("statusCode").nodeValue & vbCrLf & vbCrLf & "Message = " & objReportResponseAttributes.getNamedItem("statusMessage").nodeValue)
			
			strPOInfo(0) = "<spliter><spliter>There were no open purchase orders returned<spliter><spliter>"
			Exit Sub
		End If
		
		Dim objDataRowList As MSXML2.IXMLDOMNodeList
		objDataRowList = objXMLDoc.getElementsByTagName("DataRow")
		
		Dim intDataRows As Short
		intDataRows = objDataRowList.length
		
		Dim objColDataList As MSXML2.IXMLDOMNodeList
		Dim objColData As MSXML2.IXMLDOMNode
		
		Dim i As Short
		Dim j As Short
		If intDataRows > UBound(strPOInfo) Then
			If booGiveWarning Then
				MsgBox("This program limits the number of selected open purchase orders to " & UBound(strPOInfo) & vbCrLf & vbCrLf & "This program will display the first " & UBound(strPOInfo) & " open purchase orders returned from your query period")
			End If
			j = UBound(strPOInfo)
		Else
			j = intDataRows
		End If
		For i = 0 To j - 1
			objColDataList = objDataRowList.Item(i).selectNodes("ColData")
			If objColDataList.length = 6 Then
				strPOInfo(i + 1) = objColDataList.Item(0).Attributes.getNamedItem("value").nodeValue & "<spliter>" & objColDataList.Item(1).Attributes.getNamedItem("value").nodeValue & "<spliter>" & objColDataList.Item(2).Attributes.getNamedItem("value").nodeValue & "<spliter>" & objColDataList.Item(3).Attributes.getNamedItem("value").nodeValue & "<spliter>" & objColDataList.Item(4).Attributes.getNamedItem("value").nodeValue & "<spliter>" & objColDataList.Item(5).Attributes.getNamedItem("value").nodeValue
			Else
				strPOInfo(i + 1) = objColDataList.Item(0).Attributes.getNamedItem("value").nodeValue & "<spliter>" & objColDataList.Item(1).Attributes.getNamedItem("value").nodeValue & "<spliter><spliter>" & objColDataList.Item(2).Attributes.getNamedItem("value").nodeValue & "<spliter>" & objColDataList.Item(3).Attributes.getNamedItem("value").nodeValue & "<spliter>" & objColDataList.Item(4).Attributes.getNamedItem("value").nodeValue
			End If
		Next 
	End Sub
	
	
	Public Sub QBXMLRP_GetPOInfo(ByRef strTxnID As String, ByRef strEditSequence As String, ByRef strRefNumber As String, ByRef strTxnDate As String, ByRef strVendor As String, ByRef strPOLines() As String, ByRef strSelectedPOInfo As String)
		
		Dim strPOInfoSplit() As String
		
		strPOInfoSplit = Split(strSelectedPOInfo, "<spliter>")
		
		Dim strXMLRequest As String
		strXMLRequest = "<?xml version=""1.0"" ?>" & "<?qbxml version=""3.0""?>" & "<QBXML><QBXMLMsgsRq onError = ""stopOnError"">" & "<PurchaseOrderQueryRq>" & "<TxnID>" & strPOInfoSplit(5) & "</TxnID>" & "<IncludeLineItems>true</IncludeLineItems>" & "</PurchaseOrderQueryRq>" & "</QBXMLMsgsRq></QBXML>"
		
		Dim strXMLResponse As String
		strXMLResponse = objRequestProcessor.ProcessRequest(strTicket, strXMLRequest)

        Dim objDOMDoc As New MSXML2.DOMDocument60

        objDOMDoc.async = False
		objDOMDoc.loadXML(strXMLResponse)
		
		Dim objPORetList As MSXML2.IXMLDOMNodeList
		objPORetList = objDOMDoc.getElementsByTagName("PurchaseOrderRet")
		
		Dim objPORet As MSXML2.IXMLDOMNode
		objPORet = objPORetList.Item(0)
		
		strTxnID = objPORet.selectSingleNode("TxnID").Text
		strEditSequence = objPORet.selectSingleNode("EditSequence").Text
		strTxnDate = objPORet.selectSingleNode("TxnDate").Text
		
		If Not objPORet.selectSingleNode("RefNumber") Is Nothing Then
			strRefNumber = objPORet.selectSingleNode("RefNumber").Text
		End If
		
		If Not objPORet.selectSingleNode("VendorRef") Is Nothing Then
			strVendor = objPORet.selectSingleNode("VendorRef").selectSingleNode("FullName").Text
		End If
		
		Dim objChildNode As MSXML2.IXMLDOMNode
		objChildNode = objPORet.firstChild
		
		Do 
			objChildNode = objChildNode.nextSibling
		Loop Until objChildNode.nodeName = "PurchaseOrderLineRet" Or objChildNode.nodeName = "PurchaseOrderLineGroupRet"
		
		Dim objGroupLineList As MSXML2.IXMLDOMNodeList
		Dim booDone As Boolean
		Dim i As Short
		Dim j As Short
		booDone = False
		i = 1
		
		Do 
			'UPGRADE_WARNING: Couldn't resolve default property of object POLineString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			strPOLines(i) = POLineString(objChildNode, "")
			
			If objChildNode.nodeName = "PurchaseOrderLineGroupRet" Then
				objGroupLineList = objChildNode.selectNodes("PurchaseOrderLineRet")
				For j = 0 To objGroupLineList.length - 1
					i = i + 1
					'UPGRADE_WARNING: Couldn't resolve default property of object POLineString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					strPOLines(i) = POLineString(objGroupLineList.Item(j), "GroupSubItem")
				Next j
			End If
			
			objChildNode = objChildNode.nextSibling
			i = i + 1
			booDone = (objChildNode Is Nothing)
		Loop Until booDone
		
	End Sub
	
	
	Private Function POLineString(ByRef objLineNode As MSXML2.IXMLDOMNode, ByRef strLineType As String) As Object
		
		Dim strOrdered As String
		Dim strReceived As String
		
		Dim strOutput As String
		
		If objLineNode.nodeName = "PurchaseOrderLineRet" Then
			If Not objLineNode.selectSingleNode("ItemRef") Is Nothing Then
				strOutput = objLineNode.selectSingleNode("ItemRef").selectSingleNode("FullName").Text
			End If
		Else
			strOutput = objLineNode.selectSingleNode("ItemGroupRef").selectSingleNode("FullName").Text
		End If
		strOutput = strOutput & "<spliter>"
		
		If Not objLineNode.selectSingleNode("Desc") Is Nothing Then
			strOutput = strOutput & objLineNode.selectSingleNode("Desc").Text
		End If
		strOutput = strOutput & "<spliter>"
		
		If Not objLineNode.selectSingleNode("Quantity") Is Nothing Then
			strOrdered = objLineNode.selectSingleNode("Quantity").Text
			strOutput = strOutput & strOrdered
		Else
			strOrdered = CStr(Nothing)
		End If
		strOutput = strOutput & "<spliter>"
		
		If Not objLineNode.selectSingleNode("Rate") Is Nothing Then
			strOutput = strOutput & objLineNode.selectSingleNode("Rate").Text
		End If
		strOutput = strOutput & "<spliter>"
		
		If Not objLineNode.selectSingleNode("CustomerRef") Is Nothing Then
			strOutput = strOutput & objLineNode.selectSingleNode("CustomerRef").selectSingleNode("FullName").Text
		End If
		strOutput = strOutput & "<spliter>"
		
		If Not objLineNode.selectSingleNode("Amount") Is Nothing Then
			strOutput = strOutput & objLineNode.selectSingleNode("Amount").Text
		End If
		strOutput = strOutput & "<spliter>"
		
		If Not objLineNode.selectSingleNode("ReceivedQuantity") Is Nothing Then
			strReceived = objLineNode.selectSingleNode("ReceivedQuantity").Text
			strOutput = strOutput & strReceived
		Else
			strReceived = CStr(Nothing)
		End If
		strOutput = strOutput & "<spliter>"
		
		If Not objLineNode.selectSingleNode("IsManuallyClosed") Is Nothing Then
			'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If objLineNode.selectSingleNode("IsManuallyClosed").Text = "true" Or (Not IsNothing(strOrdered) And strOrdered = strReceived) Or (strOrdered = "0") Then
				strOutput = strOutput & "X"
			End If
		End If
		strOutput = strOutput & "<spliter>"
		
		strOutput = strOutput & objLineNode.selectSingleNode("TxnLineID").Text & "<spliter>"
		
		'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If IsNothing(strLineType) Then
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
			strOutput = strOutput & objLineNode.selectSingleNode("IsManuallyClosed").Text
		End If
		strOutput = strOutput & "<spliter>"
		
		'UPGRADE_WARNING: Couldn't resolve default property of object POLineString. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		POLineString = strOutput
	End Function
	
	Public Sub QBXMLRP_ClosePO(ByRef strTxnID As String, ByRef strEditSequence As String)
		
		Dim strXMLRequest As String
		strXMLRequest = "<?xml version=""1.0"" ?>" & "<?qbxml version=""3.0""?>" & "<QBXML><QBXMLMsgsRq onError = ""stopOnError"">" & "<PurchaseOrderModRq><PurchaseOrderMod>" & "<TxnID>" & strTxnID & "</TxnID>" & "<EditSequence>" & strEditSequence & "</EditSequence>" & "<IsManuallyClosed>true</IsManuallyClosed>" & "</PurchaseOrderMod></PurchaseOrderModRq>" & "</QBXMLMsgsRq></QBXML>"
		
		Dim strXMLResponse As String
		strXMLResponse = objRequestProcessor.ProcessRequest(strTicket, strXMLRequest)

        Dim objXMLDoc As New MSXML2.DOMDocument60

        objXMLDoc.async = False
		objXMLDoc.loadXML(strXMLResponse)
		
		Dim objReportResponseNodeList As MSXML2.IXMLDOMNodeList
		objReportResponseNodeList = objXMLDoc.getElementsByTagName("PurchaseOrderModRs")
		
		Dim objReportResponseNode As MSXML2.IXMLDOMNode
		objReportResponseNode = objReportResponseNodeList.Item(0)
		
		Dim objReportResponseAttributes As MSXML2.IXMLDOMNamedNodeMap
		objReportResponseAttributes = objReportResponseNode.Attributes
		
		If objReportResponseAttributes.getNamedItem("statusCode").nodeValue <> "0" Then
			MsgBox("Error closing purchase order" & vbCrLf & vbCrLf & "Error = " & objReportResponseAttributes.getNamedItem("statusCode").nodeValue & vbCrLf & vbCrLf & "Message = " & objReportResponseAttributes.getNamedItem("statusMessage").nodeValue)
		Else
			MsgBox("Purchase order successfully closed")
		End If
	End Sub
	
	
	Public Sub QBXMLRP_ChangePOLine(ByRef strAction As String, ByRef strTxnID As String, ByRef strEditSequence As String, ByRef strPOLines() As String, ByRef intSelectedPOLine As Short, ByRef intGroupLine As Short)
		
		Dim strXMLRequest As String
		strXMLRequest = "<?xml version=""1.0"" ?>" & "<?qbxml version=""3.0""?>" & "<QBXML><QBXMLMsgsRq onError = ""stopOnError"">" & "<PurchaseOrderModRq><PurchaseOrderMod>" & "<TxnID>" & strTxnID & "</TxnID>" & "<EditSequence>" & strEditSequence & "</EditSequence>"
		
		Dim Splits() As String
		
		Dim i As Short
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
				strXMLRequest = strXMLRequest & "<PurchaseOrderLineGroupMod>" & "<TxnLineID>" & Splits(8) & "</TxnLineID>"
				booInGroup = True
			ElseIf (Splits(9) = "GroupSubItem" And (booInSelectedGroup Or strAction = "ReceiveAll")) Or Splits(9) = "Item" Then 
				strXMLRequest = strXMLRequest & "<PurchaseOrderLineMod>" & "<TxnLineID>" & Splits(8) & "</TxnLineID>"
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
					strXMLRequest = strXMLRequest & "<IsManuallyClosed>true</IsManuallyClosed>"
					'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				ElseIf Not IsNothing(Splits(2)) Then 
					strXMLRequest = strXMLRequest & "<>" & Splits(2) & "</>"
				End If
			End If
			
			If (Splits(9) = "GroupSubItem" And booInSelectedGroup) Or Splits(9) = "Item" Then
				strXMLRequest = strXMLRequest & "</PurchaseOrderLineMod>"
			End If
			
			If i = UBound(strPOLines) Then
				booDone = True
				'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			ElseIf IsNothing(strPOLines(i + 1)) Then 
				booDone = True
				If booInGroup Then
					strXMLRequest = strXMLRequest & "</PurchaseOrderLineGroupMod>"
				End If
			End If
			
			i = i + 1
		Loop Until booDone
		
		strXMLRequest = strXMLRequest & "</PurchaseOrderMod></PurchaseOrderModRq>" & "</QBXMLMsgsRq></QBXML>"
		
		'  PrettyPrintXMLToFile strXMLRequest, "C:\Temp\PurchaseOrderModify.xml"
		'  EndSessionCloseConnection
		'  End
		
		Dim strXMLResponse As String
		strXMLResponse = objRequestProcessor.ProcessRequest(strTicket, strXMLRequest)

        Dim objXMLDoc As New MSXML2.DOMDocument60

        objXMLDoc.async = False
		objXMLDoc.loadXML(strXMLResponse)
		
		Dim objReportResponseNodeList As MSXML2.IXMLDOMNodeList
		objReportResponseNodeList = objXMLDoc.getElementsByTagName("PurchaseOrderModRs")
		
		Dim objReportResponseNode As MSXML2.IXMLDOMNode
		objReportResponseNode = objReportResponseNodeList.Item(0)
		
		Dim objReportResponseAttributes As MSXML2.IXMLDOMNamedNodeMap
		objReportResponseAttributes = objReportResponseNode.Attributes
		
		If objReportResponseAttributes.getNamedItem("statusCode").nodeValue <> "0" Then
			MsgBox("Error getting PO list" & vbCrLf & vbCrLf & "Error = " & objReportResponseAttributes.getNamedItem("statusCode").nodeValue & vbCrLf & vbCrLf & "Message = " & objReportResponseAttributes.getNamedItem("statusMessage").nodeValue)
		Else
			MsgBox("Purchase Order Line successfully closed")
		End If
	End Sub
	
	
	Public Sub QBXMLRP_SetQuantitiesAndBillForRemainingItems(ByRef strTxnID As String, ByRef strEditSequence As String, ByRef strVendor As String, ByRef strRefNumber As String, ByRef strTxnDate As String, ByRef strPOLines() As String, ByRef intPOLine As Short)

        Dim objDOMDoc As New MSXML2.DOMDocument60
        Dim objRootNode As MSXML2.IXMLDOMNode
		Dim objRequestNode As MSXML2.IXMLDOMNode
		Dim objModNode As MSXML2.IXMLDOMNode
		Dim objNode As MSXML2.IXMLDOMNode
		Dim objElement As MSXML2.IXMLDOMElement
		Dim objAttribute As MSXML2.IXMLDOMAttribute
		
		'We're creating a standard request message set node setting the
		'newMessageSetID because this includes a modify and an add request and
		'setting the onError attribute to stopOnError so we don't create a bill
		'if our purchase order modify fails
		
		'UPGRADE_WARNING: Couldn't resolve default property of object objDOMDoc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CreateStandardRequestNode(True, "stopOnError", objDOMDoc, objRootNode, objRequestNode, objAttribute)
		
		'Add the PurchaseOrderModRq and PurchaseOrderMod nodes
		'UPGRADE_WARNING: Couldn't resolve default property of object objDOMDoc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		AddMSXMLNode("PurchaseOrderModRq", objDOMDoc, objRequestNode, objNode)
		'UPGRADE_WARNING: Couldn't resolve default property of object objDOMDoc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		AddMSXMLAttribute("requestID", "0", objDOMDoc, objNode, objAttribute)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object objDOMDoc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		AddMSXMLNode("PurchaseOrderMod", objDOMDoc, objNode, objModNode)
		
		'Add the TxnID and EditSequence of the purchase order we're modifying
		'UPGRADE_WARNING: Couldn't resolve default property of object objDOMDoc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		AddMSXMLElement("TxnID", strTxnID, objDOMDoc, objModNode, objElement)
		'UPGRADE_WARNING: Couldn't resolve default property of object objDOMDoc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		AddMSXMLElement("EditSequence", strEditSequence, objDOMDoc, objModNode, objElement)
		
		'Loop through the PO lines and change the ordered quantity to match the
		'received quantity
		Dim i As Short
		Dim booDone As Boolean
		Dim booPOModified As Boolean
		Dim objParentNode As MSXML2.IXMLDOMNode
		Dim strSplits() As String
		
		booPOModified = False
		objParentNode = objModNode
		i = 1
		Do 
			strSplits = Split(strPOLines(i), "<spliter>")
			
			'The quantity ordered or received can be empty, set then to zero if they're empty
			If strSplits(6) = "" Then strSplits(6) = "0"
			If strSplits(2) = "" Then strSplits(2) = "0"
			
			If strSplits(9) <> "GroupSubItem" Then
				objParentNode = objModNode
			End If
			
			If strSplits(9) = "GroupItem" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object objDOMDoc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				AddMSXMLNode("PurchaseOrderLineGroupMod", objDOMDoc, objModNode, objNode)
				'UPGRADE_WARNING: Couldn't resolve default property of object objDOMDoc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				AddMSXMLElement("TxnLineID", strSplits(8), objDOMDoc, objNode, objElement)
				objParentNode = objNode
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object objDOMDoc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				AddMSXMLNode("PurchaseOrderLineMod", objDOMDoc, objParentNode, objNode)
				'UPGRADE_WARNING: Couldn't resolve default property of object objDOMDoc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				AddMSXMLElement("TxnLineID", strSplits(8), objDOMDoc, objNode, objElement)
				
				'If the PO line hasn't been manually closed and the number of items
				'received is less than the number ordered change the number ordered
				'to the number received
				If strSplits(7) <> "X" And Int(CDbl(strSplits(6))) < Int(CDbl(strSplits(2))) And ((intPOLine = 0) Or (intPOLine = i)) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object objDOMDoc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddMSXMLElement("Quantity", strSplits(6), objDOMDoc, objNode, objElement)
					booPOModified = True
				End If
			End If
			
			If i = UBound(strPOLines) Then
				booDone = True
				'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			ElseIf IsNothing(strPOLines(i + 1)) Then 
				booDone = True
			Else
				i = i + 1
			End If
		Loop Until booDone
		
		If Not booPOModified Then
			MsgBox("The purchase order selected has either has all lines " & "manually closed or all lines received.  Ending processing " & "on the existing purchase order and not creating a new bill " & "for missing items.")
			Exit Sub
		End If
		
		'Now add the bill for the items we modified in the purchase order
		'UPGRADE_WARNING: Couldn't resolve default property of object objDOMDoc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		AddMSXMLNode("BillAddRq", objDOMDoc, objRequestNode, objNode)
		'UPGRADE_WARNING: Couldn't resolve default property of object objDOMDoc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		AddMSXMLAttribute("requestID", "1", objDOMDoc, objNode, objAttribute)
		
		Dim objBillAddNode As MSXML2.IXMLDOMNode
		'UPGRADE_WARNING: Couldn't resolve default property of object objDOMDoc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		AddMSXMLNode("BillAdd", objDOMDoc, objNode, objBillAddNode)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object objDOMDoc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		AddMSXMLNode("VendorRef", objDOMDoc, objBillAddNode, objNode)
		'UPGRADE_WARNING: Couldn't resolve default property of object objDOMDoc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		AddMSXMLElement("FullName", strVendor, objDOMDoc, objNode, objElement)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object objDOMDoc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		AddMSXMLElement("Memo", "Created by Purchase Order Modify Sample for modified PO " & strRefNumber & "; TxnDate " & strTxnDate & "; TxnID " & strTxnID & " .  Quantity ordered " & "reduced to quantity recieved for one or all lines, differences " & "reflected in this bill.", objDOMDoc, objBillAddNode, objElement)
		
		Dim objItemLineNode As MSXML2.IXMLDOMNode
		Dim j As Short
		For j = 1 To i
			strSplits = Split(strPOLines(j), "<spliter>")
			
			'The quantity ordered or received can be empty, set then to zero if they're empty
			If strSplits(6) = "" Then strSplits(6) = "0"
			If strSplits(2) = "" Then strSplits(2) = "0"
			
			If strSplits(9) <> "GroupItem" And strSplits(7) <> "X" And Int(CDbl(strSplits(6))) < Int(CDbl(strSplits(2))) And ((intPOLine = 0) Or (intPOLine = j)) Then
				
				'UPGRADE_WARNING: Couldn't resolve default property of object objDOMDoc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				AddMSXMLNode("ItemLineAdd", objDOMDoc, objBillAddNode, objItemLineNode)
				
				'UPGRADE_WARNING: Couldn't resolve default property of object objDOMDoc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				AddMSXMLNode("ItemRef", objDOMDoc, objItemLineNode, objNode)
				'UPGRADE_WARNING: Couldn't resolve default property of object objDOMDoc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				AddMSXMLElement("FullName", strSplits(0), objDOMDoc, objNode, objElement)
				
				'UPGRADE_WARNING: Couldn't resolve default property of object objDOMDoc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				AddMSXMLElement("Quantity", Str(Int(CDbl(strSplits(2))) - Int(CDbl(strSplits(6)))), objDOMDoc, objItemLineNode, objElement)

                'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                If Not String.IsNullOrEmpty(strSplits(3)) Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object objDOMDoc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    AddMSXMLElement("Cost", strSplits(3), objDOMDoc, objItemLineNode, objElement)
                    'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                ElseIf Not String.IsNullOrEmpty(strSplits(5)) Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object objDOMDoc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    AddMSXMLElement("Amount", strSplits(5), objDOMDoc, objItemLineNode, objElement)
				End If

                'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                If Not String.IsNullOrEmpty(strSplits(4)) Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object objDOMDoc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    AddMSXMLNode("CustomerRef", objDOMDoc, objItemLineNode, objNode)
                    'UPGRADE_WARNING: Couldn't resolve default property of object objDOMDoc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    AddMSXMLElement("FullName", strSplits(4), objDOMDoc, objNode, objElement)
                End If
            End If
		Next j
		
		Dim strXMLRequest As String
		strXMLRequest = "<?xml version=""1.0""?>" & "<?qbxml version=""3.0""?>" & objRootNode.xml
		Dim fnum As Short
		fnum = FreeFile
		
		Dim strXMLResponse As String
		strXMLResponse = objRequestProcessor.ProcessRequest(strTicket, strXMLRequest)

        Dim objXMLDoc As New MSXML2.DOMDocument60

        objXMLDoc.async = False
		objXMLDoc.loadXML(strXMLResponse)
		
		Dim objResponseNodeList As MSXML2.IXMLDOMNodeList
		objResponseNodeList = objXMLDoc.getElementsByTagName("PurchaseOrderModRs")
		
		Dim objResponseNode As MSXML2.IXMLDOMNode
		objResponseNode = objResponseNodeList.Item(0)
		
		If objResponseNode Is Nothing Then
			MsgBox("Error from QuickBooks processing request: " & strXMLResponse)
		End If
		
		
		Dim objResponseAttributes As MSXML2.IXMLDOMNamedNodeMap
		objResponseAttributes = objResponseNode.Attributes
		
		If objResponseAttributes.getNamedItem("statusCode").nodeValue <> "0" Then
			MsgBox("Error modifying PO" & vbCrLf & vbCrLf & "Error = " & objResponseAttributes.getNamedItem("statusCode").nodeValue & vbCrLf & vbCrLf & "Message = " & objResponseAttributes.getNamedItem("statusMessage").nodeValue)
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
		
		objResponseNodeList = objXMLDoc.getElementsByTagName("BillAddRs")
		objResponseNode = objResponseNodeList.Item(0)
		objResponseAttributes = objResponseNode.Attributes
		
		If objResponseAttributes.getNamedItem("statusCode").nodeValue <> "0" Then
			MsgBox("Error adding Bill" & vbCrLf & vbCrLf & "Error = " & objResponseAttributes.getNamedItem("statusCode").nodeValue & vbCrLf & vbCrLf & "Message = " & objResponseAttributes.getNamedItem("statusMessage").nodeValue)
			Exit Sub
		Else
			If intPOLine = 0 Then
				MsgBox("Successfully set PO lines and billed for items")
			Else
				MsgBox("Successfully set PO line and billed for items")
			End If
		End If
		
		'Now let's update the PO Memo to include the manipulation we performed
		Dim strBillTxnID As String
		Dim strBillRefNumber As String
		Dim strBillTxnDate As String
		
		Dim objBillRet As MSXML2.IXMLDOMNode
		
		objBillRet = objResponseNode.selectSingleNode("BillRet")
		
		strBillTxnID = objBillRet.selectSingleNode("TxnID").Text
		strBillTxnDate = objBillRet.selectSingleNode("TxnDate").Text
		
		If Not objBillRet.selectSingleNode("RefNumber") Is Nothing Then
			strBillRefNumber = objBillRet.selectSingleNode("RefNumber").Text
		End If
		
		strMemo = strMemo & " >>> Purchase order modified by ""Purchase Order Modify Sample"": " & "Quantity ordered changed to quantity received value "
		If intPOLine = 0 Then
			strMemo = strMemo & "on all lines not closed or fully recieved.  "
		Else
			strMemo = strMemo & "on one line.  "
		End If
		strMemo = strMemo & "Created bill " & strBillRefNumber & " with TxnDate " & strBillTxnDate & "; TxnID " & strBillTxnID & " for these items."
		
		strXMLRequest = "<?xml version=""1.0"" ?>" & "<?qbxml version=""3.0""?>" & "<QBXML><QBXMLMsgsRq onError = ""stopOnError"">" & "<PurchaseOrderModRq><PurchaseOrderMod>" & "<TxnID>" & strTxnID & "</TxnID>" & "<EditSequence>" & strNewEditSequence & "</EditSequence>" & "<Memo>" & strMemo & "</Memo>" & "</PurchaseOrderMod></PurchaseOrderModRq>" & "</QBXMLMsgsRq></QBXML>"
		
		strXMLResponse = objRequestProcessor.ProcessRequest(strTicket, strXMLRequest)
		
		objXMLDoc.async = False
		objXMLDoc.loadXML(strXMLResponse)
		
		objResponseNodeList = objXMLDoc.getElementsByTagName("PurchaseOrderModRs")
		objResponseNode = objResponseNodeList.Item(0)
		objResponseAttributes = objResponseNode.Attributes
		
		If objResponseAttributes.getNamedItem("statusCode").nodeValue <> "0" Then
			MsgBox("Error modifying PO Memo field" & vbCrLf & vbCrLf & "Error = " & objResponseAttributes.getNamedItem("statusCode").nodeValue & vbCrLf & vbCrLf & "Message = " & objResponseAttributes.getNamedItem("statusMessage").nodeValue)
		End If
	End Sub
End Module