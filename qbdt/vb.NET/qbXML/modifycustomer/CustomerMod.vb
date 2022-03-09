Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class CustomerMod
	Inherits System.Windows.Forms.Form
	'-----------------------------------------------------------
	' Form Module: CustomerMod
	'
	' Description:  This form allows the user to modify customer
	'               information.
	'
	' Created On: 11/08/2001
	' Updated to SDK 2.0: 08/05/2002
	'
	' Copyright © 2021-2022 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	'----------------------------------------------------------
	
	
	Dim requestXML As String
	Dim responseXML As String
	
	Dim currEditSeq As String ' Current EditSequesce
	Dim currName As String ' Current Customer Name
	Dim currFullName As String ' Current FullName
	Dim currListID As String ' Current List ID
	Dim currTimeModified As String ' Current Modified Time
	
	
	' Existing Customer object
	Dim existCustObj As Customer
	
	' User entered data
	Dim modifiedCustObj As Customer
	
	
	
	Private Sub Comm_Submit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Comm_Submit.Click
		
		On Error GoTo ErrHandler
		
		requestXML = ""
		responseXML = ""
		
		' Get data
		If Not CollectFormData Then
			Exit Sub
		End If
		
		' Build XML
		BuildXML()
		
		' Send request to QuickBooks
		DoRequest()
		
		If Not ParseResponseXML Then
			Exit Sub
		End If
		
		' Repaint the form
		Me.Text = "qbXML Sample: Customer " & currName & " Information"
		
		L_FullName.Text = currFullName
		L_ListID.Text = currListID
		L_ModifiedTime.Text = TimeFormat(currTimeModified)
		
		MsgBox("Customer  " & currFullName & " has been successfully modified in QuickBooks.")
		
		Exit Sub
		
ErrHandler: 
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
		Exit Sub
	End Sub
	
	Private Sub Comm_View_Req_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Comm_View_Req.Click
		
		On Error GoTo ErrHandler
		
		Dim reqFrm As New Display
		
		reqFrm.Text_Content.Text = requestXML
		reqFrm.Text = "Request XML"
		VB6.ShowForm(reqFrm, VB6.FormShowConstants.Modal, Me)
		
		Exit Sub
		
ErrHandler: 
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
		Exit Sub
		
	End Sub
	
	Private Sub Comm_View_Res_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Comm_View_Res.Click
		
		On Error GoTo ErrHandler
		
		Dim resFrm As New Display
		
		Dim tmpRes As String
		If responseXML <> "" Then
			tmpRes = Replace(responseXML, vbLf, vbCrLf)
			resFrm.Text_Content.Text = tmpRes
		Else
			resFrm.Text_Content.Text = ""
		End If
		
		resFrm.Text = "Response XML"
		VB6.ShowForm(resFrm, VB6.FormShowConstants.Modal, Me)
		
		Exit Sub
		
ErrHandler: 
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
		Exit Sub
		
	End Sub
	
	Private Sub Comm_Exit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Comm_Exit.Click
		Me.Close()
	End Sub
	
	
	' Form initialization
	Private Sub CustomerMod_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		On Error GoTo ErrHandler
		
		' allocate objs
		modifiedCustObj = New Customer
		
		With existCustObj
			
			' Set the default values
			
			' Label fields
			L_FullName.Text = .FullName
			L_ListID.Text = .ListID
			L_CreatedTime.Text = TimeFormat(.TimeCreated)
			L_ModifiedTime.Text = TimeFormat(.TimeModified)
			
			' Text fields
			Text_Phone.Text = .Phone
			Text_LastName.Text = .LastName
			Text_FirstName.Text = .FirstName
			Text_CompanyName.Text = .CompanyName
			Text_Name.Text = .Name
			
			' Set the modifiedCustObj
			modifiedCustObj.ListID = .ListID
			modifiedCustObj.FullName = .FullName
			modifiedCustObj.Name = .Name
			modifiedCustObj.CompanyName = .CompanyName
			modifiedCustObj.EditSequence = .EditSequence
			
			currEditSeq = .EditSequence
			
		End With
		
		' Load the banner
		Dim appPath As Object
        appPath = My.Application.Info.DirectoryPath
        Image_QBBANNER.Image = System.Drawing.Image.FromFile(appPath & "/qbbanner.bmp")

        Exit Sub
		
ErrHandler: 
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
		Exit Sub
		
	End Sub
	
	
	'----------------------------------------------------------
	'Build XML request
	'----------------------------------------------------------
	
	Private Sub BuildXML()

        Dim doc As New MSXML2.DOMDocument60
        Dim QBXML As MSXML2.IXMLDOMNode
		Dim MsgsRq As MSXML2.IXMLDOMElement
        Dim CustomerMod_Renamed As MSXML2.IXMLDOMElement

        '
        ' Set up the basic request frame
        '
        QBXML = doc.appendChild(doc.createElement("QBXML"))
        MsgsRq = QBXML.appendChild(doc.createElement("QBXMLMsgsRq"))
        MsgsRq.setAttribute("onError", "continueOnError")
        '
        ' Note that after we create the CustomerModRq we don't need to reference it other
        ' than to add in the CustomerMod element, so we re-use the CustomerMod var
        '
        CustomerMod_Renamed = MsgsRq.appendChild(doc.createElement("CustomerModRq"))
        CustomerMod_Renamed = CustomerMod_Renamed.appendChild(doc.createElement("CustomerMod"))

        '
        ' Now we just need to stuff in the data we are changing.
        ' Note again that we just re-use the dataelement since everything is getting
        ' appended into the structure in the right place.  Order IS important in
        ' qbXML.
        '
        Dim DataElement As MSXML2.IXMLDOMElement
		With modifiedCustObj
            DataElement = CustomerMod_Renamed.appendChild(doc.createElement("ListID"))
            DataElement.appendChild(doc.createTextNode(.ListID))
            DataElement = CustomerMod_Renamed.appendChild(doc.createElement("EditSequence"))
            DataElement.appendChild(doc.createTextNode(currEditSeq))
            DataElement = CustomerMod_Renamed.appendChild(doc.createElement("Name"))
            DataElement.appendChild(doc.createTextNode(.Name))
            DataElement = CustomerMod_Renamed.appendChild(doc.createElement("CompanyName"))
            DataElement.appendChild(doc.createTextNode(.CompanyName))
            DataElement = CustomerMod_Renamed.appendChild(doc.createElement("FirstName"))
            DataElement.appendChild(doc.createTextNode(.FirstName))
            DataElement = CustomerMod_Renamed.appendChild(doc.createElement("LastName"))
            DataElement.appendChild(doc.createTextNode(.LastName))
            DataElement = CustomerMod_Renamed.appendChild(doc.createElement("Phone"))
            DataElement.appendChild(doc.createTextNode(.Phone))
        End With
		
		requestXML = doc.xml
		
	End Sub
	
	' Send request to QuickBooks
	' We use the existing session that was initiated in Main
	Private Sub DoRequest()
		
		Dim success As Boolean
		Dim errNumber As Integer
		Dim errMsg As String
		
		success = Main_Renamed.rpWrapper.DoRequest(requestXML, responseXML)
		If Not success Then
			Main_Renamed.rpWrapper.GetErrorInfo(errNumber, errMsg)
			Err.Raise(Number:=errNumber, Description:=errMsg)
		End If
		
	End Sub
	
	' Parse the response XML
	Private Function ParseResponseXML() As Boolean
		
		On Error GoTo ErrHandler

        Dim xmlDoc As New MSXML2.DOMDocument60
        Dim objNodeList As MSXML2.IXMLDOMNodeList
		Dim objChild As MSXML2.IXMLDOMNode
		Dim custChildNode As MSXML2.IXMLDOMNode
		
		xmlDoc.async = False
		
		' Attributes Name Mapping
		Dim attrNamedNodeMap As MSXML2.IXMLDOMNamedNodeMap
		
		' QBXML values
		Dim retStatusCode As String
		Dim retStatusMessage As String
		Dim retStatusSeverity As String
		
		Dim i As Short
		Dim ret As Boolean
		Dim errorMsg As String
		
		errorMsg = ""
		
		' Load xml Doc
		ret = xmlDoc.loadXML(responseXML)
		
		If Not ret Then
			errorMsg = "loadXML failed, reason: " & xmlDoc.parseError.reason
			GoTo ErrHandler
		End If
		
		objNodeList = xmlDoc.getElementsByTagName("CustomerModRs")
		
		' The following loop is actually unnecessary for this case.
		' We have only one request, so we should only have one response.
		For i = 0 To (objNodeList.length - 1)
			
			attrNamedNodeMap = objNodeList.Item(i).attributes

            retStatusCode = attrNamedNodeMap.getNamedItem("statusCode").nodeValue
            retStatusSeverity = attrNamedNodeMap.getNamedItem("statusSeverity").nodeValue
            retStatusMessage = attrNamedNodeMap.getNamedItem("statusMessage").nodeValue

            ' Check status code to see if there is error or warning
            If retStatusCode <> "0" Then
				If retStatusSeverity = "Warning" Then
					' Show the warning, then continue normal processing
					MsgBox(retStatusMessage, MsgBoxStyle.Exclamation, "Warning from QuickBooks")
				ElseIf retStatusSeverity = "Error" Then 
					MsgBox(retStatusMessage, MsgBoxStyle.Exclamation, "Error from QuickBooks")
					' We have only one response thus we will exit.  If we have multiple
					' responses, then we need to continue looping.
					ParseResponseXML = False
					Exit Function
				End If
			End If
			
			
			' Walk through the child nodes to get the detail Customer info
			For	Each objChild In objNodeList.Item(i).childNodes
				
				' Get the CustomerRet block
				If objChild.nodeName = "CustomerRet" Then
					
					For	Each custChildNode In objChild.childNodes
						If custChildNode.nodeName = "EditSequence" Then
							currEditSeq = custChildNode.Text
						ElseIf custChildNode.nodeName = "FullName" Then 
							currFullName = custChildNode.Text
						ElseIf custChildNode.nodeName = "TimeModified" Then 
							currTimeModified = custChildNode.Text
						ElseIf custChildNode.nodeName = "ListID" Then 
							currListID = custChildNode.Text
						ElseIf custChildNode.nodeName = "Name" Then 
							currName = custChildNode.Text
						End If
						
					Next custChildNode
					
				End If
				
			Next objChild ' end of CustomerRet loop
		Next  ' end of CustomerModRs loop
		
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
	
	
	Private Function CollectFormData() As Boolean
		
		On Error GoTo ErrHandler
		
		' Get data from the form
		With modifiedCustObj
			.Name = Text_Name.Text
			.Phone = Text_Phone.Text
			.LastName = Text_LastName.Text
			.FirstName = Text_FirstName.Text
			.CompanyName = Text_CompanyName.Text
		End With
		
		CollectFormData = True
		Exit Function
		
ErrHandler: 
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
		CollectFormData = False
		Exit Function
		
	End Function
	
	Public Sub Customer(ByRef cust As Customer)
		existCustObj = cust
	End Sub
	
	
	' Convert qbXML time format to a more a readable format for display purpose
	Private Function TimeFormat(ByRef inTime As String) As String
		Dim outTime As String
		Dim ymdStr As String
		Dim hmsStr As String
		Dim tmp As String
		If inTime = "" Or Len(inTime) < 20 Then
			' Unknow format
			outTime = ""
		Else
			
			tmp = VB.Left(inTime, 19)
			ymdStr = VB.Left(tmp, 10)
			hmsStr = VB.Right(tmp, 8)
			
			outTime = ymdStr & " " & hmsStr
		End If
		TimeFormat = outTime
		
	End Function

    ' Clean up
    Private Sub Form_Terminate_Renamed()
        modifiedCustObj = Nothing
    End Sub
End Class