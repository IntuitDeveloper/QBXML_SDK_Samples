Option Strict Off
Option Explicit On
Friend Class AddCust
	Inherits System.Windows.Forms.Form
	'-----------------------------------------------------------
	' Form Module: AddCust
	'
	' Description:  This sample application demonstrates how to
	'               construct a qbfc request, send it to QuickBooks,
	'               and parse the qbfc response.  It prompts the
	'               user for customer information and adds a new
	'               customer to QuickBooks.
	'
	'               QuickBooks must be running with a data file open.
	'               The current data file is used.
	'
	' Copyright © 2002-2020 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/rdmgr/?ID=100
	'
	'----------------------------------------------------------
	
	
	' Request and response strings
	Public requestXML As String
	Public responseXML As String
	
	' Customer info
	Public firstName As String
	Public lastName As String
	Public phoneNumber As String
	Public custName As String
	Public Addr1 As String
	Public Addr2 As String
	Public Addr3 As String
	Public Addr4 As String
	Public City As String
	Public State As String
	Public PostalCode As String
	Public Country As String
	Public strDataExtName As String
	Public strDataExtValue As String
	
	' Submit button
	Private Sub Comm_Submit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Comm_Submit.Click
		
		On Error GoTo ErrHandler
		
		' Initialize
		requestXML = ""
		responseXML = ""
		OpenConnectionBeginSession()
		' Get input data
		If Not CollectFormData Then
			Exit Sub
		End If
		
		' Send request to QuickBooks
		SendCustomerAddRequest()
		EndSessionCloseConnection()
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
	
	' Form Data Collection
	'
	Private Function CollectFormData() As Boolean
		
		On Error GoTo ErrHandler
		
		' Get data from the form
		firstName = Text_FirstName.Text
		lastName = Text_LastName.Text
		phoneNumber = Text_Phone.Text
		custName = Text_CustomerName.Text
		Addr1 = Text_Addr1.Text
		Addr2 = Text_Addr2.Text
		Addr3 = Text_Addr3.Text
		Addr4 = Text_Addr4.Text
		City = Text_City.Text
		State = Text_State.Text
		PostalCode = Text_PostalCode.Text
		Country = Text_Country.Text
		strDataExtName = Text_DataExtName.Text
		strDataExtValue = Text_DataExtValue.Text
		
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
End Class