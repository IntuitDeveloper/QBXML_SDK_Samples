Option Strict Off
Option Explicit On
Friend Class CustomerList
	Inherits System.Windows.Forms.Form
	'-----------------------------------------------------------
	' Form Module: CustomerList
	'
	' Description:  This form allows the user to select a customer to
	'               be modified.
	'
	' Created On: 11/08/2001
	' Updated to SDK 2.0: 08/05/2002
	'
	' Copyright © 2021-2022 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	'----------------------------------------------------------
	
	
	'Selected Customer Obj
	Dim selectedCust As Customer
	
	
	' Submit button
	Private Sub Comm_Submit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Comm_Submit.Click
		
		On Error GoTo ErrHandler
		
		' Collect Form Data
		Dim CustModForm As New CustomerMod
		If CollectFormData Then
			
			' Load CustomerMod Form
			
			' Pass the selected Customer to CustomerMod Form
			CustModForm.Customer(selectedCust)
			
			' Set up the Form Caption with this customer's Name
			CustModForm.Text = "qbXML Sample: Customer " & selectedCust.Name & " Information"
			
			' Show the CustomerMod form
			VB6.ShowForm(CustModForm, VB6.FormShowConstants.Modal, Me)
			
			' Close this form to go back to the main form
			Me.Close()
			
		End If
		
		Exit Sub
		
ErrHandler: 
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
		Exit Sub
		
	End Sub
	
	' Display request XML
	Private Sub Comm_View_Req_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Comm_View_Req.Click
		
		On Error GoTo ErrHandler
		
		' Instantiate a RequestXML Obj
		Dim reqFrm As Display
		reqFrm = New Display
		
		reqFrm.Text_Content.Text = Main_Renamed.GetRequestXML
		reqFrm.Text = "Request XML"
		VB6.ShowForm(reqFrm, VB6.FormShowConstants.Modal, Me)
		
		Exit Sub
		
ErrHandler: 
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
		Exit Sub
		
	End Sub
	
	' Display response XML
	Private Sub Comm_View_Res_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Comm_View_Res.Click
		
		On Error GoTo ErrHandler
		
		Dim resFrm As New Display
		
		Dim tmpRes As String
		If Main_Renamed.GetResponseXML <> "" Then
			tmpRes = Replace(Main_Renamed.GetResponseXML, vbLf, vbCrLf)
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
	
	
	' Form Data Collection
	Private Function CollectFormData() As Boolean
		
		On Error GoTo ErrHandler
		
		' Customer Name
		Dim custFullName As String
		
		custFullName = Combo_CustList.Text
		If custFullName = "" Then
			MsgBox("Please select a customer")
			CollectFormData = False
			Exit Function
		End If
		
		' Set the customer obj
		selectedCust = Main_Renamed.customerCollection.Item(custFullName)
		
		CollectFormData = True
		Exit Function
		
ErrHandler: 
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
		CollectFormData = False
		Exit Function
		
	End Function
	
	
	Private Sub CustomerList_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		On Error GoTo ErrHandler
		
		' Load the image
		Dim appPath As Object
        appPath = My.Application.Info.DirectoryPath
        Image_QBBANNER.Image = System.Drawing.Image.FromFile(appPath & "/qbbanner.bmp")

        ' Initialize the customer combo list
        Dim cust As Customer
		For	Each cust In Main_Renamed.customerCollection
			Combo_CustList.Items.Add(cust.FullName)
		Next cust
		
		Exit Sub
		
ErrHandler: 
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
		Exit Sub
		
	End Sub
	
	
	Public ReadOnly Property SelectedCustomer() As Customer
		Get
			SelectedCustomer = selectedCust
		End Get
	End Property
End Class