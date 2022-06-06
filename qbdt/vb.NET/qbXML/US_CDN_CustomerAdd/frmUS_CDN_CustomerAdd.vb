Option Strict Off
Option Explicit On
Friend Class frmUS_CDN_CustomerAdd
	Inherits System.Windows.Forms.Form
	'-----------------------------------------------------------
	' Form: frmUS_CDN_CustomerAdd
	'
	' Description: Displays the nationality of the QuickBooks version
	'              connected to and allows the input of information for
	'              adding a customer to the currently open company file
	'
	' Copyright © 2021-2022 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	'----------------------------------------------------------
	Dim QBNationality As String
	
	Private Sub cmdAddCustomer_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAddCustomer.Click
		'Check to make sure that the Addr fields are filled in properly
		If txtAddr1.Text = "" Then
			If txtAddr2.Text <> "" Then
				MsgBox("The address fields must be filled in order.  Addr2 must be blank if Addr1 isn't filled")
				Exit Sub
			Else
				If txtAddr3.Text <> "" Then
					MsgBox("The address fields must be filled in order.  Addr3 must be blank if Addr1 isn't filled")
					Exit Sub
				Else
					If txtAddr4.Text <> "" Then
						MsgBox("The address fields must be filled in order.  Addr4 must be blank if Addr1 isn't filled")
						Exit Sub
					End If
				End If
			End If
		Else
			If txtAddr2.Text = "" Then
				If txtAddr3.Text <> "" Then
					MsgBox("The address fields must be filled in order.  Addr3 must be blank if Addr2 isn't filled")
					Exit Sub
				Else
					If txtAddr4.Text <> "" Then
						MsgBox("The address fields must be filled in order.  Addr4 must be blank if Addr2 isn't filled")
						Exit Sub
					End If
				End If
			Else
				If txtAddr3.Text = "" And txtAddr4.Text <> "" Then
					MsgBox("The address fields must be filled in order.  Addr4 must be blank if Addr3 isn't filled")
					Exit Sub
				End If
			End If
		End If
		
		If AddCustomer(QBNationality, (txtCustomer.Text), (txtAddr1.Text), (txtAddr2.Text), (txtAddr3.Text), (txtAddr4.Text), (txtCity.Text), (txtStateProvince.Text), (txtPostalCode.Text)) Then
			ClearForm()
		End If
	End Sub
	
	Private Sub cmdClearForm_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClearForm.Click
		ClearForm()
	End Sub
	
	Private Sub cmdExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdExit.Click
		EndQuickBooksSession()
		End
	End Sub
	
	Private Sub frmUS_CDN_CustomerAdd_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		StartQuickBooksSession()
		QBNationality = GetQBNationality
		
		lblNationality.Text = "Running against " & QBNationality & " QuickBooks"
		
		If QBNationality = "US" Then
			lblStateProvince.Text = "State"
		Else
			lblStateProvince.Text = "Province"
		End If
	End Sub
	
	Private Sub ClearForm()
		txtCustomer.Text = ""
		txtAddr1.Text = ""
		txtAddr2.Text = ""
		txtAddr3.Text = ""
		txtAddr4.Text = ""
		txtCity.Text = ""
		txtStateProvince.Text = ""
		txtPostalCode.Text = ""
	End Sub
End Class