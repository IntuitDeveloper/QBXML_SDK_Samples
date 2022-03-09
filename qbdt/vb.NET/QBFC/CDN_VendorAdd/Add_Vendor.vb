Option Strict Off
Option Explicit On
Friend Class AddVendor
	Inherits System.Windows.Forms.Form
	' Add_Vendor.frm
	' This form is part of the Add Vendor sample program
	' for the Canadian QuickBooks SDK Version CA2.0 and US QuickBooks SDK version 2.0.
	'
	' Created October, 2002
	'
	' Copyright © 2021-2022 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/rdmgr/?ID=100
	'
	'-------------------------------------------------------------
	
	
	
	
	Private Sub Combo_Currency_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo_Currency.TextChanged
		'If not in the list, we clear the combo box
		If Combo_Currency.SelectedIndex = -1 Then
			Combo_Currency.Text = ""
		End If
		
	End Sub
	
	
	
	Private Sub Combo_State_Prov_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo_State_Prov.TextChanged
		'If not in the list, we clear the combo box
		If Combo_State_Prov.SelectedIndex = -1 Then
			Combo_State_Prov.Text = ""
		End If
		
	End Sub
	
	
	
	Private Sub Combo_Vendor_Type_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo_Vendor_Type.TextChanged
		'If not in the list, we clear the combo box
		If Combo_Vendor_Type.SelectedIndex = -1 Then
			Combo_Vendor_Type.Text = ""
		End If
		
	End Sub
	
	Private Sub Command_Add_Vendor_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command_Add_Vendor.Click
		'Verify the required fields
		If Text_First_Name.Text = "" And Text_Last_Name.Text = "" Then
			MsgBox("You need at least a first name or a last name",  , "Missing name")
			Text_First_Name.Focus()
		Else
			If strSDKVersion = "US" Then 'Save using the US SDK
				SaveWithUSSDK()
			Else 'Save using the Canadian SDK
				SaveWithCanadaSDK()
			End If
		End If
		
	End Sub
	
	Private Sub Command_Clear_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command_Clear.Click
		'Clear the form
		Text_Salutation.Text = ""
		Text_First_Name.Text = ""
		Text_Last_Name.Text = ""
		Text_Middle_Name.Text = ""
		Text_Address_1.Text = ""
		Text_Address_2.Text = ""
		Text_City.Text = ""
		Combo_State_Prov.Text = ""
		Text_Postal_Code.Text = ""
		Text_Phone.Text = ""
		Text_Alt_Phone.Text = ""
		Text_EMail.Text = ""
		Text_Name_On_Check.Text = ""
		Text_Account_Number.Text = ""
		Combo_Vendor_Type.Text = ""
		Combo_Currency.Text = ""
		
	End Sub
	
	Private Sub Command_Exit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command_Exit.Click
		If blnConnectionIsOpen = True Then 'If the connection is still open
			CloseQBFCConnection() 'The procedure to close the connection is called
		End If
		Me.Close() ' close the app
	End Sub
	
	Private Sub AddVendor_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim intIndex As Object
		On Error GoTo ErrHandler
		
		Dim appPath As Object
		
		appPath = My.Application.Info.DirectoryPath
		
		Image_QBBANNER.Image = System.Drawing.Image.FromFile(appPath & "/qbbanner.bmp")
		
		Dim blnIsOpenConnection As Boolean
		'Find Which SDK version QuickBooks supports.  When I call the OpenConnection, the SDKVersion public
		'variable is set and tell us if the app is connected to a Canadian QuickBooks or a US QuickBooks
		
		blnIsOpenConnection = OpenConnection
		
		Dim strVendorTypeString As String
		Dim VendorTypeStringArray() As String
		Dim nIndex As Short
		Dim intVendorTypeCount As Short
		Dim blnMultiCurrencyOn As Boolean
		Dim strCurrencyString As String
		Dim CurrencyStringArray() As String
		Dim intCurrencyCount As Short 'This String array will contain the list of currency returned from QuickBooks 'This String array will contain the list of Vendor Type returned from QuickBooks
		If blnIsOpenConnection = True Then 'If the connection was established, I populate the combo boxex
			
			'Populate the combo boxes
			If strSDKVersion = "Canada" Then 'List of province abbreviations
				Combo_State_Prov.Items.Add("AB")
				Combo_State_Prov.Items.Add("BC")
				Combo_State_Prov.Items.Add("MB")
				Combo_State_Prov.Items.Add("NB")
				Combo_State_Prov.Items.Add("NF")
				Combo_State_Prov.Items.Add("NS")
				Combo_State_Prov.Items.Add("NT")
				Combo_State_Prov.Items.Add("NU")
				Combo_State_Prov.Items.Add("ON")
				Combo_State_Prov.Items.Add("PE")
				Combo_State_Prov.Items.Add("QC")
				Combo_State_Prov.Items.Add("SK")
				Combo_State_Prov.Items.Add("YT")
				
				
			Else 'Partial list of the US states
				Combo_State_Prov.Items.Add("AL")
				Combo_State_Prov.Items.Add("AK")
				Combo_State_Prov.Items.Add("AZ")
				Combo_State_Prov.Items.Add("AR")
				Combo_State_Prov.Items.Add("CA")
				Combo_State_Prov.Items.Add("CO")
				Combo_State_Prov.Items.Add("DC")
				Combo_State_Prov.Items.Add("FL")
				Combo_State_Prov.Items.Add("GA")
				
				
			End If
			
			'Populate the Vendor Type combo box
			' The GetVendorTypeListCA and GetVendorTypeListUS function returns a list a vendor type that are separated by "***", We use the
			'Split function to separate the customer names in order to populate the vendor type combo box
			
			intIndex = 0
			
			If strSDKVersion = "US" Then
				strVendorTypeString = GetVendorTypeListUS
			Else
				strVendorTypeString = GetVendorTypeListCA
			End If
			VendorTypeStringArray = Split(strVendorTypeString, "***") 'Split the vendor type list string into an array
			
			intVendorTypeCount = UBound(VendorTypeStringArray)
			
			
			If intVendorTypeCount = 0 Then ' If can't find at least one vendor type, give an error message
				MsgBox("Could not find any Vendor Type in QuickBooks", MsgBoxStyle.Information, "No vendor type found")
			Else
				
				While intIndex < intVendorTypeCount
					'Adding the customer's name to the combo box
					
					Combo_Vendor_Type.Items.Add(VendorTypeStringArray(intIndex))
					
					intIndex = intIndex + 1 'Go to the next customer
				End While
			End If
			
			'Populate the currency list if multicurrency feature is selected
			
			
			intIndex = 0
			
			If strSDKVersion = "Canada" Then 'Verify first that the multicurrency feature has been turned on.  The multicurrency
				'is set in the preferences.  We query the preferences to find if multicurrency is turned on or not.
				
				blnMultiCurrencyOn = FindIfMulticurrencyOn
				'                blnMultiCurrencyOn = True
				If blnMultiCurrencyOn = True Then
					'I make the currency combo box visible and populate it
					Combo_Currency.Visible = True
					Label_Multicurrency.Visible = True
					
					strCurrencyString = GetCurrencyListCA
					CurrencyStringArray = Split(strCurrencyString, "***") 'Split the vendor type list string into an array
					
					intCurrencyCount = UBound(CurrencyStringArray)
					
					
					If intCurrencyCount = 0 Then ' If can't find at least one currency, give an error message
						MsgBox("Could not find any currency in QuickBooks", MsgBoxStyle.Information, "No currency found")
					Else
						
						While intIndex < intCurrencyCount
							'Adding the customer's name to the combo box
							
							Combo_Currency.Items.Add(CurrencyStringArray(intIndex))
							
							intIndex = intIndex + 1 'Go to the next customer
						End While
					End If
				Else
					'Making the multicurrency note field visible
					Text_Note_MultiCurrency.Visible = True
					Label_Note.Visible = True
					
				End If
			End If
			
		Else 'Either the connection could not be established, the connection was refused by user or one or more dlls are missing.  The error message has already been given to the user
			'we just need to shut down the app.
			Me.Close()
		End If
		Exit Sub
		
ErrHandler: 
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
	End Sub
	Sub SaveWithUSSDK()
		On Error GoTo ErrHandler
		
		'Create the message set request object
		Dim requestMsgSet As QBFC15Lib.IMsgSetRequest
		requestMsgSet = sessionManagerUS.CreateMsgSetRequest("US", CShort(cQBXMLMajorVersion), CShort(cQBXMLMinorVersion))
		'Initialize the request's attributes
		requestMsgSet.Attributes.OnError = QBFC15Lib.ENRqOnError.roeContinue
		
		'Add the request to the message set request object
		Dim VendorAdd As QBFC15Lib.IVendorAdd
		VendorAdd = requestMsgSet.AppendVendorAddRq
		
		If Text_Salutation.Text <> "" Then
			VendorAdd.Salutation.SetValue(Text_Salutation.Text)
		End If
		
		If Text_First_Name.Text <> "" And Text_Last_Name.Text <> "" Then
			VendorAdd.Name.SetValue(Text_First_Name.Text & " " & Text_Last_Name.Text)
			VendorAdd.FirstName.SetValue(Text_First_Name.Text)
			VendorAdd.LastName.SetValue(Text_Last_Name.Text)
		ElseIf Text_First_Name.Text <> "" Then 
			VendorAdd.Name.SetValue(Text_First_Name.Text)
			VendorAdd.FirstName.SetValue(Text_First_Name.Text)
		ElseIf Text_Last_Name.Text <> "" Then 
			VendorAdd.Name.SetValue(Text_Last_Name.Text)
			VendorAdd.LastName.SetValue(Text_Last_Name.Text)
		End If
		
		If Text_Middle_Name.Text <> "" Then
			VendorAdd.MiddleName.SetValue(Text_Middle_Name.Text)
		End If
		
		If Text_Address_1.Text <> "" Then
			VendorAdd.VendorAddress.Addr1.SetValue(Text_Address_1.Text)
		End If
		
		If Text_Address_2.Text <> "" Then
			VendorAdd.VendorAddress.Addr2.SetValue(Text_Address_2.Text)
		End If
		If Text_City.Text <> "" Then
			VendorAdd.VendorAddress.City.SetValue(Text_City.Text)
		End If
		If Combo_State_Prov.Text <> "" Then
			VendorAdd.VendorAddress.State.SetValue(Combo_State_Prov.Text)
		End If
		
		If Text_Postal_Code.Text <> "" Then
			VendorAdd.VendorAddress.PostalCode.SetValue(Text_Postal_Code.Text)
		End If
		
		If Text_Phone.Text <> "" Then
			VendorAdd.Phone.SetValue(Text_Phone.Text)
		End If
		
		If Text_Alt_Phone.Text <> "" Then
			VendorAdd.AltPhone.SetValue(Text_Alt_Phone.Text)
		End If
		
		If Text_EMail.Text <> "" Then
			VendorAdd.Email.SetValue(Text_EMail.Text)
		End If
		
		If Text_Name_On_Check.Text <> "" Then
			VendorAdd.NameOnCheck.SetValue(Text_Name_On_Check.Text)
		End If
		
		If Text_Account_Number.Text <> "" Then
			VendorAdd.AccountNumber.SetValue(Text_Account_Number.Text)
		End If
		
		If Combo_Vendor_Type.Text <> "" Then
			VendorAdd.VendorTypeRef.FullName.SetValue(Combo_Vendor_Type.Text)
		End If
		
		
		'Perform the request
		Dim responseMsgSet As QBFC15Lib.IMsgSetResponse
		responseMsgSet = sessionManagerUS.DoRequests(requestMsgSet)
		
		'Interpret the response
		Dim response As QBFC15Lib.IResponse
		
		'The response list contains only one response,
		'which corresponds to our single request
		response = responseMsgSet.ResponseList.GetAt(0)
		
		If (response.statusCode <> 0) Then
			MsgBox("Status: Code = " & CStr(response.statusCode) & ", Severity = " & response.statusSeverity & ", Message = " & response.statusMessage)
		End If
		
		'The response detail for Add and Mod requests is a 'Ret' object
		'In our case, it's IVendorRet
		Dim VendorRet As QBFC15Lib.IVendorRet
		VendorRet = response.Detail
		If (Not (VendorRet Is Nothing)) Then
			MsgBox("A new vendor has been created with ListID = " & VendorRet.ListID.GetValue,  , "Vendor added")
		End If
		'Clear the form
		Command_Clear_Click(Command_Clear, New System.EventArgs())
		
		Exit Sub
		
ErrHandler: 
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
		
		
	End Sub
	Sub SaveWithCanadaSDK()
		On Error GoTo ErrHandler
		
		'Create the message set request object
		Dim requestMsgSet As QBFC15Lib.IMsgSetRequest
		requestMsgSet = sessionManagerCA.CreateMsgSetRequest("CA", CShort(cQBXMLMajorVersion), CShort(cQBXMLMinorVersion))
		'Initialize the request's attributes
		requestMsgSet.Attributes.OnError = QBFC15Lib.ENRqOnError.roeContinue
		
		'Add the request to the message set request object
		Dim VendorAdd As QBFC15Lib.IVendorAdd
		VendorAdd = requestMsgSet.AppendVendorAddRq
		
		If Text_Salutation.Text <> "" Then
			VendorAdd.Salutation.SetValue(Text_Salutation.Text)
		End If
		
		If Text_First_Name.Text <> "" And Text_Last_Name.Text <> "" Then
			VendorAdd.Name.SetValue(Text_First_Name.Text & " " & Text_Last_Name.Text)
			VendorAdd.FirstName.SetValue(Text_First_Name.Text)
			VendorAdd.LastName.SetValue(Text_Last_Name.Text)
		ElseIf Text_First_Name.Text <> "" Then 
			VendorAdd.Name.SetValue(Text_First_Name.Text)
			VendorAdd.FirstName.SetValue(Text_First_Name.Text)
		ElseIf Text_Last_Name.Text <> "" Then 
			VendorAdd.Name.SetValue(Text_Last_Name.Text)
			VendorAdd.LastName.SetValue(Text_Last_Name.Text)
		End If
		
		If Text_Middle_Name.Text <> "" Then
			VendorAdd.MiddleName.SetValue(Text_Middle_Name.Text)
		End If
		
		If Text_Address_1.Text <> "" Then
			VendorAdd.VendorAddress.Addr1.SetValue(Text_Address_1.Text)
		End If
		
		If Text_Address_2.Text <> "" Then
			VendorAdd.VendorAddress.Addr2.SetValue(Text_Address_2.Text)
		End If
		If Text_City.Text <> "" Then
			VendorAdd.VendorAddress.City.SetValue(Text_City.Text)
		End If
		If Combo_State_Prov.Text <> "" Then
			
			VendorAdd.VendorAddress.Province.SetValue(Combo_State_Prov)
		End If
		
		If Text_Postal_Code.Text <> "" Then
			VendorAdd.VendorAddress.PostalCode.SetValue(Text_Postal_Code.Text)
		End If
		
		If Text_Phone.Text <> "" Then
			VendorAdd.Phone.SetValue(Text_Phone.Text)
		End If
		
		If Text_Alt_Phone.Text <> "" Then
			VendorAdd.AltPhone.SetValue(Text_Alt_Phone.Text)
		End If
		
		If Text_EMail.Text <> "" Then
			VendorAdd.Email.SetValue(Text_EMail.Text)
		End If
		
		If Text_Name_On_Check.Text <> "" Then
			VendorAdd.NameOnCheck.SetValue(Text_Name_On_Check.Text)
		End If
		
		If Text_Account_Number.Text <> "" Then
			VendorAdd.AccountNumber.SetValue(Text_Account_Number.Text)
		End If
		
		If Combo_Vendor_Type.Text <> "" Then
			VendorAdd.VendorTypeRef.FullName.SetValue(Combo_Vendor_Type.Text)
		End If
		
		If Combo_Currency.Text <> "" Then
			VendorAdd.CurrencyRef.FullName.SetValue(Combo_Currency.Text)
		End If
		
		'Perform the request
		Dim responseMsgSet As QBFC15Lib.IMsgSetResponse
		responseMsgSet = sessionManagerCA.DoRequests(requestMsgSet)
		
		'Interpret the response
		Dim response As QBFC15Lib.IResponse
		
		'The response list contains only one response,
		'which corresponds to our single request
		response = responseMsgSet.ResponseList.GetAt(0)
		
		If (response.statusCode <> 0) Then
			MsgBox("Status: Code = " & CStr(response.statusCode) & ", Severity = " & response.statusSeverity & ", Message = " & response.statusMessage)
		End If
		
		'The response detail for Add and Mod requests is a 'Ret' object
		'In our case, it's IVendorRet
		Dim VendorRet As QBFC15Lib.IVendorRet
		VendorRet = response.Detail
		If (Not (VendorRet Is Nothing)) Then
			MsgBox("A new vendor has been created with ListID = " & VendorRet.ListID.GetValue,  , "Vendor added")
		End If
		'Clear the form
		Command_Clear_Click(Command_Clear, New System.EventArgs())
		Exit Sub
		
ErrHandler: 
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
		
		
	End Sub
End Class