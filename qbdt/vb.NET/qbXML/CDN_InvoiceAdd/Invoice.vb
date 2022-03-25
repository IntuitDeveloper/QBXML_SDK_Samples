Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks
Friend Class Invoice
	Inherits System.Windows.Forms.Form
    ' Invoice.frm
    ' This form is part of the Invoice sample program
    ' for the QuickBooks SDK Version CA2.0.
    '
    ' Created September, 2002
    '
    ' Copyright © 2021-2022 Intuit Inc. All rights reserved.
    ' Use is subject to the terms specified at:
    '      http://developer.intuit.com/legal/devsite_tos.html
    '
    '-------------------------------------------------------------
    
    
    Private Sub Combo_ArAccount_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo_ArAccount.TextChanged
		'If not in the list, we clear the combo box
		If Combo_ArAccount.SelectedIndex = -1 Then
			Combo_ArAccount.Text = ""
		End If
	End Sub
	
	
	Private Sub Combo_ArAccount_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo_ArAccount.SelectedIndexChanged
		'Once the account is selected, We must make sure that the currency is the same as the customer selected (If multi-currency is turned on in the preferences)
		
		Dim strArAccountArray() As String
		Dim strCustomerArray() As String 'I am using the name in the combo box to retrieve the info about
		If Combo_ArAccount.Text <> "" Then
			'This ArAccount from the collection
			
			strArAccountArray = colArAccount.Item(Combo_ArAccount.SelectedIndex + 1)
			If strArAccountArray(1) <> "" Then 'If multicurrency is turned on in the preference, The currencyref is not ""
				
				
				strCustomerArray = colCustomerCurrencyList.Item(Combo_Customer) 'I am using the name in the combo box to retrieve the info about
				'This customer from the collection
				
				If strCustomerArray(1) <> strArAccountArray(1) Then ' Make sure that both account and customer use the same currency by comparing the currencyref
					MsgBox("Please select an account that has the same currency as the customer",  , "Change account")
					Combo_ArAccount.Text = ""
					Combo_ArAccount.Focus()
					
				End If
			End If
		End If
		
		
	End Sub
	
	
	
	Private Sub Combo_Customer_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo_Customer.TextChanged
		'If not in the list, we clear the combo box
		If Combo_Customer.SelectedIndex = -1 Then
			Combo_Customer.Text = ""
		End If
	End Sub
	
	
	Private Sub Combo_Customer_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo_Customer.SelectedIndexChanged
		' Once the customer is selected, We need to find the Exchange rate if multi-currency
		' To do so, the use the currency used by the customer and find the exchange rate associated with this currency
		
		'Clear the form
		Call FormClear()
		Dim strCustomerArray() As String
		If Combo_Customer.Text <> "" Then
			'Get the customer info from the collection
			
			strCustomerArray = colCustomerCurrencyList.Item(Combo_Customer.SelectedIndex + 1) 'I am using the name in the combo box to retrieve the info about
			'This customer from the collection
			
			If strCustomerArray(1) <> "" Then 'If multicurrency is turned on in the preference, The currencyref is not ""
				Text_ExRate.Text = GetCurExRate(strCustomerArray(1)) ' Call the GetCurExRate that will use the currency ListID to find the Exchange Rate
			End If
		End If
	End Sub
	
	
	
	Private Sub Combo_First_Item_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo_First_Item.TextChanged
		'If not in the list, we clear the combo box
		If Combo_First_Item.SelectedIndex = -1 Then
			Combo_First_Item.Text = ""
		End If
	End Sub
	
	
	Private Sub Combo_First_Item_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo_First_Item.SelectedIndexChanged
		'Once the Item is selected, Populate the Description, rate and tax rate fields
		
		Dim strItemArray() As String
		If Combo_First_Item.Text <> "" Then
			'Get the Item info from the collection
			
			strItemArray = colItemList.Item(Combo_First_Item.SelectedIndex + 1) 'I am using the name in the combo box to retrieve the info about
			'This item from the collection
			'strItemArray(0) is the name of the item
			'strItemArray(1) is the taxcode associated to this currency
			'strItemArray(2) is Description associated to this item
			'strItemArray(3) is the price associated to this item
			'strItemArray(4) tell us if it's a sale or a purchase desc and price
			Text_Desc_First_Item.Text = strItemArray(2)
			Text_First_Rate.Text = strItemArray(3)
			Combo_First_TaxCode.Text = strItemArray(1)
		End If
	End Sub
	
	
	
	Private Sub Combo_First_TaxCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo_First_TaxCode.TextChanged
		'If not in the list, we clear the combo box
		If Combo_First_TaxCode.SelectedIndex = -1 Then
			Combo_First_TaxCode.Text = ""
		End If
		CalculateTaxAndTotal()
	End Sub
	
	
	Private Sub Combo_First_TaxCode_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo_First_TaxCode.SelectedIndexChanged
		'Recalculate the total and taxes
		CalculateTaxAndTotal()
	End Sub
	
	
	Private Sub Combo_Second_Item_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo_Second_Item.TextChanged
		'If not in the list, we clear the combo box
		If Combo_Second_Item.SelectedIndex = -1 Then
			Combo_Second_Item.Text = ""
		End If
		CalculateTaxAndTotal()
	End Sub
	
	
	Private Sub Combo_Second_Item_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo_Second_Item.SelectedIndexChanged
		'Once the Item is selected, Populate the Description, rate and tax rate fields
		
		Dim strItemArray() As String
		If Combo_Second_Item.Text <> "" Then
			'Get the Item info from the collection
			
			strItemArray = colItemList.Item(Combo_Second_Item.SelectedIndex + 1) 'I am using the name in the combo box to retrieve the info about
			'This item from the collection
			'strItemArray(0) is the name of the item
			'strItemArray(1) is the taxcode associated to this currency
			'strItemArray(2) is Description associated to this item
			'strItemArray(3) is the price associated to this item
			'strItemArray(4) tell us if it's a sale or a purchase desc and price
			Text_Desc_Second_Item.Text = strItemArray(2)
			Text_Second_Rate.Text = strItemArray(3)
			Combo_Second_TaxCode.Text = strItemArray(1)
		End If
	End Sub
	
	
	
	Private Sub Combo_Second_TaxCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo_Second_TaxCode.TextChanged
		'If not in the list, we clear the combo box
		If Combo_Second_TaxCode.SelectedIndex = -1 Then
			Combo_Second_TaxCode.Text = ""
		End If
		CalculateTaxAndTotal()
		
	End Sub
	
	
	Private Sub Combo_Second_TaxCode_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo_Second_TaxCode.SelectedIndexChanged
		'Recalculate the total and taxes
		CalculateTaxAndTotal()
	End Sub
	
	
	
	Private Sub Combo_Terms_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo_Terms.TextChanged
		'If not in the list, we clear the combo box
		If Combo_Terms.SelectedIndex = -1 Then
			Combo_Terms.Text = ""
		End If
	End Sub
	
	
	
	Private Sub Combo_Third_Item_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo_Third_Item.TextChanged
		'If not in the list, we clear the combo box
		If Combo_Third_Item.SelectedIndex = -1 Then
			Combo_Third_Item.Text = ""
		End If
	End Sub
	
	
	Private Sub Combo_Third_Item_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo_Third_Item.SelectedIndexChanged
		'Once the Item is selected, Populate the Description, rate and tax rate fields
		
		Dim strItemArray() As String
		If Combo_Third_Item.Text <> "" Then
			'Get the Item info from the collection
			
			strItemArray = colItemList.Item(Combo_Third_Item.SelectedIndex + 1) 'I am using the name in the combo box to retrieve the info about
			'This item from the collection
			'strItemArray(0) is the name of the item
			'strItemArray(1) is the taxcode associated to this currency
			'strItemArray(2) is Description associated to this item
			'strItemArray(3) is the price associated to this item
			'strItemArray(4) tell us if it's a sale or a purchase desc and price
			Text_Desc_Third_Item.Text = strItemArray(2)
			Text_Third_Rate.Text = strItemArray(3)
			Combo_Third_TaxCode.Text = strItemArray(1)
		End If
	End Sub
	
	
	
	Private Sub Combo_Third_TaxCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo_Third_TaxCode.TextChanged
		'If not in the list, we clear the combo box
		If Combo_Third_TaxCode.SelectedIndex = -1 Then
			Combo_Third_TaxCode.Text = ""
		End If
		CalculateTaxAndTotal()
	End Sub
	
	
	Private Sub Combo_Third_TaxCode_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo_Third_TaxCode.SelectedIndexChanged
		'Recalculate the total and taxes
		CalculateTaxAndTotal()
		
	End Sub
	
	Private Sub Command_Clear_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command_Clear.Click
		Combo_Customer.Text = ""
		'I call the FormClear sub
		Call FormClear()
	End Sub
	
	Private Sub Command_Close_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command_Close.Click
		CloseConnection() ' Make sure that all connections are closed with QuickBooks
		Me.Close() ' close the window
	End Sub
	
	Private Sub Command_Save_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command_Save.Click
		'This function is used to save the new invoice
		
		Dim blnCurrencySame As Boolean
		Dim strXMLRequest As String
		If Combo_Customer.Text = "" Then 'Need to have a customer selected
			MsgBox("Please select a customer",  , "Customer missing")
			Combo_Customer.Focus()
		ElseIf Combo_First_Item.Text = "" And Combo_Second_Item.Text = "" And Combo_Third_Item.Text = "" Then 
			MsgBox("Please select at least an item",  , "Item missing")
			Combo_First_Item.Focus()
		ElseIf "" <> Text_Date.Text And Not IsDate(Text_Date.Text) Then  'Verify the date format
			MsgBox("The Date Format is wrong" & vbCrLf & "The format of the date should be:" & vbCrLf & "DD/MM/YY" & vbCrLf & "DD/MM/YYYY" & vbCrLf & "DD-MM-YYYY" & vbCrLf & vbCrLf & "Examples: 14/08/02 14/08/2002 14-08-02 14-08-2002")
			Text_Date.Focus()
			
		Else
			blnCurrencySame = True
			If Text_Date.Text <> "" Then 'make sure that the date is in the right format for QuickBooks
				Text_Date.Text = toQBDate(Text_Date.Text)
			End If
			
			
			
			strXMLRequest = GenerateInvoiceXMLRequest() 'Create the XML request using the information on the form
			'Save the invoice in QuickBooks
			FuncTionSaveInvoice((strXMLRequest)) 'Call this function to send the request to QuickBooks
		End If
	End Sub
	
	Private Sub Invoice_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        ' The FormLoad function is used to populate the different combo boxes on the form.
        ' It stores the information received from QuickBooks in order to use them later.



        Text_Date.Text = toQBDate(CStr(Today))
		' Populating the Customer combo box
		Dim intIndex As Short
		
		
		intIndex = 0
		colCustomerCurrencyList = New Collection
		'Call the GetCustomerList procedure that will populate the colCustomerCurrencyList collection
		GetCustomerList()
		Dim strCustomerArray() As String
		Dim strItemArray() As String
		Dim strTaxCodeArray() As String
		Dim intArAccountCount As Short
		Dim StringArAccountArray() As String
		Dim strArAccountArray() As String 'This String array will contain the list of item names returned from QuickBooks
		If blnIsOpenConnection = True Then ' If we could not connect with the first function, I do not continue
			
			intIndex = 1
			'We loop through the customer arrays in the colCustomerCurrencyList collection
			While intIndex <= colCustomerCurrencyList.Count()
				'Adding the customer's name to the combo box
				
				strCustomerArray = colCustomerCurrencyList.Item(intIndex)
				Combo_Customer.Items.Add(strCustomerArray(0))
				intIndex = intIndex + 1 ' Go to the next customer
			End While
			
			
			
			'' Populating the Item combo box
			colItemList = New Collection
			'Call the GetItemList procedure that will populate the colItemList collection
			GetItemList()
			
			intIndex = 1
			'We loop through the item arrays in the colItemList collection
			While intIndex <= colItemList.Count()
				'Adding the Item name to all three item combo boxes
				
				strItemArray = colItemList.Item(intIndex)
				Combo_First_Item.Items.Add(strItemArray(0))
				Combo_Second_Item.Items.Add(strItemArray(0))
				Combo_Third_Item.Items.Add(strItemArray(0))
				intIndex = intIndex + 1 ' Go to the next item
			End While
			
			
			
			
			' Populating the Tax Code combo boxes
			
			colTaxCodeList = New Collection
			'Call the GetTaxCodeList procedure that will populate the colItemList collection
			GetTaxCodeList()
			intIndex = 1
			'We loop through the taxcode arrays in the colTaxCodeList collection
			While intIndex <= colTaxCodeList.Count()
				'Adding the TaxCode name to all three TaxCode combo boxes
				
				strTaxCodeArray = colTaxCodeList.Item(intIndex)
				Combo_First_TaxCode.Items.Add(strTaxCodeArray(0))
				Combo_Second_TaxCode.Items.Add(strTaxCodeArray(0))
				Combo_Third_TaxCode.Items.Add(strTaxCodeArray(0))
				intIndex = intIndex + 1 ' Go to the next TaxCode
			End While
			
			
			
			
			'
			' Populating the ArAccount combo box
			intArAccountCount = 0
			intIndex = 1
			
			'Call the function to retrieve the list of accounts
			colArAccount = New Collection
			'Call the GetArAccountList procedure that will populate the colArAccount collection
			GetArAccountList()
			
			'We loop through the ArAccount arrays in the colArAccount collection
			While intIndex <= colArAccount.Count()
				'Adding the ArAccount name to the ArAccount combo box
				
				strArAccountArray = colArAccount.Item(intIndex)
				Combo_ArAccount.Items.Add(strArAccountArray(0))
				intIndex = intIndex + 1 ' Go to the next ArAccount
			End While
			
		End If
	End Sub
	
	Private Sub Text_First_Amount_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text_First_Amount.Leave
		'Process the entry of a new amount
		If Text_First_Amount.Text <> "" Then
			If Not IsNumeric(Text_First_QTY.Text) Then
				MsgBox("Please enter a numeric value",  , "Enter an amount")
				Text_First_Amount.Text = ""
			Else
				If Text_First_QTY.Text <> "" Then
					
					Text_First_Rate.Text = CStr(CDbl(Text_First_Amount.Text) / CDbl(Text_First_QTY.Text))
					Text_First_Rate.Text = FormatNumber(Text_First_Rate.Text)
				ElseIf Text_First_QTY.Text = "" Then 
					Text_First_Rate.Text = FormatNumber(Text_First_Amount.Text)
				End If
				Text_First_Amount.Text = FormatNumber(Text_First_Amount.Text)
			End If
			Call CalculateTaxAndTotal() 'Call the function that calculated the taxes and total
			
		End If
		
	End Sub
	
	Private Sub Text_First_QTY_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text_First_QTY.Leave
		'Process the entry of a new Quantity
		If Text_First_QTY.Text <> "" Then
			If Not IsNumeric(Text_First_QTY.Text) Then
				MsgBox("Please enter a numeric value",  , "Enter a qty")
				Text_First_QTY.Text = ""
			Else
				If Text_First_Rate.Text <> "" Then
					Text_First_Amount.Text = FormatNumber(CDbl(Text_First_Rate.Text) * CDbl(Text_First_QTY.Text))
				End If
				
			End If
			Call CalculateTaxAndTotal() 'Call the function that calculated the taxes and total
		End If
	End Sub
	
	
	Private Sub Text_First_Rate_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text_First_Rate.Leave
		'Process the entry of a new Rate
		If Text_First_Rate.Text <> "" Then
			If Not IsNumeric(Text_First_Rate.Text) Then
				MsgBox("Please enter a numeric value",  , "Enter a Rate")
				Text_First_Rate.Text = ""
			Else
				If Text_First_Amount.Text <> "" And Text_First_QTY.Text = "" Then
					Text_First_Amount.Text = FormatNumber(Text_First_Rate.Text)
				ElseIf Text_First_QTY.Text <> "" Then 
					Text_First_Amount.Text = FormatNumber(CDbl(Text_First_Rate.Text) * CDbl(Text_First_QTY.Text))
				End If
				Text_First_Rate.Text = FormatNumber(Text_First_Rate.Text)
			End If
			Call CalculateTaxAndTotal() 'Call the function that calculated the taxes and total
			
		End If
		
	End Sub
	
	
	
	Private Sub Text_Second_Amount_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text_Second_Amount.Leave
		'Process the entry of a new Amount
		If Text_Second_Amount.Text <> "" Then
			If Not IsNumeric(Text_Second_QTY.Text) Then
				MsgBox("Please enter a numeric value",  , "Enter an amount")
				Text_Second_Amount.Text = ""
			Else
				If Text_Second_QTY.Text <> "" Then
					
					Text_Second_Rate.Text = CStr(CDbl(Text_Second_Amount.Text) / CDbl(Text_Second_QTY.Text))
					Text_Second_Rate.Text = FormatNumber(Text_Second_Rate.Text)
				ElseIf Text_Second_QTY.Text = "" Then 
					Text_Second_Rate.Text = FormatNumber(Text_Second_Amount.Text)
				End If
				Text_Second_Amount.Text = FormatNumber(Text_Second_Amount.Text)
			End If
			Call CalculateTaxAndTotal() 'Call the function that calculated the taxes and total
			
		End If
		
		
	End Sub
	
	Private Sub Text_Second_QTY_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text_Second_QTY.Leave
		'Process the entry of a new Quantity
		If Text_Second_QTY.Text <> "" Then
			If Not IsNumeric(Text_Second_QTY.Text) Then
				MsgBox("Please enter a numeric value",  , "Enter a qty")
				Text_Second_QTY.Text = ""
			Else
				If Text_Second_Rate.Text <> "" Then
					Text_Second_Amount.Text = FormatNumber(CDbl(Text_Second_Rate.Text) * CDbl(Text_Second_QTY.Text))
				End If
				
			End If
			Call CalculateTaxAndTotal() 'Call the function that calculated the taxes and total
		End If
	End Sub
	Private Sub Text_Second_Rate_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text_Second_Rate.Leave
		'Process the entry of a new Rate
		If Text_Second_Rate.Text <> "" Then
			If Not IsNumeric(Text_Second_Rate.Text) Then
				MsgBox("Please enter a numeric value",  , "Enter a Rate")
				Text_Second_Rate.Text = ""
			Else
				If Text_Second_Amount.Text <> "" And Text_Second_QTY.Text = "" Then
					Text_Second_Amount.Text = FormatNumber(Text_Second_Rate.Text)
				ElseIf Text_Second_QTY.Text <> "" Then 
					Text_Second_Amount.Text = FormatNumber(CDbl(Text_Second_Rate.Text) * CDbl(Text_Second_QTY.Text))
				End If
				Text_Second_Rate.Text = FormatNumber(Text_Second_Rate.Text)
			End If
			Call CalculateTaxAndTotal() 'Call the function that calculated the taxes and total
		End If
	End Sub
	Private Sub Text_Third_Amount_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text_Third_Amount.Leave
		'Process the entry of a new Amount
		If Text_Third_Amount.Text <> "" Then
			If Not IsNumeric(Text_Third_QTY.Text) Then
				MsgBox("Please enter a numeric value",  , "Enter an amount")
				Text_Third_Amount.Text = ""
			Else
				If Text_Third_QTY.Text <> "" Then
					
					Text_Third_Rate.Text = CStr(CDbl(Text_Third_Amount.Text) / CDbl(Text_Third_QTY.Text))
					Text_Third_Rate.Text = FormatNumber(Text_Third_Rate.Text)
				ElseIf Text_Third_QTY.Text = "" Then 
					Text_Third_Rate.Text = FormatNumber(Text_Third_Amount.Text)
				End If
				Text_Third_Amount.Text = FormatNumber(Text_Third_Amount.Text)
			End If
			Call CalculateTaxAndTotal() 'Call the function that calculated the taxes and total
			
		End If
		
	End Sub
	
	Private Sub Text_Third_QTY_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text_Third_QTY.Leave
		'Process the entry of a new Quantity
		If Text_Third_QTY.Text <> "" Then
			If Not IsNumeric(Text_Third_QTY.Text) Then
				MsgBox("Please enter a numeric value",  , "Enter a qty")
				Text_Third_QTY.Text = ""
			Else
				If Text_Third_Rate.Text <> "" Then
					Text_Third_Amount.Text = FormatNumber(CDbl(Text_Third_Rate.Text) * CDbl(Text_Third_QTY.Text))
				End If
				
			End If
			Call CalculateTaxAndTotal() 'Call the function that calculated the taxes and total
		End If
		
	End Sub
	
	
	Private Sub Text_Third_Rate_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text_Third_Rate.Leave
		'Process the entry of a new Rate
		If Text_Third_Rate.Text <> "" Then
			If Not IsNumeric(Text_Third_Rate.Text) Then
				MsgBox("Please enter a numeric value",  , "Enter a Rate")
				Text_Third_Rate.Text = ""
			Else
				If Text_Third_Amount.Text <> "" And Text_Third_QTY.Text = "" Then
					Text_Third_Amount.Text = FormatNumber(Text_Third_Rate.Text)
				ElseIf Text_Third_QTY.Text <> "" Then 
					Text_Third_Amount.Text = FormatNumber(CDbl(Text_Third_Rate.Text) * CDbl(Text_Third_QTY.Text))
				End If
				Text_Third_Rate.Text = FormatNumber(Text_Third_Rate.Text)
			End If
			Call CalculateTaxAndTotal() 'Call the function that calculated the taxes and total
			
		End If
		
	End Sub
	
	
	Public Function toQBDate(ByRef theDate As String) As String
        'Make sure that the date is in a format that QuickBooks accepts
        On Error GoTo ErrHandler
        Dim dt As DateTime = Convert.ToDateTime(theDate)
        toQBDate = dt.ToString("yyyy-MM-dd")

        If toQBDate <> "" And Not IsDate(toQBDate) Then
			toQBDate = "error"
			Exit Function
		End If
		
		Exit Function
		
ErrHandler: 
		toQBDate = "error formating the date"
	End Function
	Sub SetNumberInCurrency()
		'Format the rate and amount fields into a field with 2 decimal point
		If Text_First_Rate.Text <> "" Then
			Text_First_Rate.Text = FormatNumber(Text_First_Rate.Text)
		End If
		If Text_First_Amount.Text <> "" Then
			Text_First_Amount.Text = FormatNumber(Text_First_Amount.Text)
		End If
		If Text_Second_Rate.Text <> "" Then
			Text_Second_Rate.Text = FormatNumber(Text_Second_Rate.Text)
		End If
		If Text_Second_Amount.Text <> "" Then
			Text_Second_Amount.Text = FormatNumber(Text_Second_Amount.Text)
		End If
		If Text_Third_Rate.Text <> "" Then
			Text_Third_Rate.Text = FormatNumber(Text_Third_Rate.Text)
		End If
		If Text_Third_Amount.Text <> "" Then
			Text_Third_Amount.Text = FormatNumber(Text_Third_Amount.Text)
		End If
	End Sub
	
	Sub CalculateTaxAndTotal()
		'Calculating the Taxes and total
		Text_GST.Text = ""
		Text_PST.Text = ""
		
		Dim strTaxCodeArray() As String 'I am using the name in the combo box to retrieve the info about
		'This ArAccount from the collection
		
		'strTaxCodeArray(0) is TaxCode Name
		'strTaxCodeArray (1) is GST (TaxRate1)
		'strTaxCodeArray(2) is PST (TaxRate2)
		
		
		
		
		Dim strTempTaxRate As String
		
		Dim strFirstComboTaxRate1 As String
		Dim strFirstComboTaxRate2 As String
		If Combo_First_TaxCode.Text <> "" Then
            'If tax rate selected for the first item...
            
            strTaxCodeArray = colTaxCodeList.Item(Combo_First_TaxCode.SelectedIndex + 1) 'We  retrieve the taxcode array corresponding to the taxcode selected in the combo box

            If strTaxCodeArray(1) <> "" Then
				strFirstComboTaxRate1 = CStr(CDbl(strTaxCodeArray(1)) / 100)
			End If
			If strTaxCodeArray(2) <> "" Then
				strFirstComboTaxRate2 = CStr(CDbl(strTaxCodeArray(2)) / 100)
			Else
				strFirstComboTaxRate2 = "0"
			End If
			
		End If
		
		Dim strSecondComboTaxRate1 As String
		Dim strSecondComboTaxRate2 As String
		If Combo_Second_TaxCode.Text <> "" Then
            'If tax rate selected for the second item...
            
            strTaxCodeArray = colTaxCodeList.Item(Combo_Second_TaxCode.SelectedIndex + 1) 'We  retrieve the taxcode array corresponding to the taxcode selected in the combo box

            If strTaxCodeArray(1) <> "" Then
				strSecondComboTaxRate1 = CStr(CDbl(strTaxCodeArray(1)) / 100)
			End If
			If strTaxCodeArray(2) <> "" Then
				strSecondComboTaxRate2 = CStr(CDbl(strTaxCodeArray(2)) / 100)
			Else
				strSecondComboTaxRate2 = "0"
			End If
			
		End If
		
		Dim strThirdComboTaxRate1 As String
		Dim strThirdComboTaxRate2 As String
		If Combo_Third_TaxCode.Text <> "" Then
            'If tax rate selected for the third item...
            
            strTaxCodeArray = colTaxCodeList.Item(Combo_Third_TaxCode.SelectedIndex + 1) 'We  retrieve the taxcode array corresponding to the taxcode selected in the combo box

            If strTaxCodeArray(1) <> "" Then
				strThirdComboTaxRate1 = CStr(CDbl(strTaxCodeArray(1)) / 100)
			End If
			If strTaxCodeArray(2) <> "" Then
				strThirdComboTaxRate2 = CStr(CDbl(strTaxCodeArray(2)) / 100)
			Else
				strThirdComboTaxRate2 = "0"
				
			End If
			
		End If
		'Calculate the Total GST and PST
		If Text_First_Amount.Text <> "" And Combo_First_TaxCode.Text <> "" Then
			Text_GST.Text = CStr(CDbl(Text_First_Amount.Text) * CDbl(strFirstComboTaxRate1))
			Text_PST.Text = CStr(CDbl(Text_First_Amount.Text) * CDbl(strFirstComboTaxRate2))
		End If
		
		If Text_Second_Amount.Text <> "" And Combo_Second_TaxCode.Text <> "" Then
			If Text_GST.Text = "" Then '
				Text_GST.Text = CStr(CDbl(Text_Second_Amount.Text) * CDbl(strSecondComboTaxRate1))
			Else
				Text_GST.Text = CStr(CDbl(Text_GST.Text) + CDbl(Text_Second_Amount.Text) * CDbl(strSecondComboTaxRate1))
			End If
			If Text_PST.Text = "" Then
				Text_PST.Text = CStr(CDbl(Text_Second_Amount.Text) * CDbl(strSecondComboTaxRate2))
			Else
				Text_PST.Text = CStr(CDbl(Text_PST.Text) + CDbl(Text_Second_Amount.Text) * CDbl(strSecondComboTaxRate2))
			End If
		End If
		
		If Text_Third_Amount.Text <> "" And Combo_Third_TaxCode.Text <> "" Then
			If Text_GST.Text = "" Then
				Text_GST.Text = CStr(CDbl(Text_Third_Amount.Text) * CDbl(strThirdComboTaxRate1))
			Else
				Text_GST.Text = CStr(CDbl(Text_GST.Text) + CDbl(Text_Third_Amount.Text) * CDbl(strThirdComboTaxRate1))
				
			End If
			
			If Text_PST.Text = "" Then
				Text_PST.Text = CStr(CDbl(Text_Third_Amount.Text) * CDbl(strThirdComboTaxRate2))
			Else
				Text_PST.Text = CStr(CDbl(Text_PST.Text) + CDbl(Text_Third_Amount.Text) * CDbl(strThirdComboTaxRate2))
			End If
		End If
		If Text_GST.Text <> "" Then
			'Format the GST amount with 2 decimal
			Text_GST.Text = FormatNumber(Text_GST.Text)
		End If
		If Text_PST.Text <> "" Then
			'Format the PST amount with 2 decimal
			Text_PST.Text = FormatNumber(Text_PST.Text)
		End If
		
		
		'Calculate the total
		Text_Total.Text = CStr(0)
		If Text_First_Amount.Text <> "" Then
			Text_Total.Text = CStr(CDbl(Text_Total.Text) + CDbl(Text_First_Amount.Text))
		End If
		If Text_Second_Amount.Text <> "" Then
			Text_Total.Text = CStr(CDbl(Text_Total.Text) + CDbl(Text_Second_Amount.Text))
		End If
		If Text_Third_Amount.Text <> "" Then
			Text_Total.Text = CStr(CDbl(Text_Total.Text) + CDbl(Text_Third_Amount.Text))
		End If
		If Text_GST.Text <> "" Then
			Text_Total.Text = CStr(CDbl(Text_Total.Text) + CDbl(Text_GST.Text))
		End If
		If Text_PST.Text <> "" Then
			Text_Total.Text = CStr(CDbl(Text_Total.Text) + CDbl(Text_PST.Text))
		End If
		If Text_Total.Text <> "" Then
			Text_Total.Text = FormatNumber(Text_Total.Text)
		End If
	End Sub
	Function GenerateInvoiceXMLRequest() As String
		' We generate the request qbXML in order to get the ArAccount list with filter to have only account receivables.
		Dim requestXML As String
        requestXML = "<?xml version=""1.0"" ?>" & "<?qbxml version=""2.0""?>" & "<QBXML><QBXMLMsgsRq onError=""continueOnError"">" & "<InvoiceAddRq requestID=""1"">" & "<InvoiceAdd>" & "<CustomerRef>" & "<FullName>" & Combo_Customer.Text & "</FullName>" & "</CustomerRef>"
        If Combo_ArAccount.Text <> "" Then
			requestXML = requestXML & "<ARAccountRef>" & "<FullName>" & Combo_ArAccount.Text & "</FullName>" & "</ARAccountRef>"
		End If
		
		If Text_Date.Text <> "" Then
			requestXML = requestXML & "<TxnDate>" & Text_Date.Text & "</TxnDate>"
		End If
		'If the first item is selected...
		If Combo_First_Item.Text <> "" Then
			requestXML = requestXML & "<InvoiceLineAdd>" & "<ItemRef>" & "<FullName>" & Combo_First_Item.Text & "</FullName>" & "</ItemRef>"
			If Text_Desc_First_Item.Text <> "" Then
				requestXML = requestXML & "<Desc>" & Text_Desc_First_Item.Text & "</Desc>"
			End If
			
			If Text_First_QTY.Text <> "" Then
				requestXML = requestXML & "<Quantity>" & Text_First_QTY.Text & "</Quantity>"
			End If
			
			If Text_First_Rate.Text <> "" Then
				requestXML = requestXML & "<Rate>" & Text_First_Rate.Text & "</Rate>"
			End If
			
			If Combo_First_TaxCode.Text <> "" Then
				requestXML = requestXML & "<SalesTaxCodeRef>" & "<FullName>" & Combo_First_TaxCode.Text & "</FullName>" & "</SalesTaxCodeRef>"
			End If
			requestXML = requestXML & "</InvoiceLineAdd>"
		End If
		'If the second item is selected...
		If Combo_Second_Item.Text <> "" Then
			requestXML = requestXML & "<InvoiceLineAdd>" & "<ItemRef>" & "<FullName>" & Combo_Second_Item.Text & "</FullName>" & "</ItemRef>"
			If Text_Desc_Second_Item.Text <> "" Then
				requestXML = requestXML & "<Desc>" & Text_Desc_Second_Item.Text & "</Desc>"
			End If
			
			If Text_Second_QTY.Text <> "" Then
				requestXML = requestXML & "<Quantity>" & Text_Second_QTY.Text & "</Quantity>"
			End If
			
			If Text_Second_Rate.Text <> "" Then
				requestXML = requestXML & "<Rate>" & Text_Second_Rate.Text & "</Rate>"
			End If
			
			If Combo_Second_TaxCode.Text <> "" Then
				requestXML = requestXML & "<SalesTaxCodeRef>" & "<FullName>" & Combo_Second_TaxCode.Text & "</FullName>" & "</SalesTaxCodeRef>"
			End If
			requestXML = requestXML & "</InvoiceLineAdd>"
			
			
		End If
		'If the third item is selected...
		If Combo_Third_Item.Text <> "" Then
			requestXML = requestXML & "<InvoiceLineAdd>" & "<ItemRef>" & "<FullName>" & Combo_Third_Item.Text & "</FullName>" & "</ItemRef>"
			If Text_Desc_Third_Item.Text <> "" Then
				requestXML = requestXML & "<Desc>" & Text_Desc_Third_Item.Text & "</Desc>"
			End If
			
			If Text_Third_QTY.Text <> "" Then
				requestXML = requestXML & "<Quantity>" & Text_Third_QTY.Text & "</Quantity>"
			End If
			
			If Text_Third_Rate.Text <> "" Then
				requestXML = requestXML & "<Rate>" & Text_Third_Rate.Text & "</Rate>"
			End If
			
			
			If Combo_Third_TaxCode.Text <> "" Then
				requestXML = requestXML & "<SalesTaxCodeRef>" & "<FullName>" & Combo_Third_TaxCode.Text & "</FullName>" & "</SalesTaxCodeRef>"
			End If
			requestXML = requestXML & "</InvoiceLineAdd>"
			
		End If
		'Add the exchange rate to the request
		If Text_ExRate.Text <> "" Then
			requestXML = requestXML & "<ExchangeRate>" & Text_ExRate.Text & "</ExchangeRate>"
		End If
		
		requestXML = requestXML & "</InvoiceAdd>" & "</InvoiceAddRq>" & "</QBXMLMsgsRq></QBXML>"
		
		GenerateInvoiceXMLRequest = requestXML 'return the XML request
		
		
	End Function
	Sub FormClear()
		'Clear the form
		Combo_ArAccount.Text = ""
		Combo_First_Item.Text = ""
		Text_Desc_First_Item.Text = ""
		Text_First_QTY.Text = ""
		Text_First_Rate.Text = ""
		Text_First_Amount.Text = ""
		Combo_First_TaxCode.Text = ""
		
		Combo_Second_Item.Text = ""
		Text_Desc_Second_Item.Text = ""
		Text_Second_QTY.Text = ""
		Text_Second_Rate.Text = ""
		Text_Second_Amount.Text = ""
		Combo_Second_TaxCode.Text = ""
		
		Combo_Third_Item.Text = ""
		Text_Desc_Third_Item.Text = ""
		Text_Third_QTY.Text = ""
		Text_Third_Rate.Text = ""
		Text_Third_Amount.Text = ""
		Combo_Third_TaxCode.Text = ""
		
		Text_ExRate.Text = ""
		Text_GST.Text = ""
		Text_PST.Text = ""
		Text_Total.Text = ""
	End Sub
End Class