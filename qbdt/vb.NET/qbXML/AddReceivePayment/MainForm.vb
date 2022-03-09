Option Strict Off
Option Explicit On
Friend Class MainForm
	Inherits System.Windows.Forms.Form
	' MainForm.frm
	'
	' This is the form file for the sample receive payment code.
	' The user is first asked to choose a customer.  Then lists
	' are displayed of the customer's open invoices and credit
	' memos as retrieved from QuickBooks.  The user can choose
	' an invoice and optionally a credit memo, fill in all of the
	' payment information, and then send the payment to QuickBooks.
	'
	' This application works with a few selected customers. A sample
	' QuickBooks company file: ReceivePaymentsCompanyFile.qbw has
	' been provided with this application, which has the relevant
	' information for invoices and credit memos set up for this
	' purpose.
	'
	'
	'
	' Created February, 2002
	' Updated to SDK 2.0 July, 2002
	'
	' Copyright © 2021-2022 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	'-------------------------------------------------------------
	
	
	' Global variable holds the customer information so it will not
	' be inadvertantly modified between when "Display Invoices"
	' is clicked and when "Apply Payment" is clicked
	Dim customer As String
	Public connType As QBXMLRP2Lib.QBXMLRPConnectionType
	
	' Hold info about all of the current customer's open invoices
	Dim invoice_TxnIds() As String
	Dim invoice_AppliedAmts() As Decimal
	Dim invoice_Balances() As Decimal
	Dim invoice_RefNumbers() As String
	
	' Hold info about all of the current customer's credit memos
	Dim credit_TxnIds() As String
	Dim credit_TotalAmts() As Decimal
	Dim credit_AmtLeft() As Decimal
	'
	
	' After the user has filled in all appropriate information,
	' we can collect this info and finally make a payment on the
	' appropriate invoice.
	' Most of this subroutine is concerned with error checking.
	' The user can not specify a credit memo amount if no credit
	' memo is selected.  The user can't make a payment larger than
	' the total amount of the invoice.  There are a number of
	' required fields as well.
	'
	Private Sub Command_Apply_Payment_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command_Apply_Payment.Click
		
		Dim oldDate As String
		Dim payDate As String
		Dim refNum As String
		Dim payMethod As String
		Dim payAmt As Decimal
		Dim memoAmt As Decimal
		Dim discountAmt As Decimal
		Dim invoice As String
		Dim credMemo As String
		
		' We need to find the invoice and credit memo data
		' for the correct invoice and credit memo in the
		' global arrays
		Dim invNumber As Short
		Dim credNumber As Short
		Dim invTxnID As String
		Dim credTxnID As String
		
		Dim invNumString() As String
		Dim credNumString() As String
		
		' Make sure an invoice is selected
		If "" = Combo_Invoices.Text Then
			MsgBox("You must select an invoice.")
			Exit Sub
		End If
		
		' Figure out which invoice the user has selected
		invoice = Combo_Invoices.Text
		invNumString = Split(invoice, ">>")
		
		' InvNumber is the array index of this invoice in
		' the four invoice related global variables
		invNumber = CShort(invNumString(0)) - 1
		
		' See if the user has selected a credit memo, if
		' so, figure out which it is
		If "" = Combo_Credit_Memos.Text Then
			credMemo = "" ' No credit memo selected
		Else
			' Same way we figured out which invoice was selected
			credMemo = Combo_Credit_Memos.Text
			credNumString = Split(credMemo, ">>")
			credNumber = CShort(credNumString(0)) - 1
		End If
		
		
		' Get the amount that the customer wants to apply
		' to the credit memo and make sure that it is not larger
		' than the amount of the credit memo the customer is using
		If Not ("" = Text_Credit_Memo.Text) Then
			memoAmt = CDec(Text_Credit_Memo.Text)
			
			If ("" = credMemo) Then
				' If user has tried to specify a credit amount without
				' choosing a credit memo, go back to form
				MsgBox("You may not specify a credit memo amount unless you " & "choose a credit memo.  Please fix this error and " & "hit Apply Payment again.")
				Exit Sub
				
			Else
				
				If credit_AmtLeft(credNumber) = 0 Then
					' The remaining balance for the credit memo is $0.
					' Another credit memo will need to be selected.
					MsgBox("There is no remaining balance available on this credit memo to apply a payment. " & "Please select another credit memo.")
					Exit Sub
					
				End If
				
				If memoAmt < 0 Then
					MsgBox("Please specify a positive credit memo amount.")
					Exit Sub
				End If
				
				
				If memoAmt > credit_AmtLeft(credNumber) Then
					' Can not use more credit than they have!
					' Go back to form if user tries to do this
					MsgBox("In order to use this credit memo, you must " & "select a credit amount less than the remaining balance: $" & credit_AmtLeft(credNumber))
					Exit Sub
					
				End If
			End If
			
		Else
			' If the customer is not applying a credit memo..
			memoAmt = 0
		End If
		
		
		' Get the discount amount if there is one filled in
		If Not ("" = Text_Discount.Text) Then
			discountAmt = CDec(Text_Discount.Text)
			If discountAmt < 0 Then
				MsgBox("Please specify a positive amount for discount.")
				Exit Sub
			End If
		Else
			discountAmt = 0
		End If
		
		
		' Get the amount of the cash/credit/check/etc payment
		If Not ("" = Text_Amt_Paid.Text) Then
			payAmt = CDec(Text_Amt_Paid.Text)
			If payAmt < 0 Then
				MsgBox("Please specify a positive payment amount.")
				Exit Sub
			End If
		Else
			payAmt = 0
		End If
		
		
		' Make sure that the total of the memo amount, the discount
		' amount, and the regular payment amount is no larger than
		' the total amount of the invoice -- Return to the form
		' if it is.
		If ((memoAmt + discountAmt + payAmt) > invoice_Balances(invNumber)) Then
			MsgBox("You may not pay more than the remaining balance amount " & "of the selected invoice.  Please make sure the total of " & "your payment amounts is less than $" & invoice_Balances(invNumber) & " and hit Apply again.")
			Exit Sub
		End If
		
		
		' Get date info from the form and convert it to a date
		' that QuickBooks can read.  Return to form if date not valid.
		oldDate = Text_Pay_Date.Text
		payDate = toQBDate(oldDate)
		If ("error" = payDate) Then
			MsgBox("The date you entered is not valid.")
			Exit Sub
		End If
		
		' Get the reference number and payment method from the form
		refNum = Text_Ref_Number.Text
		payMethod = Combo_Pay_Method.Text
		
		' Force all mandatory fields to be entered -- The invoice
		' field is not listed here because it should have already
		' been checked above
		If ("" = payDate) Or ("" = refNum) Or ("" = payMethod) Or ("" = customer) Or (0 = payAmt) Then
			
			' Go back to form if not all required fields are filled in
			MsgBox("The following fields are required before you apply a payment to an invoice:" & vbCrLf & "Payment Date" & vbCrLf & "Ref/Check Number" & vbCrLf & "Payment Method" & vbCrLf & "Payment Amount" & vbCrLf & "Invoice")
			Exit Sub
			
		Else
			' Send the payment to QuickBooks!
			
			If (memoAmt > 0) Then
				
				SendPaymentToQB(customer, payDate, refNum, payMethod, payAmt, memoAmt, discountAmt, invoice_TxnIds(invNumber), credit_TxnIds(credNumber))
				
			Else
				
				' Avoid out of range error if there is no credit memo
				' being used
				SendPaymentToQB(customer, payDate, refNum, payMethod, payAmt, memoAmt, discountAmt, invoice_TxnIds(invNumber), "")
				
			End If
			
			' Now that we have applied the payment, clear all of the fields
			' so that the data will not be stale
			Combo_Invoices.Items.Clear()
			Combo_Credit_Memos.Items.Clear()
			Text_Pay_Date.Text = CStr(Today)
			Text_Ref_Number.Text = ""
			Text_Amt_Paid.Text = ""
			Text_Credit_Memo.Text = ""
			Text_Discount.Text = ""
			
		End If
		
	End Sub
	
	
	' Once a customer has been selected and the user clicks the button,
	' we need to fetch all of the users open invoices and credit memos
	' and display them in the appropriate list boxes.
	'
	Private Sub Command_Display_Invoices_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command_Display_Invoices.Click
		
		' Set global variable to make sure that its value
		' will remain the same even if the data is somehow
		' deleted or changed before "Apply Payment" is clicked
		customer = Combo_Customer.Text
		
		' Force mandatory field to be entered
		If ("" = customer) Then
			MsgBox("You must specify a customer before invoices can be displayed.")
		Else
			
			' Before we can display the new lists, we have to clear
			' the old ones in case there were choices left over from
			' another customer
			Combo_Invoices.Items.Clear()
			Combo_Credit_Memos.Items.Clear()
			
			' Clear all fields that appear below on the form also
			Text_Pay_Date.Text = CStr(Today)
			Text_Ref_Number.Text = ""
			Text_Amt_Paid.Text = ""
			Text_Credit_Memo.Text = ""
			Text_Discount.Text = ""
			
			' This is the first time we communicate with QuickBooks.  A
			' connection is opened and we ask for a list of invoices for
			' the currently selected customer.  These invoices are then
			' displayed on the form.  The same is done for credit memos.
			displayLists((customer))
			
		End If
		
	End Sub
	
	
	' Add original lists of customers and payment methods.
	' Also set the date field to the current date.
	'
	Private Sub MainForm_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Text_Pay_Date.Text = CStr(Today)
		connType = QBXMLRP2Lib.QBXMLRPConnectionType.localQBDLaunchUI
		If (MsgBox("Do you want to communicate with QuickBooks Online Edition?", MsgBoxStyle.YesNo, "Use QuickBooks Online Edition") = MsgBoxResult.Yes) Then
			connType = QBXMLRP2Lib.QBXMLRPConnectionType.remoteQBOE
		End If
		OpenConnection()
		fillCustomerList()
		
		
		
		Combo_Pay_Method.Items.Add("American Express")
		Combo_Pay_Method.Items.Add("Barter")
		Combo_Pay_Method.Items.Add("Cash")
		Combo_Pay_Method.Items.Add("Check")
		Combo_Pay_Method.Items.Add("Discover")
		Combo_Pay_Method.Items.Add("Master Card")
		Combo_Pay_Method.Items.Add("VISA")
		
	End Sub
	
	Private Sub MainForm_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		
		If blnIsOpenConnection Then
			'Call CloseConnection
			CloseConnection()
		End If
		
	End Sub
	
	' Need a public sub that allows qbooks.bas to set the size
	' of the invoice related arrays based on
	' how many open invoices are found
	'
	Public Sub SetInvoiceArrayLengths(ByRef numInvoices As Short)
		
		ReDim invoice_TxnIds(numInvoices)
		ReDim invoice_AppliedAmts(numInvoices)
		ReDim invoice_Balances(numInvoices)
		Dim invoice_SuggDiscounts(numInvoices) As Decimal
		ReDim invoice_RefNumbers(numInvoices)
		
	End Sub
	
	
	
	
	' Need a public sub that allows qbooks.bas to set the size
	' of the credit memo related arrays based on
	' how many open credit memos are found
	'
	Public Sub SetCreditMemoArrayLengths(ByRef numCreditMemos As Short)
		
		ReDim credit_TxnIds(numCreditMemos)
		ReDim credit_TotalAmts(numCreditMemos)
		ReDim credit_AmtLeft(numCreditMemos)
		
	End Sub
	
	Public Sub AddCustomerToList(ByRef FullName As String)
		Combo_Customer.Items.Add((FullName))
	End Sub
	
	' Need a public sub so that qbooks.bas can add
	' invoices to the list when getInvoiceList() is called
	'
	Public Sub AddInvoiceToList(ByRef txnId As String, ByRef appliedAmt As Decimal, ByRef balanceRemaining As Decimal, ByRef suggDiscount As Decimal, ByRef invRefNum As String, ByRef number As Short)
		
		Dim invoiceData As String
		Dim index As Short
		
		' Start numbering the list of invoices at 1 instead of 0
		index = number + 1
		
		' Store info in all of the global invoice arrays so that
		' we will have it later once the user chooses an invoice
		invoice_TxnIds(number) = txnId
		invoice_AppliedAmts(number) = appliedAmt
		invoice_Balances(number) = balanceRemaining
		invoice_RefNumbers(number) = invRefNum
		
		invoiceData = index & ">> " & " Ref Num: " & invRefNum & "  /  Remaining Balance: $" & currToString(balanceRemaining) & "  /  Applied: $" & currToString(appliedAmt)
		
		' Add invoice to the list of choices on the form
		Combo_Invoices.Items.Add(invoiceData)
		
		' Want the first invoice to be selected initially
		If (1 = index) Then
			Combo_Invoices.Text = invoiceData
		End If
	End Sub
	
	
	' Need a public sub so that qbooks.bas can add
	' credit memos to the list when getCreditMemoList() is called
	'
	Public Sub AddCreditMemoToList(ByRef txnId As String, ByRef TotalAmount As Decimal, ByRef amountLeft As Decimal, ByRef num As Short)
		
		Dim cmData As String
		Dim index As Short
		
		' Want to start numbering the list of credit memos that is
		' displayed at 1 instead of 0
		index = num + 1
		
		' Add the information about these credit memos to the global
		' variables so that it will be available when the user
		' chooses a memo later.
		credit_TxnIds(num) = txnId
		credit_TotalAmts(num) = TotalAmount
		credit_AmtLeft(num) = amountLeft
		
		cmData = index & ">> Total Amount: $" & currToString(TotalAmount) & "  /  Remaining Balance: $" & currToString(amountLeft)
		
		' Add credit memo to the list of choices on the form
		Combo_Credit_Memos.Items.Add(cmData)
		
		' Want the first credit memo to be selected initially
		If (1 = index) Then
			Combo_Credit_Memos.Text = cmData
		End If
	End Sub
	
	
	' Adds credit memo boxes to form if they have been removed --
	' this is used when a customer who has credit memos open is
	' selected after a previous one who did not have open credit memos
	'
	Public Sub TurnOnCreditMemos()
		
		Combo_Credit_Memos.Visible = True
		Text_Credit_Memo.Visible = True
		Label_Credit_Memo.Visible = True
		Label_Credit_Amt.Visible = True
		
	End Sub
	
	' Removes text boxes and labels dealing with credit memos --
	' this is used when a customer who does not have any open
	' credit memos is selected
	'
	Public Sub TurnOffCreditMemos()
		
		Combo_Credit_Memos.Visible = False
		Text_Credit_Memo.Visible = False
		Label_Credit_Memo.Visible = False
		Label_Credit_Amt.Visible = False
		
	End Sub
End Class