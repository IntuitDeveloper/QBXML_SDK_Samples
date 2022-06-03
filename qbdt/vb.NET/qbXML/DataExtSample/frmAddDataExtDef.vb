Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmAddDataExtDef
	Inherits System.Windows.Forms.Form
    '----------------------------------------------------------
    ' Form: frmAddDataExtDef
    '
    ' Description: This form allows the user to type in the name of a
    '              new data extension definition, select the type for it,
    '              select the items and transactions to associate it with
    '              and the add it to the currently open QuickBooks company
    '              file.
    '
    ' Copyright © 2021-2022 Intuit Inc. All rights reserved.
    ' Use is subject to the terms specified at:
    '      http://developer.intuit.com/legal/devsite_tos.html
    '
    '----------------------------------------------------------

    Private Sub cmdAdd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAdd.Click
		
		Dim strObjects As String
		
		If txtDataExtName.Text = "" Then
			MsgBox("You must enter a Data Extension Name")
			Exit Sub
		End If
		
		If lstDataExtType.SelectedIndex < 0 Then
			MsgBox("You must select a Data Extension Type")
			Exit Sub
		End If
		
		strObjects = ""
		If chkCustomer.CheckState = 1 Then strObjects = strObjects & " Customer"
		If chkVendor.CheckState = 1 Then strObjects = strObjects & " Vendor"
		If chkEmployee.CheckState = 1 Then strObjects = strObjects & " Employee"
		If chkOtherName.CheckState = 1 Then strObjects = strObjects & " OtherName"
		If chkItem.CheckState = 1 Then strObjects = strObjects & " Item"
		If chkAccount.CheckState = 1 Then strObjects = strObjects & " Account"
		If chkBill.CheckState = 1 Then strObjects = strObjects & " Bill"
		If chkBillPaymentCheck.CheckState = 1 Then strObjects = strObjects & " BillPaymentCheck"
		If chkBillPaymentCreditCard.CheckState = 1 Then strObjects = strObjects & " BillPaymentCreditCard"
		If chkCharge.CheckState = 1 Then strObjects = strObjects & " Charge"
		If chkCheck.CheckState = 1 Then strObjects = strObjects & " Check"
		If chkCreditCardCharge.CheckState = 1 Then strObjects = strObjects & " CreditCardCharge"
		If chkCreditCardCredit.CheckState = 1 Then strObjects = strObjects & " CreditCardCredit"
		If chkCreditMemo.CheckState = 1 Then strObjects = strObjects & " CreditMemo"
		If chkDeposit.CheckState = 1 Then strObjects = strObjects & " Deposit"
		If chkEstimate.CheckState = 1 Then strObjects = strObjects & " Estimate"
		If chkInventoryAdjustment.CheckState = 1 Then strObjects = strObjects & " InventoryAdjustment"
		If chkInvoice.CheckState = 1 Then strObjects = strObjects & " Invoice"
		If chkJournalEntry.CheckState = 1 Then strObjects = strObjects & " JournalEntry"
		If chkPurchaseOrder.CheckState = 1 Then strObjects = strObjects & " PurchaseOrder"
		If chkReceivePayment.CheckState = 1 Then strObjects = strObjects & " ReceivePayment"
		If chkSalesReceipt.CheckState = 1 Then strObjects = strObjects & " SalesReceipt"
		If chkSalesTaxPaymentCheck.CheckState = 1 Then strObjects = strObjects & " SalesTaxPaymentCheck"
		If chkVendorCredit.CheckState = 1 Then strObjects = strObjects & " VendorCredit"
		
		If strObjects = "" Then
			MsgBox("You need to pick at least one target for the data extension")
			Exit Sub
		End If
		
		'Remove the leading space from strObjects
		strObjects = VB.Right(strObjects, Len(strObjects) - 1)
		
		AddDataExtDef((txtDataExtName.Text), VB6.GetItemString(lstDataExtType, lstDataExtType.SelectedIndex), strObjects)
		
		MsgBox("Succesfully added Data Extension Definition """ & txtDataExtName.Text & """ of type """ & VB6.GetItemString(lstDataExtType, lstDataExtType.SelectedIndex) & """ to objects " & strObjects)
		
		cmdShowRequest.Enabled = True
		cmdShowResponse.Enabled = True
	End Sub
	
	Private Sub cmdAll_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAll.Click
		chkCustomer.CheckState = System.Windows.Forms.CheckState.Checked
		chkVendor.CheckState = System.Windows.Forms.CheckState.Checked
		chkEmployee.CheckState = System.Windows.Forms.CheckState.Checked
		chkOtherName.CheckState = System.Windows.Forms.CheckState.Checked
		chkItem.CheckState = System.Windows.Forms.CheckState.Checked
		chkAccount.CheckState = System.Windows.Forms.CheckState.Checked
		chkBill.CheckState = System.Windows.Forms.CheckState.Checked
		chkBillPaymentCheck.CheckState = System.Windows.Forms.CheckState.Checked
		chkBillPaymentCreditCard.CheckState = System.Windows.Forms.CheckState.Checked
		chkCharge.CheckState = System.Windows.Forms.CheckState.Checked
		chkCheck.CheckState = System.Windows.Forms.CheckState.Checked
		chkCreditCardCharge.CheckState = System.Windows.Forms.CheckState.Checked
		chkCreditCardCredit.CheckState = System.Windows.Forms.CheckState.Checked
		chkCreditMemo.CheckState = System.Windows.Forms.CheckState.Checked
		chkDeposit.CheckState = System.Windows.Forms.CheckState.Checked
		chkEstimate.CheckState = System.Windows.Forms.CheckState.Checked
		chkInventoryAdjustment.CheckState = System.Windows.Forms.CheckState.Checked
		chkInvoice.CheckState = System.Windows.Forms.CheckState.Checked
		chkJournalEntry.CheckState = System.Windows.Forms.CheckState.Checked
		chkPurchaseOrder.CheckState = System.Windows.Forms.CheckState.Checked
		chkReceivePayment.CheckState = System.Windows.Forms.CheckState.Checked
		chkSalesReceipt.CheckState = System.Windows.Forms.CheckState.Checked
		chkSalesTaxPaymentCheck.CheckState = System.Windows.Forms.CheckState.Checked
		chkVendorCredit.CheckState = System.Windows.Forms.CheckState.Checked
	End Sub
	
	Private Sub cmdAllEntities_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAllEntities.Click
		chkCustomer.CheckState = System.Windows.Forms.CheckState.Checked
		chkVendor.CheckState = System.Windows.Forms.CheckState.Checked
		chkEmployee.CheckState = System.Windows.Forms.CheckState.Checked
		chkOtherName.CheckState = System.Windows.Forms.CheckState.Checked
	End Sub
	
	Private Sub cmdAllTransactions_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAllTransactions.Click
		chkBill.CheckState = System.Windows.Forms.CheckState.Checked
		chkBillPaymentCheck.CheckState = System.Windows.Forms.CheckState.Checked
		chkBillPaymentCreditCard.CheckState = System.Windows.Forms.CheckState.Checked
		chkCharge.CheckState = System.Windows.Forms.CheckState.Checked
		chkCheck.CheckState = System.Windows.Forms.CheckState.Checked
		chkCreditCardCharge.CheckState = System.Windows.Forms.CheckState.Checked
		chkCreditCardCredit.CheckState = System.Windows.Forms.CheckState.Checked
		chkCreditMemo.CheckState = System.Windows.Forms.CheckState.Checked
		chkDeposit.CheckState = System.Windows.Forms.CheckState.Checked
		chkEstimate.CheckState = System.Windows.Forms.CheckState.Checked
		chkInventoryAdjustment.CheckState = System.Windows.Forms.CheckState.Checked
		chkInvoice.CheckState = System.Windows.Forms.CheckState.Checked
		chkJournalEntry.CheckState = System.Windows.Forms.CheckState.Checked
		chkPurchaseOrder.CheckState = System.Windows.Forms.CheckState.Checked
		chkReceivePayment.CheckState = System.Windows.Forms.CheckState.Checked
		chkSalesReceipt.CheckState = System.Windows.Forms.CheckState.Checked
		chkSalesTaxPaymentCheck.CheckState = System.Windows.Forms.CheckState.Checked
		chkVendorCredit.CheckState = System.Windows.Forms.CheckState.Checked
	End Sub
	
	Private Sub cmdClearAll_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClearAll.Click
		chkCustomer.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkVendor.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkEmployee.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkOtherName.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkItem.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkAccount.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkBill.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkBillPaymentCheck.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkBillPaymentCreditCard.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkCharge.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkCheck.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkCreditCardCharge.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkCreditCardCredit.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkCreditMemo.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkDeposit.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkEstimate.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkInventoryAdjustment.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkInvoice.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkJournalEntry.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkPurchaseOrder.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkReceivePayment.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkSalesReceipt.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkSalesTaxPaymentCheck.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkVendorCredit.CheckState = System.Windows.Forms.CheckState.Unchecked
	End Sub
	
	Private Sub cmdQuit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdQuit.Click
		Me.Close()
	End Sub
	
	Private Sub cmdShowRequest_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdShowRequest.Click
		frmqbXMLDisplay.txtqbXML.Text = PrettyXMLString((modDataExtSample.strAddRequest))
        frmqbXMLDisplay.ShowDialog()
    End Sub
	
	Private Sub cmdShowResponse_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdShowResponse.Click
		frmqbXMLDisplay.txtqbXML.Text = PrettyXMLString((modDataExtSample.strAddResponse))
        frmqbXMLDisplay.ShowDialog()
    End Sub
	
	Private Sub frmAddDataExtDef_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'Set up the data extension type list box with the accepted types
		lstDataExtType.Items.Clear()
		lstDataExtType.Items.Add("INTTYPE")
		lstDataExtType.Items.Add("AMTTYPE")
		lstDataExtType.Items.Add("PRICETYPE")
		lstDataExtType.Items.Add("QUANTYPE")
		lstDataExtType.Items.Add("PERCENTTYPE")
		lstDataExtType.Items.Add("DATETIMETYPE")
		lstDataExtType.Items.Add("STR255TYPE")
		lstDataExtType.Items.Add("STR1024TYPE")
	End Sub
End Class