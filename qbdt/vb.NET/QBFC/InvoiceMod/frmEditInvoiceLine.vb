Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmEditInvoiceLine
	Inherits System.Windows.Forms.Form
	'----------------------------------------------------------
	' Copyright © 2003-2013 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/rdmgr/?ID=100
	'
	'----------------------------------------------------------
	
	
	'Variables to track changes made on the form
	Dim booItemChanged As Boolean
	Dim booDescChanged As Boolean
	Dim booQuantityChanged As Boolean
	Dim booRatePercentChanged As Boolean
	Dim booRateChanged As Boolean
	Dim booAmountChanged As Boolean
	Dim booClassChanged As Boolean
	Dim booServiceDateChanged As Boolean
	Dim booSalesTaxCodeChanged As Boolean
	Dim booOverrideItemAccountChanged As Boolean
	
	Dim booGroupLine As Boolean
	
	'Variables to keep the type and state of the line we're editing
	Dim strLineType As String
	Dim strLineState As String
	
	'Variable to track the type of line the content of the items combo box
	'is currently set up for
	Dim strItemListType As String
	
	Private Sub frmEditInvoiceLine_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		ClearForm()
		
		frmPatience.Show()
		FillComboBox(cmbClass, "Class", "FullName", "", False)
		FillComboBox(cmbItem, "ItemService,ItemInventory,ItemInventoryAssembly,ItemNonInventory,ItemOtherCharge,ItemSubtotal,ItemGroup,ItemDiscount,ItemPayment,ItemSalesTax", "FullName,FullName,FullName,FullName,FullName,Name,Name,FullName,Name,Name", "", True)
		strItemListType = "Item"
		
		FillComboBox(cmbSalesTaxCode, "SalesTaxCode", "Name", "", False)
		FillComboBox(cmbOverrideItemAccount, "Account", "FullName", "<AccountType>Income</AccountType>", False)
		frmPatience.Hide()
		
		DisplayLineInfo()
	End Sub
	
	
	'UPGRADE_WARNING: Form event frmEditInvoiceLine.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmEditInvoiceLine_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		ClearForm()
		
		DisplayLineInfo()
	End Sub
	
	
	Private Sub cmdFinish_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdFinish.Click
		'Don't save any changes, just unload the form and return to the invoice
		'modify form
		Me.Hide()
	End Sub
	
	
	Private Sub cmdSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSave.Click
		If AnyValueChanged Then SaveLineInfo()
		Me.Hide()
	End Sub
	
	
	Private Sub cmdUndo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdUndo.Click
		ClearForm()
		DisplayLineInfo()
	End Sub
	
	
	'UPGRADE_WARNING: Event cmbClass.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	'UPGRADE_WARNING: ComboBox event cmbClass.Change was upgraded to cmbClass.TextChanged which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
	Private Sub cmbClass_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbClass.TextChanged
		booClassChanged = True
	End Sub
	
	'UPGRADE_WARNING: Event cmbItem.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cmbItem_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbItem.SelectedIndexChanged
		booItemChanged = True
		
		If Not booGroupLine Then
			If InStr(1, cmbItem.Text, " - Group Item") > 0 Then
				SetFormForItemLine(False)
			ElseIf txtDesc.ReadOnly = True Then 
				SetFormForItemLine(True)
			End If
		End If
		
		Dim strDesc As String
		Dim strRate As String
		Dim strRatePercent As String
		Dim strSalesTaxCode As String
		
		GetItemInfo(Replace(cmbItem.Text, " - Group Item", ""), strDesc, strRate, strRatePercent, strSalesTaxCode)
		
		Dim strDescSplits() As String
		Dim i As Short
		
		strDescSplits = Split(strDesc, Chr(10))
		
		If UBound(strDescSplits) > 0 Then
			txtDesc.Text = strDescSplits(0)
			For i = 1 To UBound(strDescSplits)
				txtDesc.Text = txtDesc.Text & vbCrLf & strDescSplits(i)
			Next i
		Else
			txtDesc.Text = strDesc
		End If
		txtSavedDesc.Text = txtDesc.Text
		
		txtRate.Text = strRate
		cmbSalesTaxCode.Text = strSalesTaxCode
		If strRatePercent = "Rate" Then
			optRate.Checked = True
			optRatePercent.Checked = False
		Else
			optRate.Checked = False
			optRatePercent.Checked = True
		End If
	End Sub
	
	'UPGRADE_WARNING: Event cmbOverrideItemAccount.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	'UPGRADE_WARNING: ComboBox event cmbOverrideItemAccount.Change was upgraded to cmbOverrideItemAccount.TextChanged which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
	Private Sub cmbOverrideItemAccount_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbOverrideItemAccount.TextChanged
		booOverrideItemAccountChanged = True
	End Sub
	
	'UPGRADE_WARNING: Event cmbSalesTaxCode.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	'UPGRADE_WARNING: ComboBox event cmbSalesTaxCode.Change was upgraded to cmbSalesTaxCode.TextChanged which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
	Private Sub cmbSalesTaxCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbSalesTaxCode.TextChanged
		booSalesTaxCodeChanged = True
	End Sub
	
	'UPGRADE_WARNING: Event optRate.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optRate_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optRate.CheckedChanged
		If eventSender.Checked Then
			booRatePercentChanged = True
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optRatePercent.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optRatePercent_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optRatePercent.CheckedChanged
		If eventSender.Checked Then
			booRatePercentChanged = True
		End If
	End Sub
	
	Private Sub txtAmount_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtAmount.Validating
		Dim Cancel As Boolean = eventArgs.Cancel
		
		txtAmount.Text = Trim(txtAmount.Text)
		
		Dim strSplits() As String
		strSplits = Split(GetInvoiceLineInfo, "<spliter>")
		
		Dim strOldAmount As String
		
		If UBound(strSplits) = 0 Then
			strOldAmount = ""
		Else
			strOldAmount = strSplits(5)
		End If
        'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If String.IsNullOrEmpty(Trim(txtAmount.Text)) Then
            'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            If Not String.IsNullOrEmpty(Trim(strOldAmount)) Then booAmountChanged = True
            GoTo EventExitSub
        End If

        'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If (Not IsNumeric(txtAmount.Text)) And (String.IsNullOrEmpty(Trim(txtAmount.Text))) Then
            MsgBox("Amount must be numeric, please re-enter Amount")
            txtAmount.Text = strOldAmount
            txtAmount.Focus()
            booAmountChanged = False
            GoTo EventExitSub
        ElseIf Not InStr(1, txtAmount.Text, ".") > 0 Then
            txtAmount.Text = txtAmount.Text & ".00"
        ElseIf Len(txtAmount.Text) - InStr(1, txtAmount.Text, ".") = 1 Then
            txtAmount.Text = txtAmount.Text & "0"
        ElseIf Len(txtAmount.Text) - InStr(1, txtAmount.Text, ".") > 2 Then
            MsgBox("Amount can only have two decimal places")
			txtAmount.Text = VB.Left(txtAmount.Text, Len(txtAmount.Text) + 3 - InStr(1, txtAmount.Text, "."))
		End If
		
		booAmountChanged = Trim(txtAmount.Text) <> Trim(strOldAmount)

        'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If String.IsNullOrEmpty(Trim(txtRate.Text)) Then GoTo EventExitSub

        If CDbl(txtRate.Text) = 0 Then
			txtRate.Text = CStr(Nothing)
			GoTo EventExitSub
		End If
		
		Dim objMsgBoxAnswer As MsgBoxResult
        'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If Not String.IsNullOrEmpty(Trim(txtRate.Text)) And chkDisable.CheckState = 0 And booAmountChanged Then
            objMsgBoxAnswer = MsgBox("Changing Amount requires Rate to be cleared" & vbCrLf & vbCrLf & "Do you still wish to change Amount and clear Rate?", MsgBoxStyle.YesNo)
        End If

        If (objMsgBoxAnswer = MsgBoxResult.Yes Or chkDisable.CheckState <> 0) And booAmountChanged Then
			txtRate.Text = CStr(Nothing)
			booRateChanged = True
			optRate.Checked = 1
			optRatePercent.Checked = 0
		Else
			txtAmount.Text = Trim(strOldAmount)
		End If
EventExitSub: 
		eventArgs.Cancel = Cancel
	End Sub
	
	'UPGRADE_WARNING: Event txtDesc.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtDesc_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDesc.TextChanged
		booDescChanged = True
	End Sub
	
	Private Sub txtQuantity_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtQuantity.Validating
		Dim Cancel As Boolean = eventArgs.Cancel
		
		Dim strOldQuantity As String
		
		txtQuantity.Text = Trim(txtQuantity.Text)
		
		Dim strSplits() As String
		strSplits = Split(GetInvoiceLineInfo, "<spliter>")
		
		If UBound(strSplits) = 0 Then
			strOldQuantity = ""
		Else
			strOldQuantity = strSplits(1)
		End If

        'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If String.IsNullOrEmpty(Trim(txtQuantity.Text)) Then
            'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            If Not String.IsNullOrEmpty(Trim(strOldQuantity)) Then booQuantityChanged = True
            GoTo EventExitSub
        End If

        'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If (Not IsNumeric(txtQuantity.Text)) And (Not String.IsNullOrEmpty(Trim(txtQuantity.Text))) Then
            MsgBox("Quantity must be numeric, please re-enter Quantity")
            txtQuantity.Text = strOldQuantity
            txtQuantity.Focus()
            booQuantityChanged = False
            GoTo EventExitSub
        End If

        booQuantityChanged = txtQuantity.Text <> Trim(strOldQuantity)

        'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If String.IsNullOrEmpty(Trim(txtAmount.Text)) Then GoTo EventExitSub

        Dim objMsgBoxAnswer As MsgBoxResult
        'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If Not String.IsNullOrEmpty(Trim(txtRate.Text)) And chkDisable.CheckState = 0 And booQuantityChanged Then
            objMsgBoxAnswer = MsgBox("Changing Quantity requires Amount to be cleared" & vbCrLf & vbCrLf & "Do you still wish to change Quantity and clear Amount?", MsgBoxStyle.YesNo)
        End If

        If (objMsgBoxAnswer = MsgBoxResult.Yes Or chkDisable.CheckState <> 0) And booQuantityChanged Then
			txtAmount.Text = CStr(Nothing)
			booAmountChanged = True
		Else
			txtQuantity.Text = Trim(strOldQuantity)
		End If
EventExitSub: 
		eventArgs.Cancel = Cancel
	End Sub
	
	Private Sub txtRate_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtRate.Validating
		Dim Cancel As Boolean = eventArgs.Cancel
		txtRate.Text = Trim(txtRate.Text)
		
		Dim strSplits() As String
		strSplits = Split(GetInvoiceLineInfo, "<spliter>")
		
		Dim strOldRate As String
		
		If UBound(strSplits) = 0 Then
			strOldRate = ""
		Else
			strOldRate = strSplits(4)
		End If

        'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If String.IsNullOrEmpty(Trim(txtRate.Text)) Then
            'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            If Not String.IsNullOrEmpty(Trim(strOldRate)) Then booRateChanged = True
            GoTo EventExitSub
        End If

        If Not IsNumeric(txtRate.Text) Then
			MsgBox("Rate must be numeric, please re-enter Rate")
			txtRate.Text = strOldRate
			txtRate.Focus()
			booRateChanged = False
			GoTo EventExitSub
		End If
		
		booRateChanged = txtRate.Text <> Trim(strOldRate)

        'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If String.IsNullOrEmpty(Trim(txtAmount.Text)) Then GoTo EventExitSub

        Dim objMsgBoxAnswer As MsgBoxResult
        'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If Not String.IsNullOrEmpty(Trim(txtRate.Text)) And Not IsNothing(Trim(txtAmount.Text)) And chkDisable.CheckState = 0 And booRateChanged Then
            objMsgBoxAnswer = MsgBox("Changing Rate requires Amount to be cleared" & vbCrLf & vbCrLf & "Do you still wish to change Rate and clear Amount?", MsgBoxStyle.YesNo)
        End If

        If (objMsgBoxAnswer = MsgBoxResult.Yes Or chkDisable.CheckState <> 0) And booRateChanged Then
			txtAmount.Text = CStr(Nothing)
			booAmountChanged = True
		Else
			txtRate.Text = Trim(strOldRate)
		End If
EventExitSub: 
		eventArgs.Cancel = Cancel
	End Sub
	
	'UPGRADE_WARNING: Event txtServiceDate.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtServiceDate_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtServiceDate.TextChanged
		booServiceDateChanged = True
	End Sub
	
	Private Sub txtDesc_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDesc.Click
		If txtDesc.ReadOnly Then MsgBox("You cannot change the description of a group item")
	End Sub
	
	
	Private Sub ClearForm()
		txtTxnLineID.Text = CStr(Nothing)
		cmbItem.Text = CStr(Nothing)
		txtDesc.Text = CStr(Nothing)
		txtQuantity.Text = CStr(Nothing)
		optRate.Checked = False
		optRatePercent.Checked = False
		txtRate.Text = CStr(Nothing)
		txtAmount.Text = CStr(Nothing)
		cmbClass.Text = CStr(Nothing)
		txtServiceDate.Text = CStr(Nothing)
		cmbSalesTaxCode.Text = CStr(Nothing)
		cmbOverrideItemAccount.Text = CStr(Nothing)
		
		booItemChanged = False
		booDescChanged = False
		booQuantityChanged = False
		booRatePercentChanged = False
		booRateChanged = False
		booAmountChanged = False
		booClassChanged = False
		booServiceDateChanged = False
		booSalesTaxCodeChanged = False
		booOverrideItemAccountChanged = False
		
		SetFormForItemLine(True)
	End Sub
	
	
	Private Sub DisplayLineInfo()
		
		Dim strPassedLine As String
		Dim strSplits() As String
		
		'We could be editing a new line in which case the invoice line info will
		'be empty and we don't have anything to display except for a -1 for the
		'TxnLineID
		strPassedLine = GetInvoiceLineInfo
		If VB.Right(strPassedLine, 4) = ",New" Then
			strLineType = VB.Left(strPassedLine, InStr(1, strPassedLine, ",") - 1)
			strLineState = "New"
			txtTxnLineID.Text = "-1 (New " & strLineType & ")"
			booGroupLine = False
			
			If strLineType = "SubItem" And strItemListType <> "SubItem" Then
				FillComboBox(cmbItem, "ItemService,ItemInventory,ItemInventoryAssembly,ItemNonInventory,ItemOtherCharge,ItemSubtotal,ItemDiscount,ItemPayment,ItemSalesTax", "FullName,FullName,FullName,FullName,FullName,Name,FullName,Name,Name", "", False)
				strItemListType = "SubItem"
				cmbItem.Refresh()
			End If
			Exit Sub
		End If
		
		strSplits = Split(strPassedLine, "<spliter>")
		
		If strSplits(11) = "SubItem" And strItemListType <> "SubItem" Then
			FillComboBox(cmbItem, "ItemService,ItemInventory,ItemInventoryAssembly,ItemNonInventory,ItemOtherCharge,ItemSubtotal,ItemDiscount,ItemPayment,ItemSalesTax", "FullName,FullName,FullName,FullName,FullName,Name,FullName,Name,Name", "", False)
			strItemListType = "SubItem"
			cmbItem.Refresh()
			booGroupLine = False
		ElseIf strSplits(11) = "Group" And strItemListType <> "Group" Then 
			FillComboBox(cmbItem, "ItemGroup", "Name", "", False)
			strItemListType = "Group"
			cmbItem.Refresh()
			booGroupLine = True
		ElseIf strSplits(11) = "Item" And strItemListType <> "Item" Then 
			FillComboBox(cmbItem, "ItemService,ItemInventory,ItemInventoryAssembly,ItemNonInventory,ItemOtherCharge,ItemSubtotal,ItemGroup,ItemDiscount,ItemPayment,ItemSalesTax", "FullName,FullName,FullName,FullName,FullName,Name,Name,FullName,Name,Name", "", True)
			strItemListType = "Item"
			cmbItem.Refresh()
			booGroupLine = False
		End If
		
		SetFormForItemLine((strSplits(11) <> "Group"))
		
		txtTxnLineID.Text = strSplits(0)
		cmbItem.Text = strSplits(2)
		
		Dim strDescSplits() As String
		Dim i As Short
		
		strDescSplits = Split(strSplits(3), Chr(10))
		
		If UBound(strDescSplits) > 0 Then
			txtDesc.Text = strDescSplits(0)
			For i = 1 To UBound(strDescSplits)
				txtDesc.Text = txtDesc.Text & vbCrLf & strDescSplits(i)
			Next i
		Else
			txtDesc.Text = strSplits(3)
		End If
		txtSavedDesc.Text = txtDesc.Text
		
		txtQuantity.Text = strSplits(1)
		
		If strSplits(9) = "Rate" Then
			optRate.Checked = True
		ElseIf strSplits(9) = "RatePercent" Then 
			optRatePercent.Checked = True
		End If
		
		txtRate.Text = strSplits(4)
		txtAmount.Text = strSplits(5)
		cmbClass.Text = strSplits(6)
		txtServiceDate.Text = strSplits(7)
		cmbSalesTaxCode.Text = strSplits(8)
		
		strLineType = strSplits(11)
		strLineState = strSplits(12)
		If strLineState <> "New" Then strLineState = "Changed"
	End Sub
	
	
	Private Function AnyValueChanged() As Boolean
		
		Dim strSplits() As String
		If booItemChanged Or booDescChanged Or booQuantityChanged Or booRatePercentChanged Or booRateChanged Or booAmountChanged Or booClassChanged Or booServiceDateChanged Or booSalesTaxCodeChanged Or booOverrideItemAccountChanged Then
			
			If VB.Right(GetInvoiceLineInfo, 3) = "New" Then
                'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                If Not String.IsNullOrEmpty(Trim(txtTxnLineID.Text)) Or Not String.IsNullOrEmpty(Trim(txtQuantity.Text)) Or Not String.IsNullOrEmpty(Trim(cmbItem.Text)) Or Not String.IsNullOrEmpty(Trim(txtDesc.Text)) Or Not String.IsNullOrEmpty(Trim(txtRate.Text)) Or Not String.IsNullOrEmpty(Trim(txtAmount.Text)) Or Not String.IsNullOrEmpty(Trim(cmbClass.Text)) Or Not String.IsNullOrEmpty(Trim(txtServiceDate.Text)) Or Not String.IsNullOrEmpty(Trim(cmbSalesTaxCode.Text)) Or optRate.Checked = True Or optRatePercent.Checked = True Or Not String.IsNullOrEmpty(cmbOverrideItemAccount.Text) Then
                    AnyValueChanged = True
                Else
                    AnyValueChanged = False
				End If
				
				Exit Function
			End If
			
			strSplits = Split(GetInvoiceLineInfo, "<spliter>")

            'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            If Trim(txtTxnLineID.Text) <> Trim(strSplits(0)) Or Trim(txtQuantity.Text) <> Trim(strSplits(1)) Or Trim(cmbItem.Text) <> Trim(strSplits(2)) Or Trim(txtDesc.Text) <> Trim(txtSavedDesc.Text) Or Trim(txtRate.Text) <> Trim(strSplits(4)) Or Trim(txtAmount.Text) <> Trim(strSplits(5)) Or Trim(cmbClass.Text) <> Trim(strSplits(6)) Or Trim(txtServiceDate.Text) <> Trim(strSplits(7)) Or Trim(cmbSalesTaxCode.Text) <> Trim(strSplits(8)) Or (strSplits(9) = "Rate" And optRate.Checked <> True) Or (strSplits(9) = "RatePercent" And optRatePercent.Checked <> True) Or Not String.IsNullOrEmpty(cmbOverrideItemAccount.Text) Then
                AnyValueChanged = True
            Else
                AnyValueChanged = False
			End If
		Else
			AnyValueChanged = False
		End If
	End Function
	
	
	Private Sub SaveLineInfo()
		
		Dim strInvoiceLineInfo As String
		
		strInvoiceLineInfo = txtTxnLineID.Text & "<spliter>" & txtQuantity.Text & "<spliter>" & cmbItem.Text & "<spliter>" & Replace(txtDesc.Text, vbCrLf, vbCr) & "<spliter>" & txtRate.Text & "<spliter>" & txtAmount.Text & "<spliter>" & cmbClass.Text & "<spliter>" & txtServiceDate.Text & "<spliter>" & cmbSalesTaxCode.Text & "<spliter>"
		
		
		If optRate.Checked = True Then
			strInvoiceLineInfo = strInvoiceLineInfo & "Rate<spliter>"
		ElseIf optRatePercent.Checked = True Then 
			strInvoiceLineInfo = strInvoiceLineInfo & "RatePercent<spliter>"
		Else
			strInvoiceLineInfo = strInvoiceLineInfo & "Neither<spliter>"
		End If
		
		strInvoiceLineInfo = strInvoiceLineInfo & cmbOverrideItemAccount.Text & "<spliter>" & strLineType & "<spliter>" & strLineState
		
		SetInvoiceLineInfo(strInvoiceLineInfo)
	End Sub
	
	
	Public Sub SetFormForItemLine(ByRef Value As Boolean)
		txtDesc.ReadOnly = Not Value
		optRate.Enabled = Value
		optRatePercent.Enabled = Value
		txtRate.Enabled = Value
		txtAmount.Enabled = Value
		cmbClass.Enabled = Value
		txtServiceDate.Enabled = Value
		cmbSalesTaxCode.Enabled = Value
		cmbOverrideItemAccount.Enabled = Value
		lblAmount.Enabled = Value
		lblClass.Enabled = Value
		lblServiceDate.Enabled = Value
		lblSalesTaxCode.Enabled = Value
		lblOverrideItemAccount.Enabled = Value
	End Sub
End Class