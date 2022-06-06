Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmInvoiceModify
	Inherits System.Windows.Forms.Form
	'----------------------------------------------------------
	' Copyright © 2003-2013 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/rdmgr/?ID=100
	'
	'----------------------------------------------------------
	
	
	Dim strLineArray(200) As String
	Dim strLastCommand As String
	
	Dim booEditInvoiceLineFormLoaded As Boolean
	Dim booCheckForChanges As Boolean
	Dim booCheckForLineChanges As Boolean
	
	Dim intNumInvoiceLines As Short
	Dim intTopDisplayedLine As Short
	Dim intHighlightedLine As Short
	Dim intActualHighlightedLine As Short
	
	Dim strTxnID As String
	Dim strEditSequence As String
	Dim strRefNumber As String
	Dim strTxnDate As String
	Dim intPending As Short
	Dim intToBePrinted As Short
	Dim strCustomer As String
	Dim strClass As String
	Dim strBillAddr1 As String
	Dim strBillAddr2 As String
	Dim strBillAddr3 As String
	Dim strBillAddr4 As String
	Dim strBillCity As String
	Dim strBillState As String
	Dim strBillPostalCode As String
	Dim strBillCountry As String
	Dim strShipAddr1 As String
	Dim strShipAddr2 As String
	Dim strShipAddr3 As String
	Dim strShipAddr4 As String
	Dim strShipCity As String
	Dim strShipState As String
	Dim strShipPostalCode As String
	Dim strShipCountry As String
	Dim strARAccount As String
	Dim strTerms As String
	Dim strPONumber As String
	Dim strDueDate As String
	Dim strShipDate As String
	Dim strFOB As String
	Dim strSalesRep As String
	Dim strShipMethod As String
	Dim strItemSalesTax As String
	Dim strCustTaxCode As String
	Dim strCustomerMsg As String
	Dim strMemo As String
	
	Dim booInvoiceBodyChanged As Boolean
	Dim booBillAddressChanged As Boolean
	Dim booShipAddressChanged As Boolean
	Dim booInvoiceLinesChanged As Boolean
	
	Private Sub frmInvoiceModify_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		strLastCommand = ""
		
		ClearForm()
		
		frmPatience.Show()
		FillComboBox(cmbCustomer, "Customer", "FullName", "", False)
		FillComboBox(cmbClass, "Class", "FullName", "", False)
		FillComboBox(cmbARAccount, "Account", "FullName", "<AccountType>AccountsReceivable</AccountType>", False)
		FillComboBox(cmbTerms, "StandardTerms", "Name", "", False)
		FillComboBox(cmbSalesRep, "SalesRep", "FullName", "", False)
		FillComboBox(cmbShipMethod, "ShipMethod", "Name", "", False)
		FillComboBox(cmbItemSalesTax, "ItemSalesTax,ItemSalesTaxGroup", "Name", "", False)
		FillComboBox(cmbCustTaxCode, "SalesTaxCode", "Name", "", False)
		FillComboBox(cmbCustomerMsg, "CustomerMsg", "Name", "", False)
		frmPatience.Hide()
		
		intHighlightedLine = -1
		intActualHighlightedLine = -1
		
		LoadInvoiceModifyForm()
		SaveOriginalValues()
		LoadInvoiceLineArray(strLineArray)
		CountInvoiceLines()
		DisplayInvoiceLines(1)
		booEditInvoiceLineFormLoaded = False
		booCheckForChanges = False
		booCheckForLineChanges = False
	End Sub
	
	
	'UPGRADE_WARNING: Form event frmInvoiceModify.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmInvoiceModify_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		Dim intOldHighlightedLine As Short
        If strLastCommand = "ReturnToInvoiceSelection" Or strLastCommand = "ModifyInvoice" Then


            ClearForm()

            If intHighlightedLine > -1 Then
                UnHighlightLine(intHighlightedLine)
            End If

            intHighlightedLine = -1
            intActualHighlightedLine = -1

            LoadInvoiceModifyForm()
            SaveOriginalValues()
            LoadInvoiceLineArray(strLineArray)
            CountInvoiceLines()
            If intNumInvoiceLines < 11 Or vscInvoiceLineScroll.Value = 1 Then
                DisplayInvoiceLines(1)
            Else
                vscInvoiceLineScroll.Value = 1
            End If
            booEditInvoiceLineFormLoaded = False
            booCheckForChanges = False
            booCheckForLineChanges = False
        ElseIf strLastCommand = "EditLine" Then
            If GetInvoiceLineInfo() <> strLineArray(intActualHighlightedLine) Then
                strLineArray(intActualHighlightedLine) = GetInvoiceLineInfo()
                DisplayLine(strLineArray(intActualHighlightedLine), intHighlightedLine)
                booCheckForLineChanges = True
            End If
        ElseIf strLastCommand = "NewLineAfter" Then
            If InStr(1, GetInvoiceLineInfo, "<spliter>") = 0 Then Exit Sub

            booCheckForLineChanges = True
            InsertInvoiceLine(GetInvoiceLineInfo, intActualHighlightedLine + 1)

            If intNumInvoiceLines > 10 Then
                If intHighlightedLine = 9 Then
                    vscInvoiceLineScroll.Value = intTopDisplayedLine + 1
                Else
                    DisplayInvoiceLines(intTopDisplayedLine)
                End If
            Else
                DisplayInvoiceLines(intTopDisplayedLine)
            End If

        ElseIf strLastCommand = "NewLineBefore" Then
            If InStr(1, GetInvoiceLineInfo, "<spliter>") = 0 Then Exit Sub
			
			booCheckForLineChanges = True
			InsertInvoiceLine(GetInvoiceLineInfo, intActualHighlightedLine)
			
			intOldHighlightedLine = intHighlightedLine
			UnHighlightLine(intHighlightedLine)
			intActualHighlightedLine = intActualHighlightedLine + 1
			
			If intNumInvoiceLines > 10 Then
				If intOldHighlightedLine = 9 Then
					vscInvoiceLineScroll.Value = intTopDisplayedLine + 1
				Else
					DisplayInvoiceLines(intTopDisplayedLine)
				End If
			Else
				DisplayInvoiceLines(intTopDisplayedLine)
			End If
			
		End If
	End Sub
	
	Private Sub cmdDoModify_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDoModify.Click
		
		If Not InvoiceChanged Then
			If MsgBox(vbCrLf & "This invoice appears to be unchanged" & vbCrLf & vbCrLf & "Close this window and return to the invoice selection window?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
				Me.Close()
			Else
				Exit Sub
			End If
		End If
		
		strLastCommand = "ModifyInvoice"
		frmModifying.Show()
		ModifyInvoice(InvoiceChangeString)
		frmModifying.Hide()
		If chkShow.CheckState = 1 Then
            'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
            frmShowRequest.ShowDialog()
            Me.Hide()
		Else
			Me.Hide()
		End If
	End Sub
	
	Private Sub cmdUndo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdUndo.Click
		UnHighlightLine(intHighlightedLine)
		intActualHighlightedLine = -1
		ClearForm()
		ReloadInvoiceModifyForm()
		LoadInvoiceLineArray(strLineArray)
		CountInvoiceLines()
		DisplayInvoiceLines(1)
		booCheckForChanges = False
		booCheckForLineChanges = False
	End Sub
	
	Private Sub cmdDeleteUndelete_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDeleteUndelete.Click
		
		booCheckForLineChanges = True
		
		Dim strAction As String
		Dim strSplits() As String
		strSplits = Split(strLineArray(intActualHighlightedLine), "<spliter>")
		
		Dim strPrevSplits() As String
		Dim i As Short
		If strSplits(11) = "SubItem" Then
			
			i = intActualHighlightedLine - 1
			Do 
				strPrevSplits = Split(strLineArray(i), "<spliter>")
				i = i - 1
			Loop Until strPrevSplits(11) = "Group"
			
			If InStr(1, strPrevSplits(12), "Deleted") > 0 Then
				MsgBox("You can't undelete a subitem in a deleted group item")
				Exit Sub
			End If
		End If
		
		If InStr(1, strSplits(12), "Deleted") > 0 Then
			strLineArray(intActualHighlightedLine) = VB.Left(strLineArray(intActualHighlightedLine), Len(strLineArray(intActualHighlightedLine)) - 7)
			cmdEditLine.Enabled = True
			cmdDeleteUndelete.Text = "Delete Line"
			strAction = "Undelete"
		Else
			strLineArray(intActualHighlightedLine) = strLineArray(intActualHighlightedLine) & "Deleted"
			cmdEditLine.Enabled = False
			cmdDeleteUndelete.Text = "Undelete Line"
			strAction = "Delete"
		End If
		
		DisplayLine(strLineArray(intActualHighlightedLine), intHighlightedLine)
		
		If strSplits(11) = "Group" Then ChangeSubLines(intActualHighlightedLine, strAction)
	End Sub
	
	
	Private Sub cmdEditLine_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdEditLine.Click
		
		strLastCommand = "EditLine"
		
		SetInvoiceLineInfo(strLineArray(intActualHighlightedLine))

        'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
        ' Dim form As New frmEditInvoiceLine
        If Not booEditInvoiceLineFormLoaded Then frmEditInvoiceLine.ShowDialog() 'Form.Show() 'Load(frmEditInvoiceLine)

    End Sub
	
	
	Private Sub cmdAddLineAfter_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAddLineAfter.Click
		strLastCommand = "NewLineAfter"
		
		If intNumInvoiceLines = 200 Then
			MsgBox("This program is limited to 200 invoice lines and sublines" & vbCrLf & vbCrLf & "You can't add any more lines")
			Exit Sub
		End If
		
		Dim strSplits() As String
		Dim strItemType As String
		Dim strNextItemType As String
		Dim booNewItem As Boolean
		
		strSplits = Split(strLineArray(intActualHighlightedLine), "<spliter>")
		strItemType = strSplits(11)
		booNewItem = (InStr(1, strSplits(12), "New") > 0)
		
		If intActualHighlightedLine = intNumInvoiceLines Then
			strNextItemType = ""
		Else
			strSplits = Split(strLineArray(intActualHighlightedLine + 1), "<spliter>")
			strNextItemType = strSplits(11)
		End If
		
		Dim vmbResponse As MsgBoxResult
		If (strItemType = "SubItem" Or (strItemType = "Group" And Not booNewItem)) Then
			If strNextItemType = "SubItem" Then
				If InDeletedGroup(intActualHighlightedLine) Then
					MsgBox("You can't add a new sub line to a group that's been deleted")
					Exit Sub
				Else
					SetInvoiceLineInfo("SubItem,New")
				End If
			Else
				vmbResponse = MsgBoxResult.No
				If Not InDeletedGroup(intActualHighlightedLine) Then
					vmbResponse = MsgBox(vbCrLf & "Add new line as a group sub line?", MsgBoxStyle.YesNo)
				End If
				
				If vmbResponse = MsgBoxResult.Yes Then
					SetInvoiceLineInfo("SubItem,New")
				Else
					SetInvoiceLineInfo("Item,New")
				End If
			End If
		Else
			SetInvoiceLineInfo("Item,New")
		End If

        'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
        If Not booEditInvoiceLineFormLoaded Then frmEditInvoiceLine.ShowDialog()
    End Sub
	
	
	Private Sub cmdAddLineBefore_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAddLineBefore.Click
		strLastCommand = "NewLineBefore"
		
		If intNumInvoiceLines = 200 Then
			MsgBox("This program is limited to 200 invoice lines and sublines" & vbCrLf & vbCrLf & "You can't add any more lines")
			Exit Sub
		End If
		
		Dim strSplits() As String
		Dim vmbResponse As MsgBoxResult
		If intActualHighlightedLine = 1 Then
			SetInvoiceLineInfo("Item,New")
		Else
			strSplits = Split(strLineArray(intActualHighlightedLine), "<spliter>")
            If (Not String.IsNullOrEmpty(strSplits(0))) Then

                If strSplits(11) = "SubItem" Then
                    If InDeletedGroup(intActualHighlightedLine) Then
                        MsgBox("You can't add a new sub line to a group that's been deleted")
                        Exit Sub
                    Else
                        SetInvoiceLineInfo("SubItem,New")
                    End If
                Else
                    strSplits = Split(strLineArray(intActualHighlightedLine - 1), "<spliter>")
                    If (strSplits(11) = "SubItem" And Not InDeletedGroup(intActualHighlightedLine - 1)) Or (strSplits(11) = "Group" And InStr(1, strSplits(12), "New") = 0) Then

                        vmbResponse = MsgBox(vbCrLf & "Add new line as a group sub line?", MsgBoxStyle.YesNo)
                        If vmbResponse = MsgBoxResult.Yes Then
                            SetInvoiceLineInfo("SubItem,New")
                        Else
                            SetInvoiceLineInfo("Item,New")
                        End If
                    Else
                        SetInvoiceLineInfo("Item,New")
                    End If
                End If 'strSplits(11) = "SubItem"
            End If
        End If

            'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
            If Not booEditInvoiceLineFormLoaded Then frmEditInvoiceLine.ShowDialog()
    End Sub


    Private Sub cmdEditMemo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdEditMemo.Click
		booCheckForChanges = True
	End Sub
	
	
	Private Sub cmdFinish_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdFinish.Click
		strLastCommand = "ReturnToInvoiceSelection"
		Me.Hide()
	End Sub
	
	
	'UPGRADE_WARNING: Event chkPending.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkPending_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkPending.CheckStateChanged
		booCheckForChanges = True
	End Sub
	
	'UPGRADE_WARNING: Event chkToBePrinted.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkToBePrinted_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkToBePrinted.CheckStateChanged
		booCheckForChanges = True
	End Sub
	
	'UPGRADE_WARNING: Event cmbARAccount.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cmbARAccount_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbARAccount.SelectedIndexChanged
		booCheckForChanges = True
	End Sub
	
	'UPGRADE_WARNING: Event cmbARAccount.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	'UPGRADE_WARNING: ComboBox event cmbARAccount.Change was upgraded to cmbARAccount.TextChanged which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
	Private Sub cmbARAccount_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbARAccount.TextChanged
		booCheckForChanges = True
	End Sub
	
	'UPGRADE_WARNING: Event cmbClass.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cmbClass_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbClass.SelectedIndexChanged
		booCheckForChanges = True
	End Sub
	
	'UPGRADE_WARNING: Event cmbClass.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	'UPGRADE_WARNING: ComboBox event cmbClass.Change was upgraded to cmbClass.TextChanged which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
	Private Sub cmbClass_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbClass.TextChanged
		booCheckForChanges = True
	End Sub
	
	'UPGRADE_WARNING: Event cmbCustomer.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cmbCustomer_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbCustomer.SelectedIndexChanged
		booCheckForChanges = True
	End Sub
	
	'UPGRADE_WARNING: Event cmbCustomer.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	'UPGRADE_WARNING: ComboBox event cmbCustomer.Change was upgraded to cmbCustomer.TextChanged which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
	Private Sub cmbCustomer_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbCustomer.TextChanged
		booCheckForChanges = True
	End Sub
	
	'UPGRADE_WARNING: Event cmbCustomerMsg.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cmbCustomerMsg_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbCustomerMsg.SelectedIndexChanged
		booCheckForChanges = True
	End Sub
	
	'UPGRADE_WARNING: Event cmbCustomerMsg.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	'UPGRADE_WARNING: ComboBox event cmbCustomerMsg.Change was upgraded to cmbCustomerMsg.TextChanged which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
	Private Sub cmbCustomerMsg_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbCustomerMsg.TextChanged
		booCheckForChanges = True
	End Sub
	
	'UPGRADE_WARNING: Event cmbCustTaxCode.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cmbCustTaxCode_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbCustTaxCode.SelectedIndexChanged
		booCheckForChanges = True
	End Sub
	
	'UPGRADE_WARNING: Event cmbCustTaxCode.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	'UPGRADE_WARNING: ComboBox event cmbCustTaxCode.Change was upgraded to cmbCustTaxCode.TextChanged which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
	Private Sub cmbCustTaxCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbCustTaxCode.TextChanged
		booCheckForChanges = True
	End Sub
	
	'UPGRADE_WARNING: Event cmbItemSalesTax.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cmbItemSalesTax_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbItemSalesTax.SelectedIndexChanged
		booCheckForChanges = True
	End Sub
	
	'UPGRADE_WARNING: Event cmbItemSalesTax.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	'UPGRADE_WARNING: ComboBox event cmbItemSalesTax.Change was upgraded to cmbItemSalesTax.TextChanged which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
	Private Sub cmbItemSalesTax_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbItemSalesTax.TextChanged
		booCheckForChanges = True
	End Sub
	
	'UPGRADE_WARNING: Event cmbSalesRep.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cmbSalesRep_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbSalesRep.SelectedIndexChanged
		booCheckForChanges = True
	End Sub
	
	'UPGRADE_WARNING: Event cmbSalesRep.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	'UPGRADE_WARNING: ComboBox event cmbSalesRep.Change was upgraded to cmbSalesRep.TextChanged which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
	Private Sub cmbSalesRep_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbSalesRep.TextChanged
		booCheckForChanges = True
	End Sub
	
	'UPGRADE_WARNING: Event cmbShipMethod.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cmbShipMethod_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbShipMethod.SelectedIndexChanged
		booCheckForChanges = True
	End Sub
	
	'UPGRADE_WARNING: Event cmbShipMethod.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	'UPGRADE_WARNING: ComboBox event cmbShipMethod.Change was upgraded to cmbShipMethod.TextChanged which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
	Private Sub cmbShipMethod_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbShipMethod.TextChanged
		booCheckForChanges = True
	End Sub
	
	'UPGRADE_WARNING: Event cmbTerms.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cmbTerms_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbTerms.SelectedIndexChanged
		booCheckForChanges = True
	End Sub
	
	'UPGRADE_WARNING: Event cmbTerms.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	'UPGRADE_WARNING: ComboBox event cmbTerms.Change was upgraded to cmbTerms.TextChanged which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
	Private Sub cmbTerms_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbTerms.TextChanged
		booCheckForChanges = True
	End Sub
	
	Private Sub txtAmount_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAmount.Click
		Dim Index As Short = txtAmount.GetIndex(eventSender)
		HighlightLine(Index)
	End Sub
	
	'UPGRADE_WARNING: Event txtBillAddr1.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtBillAddr1_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBillAddr1.TextChanged
		booCheckForChanges = True
	End Sub
	
	'UPGRADE_WARNING: Event txtBillAddr2.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtBillAddr2_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBillAddr2.TextChanged
		booCheckForChanges = True
	End Sub
	
	'UPGRADE_WARNING: Event txtBillAddr3.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtBillAddr3_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBillAddr3.TextChanged
		booCheckForChanges = True
	End Sub
	
	'UPGRADE_WARNING: Event txtBillAddr4.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtBillAddr4_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBillAddr4.TextChanged
		booCheckForChanges = True
	End Sub
	
	'UPGRADE_WARNING: Event txtBillCity.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtBillCity_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBillCity.TextChanged
		booCheckForChanges = True
	End Sub
	
	'UPGRADE_WARNING: Event txtBillCountry.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtBillCountry_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBillCountry.TextChanged
		booCheckForChanges = True
	End Sub
	
	'UPGRADE_WARNING: Event txtBillPostalCode.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtBillPostalCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBillPostalCode.TextChanged
		booCheckForChanges = True
	End Sub
	
	'UPGRADE_WARNING: Event txtBillState.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtBillState_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBillState.TextChanged
		booCheckForChanges = True
	End Sub
	
	Private Sub txtDesc_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDesc.Click
		Dim Index As Short = txtDesc.GetIndex(eventSender)
		HighlightLine(Index)
	End Sub
	
	'UPGRADE_WARNING: Event txtDueDate.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtDueDate_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDueDate.TextChanged
		booCheckForChanges = True
	End Sub
	
	'UPGRADE_WARNING: Event txtFOB.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtFOB_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFOB.TextChanged
		booCheckForChanges = True
	End Sub
	
	Private Sub txtItemFullName_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtItemFullName.Click
		Dim Index As Short = txtItemFullName.GetIndex(eventSender)
		HighlightLine(Index)
	End Sub
	
	Private Sub txtItemLineNumber_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtItemLineNumber.Click
		Dim Index As Short = txtItemLineNumber.GetIndex(eventSender)
		HighlightLine(Index)
	End Sub
	
	'UPGRADE_WARNING: Event txtPONumber.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtPONumber_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPONumber.TextChanged
		booCheckForChanges = True
	End Sub
	
	Private Sub txtQuantity_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtQuantity.Click
		Dim Index As Short = txtQuantity.GetIndex(eventSender)
		HighlightLine(Index)
	End Sub
	
	Private Sub txtRate_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRate.Click
		Dim Index As Short = txtRate.GetIndex(eventSender)
		HighlightLine(Index)
	End Sub
	
	'UPGRADE_WARNING: Event txtRefNumber.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtRefNumber_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRefNumber.TextChanged
		booCheckForChanges = True
	End Sub
	
	'UPGRADE_WARNING: Event txtShipAddr1.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtShipAddr1_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtShipAddr1.TextChanged
		booCheckForChanges = True
	End Sub
	
	'UPGRADE_WARNING: Event txtShipAddr2.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtShipAddr2_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtShipAddr2.TextChanged
		booCheckForChanges = True
	End Sub
	
	'UPGRADE_WARNING: Event txtShipAddr3.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtShipAddr3_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtShipAddr3.TextChanged
		booCheckForChanges = True
	End Sub
	
	'UPGRADE_WARNING: Event txtShipAddr4.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtShipAddr4_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtShipAddr4.TextChanged
		booCheckForChanges = True
	End Sub
	
	'UPGRADE_WARNING: Event txtShipCity.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtShipCity_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtShipCity.TextChanged
		booCheckForChanges = True
	End Sub
	
	'UPGRADE_WARNING: Event txtShipCountry.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtShipCountry_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtShipCountry.TextChanged
		booCheckForChanges = True
	End Sub
	
	'UPGRADE_WARNING: Event txtShipDate.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtShipDate_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtShipDate.TextChanged
		booCheckForChanges = True
	End Sub
	
	'UPGRADE_WARNING: Event txtShipPostalCode.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtShipPostalCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtShipPostalCode.TextChanged
		booCheckForChanges = True
	End Sub
	
	'UPGRADE_WARNING: Event txtShipState.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtShipState_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtShipState.TextChanged
		booCheckForChanges = True
	End Sub
	
	Private Sub txtSpace1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSpace1.Click
		Dim Index As Short = txtSpace1.GetIndex(eventSender)
		HighlightLine((Index Mod 10))
	End Sub
	
	'UPGRADE_WARNING: Event txtTxnDate.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtTxnDate_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTxnDate.TextChanged
		booCheckForChanges = True
	End Sub
	
	'UPGRADE_NOTE: vscInvoiceLineScroll.Change was changed from an event to a procedure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="4E2DC008-5EDA-4547-8317-C9316952674F"'
	'UPGRADE_WARNING: VScrollBar event vscInvoiceLineScroll.Change has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub vscInvoiceLineScroll_Change(ByVal newScrollValue As Integer)
		If vscInvoiceLineScroll.Enabled = False Then Exit Sub
		UnHighlightLine(intHighlightedLine)
		DisplayInvoiceLines((newScrollValue))
	End Sub
	
	Private Sub vscInvoiceLineScroll_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles vscInvoiceLineScroll.Enter
		cmdDoModify.Focus()
	End Sub
	
	Public Sub ClearForm()
		txtTxnID.Text = ""
		txtEditSequence.Text = ""
		txtRefNumber.Text = ""
		txtTxnDate.Text = ""
		chkPending.CheckState = System.Windows.Forms.CheckState.Unchecked 'Unchecked
		chkToBePrinted.CheckState = System.Windows.Forms.CheckState.Unchecked 'Unchecked
		cmbCustomer.Text = ""
		cmbClass.Text = ""
		txtBillAddr1.Text = ""
		txtBillAddr2.Text = ""
		txtBillAddr3.Text = ""
		txtBillAddr4.Text = ""
		txtBillCity.Text = ""
		txtBillState.Text = ""
		txtBillPostalCode.Text = ""
		txtBillCountry.Text = ""
		txtShipAddr1.Text = ""
		txtShipAddr2.Text = ""
		txtShipAddr3.Text = ""
		txtShipAddr4.Text = ""
		txtShipCity.Text = ""
		txtShipState.Text = ""
		txtShipPostalCode.Text = ""
		txtShipCountry.Text = ""
		cmbARAccount.Text = ""
		cmbTerms.Text = ""
		txtPONumber.Text = ""
		txtDueDate.Text = ""
		txtShipDate.Text = ""
		txtFOB.Text = ""
		cmbSalesRep.Text = ""
		cmbShipMethod.Text = ""
		cmbItemSalesTax.Text = ""
		cmbCustTaxCode.Text = ""
		cmbCustomerMsg.Text = ""
		txtMemo.Text = ""
		
		Dim i As Short
		For i = 0 To 9
			txtItemLineNumber(i).Text = ""
			txtQuantity(i).Text = ""
			txtItemFullName(i).Text = ""
			txtDesc(i).Text = ""
			txtRate(i).Text = ""
			txtAmount(i).Text = ""
		Next 
		
		For i = 1 To 200
			strLineArray(i) = CStr(Nothing)
		Next i
		
		'Disable the buttons in case we're reactivating the form since a line
		'hasn't yet been chosen
		cmdEditLine.Enabled = False
        cmdAddLineAfter.Enabled = False
        cmdDeleteUndelete.Enabled = False
	End Sub
	
	
	Private Sub CountInvoiceLines()
		
		Dim booDone As Boolean
		Dim i As Short
		booDone = False
		i = 1
		Do
            If i = 200 Then
                booDone = True
                'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            ElseIf String.IsNullOrEmpty(strLineArray(i + 1)) Then
                booDone = True
			Else
				i = i + 1
			End If
		Loop Until booDone
		
		intNumInvoiceLines = i
		
		If intNumInvoiceLines < 11 Then
			vscInvoiceLineScroll.Enabled = False
			vscInvoiceLineScroll.Visible = False
			vscInvoiceLineScroll.Minimum = 1
			vscInvoiceLineScroll.Maximum = (1 + vscInvoiceLineScroll.LargeChange - 1)
		Else
			vscInvoiceLineScroll.Enabled = True
			vscInvoiceLineScroll.Visible = True
			vscInvoiceLineScroll.Minimum = 1
			vscInvoiceLineScroll.Maximum = (intNumInvoiceLines - 9 + vscInvoiceLineScroll.LargeChange - 1)
		End If
	End Sub
	
	
	Private Sub DisplayInvoiceLines(ByRef intStartLine As Short)
		
		intTopDisplayedLine = intStartLine
		
		Dim i As Short
		Dim intLastLine As Short
		
		If intNumInvoiceLines < 10 Then
			intLastLine = intNumInvoiceLines - 1
		Else
			intLastLine = 9
		End If
		
		For i = 0 To intLastLine
			DisplayLine(strLineArray(intStartLine + i), i)
			If intStartLine + i = intActualHighlightedLine Then
				HighlightLine(i)
			End If
		Next i
	End Sub
	
	
	Private Sub DisplayLine(ByRef strLineInfo As String, ByRef intDisplayLine As Short)
		
		Dim strSplits() As String
		
		strSplits = Split(strLineInfo, "<spliter>")
		
		txtItemLineNumber(intDisplayLine).Text = strSplits(0)
		txtQuantity(intDisplayLine).Text = strSplits(1)
		txtItemFullName(intDisplayLine).Text = strSplits(2)
		txtDesc(intDisplayLine).Text = strSplits(3)
		txtRate(intDisplayLine).Text = strSplits(4)
		txtAmount(intDisplayLine).Text = strSplits(5)
		
		If InStr(1, strSplits(12), "Original") > 0 Then
			txtItemLineNumber(intDisplayLine).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
			txtQuantity(intDisplayLine).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
			txtItemFullName(intDisplayLine).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
			txtDesc(intDisplayLine).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
			txtRate(intDisplayLine).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
			txtAmount(intDisplayLine).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
		ElseIf InStr(1, strSplits(12), "New") > 0 Then 
			txtItemLineNumber(intDisplayLine).ForeColor = System.Drawing.ColorTranslator.FromOle(&HC000)
			txtQuantity(intDisplayLine).ForeColor = System.Drawing.ColorTranslator.FromOle(&HC000)
			txtItemFullName(intDisplayLine).ForeColor = System.Drawing.ColorTranslator.FromOle(&HC000)
			txtDesc(intDisplayLine).ForeColor = System.Drawing.ColorTranslator.FromOle(&HC000)
			txtRate(intDisplayLine).ForeColor = System.Drawing.ColorTranslator.FromOle(&HC000)
			txtAmount(intDisplayLine).ForeColor = System.Drawing.ColorTranslator.FromOle(&HC000)
		ElseIf InStr(1, strSplits(12), "Changed") > 0 Then 
			txtItemLineNumber(intDisplayLine).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF)
			txtQuantity(intDisplayLine).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF)
			txtItemFullName(intDisplayLine).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF)
			txtDesc(intDisplayLine).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF)
			txtRate(intDisplayLine).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF)
			txtAmount(intDisplayLine).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF)
		End If
		
		If InStr(1, strSplits(12), "Deleted") > 0 Then
			txtItemLineNumber(intDisplayLine).Font = VB6.FontChangeStrikeOut(txtItemLineNumber(intDisplayLine).Font, True)
			txtQuantity(intDisplayLine).Font = VB6.FontChangeStrikeOut(txtQuantity(intDisplayLine).Font, True)
			txtItemFullName(intDisplayLine).Font = VB6.FontChangeStrikeOut(txtItemFullName(intDisplayLine).Font, True)
			txtDesc(intDisplayLine).Font = VB6.FontChangeStrikeOut(txtDesc(intDisplayLine).Font, True)
			txtRate(intDisplayLine).Font = VB6.FontChangeStrikeOut(txtRate(intDisplayLine).Font, True)
			txtAmount(intDisplayLine).Font = VB6.FontChangeStrikeOut(txtAmount(intDisplayLine).Font, True)
		Else
			txtItemLineNumber(intDisplayLine).Font = VB6.FontChangeStrikeOut(txtItemLineNumber(intDisplayLine).Font, False)
			txtQuantity(intDisplayLine).Font = VB6.FontChangeStrikeOut(txtQuantity(intDisplayLine).Font, False)
			txtItemFullName(intDisplayLine).Font = VB6.FontChangeStrikeOut(txtItemFullName(intDisplayLine).Font, False)
			txtDesc(intDisplayLine).Font = VB6.FontChangeStrikeOut(txtDesc(intDisplayLine).Font, False)
			txtRate(intDisplayLine).Font = VB6.FontChangeStrikeOut(txtRate(intDisplayLine).Font, False)
			txtAmount(intDisplayLine).Font = VB6.FontChangeStrikeOut(txtAmount(intDisplayLine).Font, False)
		End If
		
		If strSplits(11) = "SubItem" Then
			txtItemLineNumber(intDisplayLine).Font = VB6.FontChangeBold(txtItemLineNumber(intDisplayLine).Font, False)
		Else
			txtItemLineNumber(intDisplayLine).Font = VB6.FontChangeBold(txtItemLineNumber(intDisplayLine).Font, True)
		End If
		
		If strSplits(11) = "Group" Then
			txtItemFullName(intDisplayLine).Font = VB6.FontChangeBold(txtItemFullName(intDisplayLine).Font, True)
			txtDesc(intDisplayLine).Font = VB6.FontChangeBold(txtDesc(intDisplayLine).Font, True)
		Else
			txtItemFullName(intDisplayLine).Font = VB6.FontChangeBold(txtItemFullName(intDisplayLine).Font, False)
			txtDesc(intDisplayLine).Font = VB6.FontChangeBold(txtDesc(intDisplayLine).Font, False)
		End If
	End Sub
	
	
	Private Sub HighlightLine(ByRef intDisplayLine As Short)
		If intDisplayLine >= intNumInvoiceLines Then Exit Sub
		
		If intHighlightedLine > -1 Then UnHighlightLine(intHighlightedLine)
		
		txtItemLineNumber(intDisplayLine).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFF00)
		txtQuantity(intDisplayLine).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFF00)
		txtItemFullName(intDisplayLine).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFF00)
		txtDesc(intDisplayLine).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFF00)
		txtRate(intDisplayLine).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFF00)
		txtAmount(intDisplayLine).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFF00)
		txtSpace1(intDisplayLine).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFF00)
		txtSpace1(intDisplayLine + 10).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFF00)
		txtSpace1(intDisplayLine + 20).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFF00)
		
		intActualHighlightedLine = vscInvoiceLineScroll.Value + intDisplayLine
		intHighlightedLine = intDisplayLine
		
		Dim strSplits() As String
		strSplits = Split(strLineArray(intActualHighlightedLine), "<spliter>")
		
		cmdAddLineBefore.Enabled = True
		cmdAddLineAfter.Enabled = True
		cmdDeleteUndelete.Enabled = True
		
		If InStr(1, strSplits(12), "Deleted") > 0 Then
			cmdDeleteUndelete.Text = "Undelete Line"
			cmdEditLine.Enabled = False
		Else
			cmdDeleteUndelete.Text = "Delete Line"
			cmdEditLine.Enabled = True
		End If
	End Sub
	
	
	Private Sub UnHighlightLine(ByRef intDisplayLine As Short)
		If intHighlightedLine > -1 Then
			txtItemLineNumber(intDisplayLine).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			txtQuantity(intDisplayLine).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			txtItemFullName(intDisplayLine).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			txtDesc(intDisplayLine).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			txtRate(intDisplayLine).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			txtAmount(intDisplayLine).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			txtSpace1(intDisplayLine).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			txtSpace1(intDisplayLine + 10).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			txtSpace1(intDisplayLine + 20).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			
			cmdAddLineBefore.Enabled = False
			cmdAddLineAfter.Enabled = False
			cmdDeleteUndelete.Enabled = False
			cmdEditLine.Enabled = False
			
			intHighlightedLine = -1
		End If
	End Sub
	
	
	Private Sub InsertInvoiceLine(ByRef strInvoiceLine As String, ByRef intLineNumber As Short)
		
		Dim i As Short
		If intLineNumber > intNumInvoiceLines Then
			strLineArray(intNumInvoiceLines + 1) = GetInvoiceLineInfo
		Else
			For i = intNumInvoiceLines + 1 To intLineNumber + 1 Step -1
				strLineArray(i) = strLineArray(i - 1)
			Next i
			strLineArray(intLineNumber) = strInvoiceLine
		End If
		
		intNumInvoiceLines = intNumInvoiceLines + 1
		If intNumInvoiceLines = 11 Then
			vscInvoiceLineScroll.Maximum = (2 + vscInvoiceLineScroll.LargeChange - 1)
			vscInvoiceLineScroll.Enabled = True
			vscInvoiceLineScroll.Visible = True
			vscInvoiceLineScroll.Value = intTopDisplayedLine
		ElseIf intNumInvoiceLines > 11 Then 
			vscInvoiceLineScroll.Maximum = ((vscInvoiceLineScroll.Maximum - vscInvoiceLineScroll.LargeChange + 1) + 1 + vscInvoiceLineScroll.LargeChange - 1)
		End If
	End Sub
	
	
	Private Sub SaveOriginalValues()
		strTxnID = txtTxnID.Text
		strEditSequence = txtEditSequence.Text
		strRefNumber = txtRefNumber.Text
		strTxnDate = txtTxnDate.Text
		intPending = chkPending.CheckState
		intToBePrinted = chkToBePrinted.CheckState
		strCustomer = cmbCustomer.Text
		strClass = cmbClass.Text
		strBillAddr1 = txtBillAddr1.Text
		strBillAddr2 = txtBillAddr2.Text
		strBillAddr3 = txtBillAddr3.Text
		strBillAddr4 = txtBillAddr4.Text
		strBillCity = txtBillCity.Text
		strBillState = txtBillState.Text
		strBillPostalCode = txtBillPostalCode.Text
		strBillCountry = txtBillCountry.Text
		strShipAddr1 = txtShipAddr1.Text
		strShipAddr2 = txtShipAddr2.Text
		strShipAddr3 = txtShipAddr3.Text
		strShipAddr4 = txtShipAddr4.Text
		strShipCity = txtShipCity.Text
		strShipState = txtShipState.Text
		strShipPostalCode = txtShipPostalCode.Text
		strShipCountry = txtShipCountry.Text
		strARAccount = cmbARAccount.Text
		strTerms = cmbTerms.Text
		strPONumber = txtPONumber.Text
		strDueDate = txtDueDate.Text
		strShipDate = txtShipDate.Text
		strFOB = txtFOB.Text
		strSalesRep = cmbSalesRep.Text
		strShipMethod = cmbShipMethod.Text
		strItemSalesTax = cmbItemSalesTax.Text
		strCustTaxCode = cmbCustTaxCode.Text
		strCustomerMsg = cmbCustomerMsg.Text
		strMemo = txtMemo.Text
	End Sub
	
	
	Private Sub ReloadInvoiceModifyForm()
		txtTxnID.Text = strTxnID
		txtEditSequence.Text = strEditSequence
		txtRefNumber.Text = strRefNumber
		txtTxnDate.Text = strTxnDate
		chkPending.CheckState = intPending
		chkToBePrinted.CheckState = intToBePrinted
		cmbCustomer.Text = strCustomer
		cmbClass.Text = strClass
		txtBillAddr1.Text = strBillAddr1
		txtBillAddr2.Text = strBillAddr2
		txtBillAddr3.Text = strBillAddr3
		txtBillAddr4.Text = strBillAddr4
		txtBillCity.Text = strBillCity
		txtBillState.Text = strBillState
		txtBillPostalCode.Text = strBillPostalCode
		txtBillCountry.Text = strBillCountry
		txtShipAddr1.Text = strShipAddr1
		txtShipAddr2.Text = strShipAddr2
		txtShipAddr3.Text = strShipAddr3
		txtShipAddr4.Text = strShipAddr4
		txtShipCity.Text = strShipCity
		txtShipState.Text = strShipState
		txtShipPostalCode.Text = strShipPostalCode
		txtShipCountry.Text = strShipCountry
		cmbARAccount.Text = strARAccount
		cmbTerms.Text = strTerms
		txtPONumber.Text = strPONumber
		txtDueDate.Text = strDueDate
		txtShipDate.Text = strShipDate
		txtFOB.Text = strFOB
		cmbSalesRep.Text = strSalesRep
		cmbShipMethod.Text = strShipMethod
		cmbItemSalesTax.Text = strItemSalesTax
		cmbCustTaxCode.Text = strCustTaxCode
		cmbCustomerMsg.Text = strCustomerMsg
		txtMemo.Text = strMemo
	End Sub
	
	
	Private Function InvoiceChanged() As Boolean
		
		booBillAddressChanged = False
		booShipAddressChanged = False
		booInvoiceBodyChanged = False
		booInvoiceLinesChanged = False
		
		If Not (booCheckForChanges Or booCheckForLineChanges) Then
			InvoiceChanged = False
			Exit Function
		End If
		Dim booInvoiceChanged As Boolean
		booInvoiceChanged = False
		
		If Trim(strRefNumber) <> Trim(txtRefNumber.Text) Or Trim(strTxnDate) <> Trim(txtTxnDate.Text) Or intPending <> chkPending.CheckState Or intToBePrinted <> chkToBePrinted.CheckState Or Trim(strCustomer) <> Trim(cmbCustomer.Text) Or Trim(strClass) <> Trim(cmbClass.Text) Or Trim(strARAccount) <> Trim(cmbARAccount.Text) Or Trim(strTerms) <> Trim(cmbTerms.Text) Or Trim(strPONumber) <> Trim(txtPONumber.Text) Or Trim(strDueDate) <> Trim(txtDueDate.Text) Or Trim(strShipDate) <> Trim(txtShipDate.Text) Or Trim(strFOB) <> Trim(txtFOB.Text) Or Trim(strSalesRep) <> Trim(cmbSalesRep.Text) Or Trim(strShipMethod) <> Trim(cmbShipMethod.Text) Or Trim(strItemSalesTax) <> Trim(cmbItemSalesTax.Text) Or Trim(strCustTaxCode) <> Trim(cmbCustTaxCode.Text) Or Trim(strCustomerMsg) <> Trim(cmbCustomerMsg.Text) Or Trim(strMemo) <> Trim(txtMemo.Text) Then
			booInvoiceChanged = True
			booInvoiceBodyChanged = True
		End If
		
		If Trim(strBillAddr1) <> Trim(txtBillAddr1.Text) Or Trim(strBillAddr2) <> Trim(txtBillAddr2.Text) Or Trim(strBillAddr3) <> Trim(txtBillAddr3.Text) Or Trim(strBillAddr4) <> Trim(txtBillAddr4.Text) Or Trim(strBillCity) <> Trim(txtBillCity.Text) Or Trim(strBillState) <> Trim(txtBillState.Text) Or Trim(strBillPostalCode) <> Trim(txtBillPostalCode.Text) Or Trim(strBillCountry) <> Trim(txtBillCountry.Text) Then
			booBillAddressChanged = True
			booInvoiceChanged = True
			booInvoiceBodyChanged = True
		End If
		
		If Trim(strShipAddr1) <> Trim(txtShipAddr1.Text) Or Trim(strShipAddr2) <> Trim(txtShipAddr2.Text) Or Trim(strShipAddr3) <> Trim(txtShipAddr3.Text) Or Trim(strShipAddr4) <> Trim(txtShipAddr4.Text) Or Trim(strShipCity) <> Trim(txtShipCity.Text) Or Trim(strShipState) <> Trim(txtShipState.Text) Or Trim(strShipPostalCode) <> Trim(txtShipPostalCode.Text) Or Trim(strShipCountry) <> Trim(txtShipCountry.Text) Then
			booShipAddressChanged = True
			booInvoiceChanged = True
			booInvoiceBodyChanged = True
		End If
		
		Dim strSplits() As String
		Dim i As Short
		For i = 1 To intNumInvoiceLines
			strSplits = Split(strLineArray(i), "<spliter>")
			If strSplits(12) = "New" Or strSplits(12) = "Changed" Or strSplits(12) = "OriginalDeleted" Then
				booInvoiceLinesChanged = True
			End If
		Next i
		InvoiceChanged = booInvoiceChanged Or booInvoiceLinesChanged
	End Function
	
	
	Private Function InvoiceChangeString() As String
		
		Dim strInvoiceChangeString As String
		
		strInvoiceChangeString = strInvoiceChangeString & "<TxnID>" & txtTxnID.Text & "</TxnID>" & "<EditSequence>" & Trim(txtEditSequence.Text) & "</EditSequence>"
		
		If booInvoiceBodyChanged Then
			If cmbCustomer.Text <> strCustomer Then
				strInvoiceChangeString = strInvoiceChangeString & "<CustomerRef><FullName>" & Trim(cmbCustomer.Text) & "</FullName></CustomerRef>"
			End If
			
			If cmbClass.Text <> strClass Then
				strInvoiceChangeString = strInvoiceChangeString & "<ClassRef><FullName>" & Trim(cmbClass.Text) & "</FullName></ClassRef>"
			End If
			
			If cmbARAccount.Text <> strARAccount Then
				strInvoiceChangeString = strInvoiceChangeString & "<ARAccountRef><FullName>" & Trim(cmbARAccount.Text) & "</FullName></ARAccountRef>"
			End If
			
			If txtTxnDate.Text <> strTxnDate Then
				strInvoiceChangeString = strInvoiceChangeString & "<TxnDate>" & Trim(txtTxnDate.Text) & "</TxnDate>"
			End If
			
			If txtRefNumber.Text <> strRefNumber Then
                strInvoiceChangeString = strInvoiceChangeString & "<RefNumber>" & Trim(txtRefNumber.Text) & "</RefNumber>"
            End If
			
			If booBillAddressChanged Then
				strInvoiceChangeString = strInvoiceChangeString & "<BillAddress>"
				If txtBillAddr1.Text <> strBillAddr1 Then
					strInvoiceChangeString = strInvoiceChangeString & "<Addr1>" & Trim(txtBillAddr1.Text) & "</Addr1>"
				End If
				
				If txtBillAddr2.Text <> strBillAddr2 Then
					strInvoiceChangeString = strInvoiceChangeString & "<Addr2>" & Trim(txtBillAddr2.Text) & "</Addr2>"
				End If
				
				If txtBillAddr3.Text <> strBillAddr3 Then
					strInvoiceChangeString = strInvoiceChangeString & "<Addr3>" & Trim(txtBillAddr3.Text) & "</Addr3>"
				End If
				
				If txtBillAddr4.Text <> strBillAddr4 Then
					strInvoiceChangeString = strInvoiceChangeString & "<Addr4>" & Trim(txtBillAddr4.Text) & "</Addr4>"
				End If
				
				If txtBillCity.Text <> strBillCity Then
					strInvoiceChangeString = strInvoiceChangeString & "<City>" & Trim(txtBillCity.Text) & "</City>"
				End If
				
				If txtBillState.Text <> strBillState Then
					strInvoiceChangeString = strInvoiceChangeString & "<State>" & Trim(txtBillState.Text) & "</State>"
				End If
				
				If txtBillPostalCode.Text <> strBillPostalCode Then
					strInvoiceChangeString = strInvoiceChangeString & "<PostalCode>" & Trim(txtBillPostalCode.Text) & "</PostalCode>"
				End If
				
				If txtBillCountry.Text <> strBillCountry Then
					strInvoiceChangeString = strInvoiceChangeString & "<Country>" & Trim(txtBillCountry.Text) & "</Country>"
				End If
				
				strInvoiceChangeString = strInvoiceChangeString & "</BillAddress>"
			End If ' If booBillAddressChanged
			
			If booShipAddressChanged Then
				strInvoiceChangeString = strInvoiceChangeString & "<ShipAddress>"
				If txtShipAddr1.Text <> strShipAddr1 Then
					strInvoiceChangeString = strInvoiceChangeString & "<Addr1>" & Trim(txtShipAddr1.Text) & "</Addr1>"
				End If
				
				If txtShipAddr2.Text <> strShipAddr2 Then
					strInvoiceChangeString = strInvoiceChangeString & "<Addr2>" & Trim(txtShipAddr2.Text) & "</Addr2>"
				End If
				
				If txtShipAddr3.Text <> strShipAddr3 Then
					strInvoiceChangeString = strInvoiceChangeString & "<Addr3>" & Trim(txtShipAddr3.Text) & "</Addr3>"
				End If
				
				If txtShipAddr4.Text <> strShipAddr4 Then
					strInvoiceChangeString = strInvoiceChangeString & "<Addr4>" & Trim(txtShipAddr4.Text) & "</Addr4>"
				End If
				
				If txtShipCity.Text <> strShipCity Then
					strInvoiceChangeString = strInvoiceChangeString & "<City>" & Trim(txtShipCity.Text) & "</City>"
				End If
				
				If txtShipState.Text <> strShipState Then
					strInvoiceChangeString = strInvoiceChangeString & "<State>" & Trim(txtShipState.Text) & "</State>"
				End If
				
				If txtShipPostalCode.Text <> strShipPostalCode Then
					strInvoiceChangeString = strInvoiceChangeString & "<PostalCode>" & Trim(txtShipPostalCode.Text) & "</PostalCode>"
				End If
				
				If txtShipCountry.Text <> strShipCountry Then
					strInvoiceChangeString = strInvoiceChangeString & "<Country>" & Trim(txtShipCountry.Text) & "</Country>"
				End If
				
				strInvoiceChangeString = strInvoiceChangeString & "</ShipAddress>"
			End If ' If booShipAddressChanged
			
			If chkPending.CheckState <> intPending Then
				strInvoiceChangeString = strInvoiceChangeString & "<IsPending>" & chkPending.CheckState & "</IsPending>"
			End If
			
			If txtPONumber.Text <> strPONumber Then
				strInvoiceChangeString = strInvoiceChangeString & "<PONumber>" & Trim(txtPONumber.Text) & "</PONumber>"
			End If
			
			If cmbTerms.Text <> strTerms Then
				strInvoiceChangeString = strInvoiceChangeString & "<TermsRef><FullName>" & Trim(cmbTerms.Text) & "</FullName></TermsRef>"
			End If
			
			If txtDueDate.Text <> strDueDate Then
				strInvoiceChangeString = strInvoiceChangeString & "<DueDate>" & Trim(txtDueDate.Text) & "</DueDate>"
			End If
			
			If cmbSalesRep.Text <> strSalesRep Then
				strInvoiceChangeString = strInvoiceChangeString & "<SalesRepRef><FullName>" & Trim(cmbSalesRep.Text) & "</FullName></SalesRepRef>"
			End If
			
			If txtFOB.Text <> strFOB Then
				strInvoiceChangeString = strInvoiceChangeString & "<FOB>" & Trim(txtFOB.Text) & "</FOB>"
			End If
			
			If txtShipDate.Text <> strShipDate Then
				strInvoiceChangeString = strInvoiceChangeString & "<ShipDate>" & Trim(txtShipDate.Text) & "</ShipDate>"
			End If
			
			If cmbShipMethod.Text <> strShipMethod Then
				strInvoiceChangeString = strInvoiceChangeString & "<ShipMethodRef><FullName>" & Trim(cmbShipMethod.Text) & "</FullName></ShipMethodRef>"
			End If
			
			If cmbItemSalesTax.Text <> strItemSalesTax Then
				strInvoiceChangeString = strInvoiceChangeString & "<ItemSalesTaxRef><FullName>" & Trim(cmbItemSalesTax.Text) & "</FullName></ItemSalesTaxRef>"
			End If
			
			If txtMemo.Text <> strMemo Then
				strInvoiceChangeString = strInvoiceChangeString & "<Memo>" & Trim(txtMemo.Text) & "</Memo>"
			End If
			
			If cmbCustomerMsg.Text <> strCustomerMsg Then
				strInvoiceChangeString = strInvoiceChangeString & "<CustomerMsgRef><FullName>" & Trim(cmbCustomerMsg.Text) & "</FullName></CustomerMsgRef>"
			End If
			
			If chkToBePrinted.CheckState <> intToBePrinted Then
				strInvoiceChangeString = strInvoiceChangeString & "<IsToBePrinted>" & chkToBePrinted.CheckState & "</IsToBePrinted>"
			End If
			
			If cmbCustTaxCode.Text <> strCustTaxCode Then
				strInvoiceChangeString = strInvoiceChangeString & "<CustomerSalesTaxCodeRef><FullName>" & Trim(cmbCustTaxCode.Text) & "</FullName></CustomerSalesTaxCodeRef>"
			End If
		End If
		
		Dim strSplits() As String
		Dim strGroupLineSplits() As String
		Dim booDone As Boolean
		Dim booLineLoopDone As Boolean
		Dim booIncludeGroupLines As Boolean
		Dim i As Short
        Dim j As Short
        Dim k As Short
        If booInvoiceLinesChanged Then
			booDone = False
            i = 1
            Do
                strSplits = Split(strLineArray(i), "<spliter>")

                If strSplits(11) = "Group" Then '
                    If i = 200 And InStr(1, strSplits(12), "Deleted") <= 0 Then
                        AddInvoiceLineInfo(strInvoiceChangeString, i, False)
                        InvoiceChangeString = strInvoiceChangeString
                        Exit Function
                    End If

                    j = i + 1

                    booLineLoopDone = False
                    booIncludeGroupLines = False
                    Do While Not booLineLoopDone
                        strGroupLineSplits = Split(strLineArray(j), "<spliter>")
                        If strGroupLineSplits(11) <> "SubItem" Then
                            booLineLoopDone = True
                        Else
                            If strGroupLineSplits(12) <> "NewDeleted" And strGroupLineSplits(12) <> "Original" Then
                                booIncludeGroupLines = True
                            End If

                            If j = UBound(strLineArray) Then
                                booLineLoopDone = True
                                'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                            ElseIf String.IsNullOrEmpty(strLineArray(j + 1)) Then
                                'set the value of j to j + 1 anyway to avoid having the last
                                'item in a group be repeated if it's the last item on an invoice
                                j = j + 1
                                booLineLoopDone = True
                            Else
                                j = j + 1
                            End If
                        End If
                    Loop

                    If InStr(1, strSplits(12), "Deleted") <= 0 Then
                        AddInvoiceLineInfo(strInvoiceChangeString, i, booIncludeGroupLines)
                    End If

                    i = j - 1
                Else ' for If strSplits(11) = "Group"
                    If strSplits(12) = "Original" Then
                        strInvoiceChangeString = strInvoiceChangeString & "<InvoiceLineMod><TxnLineID>" & strSplits(0) & "</TxnLineID></InvoiceLineMod>"
                    ElseIf strSplits(12) = "Changed" Or strSplits(12) = "New" Then
                        AddInvoiceLineInfo(strInvoiceChangeString, i, False)
                    End If ' for If strSplits(12) = "Original"
                End If ' for If strSplits(11) = "Group"

                If i = 200 Then
                    booDone = True
                Else
                    'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                    If String.IsNullOrEmpty(strLineArray(i + 1)) Then
                        booDone = True
                    Else
                        i = i + 1

                    End If
                End If
            Loop Until booDone
        End If
		
		InvoiceChangeString = strInvoiceChangeString
	End Function
	
	
	Private Sub AddInvoiceLineInfo(ByRef strInvoiceChangeString As String, ByRef intInvoiceLineNum As Short, ByRef booAddGroupLines As Boolean)
		
		Dim booGroupLine As Boolean
		booGroupLine = False
		Dim strSplits() As String
		strSplits = Split(strLineArray(intInvoiceLineNum), "<spliter>")
		
		If strSplits(11) = "Group" Then
			strInvoiceChangeString = strInvoiceChangeString & "<InvoiceLineGroupMod><TxnLineID>"
			booGroupLine = True
		Else
			If InStr(1, strSplits(2), " - Group Item") Then
				strInvoiceChangeString = strInvoiceChangeString & "<InvoiceLineGroupMod><TxnLineID>"
				booGroupLine = True
			Else
				strInvoiceChangeString = strInvoiceChangeString & "<InvoiceLineMod><TxnLineID>"
			End If
		End If
		
		If VB.Left(strSplits(0), 2) = "-1" Then
			strInvoiceChangeString = strInvoiceChangeString & "-1</TxnLineID>"
		Else
			strInvoiceChangeString = strInvoiceChangeString & strSplits(0) & "</TxnLineID>"
		End If

        'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If Not String.IsNullOrEmpty(Trim(strSplits(2))) Then

            If InStr(1, strSplits(2), " - Group Item") > 0 Or strSplits(11) = "Group" Then
                If strSplits(12) <> "Original" Then
                    strInvoiceChangeString = strInvoiceChangeString & "<ItemGroupRef><FullName>" & Replace(Trim(strSplits(2)), " - Group Item", "") & "</FullName></ItemGroupRef>"
                End If
            Else
                strInvoiceChangeString = strInvoiceChangeString & "<ItemRef><FullName>" & Trim(strSplits(2)) & "</FullName></ItemRef>"
            End If
        End If

        'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If Not String.IsNullOrEmpty(Trim(strSplits(3))) And strSplits(11) <> "Group" And InStr(1, strSplits(2), " - Group Item") = 0 Then
            strInvoiceChangeString = strInvoiceChangeString & "<Desc>" & Trim(strSplits(3)) & "</Desc>"
        End If

        'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If Not String.IsNullOrEmpty(Trim(strSplits(1))) And strSplits(12) <> "Original" Then
            strInvoiceChangeString = strInvoiceChangeString & "<Quantity>" & Trim(strSplits(1)) & "</Quantity>"
        End If

        'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If Not String.IsNullOrEmpty(Trim(strSplits(4))) And strSplits(11) <> "Group" Then
            If strSplits(9) = "RatePercent" Then
                strInvoiceChangeString = strInvoiceChangeString & "<RatePercent>" & Trim(strSplits(4)) & "</RatePercent>"
            Else
                strInvoiceChangeString = strInvoiceChangeString & "<Rate>" & Trim(strSplits(4)) & "</Rate>"
            End If
        End If

        'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If Not String.IsNullOrEmpty(Trim(strSplits(6))) And strSplits(11) <> "Group" Then
            strInvoiceChangeString = strInvoiceChangeString & "<ClassRef><FullName>" & Trim(strSplits(6)) & "</FullName></ClassRef>"
        End If

        'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If Not String.IsNullOrEmpty(Trim(strSplits(5))) And strSplits(11) <> "Group" Then
            strInvoiceChangeString = strInvoiceChangeString & "<Amount>" & Trim(strSplits(5)) & "</Amount>"
        End If

        'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If Not String.IsNullOrEmpty(Trim(strSplits(7))) And strSplits(11) <> "Group" Then
            strInvoiceChangeString = strInvoiceChangeString & "<ServiceDate>" & Trim(strSplits(7)) & "</ServiceDate>"
        End If

        'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If Not String.IsNullOrEmpty(Trim(strSplits(8))) And strSplits(11) <> "Group" Then
            strInvoiceChangeString = strInvoiceChangeString & "<SalesTaxCodeRef><FullName>" & Trim(strSplits(8)) & "</FullName></SalesTaxCodeRef>"
        End If

        'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If Not String.IsNullOrEmpty(Trim(strSplits(10))) And strSplits(11) <> "Group" Then
            strInvoiceChangeString = strInvoiceChangeString & "<OverrideItemAccountRef><FullName>" & Trim(strSplits(10)) & "</FullName></OverrideItemAccountRef>"
        End If

        Dim j As Short
		Dim booProcessGroupLines As Boolean
		Dim strGroupLineSplits() As String
		If booAddGroupLines Then
			
			If intInvoiceLineNum = 200 Then
				booProcessGroupLines = False
			ElseIf InStr(1, strLineArray(intInvoiceLineNum + 1), "SubItem") <= 0 Then 
				booProcessGroupLines = False
			Else
				booProcessGroupLines = True
				j = intInvoiceLineNum + 1
			End If
			
			Do While booProcessGroupLines
				strGroupLineSplits = Split(strLineArray(j), "<spliter>")
				If strGroupLineSplits(11) <> "SubItem" Then
					booProcessGroupLines = False
				Else
					If strGroupLineSplits(12) = "Original" Then
						strInvoiceChangeString = strInvoiceChangeString & "<InvoiceLineMod><TxnLineID>" & strGroupLineSplits(0) & "</TxnLineID></InvoiceLineMod>"
					ElseIf strGroupLineSplits(12) = "New" Or strGroupLineSplits(12) = "Changed" Then 
						AddInvoiceLineInfo(strInvoiceChangeString, j, False)
					End If

                    If j = UBound(strLineArray) Then
                        booProcessGroupLines = False
                        'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                    ElseIf String.IsNullOrEmpty(strLineArray(j + 1)) Then
                        booProcessGroupLines = False
					Else
						j = j + 1
					End If
				End If
			Loop 
		End If
		
		If booGroupLine Then
			strInvoiceChangeString = strInvoiceChangeString & "</InvoiceLineGroupMod>"
		Else
			strInvoiceChangeString = strInvoiceChangeString & "</InvoiceLineMod>"
		End If
	End Sub
	
	
	Private Sub ChangeSubLines(ByRef intGroupLine As Short, ByRef strAction As String)
		
		If intGroupLine = UBound(strLineArray) Then Exit Sub
		
		Dim i As Short
		Dim booDone As Boolean
		Dim strSplits() As String
		
		i = intGroupLine + 1
		booDone = False
		Do 
			strSplits = Split(strLineArray(i), "<spliter>")
			If strSplits(11) <> "SubItem" Then
				booDone = True
			Else
				If strAction = "Delete" Then
					strLineArray(i) = strLineArray(i) & "Deleted"
				Else
					strLineArray(i) = VB.Left(strLineArray(i), Len(strLineArray(i)) - 7)
				End If
				
				If intHighlightedLine + (i - intActualHighlightedLine) < 10 Then
					DisplayLine(strLineArray(i), intHighlightedLine + (i - intActualHighlightedLine))
				End If

                If i = UBound(strLineArray) Then
                    booDone = True
                    'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                ElseIf String.IsNullOrEmpty(strLineArray(i + 1)) Then
                    booDone = True
				Else
					i = i + 1
				End If
			End If
		Loop Until booDone
	End Sub
	
	
	Private Function InDeletedGroup(ByRef intArrayLine As Short) As Boolean
		
		Dim strSplits() As String
		Dim i As Short
		
		strSplits = Split(strLineArray(intArrayLine), "<spliter>")
		If strSplits(11) <> "SubItem" Then
			InDeletedGroup = False
			Exit Function
		End If
		
		i = intArrayLine
		Do 
			i = i - 1
			strSplits = Split(strLineArray(i), "<spliter>")
		Loop Until strSplits(11) = "Group"
		
		InDeletedGroup = (InStr(1, strSplits(12), "Deleted") > 0)
	End Function
	Private Sub vscInvoiceLineScroll_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.ScrollEventArgs) Handles vscInvoiceLineScroll.Scroll
		Select Case eventArgs.type
			Case System.Windows.Forms.ScrollEventType.EndScroll
				vscInvoiceLineScroll_Change(eventArgs.newValue)
		End Select
	End Sub
End Class
