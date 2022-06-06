Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmInvoiceQuery
	Inherits System.Windows.Forms.Form
	'----------------------------------------------------------
	' Copyright © 2003-2013 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/rdmgr/?ID=100
	'
	'----------------------------------------------------------
	
	
	'This variable determines whether the filter controls are enabled
	Dim booFiltersEnabled As Boolean
	
	'Saves the last command from the form for use in Form_Activate
	Dim strLastCommand As String
	
	'Save the state of the Invoice Modify form
	Dim booInvoiceModifyLoaded As Boolean
	
	'These are the variables which will be set to do the invoice query
	Dim strRefNumber As String
	Dim strFromDateTime As String
	Dim strToDateTime As String
	Dim strDateQueryType As String
	Dim strDateMacro As String
	Dim strCustomerJob As String
	Dim booCustomerWithChildren As Boolean
	Dim strAccount As String
	Dim booAccountWithChildren As Boolean
	Dim strFromRefNumberRange As String
	Dim strToRefNumberRange As String
	Dim strRefNumberPiece As String
	Dim strRefNumberCriteria As String
	Dim strPaidStatus As String
	
	
	Private Sub frmInvoiceQuery_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		strLastCommand = CStr(Nothing)
		booInvoiceModifyLoaded = False
		
		booFiltersEnabled = True
		If Not SupportsModify Then
			cmdModifyInvoice.Text = "Invoice Details"
		End If
		
		frmPatience.Show()
		FillComboBox(cmbCustomers, "Customer", "FullName", "", False)
		
		FillComboBox(cmbAccounts, "Account", "FullName", "<AccountType>AccountsReceivable</AccountType>", False)
		frmPatience.Hide()
	End Sub
	
	
	'UPGRADE_WARNING: Form event frmInvoiceQuery.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmInvoiceQuery_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If Not String.IsNullOrEmpty(strLastCommand) And strLastCommand <> "Query" Then
            lstInvoices.Items.Clear()
            ClearQueryVariables()
            If Not SetQueryVariables() Then
                Exit Sub
            End If

            FillInvoiceList(lstInvoices, strRefNumber, strFromDateTime, strToDateTime, strDateQueryType, strDateMacro, strCustomerJob, booCustomerWithChildren, strAccount, booAccountWithChildren, strFromRefNumberRange, strToRefNumberRange, strRefNumberPiece, strRefNumberCriteria, strPaidStatus)
        End If
    End Sub
	
	
	Private Sub cmdModifyInvoice_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdModifyInvoice.Click
		
		strLastCommand = "ModifyInvoice"
		
		If lstInvoices.SelectedIndex < 0 Then
			MsgBox("You must select an invoice to modify")
			Exit Sub
		End If
		
		If VB6.GetItemString(lstInvoices, lstInvoices.SelectedIndex) = "No invoices match the query filter used" Or lstInvoices.Items.Count = 0 Then
			MsgBox("You must query for a valid invoice before you can modify one")
			Exit Sub
		End If
		
		GetInvoice(VB.Right(VB6.GetItemString(lstInvoices, lstInvoices.SelectedIndex), Len(VB6.GetItemString(lstInvoices, lstInvoices.SelectedIndex)) - InStrRev(VB6.GetItemString(lstInvoices, lstInvoices.SelectedIndex), " ")))
		
		If Not booInvoiceModifyLoaded Then
            'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
            frmInvoiceModify.ShowDialog()
            booInvoiceModifyLoaded = True
        End If

        frmInvoiceModify.Show()
    End Sub
	
	
	Private Sub cmdClearDateSelection_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClearDateSelection.Click
		optFiscalQuarter.Checked = False
		optFiscalYear.Checked = False
		optLast.Checked = False
		optMacroDate.Checked = False
		optModified.Checked = False
		optMonth.Checked = False
		optThis.Checked = False
		optTxnDate.Checked = False
		optWeek.Checked = False
		SetToFromDatesEnabled(False)
		SetMacroDatesEnabled(False)
	End Sub
	
	
	Private Sub cmdClearRefNumber_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClearRefNumber.Click
		optRefNumberFilter.Checked = False
		optRefNumberRange.Checked = False
		txtFromRefNumber.Enabled = False
		txtToRefNumber.Enabled = False
		lblFromRefNumber.Enabled = False
		lblToRefNumber.Enabled = False
		cmbRefNumberCriteria.Enabled = False
		txtRefNumberPart.Enabled = False
		
		txtFromRefNumber.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000004)
		txtToRefNumber.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000004)
		cmbRefNumberCriteria.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000004)
		txtRefNumberPart.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000004)
	End Sub
	
	
	Private Sub cmdQuery_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdQuery.Click
		strLastCommand = "Query"
		
		lstInvoices.Items.Clear()
		ClearQueryVariables()
		If Not SetQueryVariables Then
			Exit Sub
		End If
		
		FillInvoiceList(lstInvoices, strRefNumber, strFromDateTime, strToDateTime, strDateQueryType, strDateMacro, strCustomerJob, booCustomerWithChildren, strAccount, booAccountWithChildren, strFromRefNumberRange, strToRefNumberRange, strRefNumberPiece, strRefNumberCriteria, strPaidStatus)
		
		If chkShow.CheckState = 1 Then
            'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
            'Load(frmShowRequest)
            'frmShowRequest.Show()
            'Dim form As New frmShowRequest
            'form.Show()
            frmShowRequest.ShowDialog()
        End If
	End Sub
	
	
	Private Sub cmdQuit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdQuit.Click
		EndSessionCloseConnection()
		End
	End Sub
	
	
	'UPGRADE_WARNING: Event lstInvoices.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub lstInvoices_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstInvoices.SelectedIndexChanged
		cmdModifyInvoice.Enabled = True
	End Sub
	
	'UPGRADE_WARNING: Event optMacroDate.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optMacroDate_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optMacroDate.CheckedChanged
		If eventSender.Checked Then
			SetToFromDatesEnabled(False)
			SetMacroDatesEnabled(True)
		End If
	End Sub
	
	
	'UPGRADE_WARNING: Event optModified.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optModified_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optModified.CheckedChanged
		If eventSender.Checked Then
			SetToFromDatesEnabled(True)
			SetMacroDatesEnabled(False)
		End If
	End Sub
	
	
	'UPGRADE_WARNING: Event optTxnDate.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optTxnDate_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optTxnDate.CheckedChanged
		If eventSender.Checked Then
			SetToFromDatesEnabled(True)
			SetMacroDatesEnabled(False)
		End If
	End Sub
	
	
	'UPGRADE_WARNING: Event optRefNumberFilter.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optRefNumberFilter_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optRefNumberFilter.CheckedChanged
		If eventSender.Checked Then
			cmbRefNumberCriteria.Enabled = True
			txtRefNumberPart.Enabled = True
			cmbRefNumberCriteria.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			txtRefNumberPart.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			
			lblFromRefNumber.Enabled = False
			lblToRefNumber.Enabled = False
			txtFromRefNumber.Enabled = False
			txtToRefNumber.Enabled = False
			txtFromRefNumber.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000004)
			txtToRefNumber.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000004)
		End If
	End Sub
	
	
	'UPGRADE_WARNING: Event optRefNumberRange.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optRefNumberRange_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optRefNumberRange.CheckedChanged
		If eventSender.Checked Then
			lblFromRefNumber.Enabled = True
			lblToRefNumber.Enabled = True
			txtFromRefNumber.Enabled = True
			txtToRefNumber.Enabled = True
			txtFromRefNumber.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			txtToRefNumber.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			
			cmbRefNumberCriteria.Enabled = False
			txtRefNumberPart.Enabled = False
			cmbRefNumberCriteria.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000004)
			txtRefNumberPart.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000004)
		End If
	End Sub
	
	
	'UPGRADE_WARNING: Event txtInvoiceNumber.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtInvoiceNumber_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtInvoiceNumber.TextChanged
		If booFiltersEnabled <> (Trim(txtInvoiceNumber.Text) = "") Then
			SetFiltersEnabled((Trim(txtInvoiceNumber.Text) = ""))
			booFiltersEnabled = (Trim(txtInvoiceNumber.Text) = "")
		End If
	End Sub
	
	
	Private Sub SetToFromDatesEnabled(ByRef Value As Boolean)
		lblAfter.Enabled = Value
		txtFromDate.Enabled = Value
		
		lblBefore.Enabled = Value
		txtToDate.Enabled = Value
		
		If Value And SupportsDateTime And optModified.Checked Then
			txtFromTime.Enabled = Value
			lblFromTime.Enabled = Value
			optFromAM.Enabled = Value
			optFromPM.Enabled = Value
			
			txtToTime.Enabled = Value
			lblToTime.Enabled = Value
			optToAM.Enabled = Value
			optToPM.Enabled = Value
		ElseIf Not Value Then 
			optFromAM.Enabled = Value
			optFromPM.Enabled = Value
			optToAM.Enabled = Value
			optToPM.Enabled = Value
		End If
		
		If optTxnDate.Checked Then
			txtFromTime.Enabled = False
			lblFromTime.Enabled = False
			optFromAM.Enabled = False
			optFromPM.Enabled = False
			
			txtToTime.Enabled = False
			lblToTime.Enabled = False
			optToAM.Enabled = False
			optToPM.Enabled = False
			
			txtFromTime.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000004)
			txtToTime.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000004)
		End If
		
		If Value Then
			txtFromDate.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			
			txtToDate.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			
			If SupportsDateTime And optModified.Checked Then
				txtFromTime.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
				txtToTime.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			End If
		Else
			txtFromDate.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000004)
			txtFromTime.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000004)
			
			txtToDate.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000004)
			txtToTime.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000004)
		End If
	End Sub
	
	
	Private Sub SetMacroDatesEnabled(ByRef Value As Boolean)
		optThis.Enabled = Value
		optLast.Enabled = Value
		optWeek.Enabled = Value
		optMonth.Enabled = Value
		optFiscalQuarter.Enabled = Value
		optFiscalYear.Enabled = Value
	End Sub
	
	
	Private Sub SetFiltersEnabled(ByRef Value As Boolean)
		Label2.Enabled = Value
		
		optModified.Enabled = Value
		optTxnDate.Enabled = Value
		optMacroDate.Enabled = Value
		
		If Value Then
			If (optModified.Checked Or optTxnDate.Checked) Then
				SetToFromDatesEnabled(True)
			ElseIf optMacroDate.Checked Then 
				SetMacroDatesEnabled(True)
			End If
		Else
			SetToFromDatesEnabled(False)
			SetMacroDatesEnabled(False)
		End If
		cmdClearDateSelection.Enabled = Value
		
		Label3.Enabled = Value
		cmbCustomers.Enabled = Value
		chkCustomerChildren.Enabled = Value
		
		Label4.Enabled = Value
		cmbAccounts.Enabled = Value
		chkAccountChildren.Enabled = Value
		
		optRefNumberRange.Enabled = Value
		lblFromRefNumber.Enabled = Value
		txtFromRefNumber.Enabled = Value
		lblToRefNumber.Enabled = Value
		txtToRefNumber.Enabled = Value
		
		optRefNumberFilter.Enabled = Value
		cmbRefNumberCriteria.Enabled = Value
		txtRefNumberPart.Enabled = Value
		
		cmdClearRefNumber.Enabled = Value
		
		Label7.Enabled = Value
		optAll.Enabled = Value
		optUnpaidOnly.Enabled = Value
		optPaidOnly.Enabled = Value
		Label8.Enabled = Value
		
		If Value Then
			cmbCustomers.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			cmbAccounts.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			txtFromRefNumber.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			txtToRefNumber.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			cmbRefNumberCriteria.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			txtRefNumberPart.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
		Else
			cmbCustomers.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000004)
			cmbAccounts.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000004)
			txtFromRefNumber.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000004)
			txtToRefNumber.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000004)
			cmbRefNumberCriteria.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000004)
			txtRefNumberPart.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000004)
		End If
	End Sub
	
	
	Private Sub ClearQueryVariables()
		strRefNumber = ""
		strFromDateTime = ""
		strToDateTime = ""
		strDateQueryType = ""
		strDateMacro = ""
		strCustomerJob = ""
		booCustomerWithChildren = False
		strAccount = ""
		booAccountWithChildren = False
		strFromRefNumberRange = ""
		strToRefNumberRange = ""
		strRefNumberPiece = ""
		strRefNumberCriteria = ""
		strPaidStatus = ""
	End Sub
	
	
	Private Function SetQueryVariables() As Boolean
		
		Dim strHours As String
		Dim strMinutes As String

        'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If Not String.IsNullOrEmpty(Trim(txtInvoiceNumber.Text)) Then
            strRefNumber = Trim(txtInvoiceNumber.Text)
            'If we have an invoice number that's all we need
            SetQueryVariables = True
            Exit Function
        End If

        If optModified.Checked Then
			If DateValid((txtFromDate.Text), "From") Then
				If TimeValid((txtFromTime.Text), "From") Then
					If DateValid((txtToDate.Text), "To") Then
						If TimeValid((txtToTime.Text), "To") Then

                            'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                            If Not String.IsNullOrEmpty(Trim(txtFromDate.Text)) Then
                                strFromDateTime = DateTimeString((txtFromDate.Text), (txtFromTime.Text), (optFromAM.Checked), SupportsDateTime)
                            End If

                            'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                            If Not String.IsNullOrEmpty(Trim(txtToDate.Text)) Then
                                strToDateTime = DateTimeString((txtToDate.Text), (txtToTime.Text), (optToAM.Checked), SupportsDateTime)
                            End If

                            'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                            If Not String.IsNullOrEmpty(strFromDateTime) And Not String.IsNullOrEmpty(strToDateTime) Then
                                If CDate(Replace(strFromDateTime, "T", " ")) > CDate(Replace(strToDateTime, "T", " ")) Then
                                    MsgBox("The from date must be before the to date")
                                    SetQueryVariables = False
                                    Exit Function
                                End If
                            End If
                        Else
							SetQueryVariables = False
							Exit Function
						End If
					Else
						SetQueryVariables = False
						Exit Function
					End If
				Else
					SetQueryVariables = False
					Exit Function
				End If
			Else
				SetQueryVariables = False
				Exit Function
			End If
			
			strDateQueryType = "ModifiedDateRangeFilter"
		ElseIf optTxnDate.Checked Then 
			If DateValid((txtFromDate.Text), "From") Then
				If DateValid((txtToDate.Text), "To") Then

                    'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                    If Not String.IsNullOrEmpty(Trim(txtFromDate.Text)) Then
                        strFromDateTime = DateTimeString((txtFromDate.Text), (txtFromTime.Text), (optFromAM.Checked), False)
                    End If

                    'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                    If Not String.IsNullOrEmpty(Trim(txtToDate.Text)) Then
                        strToDateTime = DateTimeString((txtToDate.Text), (txtToTime.Text), (optToAM.Checked), False)
                    End If

                    'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                    If Not String.IsNullOrEmpty(strFromDateTime) And Not String.IsNullOrEmpty(strToDateTime) Then
                        If CDate(strFromDateTime) > CDate(strToDateTime) Then
                            MsgBox("The from date must be before the to date")
                            SetQueryVariables = False
                            Exit Function
                        End If
                    End If
                    strDateQueryType = "TxnDateRangeFilter"
				Else
					SetQueryVariables = False
					Exit Function
				End If
			Else
				SetQueryVariables = False
				Exit Function
			End If
		ElseIf optMacroDate.Checked Then 
			If optThis.Checked Then
				strDateMacro = "This"
			ElseIf optLast.Checked Then 
				strDateMacro = "Last"
			Else
				MsgBox("You must select either This (to date) or Last when using the Date Macro filter")
				SetQueryVariables = False
				strDateMacro = ""
				Exit Function
			End If
			
			If optWeek.Checked Then
				strDateMacro = strDateMacro & "Week"
			ElseIf optMonth.Checked Then 
				strDateMacro = strDateMacro & "Month"
			ElseIf optFiscalQuarter.Checked Then 
				strDateMacro = strDateMacro & "FiscalQuarter"
			ElseIf optFiscalYear.Checked Then 
				strDateMacro = strDateMacro & "FiscalYear"
			Else
				MsgBox("You much choose, Week, Month, Fiscal Quarter or Fiscal Year when using the Date Macro filter")
				SetQueryVariables = False
				strDateMacro = ""
				Exit Function
			End If
			
			If optThis.Checked Then
				strDateMacro = strDateMacro & "ToDate"
			End If
		End If

        'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If Not String.IsNullOrEmpty(Trim(cmbCustomers.Text)) Then
            strCustomerJob = Trim(cmbCustomers.Text)
            booCustomerWithChildren = (chkCustomerChildren.CheckState = 1)
        End If

        'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If Not String.IsNullOrEmpty(Trim(cmbAccounts.Text)) Then
            strAccount = Trim(cmbAccounts.Text)
            booAccountWithChildren = (chkAccountChildren.CheckState = 1)
        End If

        If optRefNumberRange.Checked Then
			strFromRefNumberRange = Trim(txtFromRefNumber.Text)
			strToRefNumberRange = Trim(txtToRefNumber.Text)
		ElseIf optRefNumberFilter.Checked Then 
			strRefNumberCriteria = cmbRefNumberCriteria.Text
			strRefNumberCriteria = Replace(strRefNumberCriteria, " ", "")
			strRefNumberPiece = Trim(txtRefNumberPart.Text)
		End If
		
		If optAll.Checked Then
			strPaidStatus = "All"
		ElseIf optPaidOnly.Checked Then 
			strPaidStatus = "PaidOnly"
		Else
			strPaidStatus = "NotPaidOnly"
		End If
		
		SetQueryVariables = True
	End Function
End Class