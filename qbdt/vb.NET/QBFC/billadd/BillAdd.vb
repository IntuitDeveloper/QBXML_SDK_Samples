Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks
Friend Class BillAdd
	Inherits System.Windows.Forms.Form
	'
	' This is the primary form for this application, it communicates with QuickBooks to
	' populate list boxes with vendor and account lists and allows the user to select a
	' vendor, an accounts payable account, and an amount for the bill and adds it to
	' QuickBooks running either on the desktop or the online edition based on communication
	' preferences set by the user (see the CommPrefsDlg).
	'
	
	'
	' We are both a desktop and an Online edition application, so we have some constants
	' here that define us in either environment.
	'
	'Used for the registry and to Identify us to QuickBooks Desktop
	'
	Const appName As String = "IDN Bill Add Sample"
	'
	' Used to identify us to QuickBooks Online Edition
	'
	Const appLogin As String = "billadd.developer.intuit.com"
	Const appID As String = "60177302"
	Const appVersion As String = "1.0"
	
	'
	' Private variables to hold our connection to QuickBooks
	'
	'UPGRADE_WARNING: Arrays in structure SessionManager may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Private SessionManager As QBFC15Lib.QBSessionManager
	
	' User clicked the "Communication Preferences" button
	'
	Private Sub CommPrefs_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CommPrefs.Click
		DoCommPrefs()
		If (SessionManager Is Nothing) Then
			FillInForm()
		End If
	End Sub
	
	Private Sub DoCommPrefs()
		'
		' Show the Communication preferences dialog
		'
		CommPrefsDlg.ShowDialog()
		
		'
		' If the user clicked OK and we already had a session manger, get rid of
		' the existing session manager and then fill the form back in with the
		' data from QuickBooks
		'
		If (CommPrefsDlg.isOK And Not SessionManager Is Nothing) Then
			CommPrefsDlg.DisconnectFromQuickBooks(SessionManager)
			'UPGRADE_NOTE: Object SessionManager may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			SessionManager = Nothing
		End If
		
	End Sub
	
	'
	' This is where most of the work happens, we build the BillAddRq and send the bill to
	' QuickBooks (note that we don't have to do anything special for Online or Desktop, though
	' there are parts of the BillAddRq that are not supported by QuickBooks Online Edition (ItemLineAdd
	' and ItemGroupLineAdd) we aren't using them here so we don't have to check if we are Online or
	' on the desktop.
	'
	Private Sub EnterBill_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles EnterBill.Click
		Dim msg As Object
		Dim vendor As String
		Dim account As String
		Dim amount As String
		'
		' Do the validation of the bill information
		'
		vendor = VendorList.Text
		If ("" = vendor) Then
			MsgBox("Must specify a vendor for the bill!", MsgBoxStyle.Exclamation, "Missing parameter")
			Exit Sub
		End If
		account = AccountList.Text
		If ("" = account) Then
			MsgBox("Must specify an account for this bill!", MsgBoxStyle.Exclamation, "Missing Parameter")
		End If
		amount = BillAmount.Text
		If ("" = amount) Then
			MsgBox("Must specify an amount for the bill!", MsgBoxStyle.Exclamation, "Missing parameter")
			Exit Sub
		End If
		'
		' Make sure we have a session manager
		'
		If (SessionManager Is Nothing) Then
			SessionManager = CommPrefsDlg.ConnectToQuickBooks(appName)
		End If
		
		'
		' Now build the BillAddRq
		'
		Dim BillAddSet As QBFC15Lib.IMsgSetRequest
		Dim BillAddRq As QBFC15Lib.IBillAdd
		Dim Expense As QBFC15Lib.IExpenseLineAdd
		Dim ResponseSet As QBFC15Lib.IMsgSetResponse
		Dim response As QBFC15Lib.IResponse
		Dim billRet As QBFC15Lib.IBillRet
		If (Not SessionManager Is Nothing) Then
			BillAddSet = GetLatestMsgSetRequest(SessionManager)
			BillAddRq = BillAddSet.AppendBillAddRq
			
			BillAddRq.TxnDate.setValue(Now)
			BillAddRq.VendorRef.FullName.setValue(vendor)
			
			Expense = BillAddRq.ExpenseLineAddList.Append
			Expense.amount.setValue(CDbl(amount))
			Expense.AccountRef.FullName.setValue(account)
			Expense.Memo.setValue("Sample expense line")
			
			' Perform the request and obtain a response from QuickBooks
			ResponseSet = SessionManager.DoRequests(BillAddSet)
			
			' Uncomment the following to see the request and response XML for debugging
			' MsgBox BillAddSet.ToXMLString, vbOKOnly, "RequestXML"
			' MsgBox ResponseSet.ToXMLString, vbOKOnly, "ResponseXML"
			
			' Interpret the response
			
			' The response list contains only one response,
			' which corresponds to our single BillAdd request
			response = ResponseSet.ResponseList.GetAt(0)
			
			'UPGRADE_WARNING: Couldn't resolve default property of object msg. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			msg = "Status: Code = " & CStr(response.StatusCode) & ", Message = " & response.StatusMessage & ", Severity = " & response.StatusSeverity & vbCrLf
			
			' The Detail property of the IResponse object
			' returns a Ret object for Add and Mod requests.
			' In this case, the Ret object is IBillRet.
			
			' For help finding out the Detail's type, uncomment the following line:
			' MsgBox response.Detail.Type.GetAsString
			
			billRet = response.Detail
			'
			' We have to make sure we really have a BillRet, since if something went wrong we
			' won't.
			'
			If (Not (billRet Is Nothing)) Then
				'
				' Retrieve fields that are guaranteed to be returned in the response
				' We don't have to check that these fields exist in the response
				'
				'UPGRADE_WARNING: Couldn't resolve default property of object msg. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				msg = msg & vbCr & "BillRet: EditSequence = " & billRet.EditSequence.getValue
				'UPGRADE_WARNING: Couldn't resolve default property of object msg. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				msg = msg & ", Amount Due = " & billRet.AmountDue.GetAsString
				
				'
				' Retrieve a field that is not guaranteed to be returned from QuickBooks
				' Note that in this case we must first check whether the field exists in the
				' response before trying to use it.
				'
				If (Not billRet.DueDate Is Nothing) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object msg. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					msg = msg & ", Due Date = " & billRet.DueDate.getValue
				End If
			End If
			'
			' Display a message showing some of the interesting fields of the response
			'
			MsgBox(msg)
		End If
	End Sub
	
	'
	' Load the form, this is where we set up the information needed by the CommPrefsDlg
	'
	Private Sub BillAdd_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		CommPrefsDlg.SetAppLogin(appName, appLogin)
		CommPrefsDlg.SetAppID(appName, appID)
		CommPrefsDlg.SetAppVersion(appName, appVersion)
		DoCommPrefs()
		FillInForm()
	End Sub
	
	'
	' Fill in the list boxes in the form with the vendors and accounts payable accounts
	'
	Private Sub FillInForm()
		Dim i As Object
		'
		' Make sure we are connected to QuickBooks
		'
		If (SessionManager Is Nothing) Then
			SessionManager = CommPrefsDlg.ConnectToQuickBooks(appName)
		End If
		'
		' Now build the VendorQuery and AccountQuery requests to fill in our list boxes
		'
		Dim ListQuerySet As QBFC15Lib.IMsgSetRequest
		Dim VendorQuery As QBFC15Lib.IVendorQuery
		Dim AccountQuery As QBFC15Lib.IAccountQuery
		Dim ListResponse As QBFC15Lib.IMsgSetResponse
		Dim ListQueryXML As String
		Dim response As QBFC15Lib.IResponse
		Dim QBVendorList As QBFC15Lib.IVendorRetList
		Dim vendor As QBFC15Lib.IVendorRet
		Dim QBAccountList As QBFC15Lib.IAccountRetList
		Dim account As QBFC15Lib.IAccountRet
		If (Not SessionManager Is Nothing) Then
			'
			' Set up the MsgSet to query the lists
			'
			ListQuerySet = GetLatestMsgSetRequest(SessionManager)
			ListQuerySet.Attributes.OnError = QBFC15Lib.ENRqOnError.roeContinue
			
			'
			' Add the VendorQuery
			'
			VendorQuery = ListQuerySet.AppendVendorQueryRq
			
			'
			' Add the AccountQuery to get CostOfGoodsSold and Expense accounts
			'
			AccountQuery = ListQuerySet.AppendAccountQueryRq
			AccountQuery.ORAccountListQuery.AccountListFilter.AccountTypeList.Add(QBFC15Lib.ENAccountType.atCostOfGoodsSold)
			AccountQuery.ORAccountListQuery.AccountListFilter.AccountTypeList.Add(QBFC15Lib.ENAccountType.atExpense)
			AccountQuery = ListQuerySet.AppendAccountQueryRq
			
			'
			' Now send the query and get QBFC to parse the response
			'
			ListQueryXML = ListQuerySet.ToXMLString
			ListResponse = SessionManager.DoRequests(ListQuerySet)
			If (ListResponse.ResponseList Is Nothing) Then
				' This could happen if there is no session ticket available to QBOE
				' For example the session ticket we had expired
				Exit Sub
			End If
			
			'
			' The response list contains only two responses,
			' the first response corresponds to our vendor query, populate
			' the VendorList listbox with the results of that query.
			'
			response = ListResponse.ResponseList.GetAt(0)
			QBVendorList = response.Detail
			VendorList.Items.Clear()
			If (Not QBVendorList Is Nothing) Then
				For i = 0 To (QBVendorList.Count - 1)
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					vendor = QBVendorList.GetAt(i)
					VendorList.Items.Add(vendor.Name.getValue)
				Next i
			End If
			
			'
			' the second response corresponds to our account query, populate
			' the AccountList listbox with the results of that query
			'
			response = ListResponse.ResponseList.GetAt(1)
			QBAccountList = response.Detail
			AccountList.Items.Clear()
			If (Not QBAccountList Is Nothing) Then
				For i = 0 To (QBAccountList.Count - 1)
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					account = QBAccountList.GetAt(i)
					AccountList.Items.Add(account.FullName.getValue)
				Next i
			End If
		End If
	End Sub
	
	Function QBFCLatestVersion(ByRef SessionManager As QBFC15Lib.QBSessionManager) As String
		Dim strXMLVersions() As String
		'Should be able to use this, but there appears to be a bug that may cause 2.0 to be returned
		'when it should not.
		'strXMLVersions = SessionManager.QBXMLVersionsForSession
		
		Dim msgset As QBFC15Lib.IMsgSetRequest
		On Error GoTo use2_0msgSetforOE
		'Use oldest version to ensure that we work with any QuickBooks (US)
		msgset = SessionManager.CreateMsgSetRequest("US", 1, 0)
		GoTo gotmsgset
use2_0msgSetforOE: 
		'remove error handler to avoid a possible infinite loop
		On Error GoTo 0
		msgset = SessionManager.CreateMsgSetRequest("US", 2, 0)
		QBFCLatestVersion = "2.1"
		Exit Function
gotmsgset: 
		msgset.AppendHostQueryRq()
		Dim QueryResponse As QBFC15Lib.IMsgSetResponse
		QueryResponse = SessionManager.DoRequests(msgset)
		Dim response As QBFC15Lib.IResponse
		
		' The response list contains only one response,
		' which corresponds to our single HostQuery request
		response = QueryResponse.ResponseList.GetAt(0)
		Dim HostResponse As QBFC15Lib.IHostRet
		HostResponse = response.Detail
		Dim supportedVersions As QBFC15Lib.IBSTRList
		supportedVersions = HostResponse.SupportedQBXMLVersionList
		
		Dim i As Integer
		Dim vers As Double
		Dim LastVers As Double
		LastVers = 0
		For i = 0 To supportedVersions.Count - 1
			vers = Val(supportedVersions.GetAt(i))
			If (vers > LastVers) Then
				LastVers = vers
				QBFCLatestVersion = supportedVersions.GetAt(i)
			End If
		Next i
	End Function
	
	Public Function GetLatestMsgSetRequest(ByRef SessionManager As QBFC15Lib.QBSessionManager) As QBFC15Lib.IMsgSetRequest
		Dim supportedVersion As Double
		supportedVersion = Val(QBFCLatestVersion(SessionManager))
		If (supportedVersion >= 6#) Then
			GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 6, 0)
		ElseIf (supportedVersion >= 5#) Then 
			GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 5, 0)
		ElseIf (supportedVersion >= 4#) Then 
			GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 4, 0)
		ElseIf (supportedVersion >= 3#) Then 
			GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 3, 0)
		ElseIf (supportedVersion = 2#) Then 
			GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 2, 0)
		ElseIf (supportedVersion = 1.1) Then 
			GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 1, 1)
		Else
			MsgBox("You are apparently running QuickBooks 2002 Release 1, we strongly recommend that you use QuickBooks' online update feature to obtain the latest fixes and enhancements", MsgBoxStyle.Exclamation)
			GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 1, 0)
		End If
	End Function

    Private Sub BillAdd_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        If Not SessionManager Is Nothing Then
            SessionManager.EndSession()
            SessionManager.CloseConnection()
        End If
        CommPrefsDlg.Close()

    End Sub

    Private Sub BillAdd_Paint(sender As Object, e As PaintEventArgs) Handles MyBase.Paint
        Dim myGraphics As Graphics = Me.CreateGraphics
        Dim myPen As Pen
        myPen = New Pen(Brushes.Black, 1)

        e.Graphics.DrawLine(myPen, 2, 152, 500, 152)
        e.Graphics.DrawLine(myPen, 2, 160, 500, 160)
    End Sub
End Class