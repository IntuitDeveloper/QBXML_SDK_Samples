Option Strict Off
Option Explicit On
Module UseMacroDefMacroModule
	'----------------------------------------------------------
	' Module: UseMacroDefMacroModule
	'
	' Description: This module contains the code which illustrates using
	'               the DefMacro feature.  All other functions used in this
	'               program but not explicitly showing
	'               the DefMacro feature are located in the EstimatesModule
	'
	' Routines: OpenConnectionBeginSession
	'             Opens a connection and begins a sesson with the
	'             currently open company file.  If a company isn't open,
	'             the routine will display a message and then exit the
	'             program.
	'
	'           EndSessionCloseConnection
	'             Calls EndSession and CloseConnection if the boolean
	'             booSessionBegun is true.
	'
	'           BuildInvoiceAddRequest
	'             This function will set the invoiceAdd.defMacro field
	'             Also some other invoiceAdd fields will be set.
	'
	'           BuildReceivePaymentRequest
	'             This function will set the receivePaymentAdd
	'             appliedToTxnAdd.TxnID.SetValueUseMacro field with
	'             the defTxnID value
	'
	'           CreateDefTxnID
	'             This procedure will create a defTxnID to be used
	'             in the InvoiceAdd and ReceivePaymentAdd transactions.
	'
	'           GetSessionManager
	'             This procedure will return the sessionManager value to
	'             procedures not in this module
	'
	'           SendRequestToQB
	'             This procedure will send the request to QuickBooks using
	'             the DoRequests call.
	'
	'           SendOneQBRequest
	'             This procedure will illustrate the defMacro feature by
	'             structuring an InvoiceAdd and ReceivePaymentAdd in the
	'             same request.  The InvoiceAdd will have a defined TxnID
	'             using the defMacro feature.  The ReceivePaymentAdd will
	'             refer to the InvoiceAdd transaction by using the TxnID
	'             supplied in the invoiceAdd.defMacro field.
	'
	'           SendTwoQBRequests
	'             This procedure is essentially the same as the
	'             SendOneQBRequest except that the InvoiceAdd and
	'             ReceivePaymentAdd are done in two requests.
	'             This procedure will show that the defMacro value
	'             is persistent over requests.
	'
	' Copyright © 2002-2020 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	' Updated to QBFC 12.0 and fixed setting max QBXML version: 09/2012
	'
	'----------------------------------------------------------
	
	Dim booSessionBegun As Boolean
	Dim booConnectionOpened As Boolean
	
	'Module objects
	'UPGRADE_WARNING: Arrays in structure sessionManager may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Dim sessionManager As QBFC15Lib.QBSessionManager
	Dim msgSetRequest As QBFC15Lib.IMsgSetRequest
	
	Public Sub OpenConnectionBeginSession()
		Dim estimateCollection As Object
		
		If (booSessionBegun) Then
			Exit Sub
		End If
		
		On Error GoTo Errs
		
		estimateCollection = New Collection
		
		' create the new QBSessionManager object
		If (sessionManager Is Nothing) Then
			sessionManager = New QBFC15Lib.QBSessionManager
		End If
		
		' open the connection to QuickBooks
		If (Not booConnectionOpened) Then
			sessionManager.OpenConnection("", "IDN UseMacro DefMacro Sample - QBFC")
			booConnectionOpened = True
		End If
		
		sessionManager.BeginSession("", QBFC15Lib.ENOpenMode.omDontCare)
		booSessionBegun = True
		
		'Check to make sure the QuickBooks we're working with supports version 2 of the SDK
		Dim booSupports2dot0 As Boolean
		booSupports2dot0 = False
		Dim supportedVersion As String
        supportedVersion = CStr(Val(QBFCLatestVersion(sessionManager)))
        If (supportedVersion >= CStr(15.0#)) Then
            booSupports2dot0 = True
            msgSetRequest = sessionManager.CreateMsgSetRequest("US", 15, 0)
        ElseIf (supportedVersion >= CStr(6.0#)) Then
            booSupports2dot0 = True
            msgSetRequest = sessionManager.CreateMsgSetRequest("US", 6, 0)
        ElseIf (supportedVersion >= CStr(5.0#)) Then
            booSupports2dot0 = True
            msgSetRequest = sessionManager.CreateMsgSetRequest("US", 5, 0)
        ElseIf (supportedVersion >= CStr(4.0#)) Then
            booSupports2dot0 = True
            msgSetRequest = sessionManager.CreateMsgSetRequest("US", 4, 0)
        ElseIf (supportedVersion >= CStr(3.0#)) Then
            booSupports2dot0 = True
            msgSetRequest = sessionManager.CreateMsgSetRequest("US", 3, 0)
        ElseIf (supportedVersion >= CStr(2.0#)) Then
            booSupports2dot0 = True
            msgSetRequest = sessionManager.CreateMsgSetRequest("US", 2, 0)
        End If
		
		If Not booSupports2dot0 Then
            MsgBox("This program only runs against QuickBooks installations which support the 15.0 qbXML spec.  Your version of QuickBooks does not support qbXML 15.0")
            sessionManager.EndSession()
			sessionManager.CloseConnection()
			End
		End If
		Exit Sub
		
Errs: 
		If Err.Number = &H80040416 Then
			MsgBox("You must have QuickBooks running with the company" & vbCrLf & "file open to use this program.")
			sessionManager.CloseConnection()
			End
		ElseIf Err.Number = &H80040422 Then 
			MsgBox("This QuickBooks company file is open in single user mode and" & vbCrLf & "another application is already accessing it.  Please exit the" & vbCrLf & "other application and run this application again.")
			sessionManager.CloseConnection()
			End
		Else
			MsgBox("Error in OpenConnectionBeginSession" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)
			
			If booSessionBegun Then
				sessionManager.EndSession()
			End If
			
			If booConnectionOpened Then
				sessionManager.CloseConnection()
			End If
			End
		End If
	End Sub
	
	Public Sub EndSessionCloseConnection()
		On Error GoTo Errs
		If booSessionBegun Then
			sessionManager.EndSession()
			sessionManager.CloseConnection()
		End If
		Exit Sub
Errs: 
		MsgBox("Error in EndSessionCloseConnection" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)
		
	End Sub
	
	Public Sub SendOneQBRequest(ByRef selectedIndex As Short)
		On Error GoTo Errs
		
		Dim defTxnID As String
		Dim customerID As String
		Dim msgSetResponse As QBFC15Lib.IMsgSetResponse
		Dim paymentAmount As String
		
		msgSetRequest.ClearRequests()
		
		' create a defTxnID
		defTxnID = CreateDefTxnID()
		
		' create the InvoiceAdd request
		If (Not BuildInvoiceAddRequest(defTxnID, selectedIndex, customerID, paymentAmount)) Then
			Exit Sub
		End If
		
		' create the ReceivePayment request
		' this ReceivePayment will refer to the previous invoice using the defTxnID
		If (Not BuildReceivePaymentRequest(defTxnID, customerID, paymentAmount)) Then
			Exit Sub
		End If
		
		' send the InvoiceAdd and ReceivePayment request to QB
		msgSetResponse = SendRequestToQB(msgSetRequest)
		
		' see if the request was successful
		If (Not ProcessResponse(msgSetResponse)) Then
			Exit Sub
		End If
		
		MsgBox("Success Adding Invoice and ReceivePayment in one request")
		
		Exit Sub
Errs: 
		MsgBox("Error in SendOneQBRequest" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)
		
	End Sub
	
	Public Sub SendTwoQBRequests(ByRef selectedIndex As Short)
		On Error GoTo Errs
		
		Dim defTxnID As String
		Dim msgSetResponse As QBFC15Lib.IMsgSetResponse
		Dim customerID As String
		Dim paymentAmount As String
		
		msgSetRequest.ClearRequests()
		
		' create a defTxnID
		defTxnID = CreateDefTxnID
		
		' create the InvoiceAdd request
		If (Not BuildInvoiceAddRequest(defTxnID, selectedIndex, customerID, paymentAmount)) Then
			Exit Sub
		End If
		
		' send the InvoiceAdd request to QB
		msgSetResponse = SendRequestToQB(msgSetRequest)
		
		' see if the request was successful
		If (Not ProcessResponse(msgSetResponse)) Then
			Exit Sub
		End If
		
		msgSetRequest.ClearRequests()
		
		' create the ReceivePayment request
		' this ReceivePayment will refer to the previous invoice using the defTxnID
		If (Not BuildReceivePaymentRequest(defTxnID, customerID, paymentAmount)) Then
			Exit Sub
		End If
		
		' send the ReceivePayment request to QB
		msgSetResponse = SendRequestToQB(msgSetRequest)
		
		' see if the request was successful
		If (Not ProcessResponse(msgSetResponse)) Then
			Exit Sub
		End If
		
		MsgBox("Success Adding Invoice and ReceivePayment in two requests")
		
		Exit Sub
Errs: 
		MsgBox("Error in SendTwoQBRequests" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)
		
	End Sub
	
	Public Function BuildInvoiceAddRequest(ByRef defTxnID As String, ByRef selectedIndex As Short, ByRef customerID As String, ByRef paymentAmount As String) As Boolean
		On Error GoTo Errs
		BuildInvoiceAddRequest = False
		
		' Add the InvoiceAdd request
		Dim invoiceAdd As QBFC15Lib.IInvoiceAdd
		invoiceAdd = msgSetRequest.AppendInvoiceAddRq()
		
		invoiceAdd.defMacro.SetValue((defTxnID))
		
		' build the rest of the invoiceAdd request by setting field values
		' from the estimate transaction selected in the list box
		FillInFieldsFromEstimate(invoiceAdd, selectedIndex, customerID, paymentAmount)
		
		BuildInvoiceAddRequest = True
		
		Exit Function
Errs: 
		MsgBox("Error in BuildInvoiceAddRequest" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)
		
		BuildInvoiceAddRequest = False
	End Function
	
	
	Public Function BuildReceivePaymentRequest(ByRef defTxnID As String, ByRef customerID As String, ByRef paymentAmount As String) As Boolean
		Dim appliedToTxnAdd As Object
		
		On Error GoTo Errs
		BuildReceivePaymentRequest = False
		
		' Add the ReceivePaymentAdd request
		Dim receivePaymentAdd As QBFC15Lib.IReceivePaymentAdd
		receivePaymentAdd = msgSetRequest.AppendReceivePaymentAddRq()
		
		' Set the values in the request
		' set the customer ListID
		' we are using the same customer as in the invoiceAdd
		receivePaymentAdd.CustomerRef.ListID.SetValue(customerID)
		
		receivePaymentAdd.TotalAmount.SetValue(CDbl(paymentAmount))
		
		' set the invoice TxnID on the AppliedToTxn
		appliedToTxnAdd = receivePaymentAdd.ORApplyPayment.AppliedToTxnAddList.Append()
		'UPGRADE_WARNING: Couldn't resolve default property of object appliedToTxnAdd.TxnID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		appliedToTxnAdd.TxnID.SetValueUseMacro(defTxnID)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object appliedToTxnAdd.paymentAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		appliedToTxnAdd.paymentAmount.SetValue(CDbl(paymentAmount))
		
		BuildReceivePaymentRequest = True
		
		Exit Function
Errs: 
		MsgBox("Error in BuildReceivePaymentRequest" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)
		
		BuildReceivePaymentRequest = False
	End Function
	
	Private Function SendRequestToQB(ByRef request As QBFC15Lib.IMsgSetRequest) As QBFC15Lib.IMsgSetResponse
		On Error GoTo Errs
		
		' set the OnError attribute to stopOnError
		request.Attributes.OnError = QBFC15Lib.ENRqOnError.roeStop
		
		'MsgBox request.ToXMLString()
		
		SendRequestToQB = sessionManager.DoRequests(request)
		
		'MsgBox SendRequestToQB.ToXMLString()
		
		Exit Function
Errs: 
		MsgBox("Error in SendRequestToQB" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)
	End Function
	
	Public Function CreateDefTxnID() As String
		On Error GoTo Errs
		
		CreateDefTxnID = "TxnID:" & VB6.Format(Now, "yyyymmddhhmmss")
		
		Exit Function
Errs: 
		MsgBox("Error in CreateDefTxnID" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)
		
	End Function
	
	Public Function GetSessionManager() As QBFC15Lib.QBSessionManager
		On Error GoTo Errs
		
		GetSessionManager = sessionManager
		
		Exit Function
Errs: 
		MsgBox("Error in GetSessionManager" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)
	End Function
	
	Function QBFCLatestVersion(ByRef sessionManager As QBFC15Lib.QBSessionManager) As String
		Dim strXMLVersions() As String
		'Should be able to use this, but there appears to be a bug that may cause 2.0 to be returned
		'when it should not.
		'strXMLVersions = SessionManager.QBXMLVersionsForSession
		
		Dim msgset As QBFC15Lib.IMsgSetRequest
		'Use oldest version to ensure that we work with any QuickBooks (US)
		msgset = sessionManager.CreateMsgSetRequest("US", 1, 0)
		msgset.AppendHostQueryRq()
		Dim QueryResponse As QBFC15Lib.IMsgSetResponse
		QueryResponse = sessionManager.DoRequests(msgset)
		Dim response As QBFC15Lib.IResponse
		
		' The response list contains only one response,
		' which corresponds to our single HostQuery request
		response = QueryResponse.responseList.GetAt(0)
		Dim HostResponse As QBFC15Lib.IHostRet
		HostResponse = response.Detail
		Dim supportedVersions As QBFC15Lib.IBSTRList
		supportedVersions = HostResponse.SupportedQBXMLVersionList
		
		Dim i As Integer
		Dim vers As Double
		Dim LastVers As Double
		LastVers = 0
		For i = 0 To supportedVersions.count - 1
			vers = Val(supportedVersions.GetAt(i))
			If (vers > LastVers) Then
				LastVers = vers
				QBFCLatestVersion = supportedVersions.GetAt(i)
			End If
		Next i
	End Function
End Module