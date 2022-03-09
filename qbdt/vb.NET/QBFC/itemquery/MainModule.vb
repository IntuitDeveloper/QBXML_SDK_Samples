Option Strict Off
Option Explicit On
Module MainModule
	'-----------------------------------------------------------
	' Module: MainModule
	'
	' Description:  This module demonstrates querying items using QBFC.
	'               Querying items is similar to querying terms or entities
	'               in the sense that the returned detail object is a list of
	'               'OR' objects  - ORItemRet in this case.
	'               It includes examples of the following:
	'                   - Constructing a query
	'                   - Looping through the list of items in a response
	'                   - Reading data from an OR object
	'                   - Obtaining returned fields, also checking for
	'                     fields that may not exist in the response
	'
	' Created On: 1/22/2002
	' Updated to QBFC 2.0: 08/2002
	' Updated to QBFC 5.0: 09/2005
	'
	' Copyright © 2021-2022 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	' Updated to QBFC 12.0 and fixed setting max QBXML version: 09/2012
	'----------------------------------------------------------
	
	
	Const cAppID As String = ""
	Const cAppName As String = "IDN Desktop VB QBFC ItemQuery"
	
	
	
	
	Public Sub Main()
		
		On Error GoTo ErrHandler
		
		Dim msg As String
		Dim okCancel As MsgBoxResult
		
		msg = "This sample queries for the first 30 items." & vbCr & vbCr & "QuickBooks must be running with a data file open." & vbCr
		
		okCancel = MsgBox(msg, MsgBoxStyle.OKCancel)
		If okCancel = MsgBoxResult.Cancel Then
			Exit Sub
		End If
		
		' Create the session manager object
		Dim SessionManager As New QBFC15Lib.QBSessionManager
		SessionManager.OpenConnection(cAppID, cAppName)
		SessionManager.BeginSession("", QBFC15Lib.ENOpenMode.omDontCare)
		
		' Create the message set request object
		Dim requestSet As QBFC15Lib.IMsgSetRequest
		requestSet = GetLatestMsgSetRequest(SessionManager)
		
		' Initialize the message set request's attributes
		requestSet.Attributes.OnError = QBFC15Lib.ENRqOnError.roeContinue
		
		' Append the request to the message set request object
		Dim ItemQ As QBFC15Lib.IItemQuery
		ItemQ = requestSet.AppendItemQueryRq
		ItemQ.ORListQuery.ListFilter.MaxReturned.SetValue(30)
		
		' Perform the request
		Dim responseSet As QBFC15Lib.IMsgSetResponse
		responseSet = SessionManager.DoRequests(requestSet)
		
		' Uncomment the following to see the request and response XML for debugging
		'MsgBox requestSet.ToXMLString, vbOKOnly, "RequestXML"
		'MsgBox responseSet.ToXMLString, vbOKOnly, "ResponseXML"
		
		' Interpret the response
		Dim response As QBFC15Lib.IResponse
		Dim statusMessage, statusCode, statusSeverity As Object
		
		' The response list contains only one response,
		' which corresponds to our single request
		response = responseSet.ResponseList.GetAt(0)
		
		statusCode = response.statusCode
		
		statusMessage = response.statusMessage
		
		statusSeverity = response.statusSeverity
		
		
		
		
		msg = "Status: Code = " & CStr(statusCode) & ", Message = " & statusMessage & ", Severity = " & statusSeverity & vbCrLf
		
		' The detail property of the IResponse interface is a list for most queries
		' except CompanyQuery, HostQuery, or PreferencesQuery, which don't return a list,
		' but a single Ret object.
		' In this case, the list in detail is an IORItemRetList (a list of IORItemRet objects)
		
		' For help finding out the detail's type, uncomment the following line:
		'MsgBox response.Detail.Type.GetAsString
		
		Dim orItemRetList As QBFC15Lib.IORItemRetList
		orItemRetList = response.Detail
		Dim ndx As Object
		Dim orItemRet As QBFC15Lib.IORItemRet
		Dim ItemServiceRet As QBFC15Lib.IItemServiceRet
		Dim ItemInventoryRet As QBFC15Lib.IItemInventoryRet
		Dim ItemNonInventoryRet As QBFC15Lib.IItemNonInventoryRet
		If (Not (orItemRetList Is Nothing)) Then
			
			For ndx = 0 To (orItemRetList.Count - 1)
				
				orItemRet = orItemRetList.GetAt(ndx)
				
				' The ortype property returns an enum
				' of the elements that can be contained in the OR object
				Select Case orItemRet.ortype
					
					Case QBFC15Lib.ENORItemRet.orirItemServiceRet '"orir" prefix comes from OR + Item + Ret
						ItemServiceRet = orItemRet.ItemServiceRet
						msg = msg & vbCrLf & "ItemServiceRet: FullName = " & ItemServiceRet.FullName.GetValue
						
						' Retrieving a field that may not be set
						' in this case, a reference object (IQBBaseRef)
						If (Not ItemServiceRet.SalesTaxCodeRef Is Nothing) Then
							msg = msg & ", SalesTaxCodeRef = " & ItemServiceRet.SalesTaxCodeRef.FullName.GetValue
						End If
						
					Case QBFC15Lib.ENORItemRet.orirItemInventoryRet
						ItemInventoryRet = orItemRet.ItemInventoryRet
						msg = msg & vbCrLf & "ItemInventoryRet: FullName = " & ItemInventoryRet.FullName.GetValue
						
						' Retrieving a field that may not be set
						' in this case, a quantity
						If (Not ItemInventoryRet.QuantityOnHand Is Nothing) Then
							msg = msg & ", QuantityOnHand = " & ItemInventoryRet.QuantityOnHand.GetAsString
						End If
						
					Case QBFC15Lib.ENORItemRet.orirItemNonInventoryRet
						
						ItemNonInventoryRet = orItemRet.ItemNonInventoryRet
						msg = msg & vbCrLf & "ItemNonInventoryRet: FullName = " & ItemNonInventoryRet.FullName.GetValue
						
					Case Else
						'...
						' Could have specific code for the other item types.
						msg = msg & vbCrLf & "Other Item Type"
						
				End Select
				
				
			Next 
			
		End If
		
		SessionManager.EndSession()
		SessionManager.CloseConnection()
		
		ShowItemList(msg)
		Exit Sub
		
ErrHandler: 
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
		
	End Sub
	
	
	
	Private Sub ShowItemList(ByRef msg As Object)
		
		Dim frmDisplay As New ShowItems
		
		frmDisplay.ItemsList.Text = msg
		frmDisplay.ShowDialog()
		
	End Sub
	
	Public Function GetLatestMsgSetRequest(ByRef SessionManager As QBFC15Lib.QBSessionManager) As QBFC15Lib.IMsgSetRequest
		Dim supportedVersion As String
		supportedVersion = CStr(Val(QBFCLatestVersion(SessionManager)))
		If (supportedVersion >= CStr(6#)) Then
			GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 6, 0)
		ElseIf (supportedVersion >= CStr(5#)) Then 
			GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 5, 0)
		ElseIf (supportedVersion >= CStr(4#)) Then 
			GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 4, 0)
		ElseIf (supportedVersion >= CStr(3#)) Then 
			GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 3, 0)
		ElseIf (supportedVersion >= CStr(2#)) Then 
			GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 2, 0)
		ElseIf (supportedVersion = CStr(1.1)) Then 
			GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 1, 1)
		Else
			MsgBox("You are apparently running QuickBooks 2002 Release 1, we strongly recommend that you use QuickBooks' online update feature to obtain the latest fixes and enhancements", MsgBoxStyle.Exclamation)
			GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 1, 0)
		End If
	End Function
	
	Function QBFCLatestVersion(ByRef SessionManager As QBFC15Lib.QBSessionManager) As String
		Dim strXMLVersions() As String
		'Should be able to use this, but there appears to be a bug that may cause 2.0 to be returned
		'when it should not.
		'strXMLVersions = SessionManager.QBXMLVersionsForSession
		
		Dim msgset As QBFC15Lib.IMsgSetRequest
		'Use oldest version to ensure that we work with any QuickBooks (US)
		msgset = SessionManager.CreateMsgSetRequest("US", 1, 0)
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
End Module