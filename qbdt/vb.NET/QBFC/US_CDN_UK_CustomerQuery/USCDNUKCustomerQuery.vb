Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class Form_USCDNUKCustomerQuery
	Inherits System.Windows.Forms.Form
	'----------------------------------------------------------
	' Sample name: US_CDN_UK_CustomerQuery
	'
	' Description:
	' This sample demonstrates how to call HostQuery to determine
	' version of QuickBooks running and thereby use appropriate
	' QBFC library to communicate with it.
	' The sample was intentionally kept simple in order to
	' keep the HostQuery implementation clean and easy to follow.
	'
	' Created On: 08/15/2006
	'
	' Copyright © 2006-2013 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/
	' Updated to QBFC 12.0 and fixed setting max QBXML version: 09/2012
	' updated to qbfc15
	'----------------------------------------------------------
	
	Const cAppID As String = "123"
	Const cAppName As String = "IDN Sample - US_CDN_UK_CustomerQuery"
	Const cRequestID As String = "1"
	
	Dim sessionManager As Object 'QBFCx.QBSessionManager
	Dim requestMsgSet As Object 'IMsgSetRequest
	Dim QueryResponse As Object 'IMsgSetResponse
	Dim HostResponse As Object 'IHostRet
	Dim response As Object 'IResponse
	Dim supportedVersions As Object 'IBSTRList
	Dim msgset As Object 'IMsgSetRequest
	
	
	Dim requestXML As String
	Dim responseXML As String
	Dim QBXMLLatestVersion As String
	Dim strXMLVersions() As String
	
	Dim i As Integer
	Dim vers As Double
	Dim LastVers As Double
	
	
	
	
	Private Sub Command1_Submit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1_Submit.Click
		On Error GoTo ErrHandler
		'Start with a US HostQuery
		'If it succeeds, app is talking to QB_US
		'If it fails with a specific parse error, catch it and then retry a CDN HostQuery
		HostQuery_US()
		
		'UPGRADE_WARNING: Couldn't resolve default property of object requestMsgSet.Attributes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		requestMsgSet.Attributes.OnError = QBFC15Lib.ENRqOnError.roeContinue
		If (requestMsgSet Is Nothing) Then
			Exit Sub
		End If
		
		'Add the request to the message set request object.
		Dim customerQuery As Object 'ICustomerQuery
		'UPGRADE_WARNING: Couldn't resolve default property of object requestMsgSet.AppendCustomerQueryRq. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		customerQuery = requestMsgSet.AppendCustomerQueryRq
		
		'Set the elements of ICustomerQuery.
		'UPGRADE_WARNING: Couldn't resolve default property of object customerQuery.ORCustomerListQuery. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		customerQuery.ORCustomerListQuery.CustomerListFilter.MaxReturned.SetValue(1)
		
		' Uncomment the following to see the response XML for debugging
		'UPGRADE_WARNING: Couldn't resolve default property of object requestMsgSet.ToXMLString. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		MsgBox(requestMsgSet.ToXMLString, MsgBoxStyle.OKOnly, "RequestXML")
		
		' Perform the request
		Dim responseMsgSet As Object 'IMsgSetResponse
		'UPGRADE_WARNING: Couldn't resolve default property of object sessionManager.DoRequests. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		responseMsgSet = sessionManager.DoRequests(requestMsgSet)
		
		' Uncomment the following to see the response XML for debugging
		'UPGRADE_WARNING: Couldn't resolve default property of object responseMsgSet.ToXMLString. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		MsgBox(responseMsgSet.ToXMLString, MsgBoxStyle.OKOnly, "ResponseXML")
		
		' Interpret the response
		Dim response As Object 'IResponse
		Dim statusMessage, statusCode, statusSeverity As Object
		
		' The response list contains only one response,
		' which corresponds to our single request
		'UPGRADE_WARNING: Couldn't resolve default property of object responseMsgSet.ResponseList. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		response = responseMsgSet.ResponseList.GetAt(0)
		'UPGRADE_WARNING: Couldn't resolve default property of object response.statusCode. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object statusCode. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		statusCode = response.statusCode
		'UPGRADE_WARNING: Couldn't resolve default property of object response.statusMessage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object statusMessage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		statusMessage = response.statusMessage
		'UPGRADE_WARNING: Couldn't resolve default property of object response.statusSeverity. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object statusSeverity. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		statusSeverity = response.statusSeverity
		
		'UPGRADE_WARNING: Couldn't resolve default property of object statusSeverity. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object statusMessage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object statusCode. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		MsgBox("Status: Code = " & CStr(statusCode) & ", Message = " & statusMessage & ", Severity = " & statusSeverity & vbCrLf)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object sessionManager.EndSession. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sessionManager.EndSession()
		'UPGRADE_WARNING: Couldn't resolve default property of object sessionManager.CloseConnection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sessionManager.CloseConnection()
		Exit Sub
		
ErrHandler: 
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
		Exit Sub
	End Sub
	
	
	
	
	
	
	
	Public Sub HostQuery_US()
		On Error GoTo ErrHandler1
		'Start with a US HostQuery
		'If it succeeds, app is talking to QB_US
		'If it fails with a specific parse error, catch it and then retry a CDN HostQuery
		sessionManager = New QBFC15Lib.QBSessionManager
		'UPGRADE_WARNING: Couldn't resolve default property of object sessionManager.OpenConnection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sessionManager.OpenConnection(cAppID, cAppName)
		'UPGRADE_WARNING: Couldn't resolve default property of object sessionManager.BeginSession. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sessionManager.BeginSession("", QBFC15Lib.ENOpenMode.omDontCare)
		'UPGRADE_WARNING: Couldn't resolve default property of object sessionManager.CreateMsgSetRequest. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		msgset = sessionManager.CreateMsgSetRequest("US", 1, 0)
		'UPGRADE_WARNING: Couldn't resolve default property of object msgset.AppendHostQueryRq. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		msgset.AppendHostQueryRq()
		'UPGRADE_WARNING: Couldn't resolve default property of object sessionManager.DoRequests. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		QueryResponse = sessionManager.DoRequests(msgset)
		'UPGRADE_WARNING: Couldn't resolve default property of object QueryResponse.ResponseList. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		response = QueryResponse.ResponseList.GetAt(0)
		'UPGRADE_WARNING: Couldn't resolve default property of object response.Detail. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		HostResponse = response.Detail
		'UPGRADE_WARNING: Couldn't resolve default property of object HostResponse.SupportedQBXMLVersionList. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		supportedVersions = HostResponse.SupportedQBXMLVersionList
		LastVers = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object supportedVersions.Count. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For i = 0 To supportedVersions.Count - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object supportedVersions.GetAt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			vers = Val(supportedVersions.GetAt(i))
			If (vers > LastVers) Then
				LastVers = vers
				'UPGRADE_WARNING: Couldn't resolve default property of object supportedVersions.GetAt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				QBXMLLatestVersion = supportedVersions.GetAt(i)
			End If
		Next i
		Dim supportedVersion As Double
		supportedVersion = Val(QBXMLLatestVersion)
		
		If (supportedVersion >= 6#) Then
			SwitchQBFC()
			'UPGRADE_WARNING: Couldn't resolve default property of object sessionManager.CreateMsgSetRequest. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			requestMsgSet = sessionManager.CreateMsgSetRequest("", 6, 0)
			' Note: The following is also correct as of 6.0
			' Set requestMsgSet = sessionManager.CreateMsgSetRequest("US", 6, 0)
		ElseIf (supportedVersion >= 5#) Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object sessionManager.CreateMsgSetRequest. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			requestMsgSet = sessionManager.CreateMsgSetRequest("US", 5, 0)
		ElseIf (supportedVersion >= 4#) Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object sessionManager.CreateMsgSetRequest. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			requestMsgSet = sessionManager.CreateMsgSetRequest("US", 4, 0)
		ElseIf (supportedVersion >= 3#) Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object sessionManager.CreateMsgSetRequest. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			requestMsgSet = sessionManager.CreateMsgSetRequest("US", 3, 0)
		ElseIf (supportedVersion >= 2#) Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object sessionManager.CreateMsgSetRequest. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			requestMsgSet = sessionManager.CreateMsgSetRequest("US", 2, 0)
		ElseIf (supportedVersion = 1.1) Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object sessionManager.CreateMsgSetRequest. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			requestMsgSet = sessionManager.CreateMsgSetRequest("US", 1, 1)
		Else
			MsgBox("You are apparently running QuickBooks 2002 Release 1, we strongly recommend that you use QuickBooks' online update feature to obtain the latest fixes and enhancements", MsgBoxStyle.Exclamation)
			'UPGRADE_WARNING: Couldn't resolve default property of object sessionManager.CreateMsgSetRequest. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			requestMsgSet = sessionManager.CreateMsgSetRequest("US", 1, 0)
		End If
		Exit Sub
		
ErrHandler1: 
		If Err.Description = "The version of QBXML that was requested is not supported or is unknown." Then
			'If it fails with a specific parse error, catch it and then retry a CDN HostQuery
			HostQuery_CA()
		Else
			MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
			End
		End If
	End Sub
	
	
	
	
	Public Sub HostQuery_CA()
		On Error GoTo ErrHandler2
		'Execute a CDN HostQuery
		'If it succeeds, app is talking to QB_CDN
		'If it fails with a specific parse error, catch it and then retry a UK HostQuery
		'UPGRADE_WARNING: Couldn't resolve default property of object sessionManager.CreateMsgSetRequest. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		msgset = sessionManager.CreateMsgSetRequest("CA", 2, 0)
		'UPGRADE_WARNING: Couldn't resolve default property of object msgset.AppendHostQueryRq. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		msgset.AppendHostQueryRq()
		'UPGRADE_WARNING: Couldn't resolve default property of object sessionManager.DoRequests. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		QueryResponse = sessionManager.DoRequests(msgset)
		'UPGRADE_WARNING: Couldn't resolve default property of object QueryResponse.ResponseList. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		response = QueryResponse.ResponseList.GetAt(0)
		'UPGRADE_WARNING: Couldn't resolve default property of object response.Detail. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		HostResponse = response.Detail
		'UPGRADE_WARNING: Couldn't resolve default property of object HostResponse.SupportedQBXMLVersionList. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		supportedVersions = HostResponse.SupportedQBXMLVersionList
		LastVers = 0
		Dim ca_vers As String
		'UPGRADE_WARNING: Couldn't resolve default property of object supportedVersions.Count. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For i = 0 To supportedVersions.Count - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object supportedVersions.GetAt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ca_vers = supportedVersions.GetAt(i)
			ca_vers = Mid(ca_vers, 3)
			vers = Val(ca_vers)
			If (vers > LastVers) Then
				LastVers = vers
				'UPGRADE_WARNING: Couldn't resolve default property of object supportedVersions.GetAt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				QBXMLLatestVersion = supportedVersions.GetAt(i)
			End If
		Next i
		If (VB.Left(QBXMLLatestVersion, 2) = "CA") Then
			QBXMLLatestVersion = Mid(QBXMLLatestVersion, 3, 3)
		End If
		Dim supportedVersion As Double
		supportedVersion = Val(QBXMLLatestVersion)
		If (supportedVersion >= 6#) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object sessionManager.CreateMsgSetRequest. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			requestMsgSet = sessionManager.CreateMsgSetRequest("", 6, 0)
			' Note: The following is also correct as of 6.0
			' Set requestMsgSet = sessionManager.CreateMsgSetRequest("CA", 6, 0)
		ElseIf (supportedVersion >= 3#) Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object sessionManager.CreateMsgSetRequest. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			requestMsgSet = sessionManager.CreateMsgSetRequest("CA", 3, 0)
		ElseIf (supportedVersion >= 2#) Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object sessionManager.CreateMsgSetRequest. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			requestMsgSet = sessionManager.CreateMsgSetRequest("CA", 2, 0)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object sessionManager.CreateMsgSetRequest. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			requestMsgSet = sessionManager.CreateMsgSetRequest("CA", 2, 0)
		End If
		Exit Sub
ErrHandler2: 
		If Err.Description = "The version of QBXML that was requested is not supported or is unknown." Then
			'If it fails with a specific parse error, catch it and then retry a UK HostQuery
			HostQuery_UK()
		Else
			MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
			End
		End If
	End Sub
	
	
	
	
	Public Sub HostQuery_UK()
		On Error GoTo ErrHandler3
		'qbXML UK is introduced in QBFC5
		' Updated to QBFC13
		'Execute a UK HostQuery
		'If it succeeds, app is talking to QB_UK
		'If it fails, catch it and throw error to user
		'UPGRADE_WARNING: Couldn't resolve default property of object sessionManager.CreateMsgSetRequest. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		msgset = sessionManager.CreateMsgSetRequest("UK", 2, 0)
		'UPGRADE_WARNING: Couldn't resolve default property of object msgset.AppendHostQueryRq. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		msgset.AppendHostQueryRq()
		'UPGRADE_WARNING: Couldn't resolve default property of object sessionManager.DoRequests. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		QueryResponse = sessionManager.DoRequests(msgset)
		'UPGRADE_WARNING: Couldn't resolve default property of object QueryResponse.ResponseList. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		response = QueryResponse.ResponseList.GetAt(0)
		'UPGRADE_WARNING: Couldn't resolve default property of object response.Detail. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		HostResponse = response.Detail
		'UPGRADE_WARNING: Couldn't resolve default property of object HostResponse.SupportedQBXMLVersionList. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		supportedVersions = HostResponse.SupportedQBXMLVersionList
		LastVers = 0
		Dim uk_vers As String
		'UPGRADE_WARNING: Couldn't resolve default property of object supportedVersions.Count. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For i = 0 To supportedVersions.Count - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object supportedVersions.GetAt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			uk_vers = supportedVersions.GetAt(i)
			uk_vers = Mid(uk_vers, 3)
			vers = Val(uk_vers)
			If (vers > LastVers) Then
				LastVers = vers
				'UPGRADE_WARNING: Couldn't resolve default property of object supportedVersions.GetAt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				QBXMLLatestVersion = supportedVersions.GetAt(i)
			End If
		Next i
		If (VB.Left(QBXMLLatestVersion, 2) = "UK") Then
			QBXMLLatestVersion = Mid(QBXMLLatestVersion, 3, 3)
		End If
		Dim supportedVersion As Double
		supportedVersion = Val(QBXMLLatestVersion)
		If (supportedVersion >= 6#) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object sessionManager.CreateMsgSetRequest. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			requestMsgSet = sessionManager.CreateMsgSetRequest("", 6, 0)
			' Note: The following is also correct as of 6.0
			' Set requestMsgSet = sessionManager.CreateMsgSetRequest("UK", 6, 0)
		ElseIf (supportedVersion >= 3#) Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object sessionManager.CreateMsgSetRequest. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			requestMsgSet = sessionManager.CreateMsgSetRequest("UK", 3, 0)
		ElseIf (supportedVersion >= 2#) Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object sessionManager.CreateMsgSetRequest. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			requestMsgSet = sessionManager.CreateMsgSetRequest("UK", 2, 0)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object sessionManager.CreateMsgSetRequest. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			requestMsgSet = sessionManager.CreateMsgSetRequest("UK", 2, 0)
		End If
		Exit Sub
ErrHandler3: 
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
		End
	End Sub
	
	
	
	
	
	Private Sub SwitchQBFC()
		On Error GoTo ErrHandler4
		'This routine performs a simple switching of QBFC library
		'and instantiates a new QBSessionManager object to use forward
		'UPGRADE_WARNING: Couldn't resolve default property of object sessionManager.EndSession. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sessionManager.EndSession()
		'UPGRADE_WARNING: Couldn't resolve default property of object sessionManager.CloseConnection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sessionManager.CloseConnection()
		sessionManager = New QBFC15Lib.QBSessionManager
		'UPGRADE_WARNING: Couldn't resolve default property of object sessionManager.OpenConnection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sessionManager.OpenConnection(cAppID, cAppName)
		'UPGRADE_WARNING: Couldn't resolve default property of object sessionManager.BeginSession. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sessionManager.BeginSession("", QBFC15Lib.ENOpenMode.omDontCare)
		Exit Sub
ErrHandler4: 
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
		End
	End Sub
End Class