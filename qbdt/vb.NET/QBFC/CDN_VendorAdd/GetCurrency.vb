Option Strict Off
Option Explicit On
Module GetCurrency
	'-----------------------------------------------------------
	' Module: GetCurrency.bas
	'
	' Description:  This module retrieve the currency list from the Canadian
	'               version of QuickBooks.  It does check as well if the multicurrency
	'               has been set in the preferences
	'
	' Created On: 10/15/2002
	'
	' Copyright © 2002-2020 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/rdmgr/?ID=100
	'
	'----------------------------------------------------------
	
	Function GetCurrencyListCA() As String
		'This function returned the currency list in a string
		'Only the Canadian version of SDK has multicurrency feature.
		On Error GoTo ErrHandler
		
		'Create the message set request object
		Dim requestMsgSet As QBFC15Lib.IMsgSetRequest
		requestMsgSet = sessionManagerCA.CreateMsgSetRequest("CA", CShort(cQBXMLMajorVersion), CShort(cQBXMLMinorVersion))
		
		'Initialize the request's attributes
		requestMsgSet.Attributes.OnError = QBFC15Lib.ENRqOnError.roeContinue
		
		'Add the request to the message set request object
		requestMsgSet.AppendCurrencyQueryRq()
		'MsgBox requestMsgSet.ToXMLString, vbOKOnly, "RequestXML"
		
		Dim responseMsgSet As QBFC15Lib.IMsgSetResponse
		'Perform the request
		responseMsgSet = sessionManagerCA.DoRequests(requestMsgSet)
		'    MsgBox responseMsgSet.ToXMLString, vbOKOnly, "ResponseXML"
		Dim rsList As QBFC15Lib.IResponseList
		
		
		rsList = responseMsgSet.ResponseList
		Dim response As QBFC15Lib.IResponse
		
		'The response list contains only one response,
		'which corresponds to our single request
		response = rsList.GetAt(0)
		
		'Interpret the response
		GetCurrencyListCA = InterpretCurrencyQueryResponseCA(response)
		
		Exit Function
ErrHandler: 
		' In case something goes wrong along the way..
		MsgBox("Error sending request to QuickBooks: " & Err.Description)
	End Function
	
	Private Function InterpretCurrencyQueryResponseCA(ByRef response As QBFC15Lib.IResponse) As String
		'This function reads the details of the response and retrieve the list of currency names from it
		On Error GoTo ErrHandler
		
		'read the response details
		Dim statusSeverity, statusCode, statusMessage, msg As Object
		
		'UPGRADE_WARNING: Couldn't resolve default property of object statusCode. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		statusCode = response.statusCode
		'UPGRADE_WARNING: Couldn't resolve default property of object statusMessage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		statusMessage = response.statusMessage
		'UPGRADE_WARNING: Couldn't resolve default property of object statusSeverity. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		statusSeverity = response.statusSeverity
		
		Dim CurrencyRetList As QBFC15Lib.ICurrencyRetList
		CurrencyRetList = response.Detail
		
		Dim ndx, count As Object
		Dim CurrencyRet As QBFC15Lib.ICurrencyRet
		If Not (CurrencyRetList Is Nothing) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object count. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			count = CurrencyRetList.count
			'UPGRADE_WARNING: Couldn't resolve default property of object count. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For ndx = 0 To (count - 1)
				'UPGRADE_WARNING: Couldn't resolve default property of object ndx. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				CurrencyRet = CurrencyRetList.GetAt(ndx)
				
				InterpretCurrencyQueryResponseCA = InterpretCurrencyQueryResponseCA & CurrencyRet.Name.GetValue & "***"
			Next 
		End If
		
		Exit Function
		
ErrHandler: 
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
		InterpretCurrencyQueryResponseCA = "Error!"
		
	End Function
	Function FindIfMulticurrencyOn() As Boolean
		'This function verifies if the data file used has the multicurrency feature turned on
		On Error GoTo ErrHandler
		
		'Create the message set request object
		Dim requestMsgSet As QBFC15Lib.IMsgSetRequest
		requestMsgSet = sessionManagerCA.CreateMsgSetRequest("CA", CShort(cQBXMLMajorVersion), CShort(cQBXMLMinorVersion))
		
		'Add the request to the message set request object
		requestMsgSet.AppendPreferencesQueryRq()
		
		
		'MsgBox requestMsgSet.ToXMLString, vbOKOnly, "RequestXML"
		'Initialize the request's attributes
		requestMsgSet.Attributes.OnError = QBFC15Lib.ENRqOnError.roeContinue
		
		
		Dim responseMsgSet As QBFC15Lib.IMsgSetResponse
		'Perform the request
		responseMsgSet = sessionManagerCA.DoRequests(requestMsgSet)
		
		
		'MsgBox responseMsgSet.ToXMLString, vbOKOnly, "ResponseXML"
		
		Dim rsList As QBFC15Lib.IResponseList
		
		rsList = responseMsgSet.ResponseList
		
		Dim response As QBFC15Lib.IResponse
		
		
		
		'The response list contains only one response,
		'which corresponds to our single request
		response = rsList.GetAt(0)
		'Interpret the response
		
		FindIfMulticurrencyOn = InterpretAccountingPrefQueryResponseCA(response)
		
		
		
		
		Exit Function
ErrHandler: 
		' In case something goes wrong along the way..
		MsgBox("Error sending request to QuickBooks: " & Err.Description)
	End Function
	
	Function InterpretAccountingPrefQueryResponseCA(ByRef response As QBFC15Lib.IResponse) As Boolean
		'This function reads the details of the response and retrieve the list of currency names from it
		On Error GoTo ErrHandler
		
		'read the response details
		Dim statusSeverity, statusCode, statusMessage, msg As Object
		
		'UPGRADE_WARNING: Couldn't resolve default property of object statusCode. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		statusCode = response.statusCode
		'UPGRADE_WARNING: Couldn't resolve default property of object statusMessage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		statusMessage = response.statusMessage
		'UPGRADE_WARNING: Couldn't resolve default property of object statusSeverity. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		statusSeverity = response.statusSeverity
		
		Dim PreferenceRet As QBFC15Lib.IPreferencesRet
		
		PreferenceRet = response.Detail
		'Find if multicurrency is turned on (if multicurrency selected in preferences, it returns true)
		If Not (PreferenceRet Is Nothing) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object PreferenceRet.AccountingPreferences.IsUsingMulticurrency. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			InterpretAccountingPrefQueryResponseCA = PreferenceRet.AccountingPreferences.IsUsingMulticurrency.GetValue
		End If
		
		Exit Function
		
ErrHandler: 
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
		
	End Function
End Module