Option Strict Off
Option Explicit On
Module GetVendorType
	'-----------------------------------------------------------
	' Module: GetVendorType.bas
	'
	' Description:  This module retrieve the vendor type list from the Canadian
	'               version of QuickBooks or the US version of QuickBooks depending
	'               on which function is called.
	'
	' Created On: 10/15/2002
	'
	' Copyright © 2002-2020 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/rdmgr/?ID=100
	'
	'----------------------------------------------------------
	
	
	Public Const cQBXMLMajorVersion As String = "2"
	Public Const cQBXMLMinorVersion As String = "0"
	
	Function GetVendorTypeListCA() As String
		'This function sends a query to QuickBooks in order to get the vendor type list
		On Error GoTo ErrHandler
		'Create the message set request object
		Dim requestMsgSet As QBFC15Lib.IMsgSetRequest
		requestMsgSet = sessionManagerCA.CreateMsgSetRequest("CA", CShort(cQBXMLMajorVersion), CShort(cQBXMLMinorVersion))
		
		'Add the request to the message
		requestMsgSet.AppendVendorTypeQueryRq()
		
		'    MsgBox requestMsgSet.ToXMLString, vbOKOnly, "RequestXML"
		'Initialize the request's attributes
		requestMsgSet.Attributes.OnError = QBFC15Lib.ENRqOnError.roeContinue
		
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
		GetVendorTypeListCA = InterpretVendorTypeQueryResponseCA(response)
		
		
		
		
		Exit Function
ErrHandler: 
		' In case something goes wrong along the way..
		MsgBox("Error sending request to QuickBooks: " & Err.Description)
	End Function
	
	Private Function InterpretVendorTypeQueryResponseCA(ByRef response As QBFC15Lib.IResponse) As String
		'This function reads the details of the response and retrieve the list of vendor type from it
		On Error GoTo ErrHandler
		
		Dim statusSeverity, statusCode, statusMessage, msg As Object
		'read the response details
		'UPGRADE_WARNING: Couldn't resolve default property of object statusCode. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		statusCode = response.statusCode
		'UPGRADE_WARNING: Couldn't resolve default property of object statusMessage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		statusMessage = response.statusMessage
		'UPGRADE_WARNING: Couldn't resolve default property of object statusSeverity. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		statusSeverity = response.statusSeverity
		
		Dim VendorTypeRetList As QBFC15Lib.IVendorTypeRetList
		VendorTypeRetList = response.Detail
		
		Dim ndx, count As Object
		Dim VendorTypeRet As QBFC15Lib.IVendorTypeRet
		If Not (VendorTypeRetList Is Nothing) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object count. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			count = VendorTypeRetList.count
			'UPGRADE_WARNING: Couldn't resolve default property of object count. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For ndx = 0 To (count - 1)
				'UPGRADE_WARNING: Couldn't resolve default property of object ndx. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				VendorTypeRet = VendorTypeRetList.GetAt(ndx)
				
				InterpretVendorTypeQueryResponseCA = InterpretVendorTypeQueryResponseCA & VendorTypeRet.FullName.GetValue & "***"
			Next 
		End If
		
		Exit Function
		
ErrHandler: 
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
		InterpretVendorTypeQueryResponseCA = "Error!"
		
	End Function
	
	Function GetVendorTypeListUS() As String
		'This function sends a query to QuickBooks in order to get the vendor type list
		On Error GoTo ErrHandler
		'Create the message set request object
		Dim requestUSMsgSet As QBFC15Lib.IMsgSetRequest
		requestUSMsgSet = sessionManagerUS.CreateMsgSetRequest("US", CShort(cQBXMLMajorVersion), CShort(cQBXMLMinorVersion))
		
		'Add the request to the message
		requestUSMsgSet.AppendVendorTypeQueryRq()
		
		'MsgBox requestUSMsgSet.ToXMLString, vbOKOnly, "RequestXML"
		'Initialize the request's attributes
		requestUSMsgSet.Attributes.OnError = QBFC15Lib.ENRqOnError.roeContinue
		
		
		'Perform the request
		Dim responseMsgSet As QBFC15Lib.IMsgSetResponse
		responseMsgSet = sessionManagerUS.DoRequests(requestUSMsgSet)
		
		
		
		'    MsgBox responseMsgSet.ToXMLString, vbOKOnly, "ResponseXML"
		
		Dim rsList As QBFC15Lib.IResponseList
		
		rsList = responseMsgSet.ResponseList
		
		Dim response As QBFC15Lib.IResponse
		
		'The response list contains only one response,
		'which corresponds to our single request
		response = rsList.GetAt(0)
		
		'Interpret the response
		GetVendorTypeListUS = InterpretVendorTypeQueryResponseUS(response)
		
		Exit Function
ErrHandler: 
		' In case something goes wrong along the way..
		MsgBox("Error sending request to QuickBooks: " & Err.Description)
	End Function
	
	Private Function InterpretVendorTypeQueryResponseUS(ByRef response As QBFC15Lib.IResponse) As String
		'This function reads the details of the response and retrieve the list of vendor type from it
		On Error GoTo ErrHandler
		
		Dim statusSeverity, statusCode, statusMessage, msg As Object
		'read the response details
		'UPGRADE_WARNING: Couldn't resolve default property of object statusCode. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		statusCode = response.statusCode
		'UPGRADE_WARNING: Couldn't resolve default property of object statusMessage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		statusMessage = response.statusMessage
		'UPGRADE_WARNING: Couldn't resolve default property of object statusSeverity. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		statusSeverity = response.statusSeverity
		
		Dim VendorTypeRetList As QBFC15Lib.IVendorTypeRetList
		VendorTypeRetList = response.Detail
		
		Dim ndx, count As Object
		Dim VendorTypeRet As QBFC15Lib.IVendorTypeRet
		If Not (VendorTypeRetList Is Nothing) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object count. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			count = VendorTypeRetList.count
			'UPGRADE_WARNING: Couldn't resolve default property of object count. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For ndx = 0 To (count - 1)
				'UPGRADE_WARNING: Couldn't resolve default property of object ndx. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				VendorTypeRet = VendorTypeRetList.GetAt(ndx)
				
				InterpretVendorTypeQueryResponseUS = InterpretVendorTypeQueryResponseUS & VendorTypeRet.FullName.GetValue & "***"
			Next 
		End If
		
		Exit Function
		
ErrHandler: 
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
		InterpretVendorTypeQueryResponseUS = "Error!"
		
	End Function
End Module