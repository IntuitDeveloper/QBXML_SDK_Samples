Option Strict Off
Option Explicit On
Module CustomerAddModule
	'----------------------------------------------------------
	' Module: CustomerAddModule
	'
	' Description: This module contains the code which illustrates writing
	'               an applicaition which conditionally executes code
	'               depending on what qbxml version is supported by the
	'               installed copy of QuickBooks.
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
	'           BuildCustomerAddRequest
	'             This function will set fields of the CustomerAdd request.
	'
	'           BuildDataExtensionRequest
	'             This function will set fields of the DataExtDefAdd and
	'             DataExtAdd requests if qbxml version is 2.0.
	'
	'           SendCustomerAddRequest
	'             This procedure will call the functions to
	'             BuildCustomerAddRequest, BuildDataExtensionRequest, send
	'             the request to QB and then process the respons.
	'
	'           SendRequestToQB
	'             This procedure will send the request to QuickBooks using
	'             the DoRequests call.
	'
	'           ProcessResponse
	'             This procedure will process the response returned from QB.
	'             The status will be evaluated for each of the requests sent.
	'
	' Copyright © 2021-2022 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/rdmgr/?ID=100
	'
	' Updated to QBFC12 and fixed max version information 09/2012
	'----------------------------------------------------------
	
	Dim booSessionBegun As Boolean
	
	'Module objects
	
	Dim SessionManager As QBFC15Lib.QBSessionManager
	Dim msgSetRequest As QBFC15Lib.IMsgSetRequest
	
	Dim qbxmlVersion As String
	Dim bQBXMLVersion2_0 As Boolean
	Dim bQBXMLVersion1_1 As Boolean
	Dim bQBXMLVersion1_0 As Boolean
	Const cstrOwnerID As String = "{E09C86CF-9D6E-4EF2-BCBE-4D66B6B0F754}"
	
	
	Public Sub OpenConnectionBeginSession()
		
		booSessionBegun = False
		bQBXMLVersion2_0 = False
		bQBXMLVersion1_1 = False
		bQBXMLVersion1_0 = False
		
		On Error GoTo Errs
		
		SessionManager = New QBFC15Lib.QBSessionManager
		
		Dim connType As QBFC15Lib.ENConnectionType
		connType = QBFC15Lib.ENConnectionType.ctLocalQBDLaunchUI
		If (AddCust.UseQBOE.CheckState = 1) Then
			connType = QBFC15Lib.ENConnectionType.ctRemoteQBOE
		End If
		SessionManager.OpenConnection2("", "IDN VersionDependent Customer Add Sample - QBFC", connType)
		
		SessionManager.BeginSession("", QBFC15Lib.ENOpenMode.omDontCare)
		booSessionBegun = True
		
		qbxmlVersion = QBFCLatestVersion(SessionManager)
		msgSetRequest = GetLatestMsgSetRequest(SessionManager)
		
		Exit Sub
		
Errs: 
		If Err.Number = &H80040416 Then
			MsgBox("You must have QuickBooks running with the company" & vbCrLf & "file open to use this program.")
			SessionManager.CloseConnection()
			End
		ElseIf Err.Number = &H80040422 Then 
			MsgBox("This QuickBooks company file is open in single user mode and" & vbCrLf & "another application is already accessing it.  Please exit the" & vbCrLf & "other application and run this application again.")
			SessionManager.CloseConnection()
			End
		Else
			MsgBox("Error in OpenConnectionBeginSession" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)
			
			If booSessionBegun Then
				SessionManager.EndSession()
			End If
			
			SessionManager.CloseConnection()
			End
		End If
	End Sub
	
	Public Sub EndSessionCloseConnection()
		On Error GoTo Errs
		If booSessionBegun Then
			SessionManager.EndSession()
			SessionManager.CloseConnection()
		End If
		Exit Sub
Errs: 
		MsgBox("Error in EndSessionCloseConnection" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)
		
	End Sub
	
	Public Sub SendCustomerAddRequest()
		On Error GoTo Errs
		
		Dim msgSetResponse As QBFC15Lib.IMsgSetResponse
		
		msgSetRequest.ClearRequests()
		
		' create the CustomerAdd request
		If (Not BuildCustomerAddRequest()) Then
			Exit Sub
		End If
		
		' only add the data extensions if qbxml 2.0 request
		If (Val(qbxmlVersion) >= 2#) Then
			' create the Data Extension requests
			If (Not BuildDataExtensionRequest()) Then
				Exit Sub
			End If
		End If
		
		' send the requests to QB
		msgSetResponse = SendRequestToQB(msgSetRequest)
		
		' see if the request was successful
		If (Not ProcessResponse(msgSetResponse)) Then
			Exit Sub
		End If
		
		MsgBox("Success Adding Customer Information.")
		
		Exit Sub
Errs: 
		MsgBox("Error in SendCustomerAddRequest" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)
		
	End Sub
	
	Public Function BuildCustomerAddRequest() As Boolean
		On Error GoTo Errs
		BuildCustomerAddRequest = False
		
		If (msgSetRequest Is Nothing) Then
			Exit Function
		End If
		
		' Add the CustomerAdd request
		Dim customerAdd As QBFC15Lib.ICustomerAdd
		customerAdd = msgSetRequest.AppendCustomerAddRq()
		
		' set the customer name
		If (AddCust.custName <> "") Then
			customerAdd.Name.setValue(AddCust.custName)
		End If
		
		' set the customer first name
		If (AddCust.firstName <> "") Then
			customerAdd.firstName.setValue(AddCust.firstName)
		End If
		
		' set the customer last name
		If (AddCust.lastName <> "") Then
			customerAdd.lastName.setValue(AddCust.lastName)
		End If
		
		' set the customer phone number
		If (AddCust.Text_Phone.Text <> "") Then
			customerAdd.Phone.setValue(AddCust.phoneNumber)
		End If
		
		' set the customer bill address
		' Addr1
		If (AddCust.Addr1 <> "") Then
			customerAdd.BillAddress.Addr1.setValue(AddCust.Addr1)
		End If
		' Addr2
		If (AddCust.Addr2 <> "") Then
			customerAdd.BillAddress.Addr2.setValue(AddCust.Addr2)
		End If
		' Addr3
		If (AddCust.Addr3 <> "") Then
			customerAdd.BillAddress.Addr3.setValue(AddCust.Addr3)
		End If
		' only add the Addr4 if we are using qbxml v2.0 request
		' Addr4
		If (Val(qbxmlVersion) >= 2#) Then
			If (AddCust.Addr4 <> "") Then
				customerAdd.BillAddress.Addr4.setValue(AddCust.Addr4)
			End If
		End If
		' City
		If (AddCust.City <> "") Then
			customerAdd.BillAddress.City.setValue(AddCust.City)
		End If
		' State
		If (AddCust.State <> "") Then
			customerAdd.BillAddress.State.setValue(AddCust.State)
		End If
		' PostalCode
		If (AddCust.PostalCode <> "") Then
			customerAdd.BillAddress.PostalCode.setValue(AddCust.PostalCode)
		End If
		' Country
		If (AddCust.Country <> "") Then
			customerAdd.BillAddress.Country.setValue(AddCust.Country)
		End If
		
		BuildCustomerAddRequest = True
		
		Exit Function
Errs: 
		MsgBox("Error in BuildCustomerAddRequest" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)
		
		BuildCustomerAddRequest = False
	End Function
	
	Public Function BuildDataExtensionRequest() As Boolean
		
		On Error GoTo Errs
		
		BuildDataExtensionRequest = False
		
		If (msgSetRequest Is Nothing) Then
			Exit Function
		End If
		
		' do not add data extension if user did not specify one
		If (AddCust.strDataExtName = "") Or (AddCust.strDataExtValue = "") Then
			BuildDataExtensionRequest = True
			Exit Function
		End If
		
		' Add the DataExtensionDefAdd request
		Dim dataExtDefAdd As QBFC15Lib.IDataExtDefAdd
		dataExtDefAdd = msgSetRequest.AppendDataExtDefAddRq()
		
		' Set the values in the request
		'Add the OwnerID element
		dataExtDefAdd.OwnerID.setValue((cstrOwnerID))
		
		'Add the DataExtName element
		dataExtDefAdd.DataExtName.setValue(AddCust.strDataExtName)
		
		'Add the DataExtType element
		dataExtDefAdd.DataExtType.setValue(QBFC15Lib.ENDataExtType.detSTR255TYPE)
		
		
		'Add the AssignToObject element
		dataExtDefAdd.AssignToObjectList.Add(QBFC15Lib.ENAssignToObject.atoCustomer)
		
		' Add the DataExtAddRq request
		Dim dataExtAdd As QBFC15Lib.IDataExtAdd
		dataExtAdd = msgSetRequest.AppendDataExtAddRq()
		
		'Add the OwnerID element
		dataExtAdd.OwnerID.setValue((cstrOwnerID))
		
		'Add the DataExtName element
		dataExtAdd.DataExtName.setValue(AddCust.strDataExtName)
		
		'Add the ListDataExtType element
		dataExtAdd.ORListTxnWithMacro.ListDataExt.ListDataExtType.setValue(QBFC15Lib.ENListDataExtType.ldetCustomer)
		
		'Add the FullName element
		dataExtAdd.ORListTxnWithMacro.ListDataExt.ListObjRef.FullName.setValue(AddCust.custName)
		
		'Add the DataExtValue element
		dataExtAdd.DataExtValue.setValue(AddCust.strDataExtValue)
		
		BuildDataExtensionRequest = True
		
		Exit Function
Errs: 
		MsgBox("Error in BuildDataExtensionRequest" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)
		
		BuildDataExtensionRequest = False
	End Function
	
	Private Function SendRequestToQB(ByRef request As QBFC15Lib.IMsgSetRequest) As QBFC15Lib.IMsgSetResponse
		On Error GoTo Errs
		
		If (request Is Nothing) Then
			Exit Function
		End If
		' set the OnError attribute to continueOnError
		' going to select continueOnError just in case the Customer already
		' had a Data Extension defined with that name.  If the Data Extension
		' already exists, we still want to contine to add a value for that
		' data extension for this customer.
		request.Attributes.OnError = QBFC15Lib.ENRqOnError.roeContinue
		
		AddCust.requestXML = request.ToXMLString()
		
		SendRequestToQB = SessionManager.DoRequests(request)
		
		AddCust.responseXML = SendRequestToQB.ToXMLString()
		
		Exit Function
Errs: 
		MsgBox("Error in SendRequestToQB" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)
	End Function
	
	Private Function ProcessResponse(ByRef msgSetResponse As QBFC15Lib.IMsgSetResponse) As Boolean
		Dim ProcessRequest As Object
		On Error GoTo Errs
		
		' check to make sure we have objects to access first
		' and that there are responses in the list
		If (msgSetResponse Is Nothing) Then
			
			ProcessRequest = False
			Exit Function
		End If
		If (msgSetResponse.responseList Is Nothing) Then
			
			ProcessRequest = False
			Exit Function
		End If
		If (msgSetResponse.responseList.count <= 0) Then
			
			ProcessRequest = False
			Exit Function
		End If
		
		' Start parsing the response list
		Dim responseList As QBFC15Lib.IResponseList
		responseList = msgSetResponse.responseList
		
		' go thru each response and process the response.
		Dim count As Short
		Dim index As Short
		Dim response As QBFC15Lib.IResponse
		count = responseList.count
		Dim responseType As QBFC15Lib.ENResponseType
		Dim responseDetailType As QBFC15Lib.ENObjectType
		For index = 0 To count - 1
			response = responseList.GetAt(index)
			
			' first make sure we have a response object to handle
			If (response Is Nothing) Then
				
				ProcessRequest = False
				Exit Function
			End If
			If (response.Type Is Nothing) Then
				
				ProcessRequest = False
				Exit Function
			End If
			
			' figure out which response we are processing
			' CustomerAddRs, DataExtDefAddRs, or DataExtAddRs
			responseType = response.Type.getValue()
			If (responseType = QBFC15Lib.ENResponseType.rtCustomerAddRs) Then
				'report status to the user
				If response.StatusCode = CDbl("0") Then
					MsgBox("Successfully added Customer.")
				Else
					MsgBox("CustomerAdd Error - " & vbCrLf & "StatusCode = " & response.StatusCode & vbCrLf & vbCrLf & response.StatusMessage)
				End If
			ElseIf (responseType = QBFC15Lib.ENResponseType.rtDataExtDefAddRs) Then 
				'report status to the user
				If response.StatusCode = CDbl("0") Then
					MsgBox("Successfully added Data Extension Definition.")
				Else
					MsgBox("DataExtDefAdd Error - " & vbCrLf & "StatusCode = " & response.StatusCode & vbCrLf & vbCrLf & response.StatusMessage)
				End If
			ElseIf (responseType = QBFC15Lib.ENResponseType.rtDataExtAddRs) Then 
				'report status to the user
				If response.StatusCode = CDbl("0") Then
					MsgBox("Successfully added Data Extension value for this customer.")
				Else
					MsgBox("DataExtAdd Error - " & vbCrLf & "StatusCode = " & response.StatusCode & vbCrLf & vbCrLf & response.StatusMessage)
				End If
			Else
				' bail, we do not have the responses we were expecting
				
				ProcessRequest = False
				Exit Function
			End If
			
		Next 
		
		Exit Function
Errs: 
		MsgBox("Error in ProcessResponse" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)
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
		ElseIf (supportedVersion >= 2#) Then 
			GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 2, 0)
		ElseIf (supportedVersion = 1.1) Then 
			GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 1, 1)
		Else
			MsgBox("You are apparently running QuickBooks 2002 Release 1, we strongly recommend that you use QuickBooks' online update feature to obtain the latest fixes and enhancements", MsgBoxStyle.Exclamation)
			GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 1, 0)
		End If
	End Function
	
	Function QBFCLatestVersion(ByRef SessionManager As QBFC15Lib.QBSessionManager) As String
		Dim strXMLVersions() As String
		strXMLVersions = VB6.CopyArray(SessionManager.QBXMLVersionsForSession)
		
		
		Dim i As Integer
		Dim vers As String
		Dim LastVers As String
		For i = LBound(strXMLVersions) To UBound(strXMLVersions)
			vers = strXMLVersions(i)
			If (vers > LastVers) Then
				LastVers = vers
				QBFCLatestVersion = vers
			End If
		Next i
	End Function
End Module