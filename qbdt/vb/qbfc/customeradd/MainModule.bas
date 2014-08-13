Attribute VB_Name = "MainModule"
'-----------------------------------------------------------
' Module: MainModule
'
' Description:  This module demonstrates the simple use of QBFC,
'               by adding a new customer to QuickBooks.
'
' Created On: 1/22/2002
' Updated to QBFC 2.0: 08/2002
' Updated to QBFC 5.0: 09/2005
' Updated to QBFC 12.0 and fixed setting max QBXML version: 09/2012
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'----------------------------------------------------------
Option Explicit

Const cAppID = "Desktop VB QBFC AddCustomer"
Const cAppName = "IDN Desktop VB QBFC AddCustomer"
Const cCustomerName = "Jim"


Public Sub Main()

On Error GoTo ErrHandler

    Dim msg As String
    Dim okCancel As VbMsgBoxResult
    
    msg = "This application adds a new Customer, named " & cCustomerName & _
          ", to QuickBooks." & vbCr & vbCr & _
          "QuickBooks must be running with a data file open." & vbCr

    okCancel = MsgBox(msg, vbOKCancel)
    If okCancel = vbCancel Then
        Exit Sub
    End If

    ' Create the session manager object and use it to open a
    ' connection and begin a session with QuickBooks
    Dim SessionManager As New QBSessionManager
    SessionManager.OpenConnection cAppID, cAppName
    SessionManager.BeginSession "", omDontCare
    
    ' Create the message set request object.  For QBFC 2.0, the
    ' major version is 2, the minor 0.
    Dim requestMsgSet As IMsgSetRequest
    Set requestMsgSet = GetLatestMsgSetRequest(SessionManager)
    
    ' Initialize the request's attributes to continue on error
    requestMsgSet.Attributes.OnError = roeContinue
    
    ' Add the request to the message set request object
    Dim CustomerAdd As ICustomerAdd
    Set CustomerAdd = requestMsgSet.AppendCustomerAddRq
    CustomerAdd.Name.SetValue cCustomerName
    
    MsgBox requestMsgSet.ToXMLString
    
    ' Perform the request and receive a response from QuickBooks
    Dim responseMsgSet As IMsgSetResponse
    Set responseMsgSet = SessionManager.DoRequests(requestMsgSet)
    
    MsgBox responseMsgSet.ToXMLString
    
    ' Interpret the response
    Dim response As IResponse
   
    ' The response list contains only one response,
    ' which corresponds to our single request
    Set response = responseMsgSet.ResponseList.GetAt(0)
    
    If (response.StatusCode <> 0) Then
        MsgBox "Status: Code = " & CStr(response.StatusCode) & _
          ", Severity = " & response.StatusSeverity & _
          ", Message = " & response.StatusMessage
    End If
        
    ' The response detail for Add and Mod requests is a 'Ret' object
    ' In our case, it's ICustomerRet
    Dim customerRet As ICustomerRet
    Set customerRet = response.Detail
    
    ' Make sure a customerRet was returned before trying to obtain
    ' the ListID of the customer.
    If (Not (customerRet Is Nothing)) Then
        MsgBox "CustomerRet ListID = " & customerRet.ListID.GetValue
    End If
   
    ' Finally close the session and connection with QuickBooks
    SessionManager.EndSession
    SessionManager.CloseConnection
        
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"

End Sub

Public Function GetLatestMsgSetRequest(SessionManager As QBSessionManager) As IMsgSetRequest
	Dim supportedVersion As Double
    supportedVersion = Val(QBFCLatestVersion(SessionManager))
    If (supportedVersion >= 6#) Then
        Set GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 6, 0)
    ElseIf (supportedVersion >= 5#) Then
        Set GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 5, 0)
    ElseIf (supportedVersion >= 4#) Then
        Set GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 4, 0)
    ElseIf (supportedVersion >= 3#) Then
        Set GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 3, 0)
    ElseIf (supportedVersion = 2#) Then
        Set GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 2, 0)
    ElseIf (supportedVersion = 1.1) Then
        Set GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 1, 1)
    Else
        MsgBox "You are apparently running QuickBooks 2002 Release 1, we strongly recommend that you use QuickBooks' online update feature to obtain the latest fixes and enhancements", vbExclamation
        Set GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 1, 0)
    End If
End Function

Function QBFCLatestVersion(SessionManager As QBSessionManager) As String
    Dim strXMLVersions() As String
    'Should be able to use this, but there appears to be a bug that may cause 2.0 to be returned
    'when it should not.
    'strXMLVersions = SessionManager.QBXMLVersionsForSession
    
    Dim msgset As IMsgSetRequest
    'Use oldest version to ensure that we work with any QuickBooks (US)
    Set msgset = SessionManager.CreateMsgSetRequest("US", 1, 0)
    msgset.AppendHostQueryRq
    Dim QueryResponse As IMsgSetResponse
    Set QueryResponse = SessionManager.DoRequests(msgset)
    Dim response As IResponse
    
    ' The response list contains only one response,
    ' which corresponds to our single HostQuery request
    Set response = QueryResponse.ResponseList.GetAt(0)
    Dim HostResponse As IHostRet
    Set HostResponse = response.Detail
    Dim supportedVersions As IBSTRList
    Set supportedVersions = HostResponse.SupportedQBXMLVersionList
    
    Dim i As Long
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

