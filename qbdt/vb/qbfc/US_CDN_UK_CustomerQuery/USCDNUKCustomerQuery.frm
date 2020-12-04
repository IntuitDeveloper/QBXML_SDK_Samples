VERSION 5.00
Begin VB.Form Form_USCDNUKCustomerQuery 
   Caption         =   "US_CDN_UK_CustomerQuery"
   ClientHeight    =   2085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "USCDNCustomerQuery"
   ScaleHeight     =   2085
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1_Submit 
      Caption         =   "Run"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "US_CDN_UK Customer Query"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   $"USCDNUKCustomerQuery.frx":0000
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   4215
   End
End
Attribute VB_Name = "Form_USCDNUKCustomerQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
' 08/09/2013 updated to qbfc13
'----------------------------------------------------------

Option Explicit
Const cAppID = "123"
Const cAppName = "IDN Sample - US_CDN_UK_CustomerQuery"
Const cRequestID = "1"

Dim sessionManager      As Object 'QBFCx.QBSessionManager
Dim requestMsgSet       As Object 'IMsgSetRequest
Dim QueryResponse       As Object 'IMsgSetResponse
Dim HostResponse        As Object 'IHostRet
Dim response            As Object 'IResponse
Dim supportedVersions   As Object 'IBSTRList
Dim msgset              As Object 'IMsgSetRequest


Dim requestXML          As String
Dim responseXML         As String
Dim QBXMLLatestVersion  As String
Dim strXMLVersions()    As String

Dim i           As Long
Dim vers        As Double
Dim LastVers    As Double
    
    
    
    
Private Sub Command1_Submit_Click()
On Error GoTo ErrHandler
    'Start with a US HostQuery
    'If it succeeds, app is talking to QB_US
    'If it fails with a specific parse error, catch it and then retry a CDN HostQuery
    HostQuery_US
    
    requestMsgSet.Attributes.OnError = roeContinue
    If (requestMsgSet Is Nothing) Then
        Exit Sub
    End If
  
    'Add the request to the message set request object.
    Dim customerQuery As Object 'ICustomerQuery
    Set customerQuery = requestMsgSet.AppendCustomerQueryRq
    
    'Set the elements of ICustomerQuery.
    customerQuery.ORCustomerListQuery.CustomerListFilter.MaxReturned.SetValue 1
            
    ' Uncomment the following to see the response XML for debugging
    MsgBox requestMsgSet.ToXMLString, vbOKOnly, "RequestXML"
    
    ' Perform the request
    Dim responseMsgSet As Object 'IMsgSetResponse
    Set responseMsgSet = sessionManager.DoRequests(requestMsgSet)
    
    ' Uncomment the following to see the response XML for debugging
    MsgBox responseMsgSet.ToXMLString, vbOKOnly, "ResponseXML"
    
    ' Interpret the response
    Dim response As Object 'IResponse
    Dim statusCode, statusMessage, statusSeverity
    
    ' The response list contains only one response,
    ' which corresponds to our single request
    Set response = responseMsgSet.ResponseList.GetAt(0)
    statusCode = response.statusCode
    statusMessage = response.statusMessage
    statusSeverity = response.statusSeverity
      
    MsgBox "Status: Code = " & CStr(statusCode) & _
          ", Message = " & statusMessage & _
          ", Severity = " & statusSeverity & vbCrLf
          
    sessionManager.EndSession
    sessionManager.CloseConnection
    Exit Sub

ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
    Exit Sub
End Sub
    
    
    
    



Public Sub HostQuery_US()
On Error GoTo ErrHandler1
    'Start with a US HostQuery
    'If it succeeds, app is talking to QB_US
    'If it fails with a specific parse error, catch it and then retry a CDN HostQuery
    Set sessionManager = New QBFC13Lib.QBSessionManager
    sessionManager.OpenConnection cAppID, cAppName
    sessionManager.BeginSession "", omDontCare
    Set msgset = sessionManager.CreateMsgSetRequest("US", 1, 0)
    msgset.AppendHostQueryRq
    Set QueryResponse = sessionManager.DoRequests(msgset)
    Set response = QueryResponse.ResponseList.GetAt(0)
    Set HostResponse = response.Detail
    Set supportedVersions = HostResponse.SupportedQBXMLVersionList
    LastVers = 0
    For i = 0 To supportedVersions.Count - 1
        vers = Val(supportedVersions.GetAt(i))
        If (vers > LastVers) Then
            LastVers = vers
            QBXMLLatestVersion = supportedVersions.GetAt(i)
        End If
    Next i
    Dim supportedVersion As Double
    supportedVersion = Val(QBXMLLatestVersion)
    
    If (supportedVersion >= 6#) Then
        SwitchQBFC
        Set requestMsgSet = sessionManager.CreateMsgSetRequest("", 6, 0)
        ' Note: The following is also correct as of 6.0
        ' Set requestMsgSet = sessionManager.CreateMsgSetRequest("US", 6, 0)
    ElseIf (supportedVersion >= 5#) Then
        Set requestMsgSet = sessionManager.CreateMsgSetRequest("US", 5, 0)
    ElseIf (supportedVersion >= 4#) Then
        Set requestMsgSet = sessionManager.CreateMsgSetRequest("US", 4, 0)
    ElseIf (supportedVersion >= 3#) Then
        Set requestMsgSet = sessionManager.CreateMsgSetRequest("US", 3, 0)
    ElseIf (supportedVersion >= 2#) Then
        Set requestMsgSet = sessionManager.CreateMsgSetRequest("US", 2, 0)
    ElseIf (supportedVersion = 1.1) Then
        Set requestMsgSet = sessionManager.CreateMsgSetRequest("US", 1, 1)
    Else
        MsgBox "You are apparently running QuickBooks 2002 Release 1, we strongly recommend that you use QuickBooks' online update feature to obtain the latest fixes and enhancements", vbExclamation
        Set requestMsgSet = sessionManager.CreateMsgSetRequest("US", 1, 0)
    End If
    Exit Sub
    
ErrHandler1:
    If Err.Description = "The version of QBXML that was requested is not supported or is unknown." Then
        'If it fails with a specific parse error, catch it and then retry a CDN HostQuery
        HostQuery_CA
    Else
        MsgBox Err.Description, vbExclamation, "Error"
        End
    End If
End Sub




Public Sub HostQuery_CA()
On Error GoTo ErrHandler2
        'Execute a CDN HostQuery
        'If it succeeds, app is talking to QB_CDN
        'If it fails with a specific parse error, catch it and then retry a UK HostQuery
        Set msgset = sessionManager.CreateMsgSetRequest("CA", 2, 0)
        msgset.AppendHostQueryRq
        Set QueryResponse = sessionManager.DoRequests(msgset)
        Set response = QueryResponse.ResponseList.GetAt(0)
        Set HostResponse = response.Detail
        Set supportedVersions = HostResponse.SupportedQBXMLVersionList
        LastVers = 0
        Dim ca_vers As String
        For i = 0 To supportedVersions.Count - 1
            ca_vers = supportedVersions.GetAt(i)
            ca_vers = Mid(ca_vers, 3)
            vers = Val(ca_vers)
            If (vers > LastVers) Then
                LastVers = vers
                QBXMLLatestVersion = supportedVersions.GetAt(i)
            End If
        Next i
        If (Left$(QBXMLLatestVersion, 2) = "CA") Then
            QBXMLLatestVersion = Mid$(QBXMLLatestVersion, 3, 3)
        End If
        Dim supportedVersion As Double
        supportedVersion = Val(QBXMLLatestVersion)
        If (supportedVersion >= 6#) Then
            Set requestMsgSet = sessionManager.CreateMsgSetRequest("", 6, 0)
            ' Note: The following is also correct as of 6.0
            ' Set requestMsgSet = sessionManager.CreateMsgSetRequest("CA", 6, 0)
        ElseIf (supportedVersion >= 3#) Then
            Set requestMsgSet = sessionManager.CreateMsgSetRequest("CA", 3, 0)
        ElseIf (supportedVersion >= 2#) Then
            Set requestMsgSet = sessionManager.CreateMsgSetRequest("CA", 2, 0)
        Else
            Set requestMsgSet = sessionManager.CreateMsgSetRequest("CA", 2, 0)
        End If
        Exit Sub
ErrHandler2:
    If Err.Description = "The version of QBXML that was requested is not supported or is unknown." Then
        'If it fails with a specific parse error, catch it and then retry a UK HostQuery
        HostQuery_UK
    Else
        MsgBox Err.Description, vbExclamation, "Error"
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
        Set msgset = sessionManager.CreateMsgSetRequest("UK", 2, 0)
        msgset.AppendHostQueryRq
        Set QueryResponse = sessionManager.DoRequests(msgset)
        Set response = QueryResponse.ResponseList.GetAt(0)
        Set HostResponse = response.Detail
        Set supportedVersions = HostResponse.SupportedQBXMLVersionList
        LastVers = 0
        Dim uk_vers As String
        For i = 0 To supportedVersions.Count - 1
            uk_vers = supportedVersions.GetAt(i)
            uk_vers = Mid(uk_vers, 3)
            vers = Val(uk_vers)
            If (vers > LastVers) Then
                LastVers = vers
                QBXMLLatestVersion = supportedVersions.GetAt(i)
            End If
        Next i
        If (Left$(QBXMLLatestVersion, 2) = "UK") Then
            QBXMLLatestVersion = Mid$(QBXMLLatestVersion, 3, 3)
        End If
        Dim supportedVersion As Double
        supportedVersion = Val(QBXMLLatestVersion)
        If (supportedVersion >= 6#) Then
            Set requestMsgSet = sessionManager.CreateMsgSetRequest("", 6, 0)
            ' Note: The following is also correct as of 6.0
            ' Set requestMsgSet = sessionManager.CreateMsgSetRequest("UK", 6, 0)
        ElseIf (supportedVersion >= 3#) Then
            Set requestMsgSet = sessionManager.CreateMsgSetRequest("UK", 3, 0)
        ElseIf (supportedVersion >= 2#) Then
            Set requestMsgSet = sessionManager.CreateMsgSetRequest("UK", 2, 0)
        Else
            Set requestMsgSet = sessionManager.CreateMsgSetRequest("UK", 2, 0)
        End If
        Exit Sub
ErrHandler3:
        MsgBox Err.Description, vbExclamation, "Error"
        End
End Sub





Private Sub SwitchQBFC()
On Error GoTo ErrHandler4
        'This routine performs a simple switching of QBFC library
        'and instantiates a new QBSessionManager object to use forward
        sessionManager.EndSession
        sessionManager.CloseConnection
        Set sessionManager = New QBFC13Lib.QBSessionManager
        sessionManager.OpenConnection cAppID, cAppName
        sessionManager.BeginSession "", omDontCare
        Exit Sub
ErrHandler4:
        MsgBox Err.Description, vbExclamation, "Error"
        End
End Sub





