Attribute VB_Name = "ConnectionWithQuickBooks"
'-----------------------------------------------------------
' Module: ConnectionWithQuickBooks.bas
'
' Description:  This module demonstrates the use of QBFC,
'               by retrieving the supported SDK versions by
'               QuickBooks.  It will help us determining
'               If we deal with the Canadian or US version
'               of QuickBooks
'
' Created On: 10/15/2002
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/rdmgr/?ID=100
'
'----------------------------------------------------------
Option Explicit
'*******I make the session public, the connection will remain open until the app is closed
Public sessionManagerUS As New QBSessionManager 'This is the library to call when you
                                                        'want to make a connection using QBFC2.dll (for US QuickBooks)

Public sessionManagerCA As New QBSessionManager 'This is the library to call when you
                                                        'want to make a connection using QBFC2CA.dll (for Canadian QuickBooks)

Public strSDKVersion As String 'Public variable that tells us which dll I should use to communicate with QuickBooks
Const cAppID = "123"
Const cAppName = "IDN QuickBooks Sample Add Vendor"
Public blnConnectionIsOpen As Boolean
Public Function OpenConnection() As Boolean
'This function open the connection with QuickBooks and find which version of QuickBooks the app is
'talking to(US or Canada) in order to use the right dll when talking it.
On Error GoTo ErrHandler

    OpenConnection = True

    'Create the session manager object
    sessionManagerUS.OpenConnection cAppID, cAppName
    sessionManagerUS.BeginSession "", omDontCare


    Dim VersionSupportedArray() As String
    VersionSupportedArray = sessionManagerUS.QBXMLVersionsForSession ' This return an array of string containing all the version of qbXML
                                                            'supported by QuickBooks.  You could have used the sessionsManagerCA to do the same
    
    'If version CA2.0 is not returned in the list, it is QuickBooks US that is running
    Dim strCanadianVersion As String
    Dim blnCanadianVersionFound As Boolean
    Dim nArrayUpperBound As Integer
    
   
    strCanadianVersion = "CA2.0"
    blnCanadianVersionFound = False
    
    nArrayUpperBound = UBound(VersionSupportedArray) 'Finding how many elements are kept in the array
    
    Dim nIndex As Integer
    nIndex = 0
    While nIndex <= nArrayUpperBound 'Looping through the array to find if "CA2.0" version is present
        
        If strCanadianVersion = VersionSupportedArray(nIndex) Then
            blnCanadianVersionFound = True
        End If
        nIndex = nIndex + 1
    Wend
    
    If blnCanadianVersionFound = False Then
        strSDKVersion = "US"
    Else
        strSDKVersion = "Canada"
        'Close the US Session and setup the Canadian session
        sessionManagerUS.EndSession
        sessionManagerUS.CloseConnection
    
        sessionManagerCA.OpenConnection cAppID, cAppName
        sessionManagerCA.BeginSession "", omDontCare
    
    
    End If
    blnConnectionIsOpen = True
    Exit Function

ErrHandler:
    blnConnectionIsOpen = False
    OpenConnection = False
    'If QBFC2.dll or QBFC2CA.dll is missing, tell the user to reinstall them
    If Err.Number = 429 Then
        MsgBox "QBFC3.dll or QBFC2CA.dll is missing, please reinstall them by using the QBFC2 and QBFC2CA installers", , "Dll Missing"
    Else
        MsgBox Err.Description, vbExclamation, "Error"
    End If
End Function


Public Function CloseQBFCConnection()
'This function close the connection with QuickBooks
'It is making sure that I close the proper session (only one will be open everytime we use it
    If strSDKVersion = "Canada" Then
        sessionManagerCA.EndSession
        sessionManagerCA.CloseConnection
    
    Else
        sessionManagerUS.EndSession
        sessionManagerUS.CloseConnection
    End If

End Function




