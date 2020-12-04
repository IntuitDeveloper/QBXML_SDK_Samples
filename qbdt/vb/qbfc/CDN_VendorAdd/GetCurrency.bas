Attribute VB_Name = "GetCurrency"
'-----------------------------------------------------------
' Module: GetCurrency.bas
'
' Description:  This module retrieve the currency list from the Canadian
'               version of QuickBooks.  It does check as well if the multicurrency
'               has been set in the preferences
'
' Created On: 10/15/2002
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/rdmgr/?ID=100
'
'----------------------------------------------------------

Function GetCurrencyListCA() As String
'This function returned the currency list in a string
'Only the Canadian version of SDK has multicurrency feature.
On Error GoTo ErrHandler
    
    'Create the message set request object
    Dim requestMsgSet As IMsgSetRequest
    Set requestMsgSet = sessionManagerCA.CreateMsgSetRequest("CA", cQBXMLMajorVersion, cQBXMLMinorVersion)

    'Initialize the request's attributes
    requestMsgSet.Attributes.OnError = roeContinue

    'Add the request to the message set request object
     requestMsgSet.AppendCurrencyQueryRq
    'MsgBox requestMsgSet.ToXMLString, vbOKOnly, "RequestXML"
    
    Dim responseMsgSet As IMsgSetResponse
    'Perform the request
    Set responseMsgSet = sessionManagerCA.DoRequests(requestMsgSet)
'    MsgBox responseMsgSet.ToXMLString, vbOKOnly, "ResponseXML"
    Dim rsList As IResponseList
     

    Set rsList = responseMsgSet.ResponseList
    Dim response As IResponse
     
    'The response list contains only one response,
    'which corresponds to our single request
    Set response = rsList.GetAt(0)
    
    'Interpret the response
    GetCurrencyListCA = InterpretCurrencyQueryResponseCA(response)
    
    Exit Function
ErrHandler:
    ' In case something goes wrong along the way..
    MsgBox "Error sending request to QuickBooks: " & Err.Description
End Function

Private Function InterpretCurrencyQueryResponseCA(response As IResponse) As String
'This function reads the details of the response and retrieve the list of currency names from it
On Error GoTo ErrHandler

    'read the response details
    Dim statusCode, statusMessage, statusSeverity, msg
    
    statusCode = response.statusCode
    statusMessage = response.statusMessage
    statusSeverity = response.statusSeverity
      
    Dim CurrencyRetList As ICurrencyRetList
    Set CurrencyRetList = response.Detail
    
    If Not (CurrencyRetList Is Nothing) Then
        Dim ndx, count
        count = CurrencyRetList.count
        For ndx = 0 To (count - 1)
            Dim CurrencyRet As ICurrencyRet
            Set CurrencyRet = CurrencyRetList.GetAt(ndx)
            
            InterpretCurrencyQueryResponseCA = InterpretCurrencyQueryResponseCA & CurrencyRet.Name.GetValue & "***"
        Next
    End If
   
   Exit Function
    
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
    InterpretCurrencyQueryResponseCA = "Error!"
    
End Function
Function FindIfMulticurrencyOn() As Boolean
'This function verifies if the data file used has the multicurrency feature turned on
On Error GoTo ErrHandler
    
    'Create the message set request object
    Dim requestMsgSet As IMsgSetRequest
    Set requestMsgSet = sessionManagerCA.CreateMsgSetRequest("CA", cQBXMLMajorVersion, cQBXMLMinorVersion)

    'Add the request to the message set request object
    requestMsgSet.AppendPreferencesQueryRq
    
    
    'MsgBox requestMsgSet.ToXMLString, vbOKOnly, "RequestXML"
    'Initialize the request's attributes
    requestMsgSet.Attributes.OnError = roeContinue
    
    
    Dim responseMsgSet As IMsgSetResponse
    'Perform the request
    Set responseMsgSet = sessionManagerCA.DoRequests(requestMsgSet)
    

    'MsgBox responseMsgSet.ToXMLString, vbOKOnly, "ResponseXML"
     
    Dim rsList As IResponseList
     
    Set rsList = responseMsgSet.ResponseList
    
    Dim response As IResponse
    
    
     
    'The response list contains only one response,
    'which corresponds to our single request
    Set response = rsList.GetAt(0)
    'Interpret the response
    
    FindIfMulticurrencyOn = InterpretAccountingPrefQueryResponseCA(response)
    

    
    
    Exit Function
ErrHandler:
    ' In case something goes wrong along the way..
    MsgBox "Error sending request to QuickBooks: " & Err.Description
End Function

Function InterpretAccountingPrefQueryResponseCA(response As IResponse) As Boolean
'This function reads the details of the response and retrieve the list of currency names from it
On Error GoTo ErrHandler

    'read the response details
    Dim statusCode, statusMessage, statusSeverity, msg
    
    statusCode = response.statusCode
    statusMessage = response.statusMessage
    statusSeverity = response.statusSeverity
      
    Dim PreferenceRet As IPreferencesRet
    
    Set PreferenceRet = response.Detail
    'Find if multicurrency is turned on (if multicurrency selected in preferences, it returns true)
    If Not (PreferenceRet Is Nothing) Then
        InterpretAccountingPrefQueryResponseCA = PreferenceRet.AccountingPreferences.IsUsingMulticurrency.GetValue
    End If
   
   Exit Function
    
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"

End Function
