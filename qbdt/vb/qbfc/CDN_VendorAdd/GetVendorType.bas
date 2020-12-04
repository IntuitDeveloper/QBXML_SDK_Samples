Attribute VB_Name = "GetVendorType"
'-----------------------------------------------------------
' Module: GetVendorType.bas
'
' Description:  This module retrieve the vendor type list from the Canadian
'               version of QuickBooks or the US version of QuickBooks depending
'               on which function is called.
'
' Created On: 10/15/2002
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/rdmgr/?ID=100
'
'----------------------------------------------------------


Public Const cQBXMLMajorVersion = "2"
Public Const cQBXMLMinorVersion = "0"

Function GetVendorTypeListCA() As String
'This function sends a query to QuickBooks in order to get the vendor type list
On Error GoTo ErrHandler
    'Create the message set request object
    Dim requestMsgSet As IMsgSetRequest
    Set requestMsgSet = sessionManagerCA.CreateMsgSetRequest("CA", cQBXMLMajorVersion, cQBXMLMinorVersion)

    'Add the request to the message
     requestMsgSet.AppendVendorTypeQueryRq
    
'    MsgBox requestMsgSet.ToXMLString, vbOKOnly, "RequestXML"
    'Initialize the request's attributes
    requestMsgSet.Attributes.OnError = roeContinue
    
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
    GetVendorTypeListCA = InterpretVendorTypeQueryResponseCA(response)
    

    
    
    Exit Function
ErrHandler:
    ' In case something goes wrong along the way..
    MsgBox "Error sending request to QuickBooks: " & Err.Description
End Function

Private Function InterpretVendorTypeQueryResponseCA(response As IResponse) As String
'This function reads the details of the response and retrieve the list of vendor type from it
On Error GoTo ErrHandler

    Dim statusCode, statusMessage, statusSeverity, msg
    'read the response details
    statusCode = response.statusCode
    statusMessage = response.statusMessage
    statusSeverity = response.statusSeverity
      
    Dim VendorTypeRetList As IVendorTypeRetList
    Set VendorTypeRetList = response.Detail
    
    If Not (VendorTypeRetList Is Nothing) Then
        Dim ndx, count
        count = VendorTypeRetList.count
        For ndx = 0 To (count - 1)
            Dim VendorTypeRet As IVendorTypeRet
            Set VendorTypeRet = VendorTypeRetList.GetAt(ndx)
            
            InterpretVendorTypeQueryResponseCA = InterpretVendorTypeQueryResponseCA & VendorTypeRet.FullName.GetValue & "***"
        Next
    End If
   
   Exit Function
    
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
    InterpretVendorTypeQueryResponseCA = "Error!"
    
End Function

Function GetVendorTypeListUS() As String
'This function sends a query to QuickBooks in order to get the vendor type list
On Error GoTo ErrHandler
    'Create the message set request object
    Dim requestUSMsgSet As IMsgSetRequest
    Set requestUSMsgSet = sessionManagerUS.CreateMsgSetRequest("US", cQBXMLMajorVersion, cQBXMLMinorVersion)

    'Add the request to the message
    requestUSMsgSet.AppendVendorTypeQueryRq
    
    'MsgBox requestUSMsgSet.ToXMLString, vbOKOnly, "RequestXML"
    'Initialize the request's attributes
    requestUSMsgSet.Attributes.OnError = roeContinue
    
    
     'Perform the request
    Dim responseMsgSet As IMsgSetResponse
    Set responseMsgSet = sessionManagerUS.DoRequests(requestUSMsgSet)
    


'    MsgBox responseMsgSet.ToXMLString, vbOKOnly, "ResponseXML"
     
    Dim rsList As IResponseList

    Set rsList = responseMsgSet.ResponseList
    
    Dim response As IResponse
    
    'The response list contains only one response,
    'which corresponds to our single request
    Set response = rsList.GetAt(0)
    
    'Interpret the response
    GetVendorTypeListUS = InterpretVendorTypeQueryResponseUS(response)
    
    Exit Function
ErrHandler:
    ' In case something goes wrong along the way..
    MsgBox "Error sending request to QuickBooks: " & Err.Description
End Function

Private Function InterpretVendorTypeQueryResponseUS(response As IResponse) As String
'This function reads the details of the response and retrieve the list of vendor type from it
On Error GoTo ErrHandler

    Dim statusCode, statusMessage, statusSeverity, msg
    'read the response details
    statusCode = response.statusCode
    statusMessage = response.statusMessage
    statusSeverity = response.statusSeverity
      
    Dim VendorTypeRetList As IVendorTypeRetList
    Set VendorTypeRetList = response.Detail
    
    If Not (VendorTypeRetList Is Nothing) Then
        Dim ndx, count
        count = VendorTypeRetList.count
        For ndx = 0 To (count - 1)
            Dim VendorTypeRet As IVendorTypeRet
            Set VendorTypeRet = VendorTypeRetList.GetAt(ndx)
            
            InterpretVendorTypeQueryResponseUS = InterpretVendorTypeQueryResponseUS & VendorTypeRet.FullName.GetValue & "***"
        Next
    End If
   
   Exit Function
    
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
    InterpretVendorTypeQueryResponseUS = "Error!"
    
End Function

