Attribute VB_Name = "QBFCEventsCallbackModule"
Const cAppName = "Desktop VB QBFC EventsCallback"
Const cAppID = ""

Public Sub Main()
    'For simplicity, this sample shows a UI Form every time it's run.
    QBFCEventsCallbackForm.Show
    
End Sub

' Queries QuickBooks for all the customers and displays
Public Sub GetAllCustomers()
    
    On Error GoTo Errs
    
    'Clear the query result text box
    QBFCEventsCallbackForm.QueryResult.Text = ""
    
    Dim result As String
    result = "QuickBooks list of customers: " & vbCrLf
    
    'Start a session with QuickBooks.
    ' Create the session manager object
    Dim sessionManager As New QBSessionManager
    Dim bConnectionOpen As Boolean
    Dim bSessionOpen As Boolean
    
    ' Connect to QuickBooks and begin a session.
    sessionManager.OpenConnection cAppID, cAppName
    bConnectionOpen = True
    sessionManager.BeginSession "", omDontCare
    bSessionOpen = True
    
    ' Create the message set request object for 3.0 version messages.
    Dim requestMsgSet As IMsgSetRequest
    Set requestMsgSet = sessionManager.CreateMsgSetRequest("US", 3, 0)
    requestMsgSet.Attributes.OnError = roeContinue
    
    'Add the request to the message set request object.
    Dim customerQuery As ICustomerQuery
    Set customerQuery = requestMsgSet.AppendCustomerQueryRq
             
    ' Perform the request and obtain a response from QuickBooks.
    Dim responseMsgSet As IMsgSetResponse
    Set responseMsgSet = sessionManager.DoRequests(requestMsgSet)
    
    ' Close the session and connection with QuickBooks.
    If (bConnectionOpen) Then
        sessionManager.EndSession
        bSessionOpen = False
        sessionManager.CloseConnection
        bConnectionOpen = False
    End If
    
    
    Dim responseList As IResponseList
    Set responseList = responseMsgSet.responseList
    If (responseList Is Nothing) Then
            Exit Sub
    End If
    
    ' Should only expect 1 response
    Dim response As IResponse
    Set response = responseList.GetAt(0)
    
    ' Iterate through the list of Customers
    
    
    ' Check the status returned for the response, which will be a CustomerRet.
    If (response.StatusCode = 0) Then
        Dim customerRetList As ICustomerRetList
        Set customerRetList = response.Detail
        ' Should only be 1 CustomerRet object returned
        Dim customerRet As ICustomerRet
        For i = 0 To customerRetList.Count - 1
            Set customerRet = customerRetList.GetAt(i)
            'Get the fullName
            customerFullName = customerRet.FullName.GetValue
            result = result + customerFullName & vbCrLf
        Next i
    End If
            
    'Write out the customer list
    QBFCEventsCallbackForm.QueryResult.Text = result
    QBFCEventsCallbackForm.ErrorMsg.Text = ""
    
    Exit Sub
    
Errs:
   
    ' Close the session and connection with QuickBooks.
    If (bConnectionOpen) Then
        sessionManager.EndSession
        bSessionOpen = False
        sessionManager.CloseConnection
        bConnectionOpen = False
    End If

    ' Write error to ErrorMsg Text Box
    QBFCEventsCallbackForm.ErrorMsg.Text = "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description

End Sub

