VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QBFCEventsCallbackClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements QBSDKEVENTLib.IQBEventCallback
        
Option Explicit
        
' AppID and AppName sent to QuickBooks
Const cAppName = "IDN Desktop VB QBFC EventsCallback"
Const cAppID = ""
               
Dim theEventXML As String

Private Sub IQBEventCallback_inform(ByVal eventXML As String)
    
    'The eventXML contains information such as the type and
    'operation of the event that occurred.
    'To get more information, you can send a qbXML query
    'request.
    'Note, however, that you cannot send requests that add, change or delete
    'data in QuickBooks, other than DataEventRecoveryInfoDelRq
    
    On Error GoTo Errs

    ' Clear the TextBoxes. You might consider creating a log to review the history
    ClearTextBoxes
    
    'Make this available to the other methods
    theEventXML = eventXML
    
    'Convert the eventXML to a QBFC response object
    Dim eventMsgSet As IEventsMsgSet
    Dim sessionManager As New QBSessionManager
    Set eventMsgSet = sessionManager.ToEventsMsgSet(eventXML, 3, 0)

    ' Check if it's a DataEvent or UIEvent
    Dim eventRet As IOREvent
    Set eventRet = eventMsgSet.OREvent
    
    If (Not eventRet.DataEventList Is Nothing) Then
        HandleDataEvent eventRet.DataEventList.GetAt(0)
    ElseIf (Not eventRet.UIEvent Is Nothing) Then
        HandleUIEvent eventRet.UIEvent
    Else
        QBFCEventsCallbackForm.ErrorMsg.Text = "QBFCEventsCallback only handles " _
            + " CustomerAdd DataEvent and Company File Close UIEvent."
    End If
    
    Exit Sub
        
Errs:
    ' Write error to ErrorMsg Text Box
    QBFCEventsCallbackForm.ErrorMsg.Text = "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description

End Sub

Private Sub HandleDataEvent(DataEvent As IDataEvent)
      
    Dim theListEvent As IListEvent
    Set theListEvent = DataEvent.ORListTxnEvent.ListEvent
      
    'This callback only supports CustomerAdd DataEvent, so check for that
    If (theListEvent Is Nothing) Then
        UnsupportedMessageOutput "DataEvent"
        Exit Sub
    End If
    ' Make sure this is an Add Event
    If (Not theListEvent.ListEventOperation.GetValue = leoAdd) Then
        UnsupportedMessageOutput "DataEvent"
        Exit Sub
    End If
    ' Make sure it was a Customer that was added
    If (Not theListEvent.ListEventType.GetValue = letCustomer) Then
        UnsupportedMessageOutput "DataEvent"
        Exit Sub
    End If
          
      
    'Query for that customer using the ListID, get the customer's FullName
    'Get the ListID from the event
    Dim theListID As String
    theListID = theListEvent.ListID.GetValue
    
    'Get the FullName of the customer returned in response
    Dim customerName As String
    GetCustomerAddedFullName theListID, customerName
    
    'Display the FullName of the Customer returned by the CustomerQuery
    'It is best to use a NON-blocking UI to display this information.
    ' A blocking UI is not recommended because as long as the application is in
    ' the callback, it will not receive other events.  Those events will be
    ' queued up to 100 and lost after that.  In addition, the SBO will be prevented
    ' from closing the company data file, because the QBXMLRP2 session is still open.
    ' Please refer to the Events documentation for more detail.
    
    ' Inform user that Customer Add Event received
    QBFCEventsCallbackForm.DataEvent.Text = "Customer Add Event received." & vbCrLf _
        + " Customer: " + customerName + " was added. "
    Exit Sub
    
    
    
End Sub
    
    
Private Sub HandleUIEvent(UIEvent As IUIEvent)
    
    ' For the company file Close event, we will disable the 'Query' button
    If (UIEvent Is Nothing) Then
            Exit Sub
    End If
    
    ' Check to see if this is really a company file close event
    If (Not UIEvent.CompanyFileEvent.CompanyFileEventOperation.GetValue = cfeoClose) Then
        ' This is an unsupported UI event
        UnsupportedMessageOutput "UIEvent"
        Exit Sub
    End If
        
            
    ' Grey out the "Query" button, disable access to QuickBooks
    QBFCEventsCallbackForm.Query.Enabled = False
    
        
    ' Inform user that they can no longer query QuickBooks because of company close event
    'It is best to use a NON-blocking UI to display this information.
    ' A blocking UI is not recommended because as long as the application is in
    ' the callback, it will not receive other events.  Those events will be
    ' queued up to 100 and lost after that.  In addition, the SBO will be prevented
    ' from closing the company data file, because the QBXMLRP2 session is still open.
    ' Please refer to the Events documentation for more detail.
    QBFCEventsCallbackForm.UIEvent.Text = "Company File Close Event received." & vbCrLf _
        + " No longer able to query QuickBooks."
            
End Sub
' Queries for the customer fullname with the given ListID from the event
Private Sub GetCustomerAddedFullName(theListID As String, ByRef customerFullName As String)
    
    On Error GoTo Errs
    
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
    
    ' Set the value of the IORCustomerListQuery.ListIDList element
    customerQuery.ORCustomerListQuery.ListIDList.Add theListID
             
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
    
    ' Check the status returned for the response, which will be a CustomerRet.
    If (response.StatusCode = 0) Then
        Dim customerRetList As ICustomerRetList
        Set customerRetList = response.Detail
        ' Should only be 1 CustomerRet object returned
        Dim customerRet As ICustomerRet
        Set customerRet = customerRetList.GetAt(0)
        
        'Get the fullName
        customerFullName = customerRet.FullName.GetValue
    End If
      
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

Private Sub ClearTextBoxes()

'Clear the DataEvent, UIEvent Msg text boxes

    QBFCEventsCallbackForm.DataEvent.Text = ""
    QBFCEventsCallbackForm.UIEvent.Text = ""
    QBFCEventsCallbackForm.ErrorMsg.Text = ""
    
End Sub

Private Sub UnsupportedMessageOutput(typeEvent As String)
    If (typeEvent = "UIEvent") Then
        'This Callback only supports Company File Close UI Event
        QBFCEventsCallbackForm.ErrorMsg.Text = "QBFCEventsCallback only handles " & vbCrLf _
            + " Company File Close UIEvent. vbCrLf " _
            + theEventXML
    ElseIf (typeEvent = "DataEvent") Then
            'This Callback only supports CustomerAdd Data Event
        ' This is an unsupported DataEvent event
        QBFCEventsCallbackForm.ErrorMsg.Text = "QBFCEventsCallback only handles " & vbCrLf _
            + " CustomerAdd DataEvents. vbCrLf " _
            + theEventXML
    End If
    
End Sub

