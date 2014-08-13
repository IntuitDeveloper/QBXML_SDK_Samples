Attribute VB_Name = "QBFCEventsSubscriberModule"
Option Explicit

' AppID and AppName sent to QuickBooks
Public Const cAppName = "IDN Desktop VB QBFC EventsSubscriber"
Public Const cAppID = ""

' SubscriberID should be a unique identifier for this application
'  the value should be obtained by calling guidgen
Const cSubscriberID = "{2CB00476-1574-46d1-A2F6-4CFF4A5DD3E3}"

' Information about the Callback Application: "QBFCEventsCallback"
Const cCallbackAppName = "Desktop VB QBFC EventsCallback"
' ProgID is the ProjectName.ClassName
Const cCallbackProgID = "EventsCallback.QBFCEventsCallbackClass"

' Request and response strings
Public requestXMLStr  As String
Public responseXMLStr As String

Public Sub Main()
    'For simplicity, this sample shows a UI Form every time it's run.
    QBFCEventsSubscriber.Show
End Sub

Private Sub ClearTextBoxes()

'Clear the Request, Response, and Error Msg text boxes

    QBFCEventsSubscriber.RequestXML.Text = ""
    QBFCEventsSubscriber.ResponseXML.Text = ""
    QBFCEventsSubscriber.ErrorMsg.Text = ""

End Sub

Public Sub DoSubscribeEvents()

        On Error GoTo Errs
        
        ClearTextBoxes

        'We want to know if we've opened a connection so we can close it if an
        'error sends us to the exception handler.
        Dim bConnectionOpen As Boolean
        bConnectionOpen = False

        ' Create the session manager object.
        Dim sessionManager As New QBSessionManager

        ' Create the message set request object for 3.0 version messages.
        Dim requestMsgSet As ISubscriptionMsgSetRequest
        Set requestMsgSet = sessionManager.CreateSubscriptionMsgSetRequest(3, 0)

        ' Add the subscription events
        BuildDataEventSubscriptionAddRq requestMsgSet
        BuildUIEventSubscriptionAddRq requestMsgSet

        ' Display the xml request that will be sent
        requestXMLStr = requestMsgSet.ToXMLString
        ' Replace the linefeed with carriage return & linefeed for pretty output
        requestXMLStr = Replace(requestXMLStr, vbLf, vbCrLf)
        
        QBFCEventsSubscriber.RequestXML.Text = requestXMLStr

        ' Connect to QuickBooks and begin a session.
        sessionManager.OpenConnection cAppID, cAppName
        bConnectionOpen = True
        
        ' Send subscription request to QuickBooks.
        ' Please note that you don't need to call BeginSession.
        ' The subscriptions will be propagated to all the company files
        ' that are opened from this computer.
        
        ' Perform the request and obtain a response from QuickBooks.
        Dim responseMsgSet As ISubscriptionMsgSetResponse
        Set responseMsgSet = sessionManager.DoSubscriptionRequests(requestMsgSet)

        ' Close the connection with QuickBooks.
        sessionManager.CloseConnection
        bConnectionOpen = False

        ' Display the xml response
        responseXMLStr = responseMsgSet.ToXMLString
        ' Replace the linefeed with carriage return & linefeed for pretty output
        responseXMLStr = Replace(responseXMLStr, vbLf, vbCrLf)
        QBFCEventsSubscriber.ResponseXML.Text = responseXMLStr
        
        Exit Sub

Errs:
        ' Write error to ErrorMsg Text Box
        QBFCEventsSubscriber.ErrorMsg.Text = "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description

        ' Close the connection with QuickBooks.
        If (bConnectionOpen) Then
            sessionManager.CloseConnection
        End If

End Sub

Public Sub DoUnSubscribeEvents()

        On Error GoTo Errs

        ClearTextBoxes
        
        'We want to know if we've opened a connection so we can close it if an
        'error sends us to the exception handler.
        Dim bConnectionOpen As Boolean
        bConnectionOpen = False

        ' Create the session manager object.
        Dim sessionManager As New QBSessionManager

        ' Create the message set request object for 3.0 version messages.
        Dim requestMsgSet As ISubscriptionMsgSetRequest
        Set requestMsgSet = sessionManager.CreateSubscriptionMsgSetRequest(3, 0)

        ' Add unsubscribe DataEvents and UIEvents to request
        BuildDataEventSubscriptionDelRq requestMsgSet
        BuildUIEventSubscriptionDelRq requestMsgSet

        ' Display the xml request that will be sent
        requestXMLStr = requestMsgSet.ToXMLString
        QBFCEventsSubscriber.RequestXML.Text = requestXMLStr

        ' Connect to QuickBooks and begin a session.
        sessionManager.OpenConnection cAppID, cAppName
        bConnectionOpen = True
        
        ' Send subscription request to QuickBooks.
        ' Please note that you don't need to call BeginSession.
        ' The subscriptions will be propagated to all the company files
        ' that are opened from this computer.
        
        ' Perform the request and obtain a response from QuickBooks.
        Dim responseMsgSet As ISubscriptionMsgSetResponse
        Set responseMsgSet = sessionManager.DoSubscriptionRequests(requestMsgSet)

        ' Close the connection with QuickBooks.
        sessionManager.CloseConnection
        bConnectionOpen = False

        ' Display the xml response
        responseXMLStr = responseMsgSet.ToXMLString
        QBFCEventsSubscriber.ResponseXML.Text = responseXMLStr
        
        Exit Sub

Errs:
        ' Write error to ErrorMsg Text Box
        QBFCEventsSubscriber.ErrorMsg.Text = "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description

        ' Close the connection with QuickBooks.
        If (bConnectionOpen) Then
                sessionManager.CloseConnection
        End If

End Sub

Private Sub BuildDataEventSubscriptionAddRq(requestMsgSet As ISubscriptionMsgSetRequest)

        If (requestMsgSet Is Nothing) Then
                Exit Sub
        End If

        'Add the request to the message set request object
        Dim dataEventSubscriptionAdd As IDataEventSubscriptionAdd
        Set dataEventSubscriptionAdd = requestMsgSet.AppendDataEventSubscriptionAddRq

        'Set the elements of IDataEventSubscriptionAdd to subscribe to a CustomerAdd Event

        ' Set the value of the IDataEventSubscriptionAdd.SubscriberID
        ' The SubscriberID should be generated by running guidgen.exe
        dataEventSubscriptionAdd.SubscriberID.SetValue cSubscriberID

        ' Set the value of the ICOMCallbackInfo.AppName
        dataEventSubscriptionAdd.COMCallbackInfo.AppName.SetValue cCallbackAppName

        ' Set the value of the IORProgCLSID.ProgID
         dataEventSubscriptionAdd.COMCallbackInfo.ORProgCLSID.ProgId.SetValue cCallbackProgID
        
        ' Set the value of the IDataEventSubscriptionAdd.DeliveryPolicy
        dataEventSubscriptionAdd.DeliveryPolicy.SetValue dpDeliverAlways

        ' Set the value of the IDataEventSubscriptionAdd.TrackLostEvents
        dataEventSubscriptionAdd.TrackLostEvents.SetValue tleNone

        ' Subscribe to DataEvent:CustomerAdd
        Dim listEventSubscription As IListEventSubscription
        ' Append an element to the list and save the element in listEventSubscription so we can set its values.
        Set listEventSubscription = dataEventSubscriptionAdd.ListEventSubscriptionList.Append
        ' Set the value of the IListEventSubscription.ListEventTypeList
        listEventSubscription.ListEventTypeList.Add letCustomer
        ' Set the value of the IListEventSubscription.ListEventOperationList
        listEventSubscription.ListEventOperationList.Add leoAdd

End Sub

Private Sub BuildUIEventSubscriptionAddRq(requestMsgSet As ISubscriptionMsgSetRequest)

        If (requestMsgSet Is Nothing) Then
                Exit Sub
        End If

        'Add the request to the message set request object.
        Dim uIEventSubscriptionAdd As IUIEventSubscriptionAdd
        Set uIEventSubscriptionAdd = requestMsgSet.AppendUIEventSubscriptionAddRq

        'Set the elements of IUIEventSubscriptionAdd to subscribe to Company File Close event

        ' Set the value of the IUIEventSubscriptionAdd.SubscriberID
        ' The SubscriberID should be generated by running guidgen.exe
        uIEventSubscriptionAdd.SubscriberID.SetValue cSubscriberID

        ' Set the value of the ICOMCallbackInfo.AppName
        uIEventSubscriptionAdd.COMCallbackInfo.AppName.SetValue cCallbackAppName

        ' Set the value of the IORProgCLSID.ProgID
        uIEventSubscriptionAdd.COMCallbackInfo.ORProgCLSID.ProgId.SetValue cCallbackProgID
                
        ' Set the value of the IUIEventSubscriptionAdd.DeliveryPolicy
        uIEventSubscriptionAdd.DeliveryPolicy.SetValue dpDeliverAlways

        ' Subscribe to company file close events
        uIEventSubscriptionAdd.CompanyFileEventSubscription.CompanyFileEventOperationList.Add cfeoClose

End Sub

Private Sub BuildDataEventSubscriptionDelRq(requestMsgSet As ISubscriptionMsgSetRequest)

        If (requestMsgSet Is Nothing) Then
                Exit Sub
        End If

        'Add the request to the message set request object
        Dim subscriptionDel As ISubscriptionDel
        Set subscriptionDel = requestMsgSet.AppendSubscriptionDelRq

        ' Set the SubscriberID
        subscriptionDel.SubscriberID.SetValue cSubscriberID

        ' Unsubscribe to DataEvents (Customer Add)
        subscriptionDel.SubscriptionType.SetValue stData

End Sub


Private Sub BuildUIEventSubscriptionDelRq(requestMsgSet As ISubscriptionMsgSetRequest)

        If (requestMsgSet Is Nothing) Then
                Exit Sub
        End If

        'Add the request to the message set request object
        Dim subscriptionDel As ISubscriptionDel
        Set subscriptionDel = requestMsgSet.AppendSubscriptionDelRq

        ' Set the SubscriberID
        subscriptionDel.SubscriberID.SetValue cSubscriberID

        ' Unsubscribe to UIEvents (Company File Close)
        subscriptionDel.SubscriptionType.SetValue stUI

End Sub

