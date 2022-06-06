Option Strict Off
Option Explicit On
Module QBFCEventsSubscriberModule
	
	' AppID and AppName sent to QuickBooks
	Public Const cAppName As String = "IDN Desktop VB QBFC EventsSubscriber"
	Public Const cAppID As String = ""
	
	' SubscriberID should be a unique identifier for this application
	'  the value should be obtained by calling guidgen
	Const cSubscriberID As String = "{2CB00476-1574-46d1-A2F6-4CFF4A5DD3E3}"
	
	' Information about the Callback Application: "QBFCEventsCallback"
	Const cCallbackAppName As String = "Desktop VB QBFC EventsCallback"
    ' ProgID is the ProjectName.ClassName
    Const cCallbackProgID As String = "QBFCEventsCallback.QBFCEventsCallbackClass"

    ' Request and response strings
    Public requestXMLStr As String
	Public responseXMLStr As String
	
	'UPGRADE_NOTE: Main was upgraded to Main_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub Main_Renamed()
		'For simplicity, this sample shows a UI Form every time it's run.
		QBFCEventsSubscriber.Show()
	End Sub
	
	Private Sub ClearTextBoxes()
		
		'Clear the Request, Response, and Error Msg text boxes
		
		QBFCEventsSubscriber.RequestXML.Text = ""
		QBFCEventsSubscriber.ResponseXML.Text = ""
		QBFCEventsSubscriber.ErrorMsg.Text = ""
		
	End Sub
	
	Public Sub DoSubscribeEvents()
		
		On Error GoTo Errs
		
		ClearTextBoxes()
		
		'We want to know if we've opened a connection so we can close it if an
		'error sends us to the exception handler.
		Dim bConnectionOpen As Boolean
		bConnectionOpen = False
		
		' Create the session manager object.
		Dim sessionManager As New QBFC15Lib.QBSessionManager
		
		' Create the message set request object for 3.0 version messages.
		Dim requestMsgSet As QBFC15Lib.ISubscriptionMsgSetRequest
		requestMsgSet = sessionManager.CreateSubscriptionMsgSetRequest(3, 0)
		
		' Add the subscription events
		BuildDataEventSubscriptionAddRq(requestMsgSet)
		BuildUIEventSubscriptionAddRq(requestMsgSet)
		
		' Display the xml request that will be sent
		requestXMLStr = requestMsgSet.ToXMLString
		' Replace the linefeed with carriage return & linefeed for pretty output
		requestXMLStr = Replace(requestXMLStr, vbLf, vbCrLf)
		
		QBFCEventsSubscriber.RequestXML.Text = requestXMLStr
		
		' Connect to QuickBooks and begin a session.
		sessionManager.OpenConnection(cAppID, cAppName)
		bConnectionOpen = True
		
		' Send subscription request to QuickBooks.
		' Please note that you don't need to call BeginSession.
		' The subscriptions will be propagated to all the company files
		' that are opened from this computer.
		
		' Perform the request and obtain a response from QuickBooks.
		Dim responseMsgSet As QBFC15Lib.ISubscriptionMsgSetResponse
		responseMsgSet = sessionManager.DoSubscriptionRequests(requestMsgSet)
		
		' Close the connection with QuickBooks.
		sessionManager.CloseConnection()
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
			sessionManager.CloseConnection()
		End If
		
	End Sub
	
	Public Sub DoUnSubscribeEvents()
		
		On Error GoTo Errs
		
		ClearTextBoxes()
		
		'We want to know if we've opened a connection so we can close it if an
		'error sends us to the exception handler.
		Dim bConnectionOpen As Boolean
		bConnectionOpen = False
		
		' Create the session manager object.
		Dim sessionManager As New QBFC15Lib.QBSessionManager
		
		' Create the message set request object for 3.0 version messages.
		Dim requestMsgSet As QBFC15Lib.ISubscriptionMsgSetRequest
		requestMsgSet = sessionManager.CreateSubscriptionMsgSetRequest(3, 0)
		
		' Add unsubscribe DataEvents and UIEvents to request
		BuildDataEventSubscriptionDelRq(requestMsgSet)
		BuildUIEventSubscriptionDelRq(requestMsgSet)
		
		' Display the xml request that will be sent
		requestXMLStr = requestMsgSet.ToXMLString
		QBFCEventsSubscriber.RequestXML.Text = requestXMLStr
		
		' Connect to QuickBooks and begin a session.
		sessionManager.OpenConnection(cAppID, cAppName)
		bConnectionOpen = True
		
		' Send subscription request to QuickBooks.
		' Please note that you don't need to call BeginSession.
		' The subscriptions will be propagated to all the company files
		' that are opened from this computer.
		
		' Perform the request and obtain a response from QuickBooks.
		Dim responseMsgSet As QBFC15Lib.ISubscriptionMsgSetResponse
		responseMsgSet = sessionManager.DoSubscriptionRequests(requestMsgSet)
		
		' Close the connection with QuickBooks.
		sessionManager.CloseConnection()
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
			sessionManager.CloseConnection()
		End If
		
	End Sub
	
	Private Sub BuildDataEventSubscriptionAddRq(ByRef requestMsgSet As QBFC15Lib.ISubscriptionMsgSetRequest)
		
		If (requestMsgSet Is Nothing) Then
			Exit Sub
		End If
		
		'Add the request to the message set request object
		Dim dataEventSubscriptionAdd As QBFC15Lib.IDataEventSubscriptionAdd
		dataEventSubscriptionAdd = requestMsgSet.AppendDataEventSubscriptionAddRq
		
		'Set the elements of IDataEventSubscriptionAdd to subscribe to a CustomerAdd Event
		
		' Set the value of the IDataEventSubscriptionAdd.SubscriberID
		' The SubscriberID should be generated by running guidgen.exe
		dataEventSubscriptionAdd.SubscriberID.SetValue(cSubscriberID)
		
		' Set the value of the ICOMCallbackInfo.AppName
		dataEventSubscriptionAdd.COMCallbackInfo.AppName.SetValue(cCallbackAppName)
		
		' Set the value of the IORProgCLSID.ProgID
		dataEventSubscriptionAdd.COMCallbackInfo.ORProgCLSID.ProgId.SetValue(cCallbackProgID)
		
		' Set the value of the IDataEventSubscriptionAdd.DeliveryPolicy
		dataEventSubscriptionAdd.DeliveryPolicy.SetValue(QBFC15Lib.ENDeliveryPolicy.dpDeliverAlways)
		
		' Set the value of the IDataEventSubscriptionAdd.TrackLostEvents
		dataEventSubscriptionAdd.TrackLostEvents.SetValue(QBFC15Lib.ENTrackLostEvents.tleNone)
		
		' Subscribe to DataEvent:CustomerAdd
		Dim listEventSubscription As QBFC15Lib.IListEventSubscription
		' Append an element to the list and save the element in listEventSubscription so we can set its values.
		listEventSubscription = dataEventSubscriptionAdd.ListEventSubscriptionList.Append
		' Set the value of the IListEventSubscription.ListEventTypeList
		listEventSubscription.ListEventTypeList.Add(QBFC15Lib.ENListEventType.letCustomer)
		' Set the value of the IListEventSubscription.ListEventOperationList
		listEventSubscription.ListEventOperationList.Add(QBFC15Lib.ENListEventOperation.leoAdd)
		
	End Sub
	
	Private Sub BuildUIEventSubscriptionAddRq(ByRef requestMsgSet As QBFC15Lib.ISubscriptionMsgSetRequest)
		
		If (requestMsgSet Is Nothing) Then
			Exit Sub
		End If
		
		'Add the request to the message set request object.
		Dim uIEventSubscriptionAdd As QBFC15Lib.IUIEventSubscriptionAdd
		uIEventSubscriptionAdd = requestMsgSet.AppendUIEventSubscriptionAddRq
		
		'Set the elements of IUIEventSubscriptionAdd to subscribe to Company File Close event
		
		' Set the value of the IUIEventSubscriptionAdd.SubscriberID
		' The SubscriberID should be generated by running guidgen.exe
		uIEventSubscriptionAdd.SubscriberID.SetValue(cSubscriberID)
		
		' Set the value of the ICOMCallbackInfo.AppName
		uIEventSubscriptionAdd.COMCallbackInfo.AppName.SetValue(cCallbackAppName)
		
		' Set the value of the IORProgCLSID.ProgID
		uIEventSubscriptionAdd.COMCallbackInfo.ORProgCLSID.ProgId.SetValue(cCallbackProgID)
		
		' Set the value of the IUIEventSubscriptionAdd.DeliveryPolicy
		uIEventSubscriptionAdd.DeliveryPolicy.SetValue(QBFC15Lib.ENDeliveryPolicy.dpDeliverAlways)
		
		' Subscribe to company file close events
		uIEventSubscriptionAdd.CompanyFileEventSubscription.CompanyFileEventOperationList.Add(QBFC15Lib.ENCompanyFileEventOperation.cfeoClose)
		
	End Sub
	
	Private Sub BuildDataEventSubscriptionDelRq(ByRef requestMsgSet As QBFC15Lib.ISubscriptionMsgSetRequest)
		
		If (requestMsgSet Is Nothing) Then
			Exit Sub
		End If
		
		'Add the request to the message set request object
		Dim subscriptionDel As QBFC15Lib.ISubscriptionDel
		subscriptionDel = requestMsgSet.AppendSubscriptionDelRq
		
		' Set the SubscriberID
		subscriptionDel.SubscriberID.SetValue(cSubscriberID)
		
		' Unsubscribe to DataEvents (Customer Add)
		subscriptionDel.SubscriptionType.SetValue(QBFC15Lib.ENSubscriptionType.stData)
		
	End Sub
	
	
	Private Sub BuildUIEventSubscriptionDelRq(ByRef requestMsgSet As QBFC15Lib.ISubscriptionMsgSetRequest)
		
		If (requestMsgSet Is Nothing) Then
			Exit Sub
		End If
		
		'Add the request to the message set request object
		Dim subscriptionDel As QBFC15Lib.ISubscriptionDel
		subscriptionDel = requestMsgSet.AppendSubscriptionDelRq
		
		' Set the SubscriberID
		subscriptionDel.SubscriberID.SetValue(cSubscriberID)
		
		' Unsubscribe to UIEvents (Company File Close)
		subscriptionDel.SubscriptionType.SetValue(QBFC15Lib.ENSubscriptionType.stUI)
		
	End Sub
End Module