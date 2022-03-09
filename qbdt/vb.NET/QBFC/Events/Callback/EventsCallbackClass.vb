Option Strict Off
Option Explicit On
Imports System.ComponentModel
Imports System.Runtime.InteropServices

<ComClass(QBFCEventsCallbackClass.ClassId, QBFCEventsCallbackClass.InterfaceId,
          QBFCEventsCallbackClass.EventsId), ComVisible(True)>
Public Class QBFCEventsCallbackClass
    Inherits ReferenceCountedObject
    Implements QBSDKEVENTLib.IQBEventCallback
    Public Const ClassId As String = "E59808E4-CA2D-4052-9BA1-B66C12DEEE49"
    Public Const InterfaceId As String = "722ACA10-BC1F-47B4-8763-7EC12B589E2E"
    Public Const EventsId As String = "F15B8823-8FFE-4AD1-B299-598E91263F3F"


    ' AppID and AppName sent to QuickBooks
    Const cAppName As String = "IDN Desktop VB QBFC EventsCallback"
    Const cAppID As String = ""

    Dim theEventXML As String

    Public Sub IQBEventCallback_inform(ByVal eventXML As String) Implements QBSDKEVENTLib.IQBEventCallback.inform

        'The eventXML contains information such as the type and
        'operation of the event that occurred.
        'To get more information, you can send a qbXML query
        'request.
        'Note, however, that you cannot send requests that add, change or delete
        'data in QuickBooks, other than DataEventRecoveryInfoDelRq

        On Error GoTo Errs

        ' Clear the TextBoxes. You might consider creating a log to review the history
        ClearTextBoxes()

        'Make this available to the other methods
        theEventXML = eventXML

        'Convert the eventXML to a QBFC response object
        Dim eventMsgSet As QBFC15Lib.IEventsMsgSet
        Dim sessionManager As New QBFC15Lib.QBSessionManager
        eventMsgSet = sessionManager.ToEventsMsgSet(eventXML, 3, 0)

        ' Check if it's a DataEvent or UIEvent
        Dim eventRet As QBFC15Lib.IOREvent
        eventRet = eventMsgSet.OREvent

        If (Not eventRet.DataEventList Is Nothing) Then
            HandleDataEvent(eventRet.DataEventList.GetAt(0))
        ElseIf (Not eventRet.UIEvent Is Nothing) Then
            HandleUIEvent((eventRet.UIEvent))
        Else
            If frm.InvokeRequired Then
                frm.Invoke(Sub()
                               frm.ErrorMsg.Text = "QBFCEventsCallback only handles " & " CustomerAdd DataEvent and Company File Close UIEvent."
                           End Sub)
            Else
                frm.ErrorMsg.Text = "QBFCEventsCallback only handles " & " CustomerAdd DataEvent and Company File Close UIEvent."
            End If
        End If

        Exit Sub

Errs:
        ' Write error to ErrorMsg Text Box
        If frm.InvokeRequired Then
            frm.Invoke(Sub()
                           frm.ErrorMsg.Text = "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description
                       End Sub)
        Else
            frm.ErrorMsg.Text = "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description
        End If
    End Sub

    Private Sub HandleDataEvent(ByRef DataEvent As QBFC15Lib.IDataEvent)

        Dim theListEvent As QBFC15Lib.IListEvent
        theListEvent = DataEvent.ORListTxnEvent.ListEvent

        'This callback only supports CustomerAdd DataEvent, so check for that
        If (theListEvent Is Nothing) Then
            UnsupportedMessageOutput("DataEvent")
            Exit Sub
        End If
        ' Make sure this is an Add Event
        If (Not theListEvent.ListEventOperation.GetValue = QBFC15Lib.ENListEventOperation.leoAdd) Then
            UnsupportedMessageOutput("DataEvent")
            Exit Sub
        End If
        ' Make sure it was a Customer that was added
        If (Not theListEvent.ListEventType.GetValue = QBFC15Lib.ENListEventType.letCustomer) Then
            UnsupportedMessageOutput("DataEvent")
            Exit Sub
        End If


        'Query for that customer using the ListID, get the customer's FullName
        'Get the ListID from the event
        Dim theListID As String
        theListID = theListEvent.ListID.GetValue

        'Get the FullName of the customer returned in response
        Dim customerName As String
        GetCustomerAddedFullName(theListID, customerName)

        'Display the FullName of the Customer returned by the CustomerQuery
        'It is best to use a NON-blocking UI to display this information.
        ' A blocking UI is not recommended because as long as the application is in
        ' the callback, it will not receive other events.  Those events will be
        ' queued up to 100 and lost after that.  In addition, the SBO will be prevented
        ' from closing the company data file, because the QBXMLRP2 session is still open.
        ' Please refer to the Events documentation for more detail.

        ' Inform user that Customer Add Event received
        If frm.InvokeRequired Then
            frm.Invoke(Sub()
                           frm.DataEvent.Text = "Customer Add Event received." & vbCrLf & " Customer: " & customerName & " was added. "
                       End Sub)
        Else
            frm.DataEvent.Text = "Customer Add Event received." & vbCrLf & " Customer: " & customerName & " was added. "
        End If
        Exit Sub



    End Sub


    Private Sub HandleUIEvent(ByRef UIEvent As QBFC15Lib.IUIEvent)

        ' For the company file Close event, we will disable the 'Query' button
        If (UIEvent Is Nothing) Then
            Exit Sub
        End If

        ' Check to see if this is really a company file close event
        If (Not UIEvent.CompanyFileEvent.CompanyFileEventOperation.GetValue = QBFC15Lib.ENCompanyFileEventOperation.cfeoClose) Then
            ' This is an unsupported UI event
            UnsupportedMessageOutput("UIEvent")
            Exit Sub
        End If

        If frm.InvokeRequired Then
            frm.Invoke(Sub()
                           ' Grey out the "Query" button, disable access to QuickBooks
                           frm.Query.Enabled = False


                           ' Inform user that they can no longer query QuickBooks because of company close event
                           'It is best to use a NON-blocking UI to display this information.
                           ' A blocking UI is not recommended because as long as the application is in
                           ' the callback, it will not receive other events.  Those events will be
                           ' queued up to 100 and lost after that.  In addition, the SBO will be prevented
                           ' from closing the company data file, because the QBXMLRP2 session is still open.
                           ' Please refer to the Events documentation for more detail.
                           frm.UIEvent.Text = "Company File Close Event received." & vbCrLf & " No longer able to query QuickBooks."
                       End Sub)
        Else
            frm.Query.Enabled = False
            frm.UIEvent.Text = "Company File Close Event received." & vbCrLf & " No longer able to query QuickBooks."
        End If


    End Sub
    ' Queries for the customer fullname with the given ListID from the event
    Private Sub GetCustomerAddedFullName(ByRef theListID As String, ByRef customerFullName As String)

        On Error GoTo Errs

        'Start a session with QuickBooks.
        ' Create the session manager object
        Dim sessionManager As New QBFC15Lib.QBSessionManager
        Dim bConnectionOpen As Boolean
        Dim bSessionOpen As Boolean

        ' Connect to QuickBooks and begin a session.
        sessionManager.OpenConnection(cAppID, cAppName)
        bConnectionOpen = True
        sessionManager.BeginSession("", QBFC15Lib.ENOpenMode.omDontCare)
        bSessionOpen = True

        ' Create the message set request object for 3.0 version messages.
        Dim requestMsgSet As QBFC15Lib.IMsgSetRequest
        requestMsgSet = sessionManager.CreateMsgSetRequest("US", 3, 0)
        requestMsgSet.Attributes.OnError = QBFC15Lib.ENRqOnError.roeContinue

        'Add the request to the message set request object.
        Dim customerQuery As QBFC15Lib.ICustomerQuery
        customerQuery = requestMsgSet.AppendCustomerQueryRq

        ' Set the value of the IORCustomerListQuery.ListIDList element
        customerQuery.ORCustomerListQuery.ListIDList.Add(theListID)

        ' Perform the request and obtain a response from QuickBooks.
        Dim responseMsgSet As QBFC15Lib.IMsgSetResponse
        responseMsgSet = sessionManager.DoRequests(requestMsgSet)

        ' Close the session and connection with QuickBooks.
        If (bConnectionOpen) Then
            sessionManager.EndSession()
            bSessionOpen = False
            sessionManager.CloseConnection()
            bConnectionOpen = False
        End If


        Dim responseList As QBFC15Lib.IResponseList
        responseList = responseMsgSet.ResponseList
        If (responseList Is Nothing) Then
            Exit Sub
        End If

        ' Should only expect 1 response
        Dim response As QBFC15Lib.IResponse
        response = responseList.GetAt(0)

        ' Check the status returned for the response, which will be a CustomerRet.
        Dim customerRetList As QBFC15Lib.ICustomerRetList
        Dim customerRet As QBFC15Lib.ICustomerRet
        If (response.StatusCode = 0) Then
            customerRetList = response.Detail
            ' Should only be 1 CustomerRet object returned
            customerRet = customerRetList.GetAt(0)

            'Get the fullName
            customerFullName = customerRet.FullName.GetValue
        End If

        Exit Sub

Errs:

        ' Close the session and connection with QuickBooks.
        If (bConnectionOpen) Then
            sessionManager.EndSession()
            bSessionOpen = False
            sessionManager.CloseConnection()
            bConnectionOpen = False
        End If

        ' Write error to ErrorMsg Text Box
        If frm.InvokeRequired Then
            frm.Invoke(Sub()
                           frm.ErrorMsg.Text = "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description
                       End Sub)
        Else
            frm.ErrorMsg.Text = "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description
        End If
    End Sub

    Private Sub ClearTextBoxes()

        'Clear the DataEvent, UIEvent Msg text boxes
        If frm.InvokeRequired Then
            frm.Invoke(Sub()
                           frm.DataEvent.Text = ""
                           frm.UIEvent.Text = ""
                           frm.ErrorMsg.Text = ""

                       End Sub)
        Else
            frm.DataEvent.Text = ""
            frm.UIEvent.Text = ""
            frm.ErrorMsg.Text = ""
        End If
    End Sub

    Private Sub UnsupportedMessageOutput(typeEvent As String)

        If frm.InvokeRequired Then
            frm.Invoke(Sub()
                           If (typeEvent = "UIEvent") Then
                               'This Callback only supports Company File Close UI Event
                               frm.ErrorMsg.Text = "QBFCEventsCallback only handles " & vbCrLf & " Company File Close UIEvent. vbCrLf " & theEventXML
                           ElseIf (typeEvent = "DataEvent") Then
                               'This Callback only supports CustomerAdd Data Event
                               ' This is an unsupported DataEvent event
                               frm.ErrorMsg.Text = "QBFCEventsCallback only handles " & vbCrLf & " CustomerAdd DataEvents. vbCrLf " & theEventXML
                           End If
                       End Sub)
        Else
            If (typeEvent = "UIEvent") Then
                'This Callback only supports Company File Close UI Event
                frm.ErrorMsg.Text = "QBFCEventsCallback only handles " & vbCrLf & " Company File Close UIEvent. vbCrLf " & theEventXML
            ElseIf (typeEvent = "DataEvent") Then
                'This Callback only supports CustomerAdd Data Event
                ' This is an unsupported DataEvent event
                frm.ErrorMsg.Text = "QBFCEventsCallback only handles " & vbCrLf & " CustomerAdd DataEvents. vbCrLf " & theEventXML
            End If
        End If


    End Sub

    <ComRegisterFunction(), EditorBrowsable(EditorBrowsableState.Never)>
    Public Shared Sub Register(ByVal t As Type)
        Try
            COMRegister.RegisterAsLocalServer(t)
        Catch ex As Exception
            Console.WriteLine(ex.Message) ' Log the error
            Throw ex ' Re-throw the exception
        End Try
    End Sub

    <EditorBrowsable(EditorBrowsableState.Never), ComUnregisterFunction()>
    Public Shared Sub Unregister(ByVal t As Type)
        Try
            COMRegister.UnRegisterAsLocalServer(t)
        Catch ex As Exception
            Console.WriteLine(ex.Message) ' Log the error
            Throw ex ' Re-throw the exception
        End Try
    End Sub
End Class