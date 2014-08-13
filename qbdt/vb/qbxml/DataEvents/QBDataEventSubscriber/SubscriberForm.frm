VERSION 5.00
Begin VB.Form SubscriberForm 
   Caption         =   "QBDataEventSubscriber"
   ClientHeight    =   2145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   ScaleHeight     =   2145
   ScaleWidth      =   4980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Unsubscribe 
      Caption         =   "UNSubscribe From Events!"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   4215
   End
   Begin VB.CommandButton SubscribeBtn 
      Caption         =   "Subscribe To Events!"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "SubscriberForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' This application is made up of 3 executables:
' QBDataEventApp        -- Provides the UI for this application
' QBDataEventSubscriber -- Subscribes the application to receive data events
'                          from QuickBooks.  This is the sort of application you
'                          would build as a Custom Action for an MSI installer to
'                          subscribe your application to events when you install your
'                          app and unsubscribe when your app is uninstalled.
' QBDataEventManager    -- Handles events from QuickBooks and supplies them to the main application
'                          via a second COM interface -- also has a UI to display the
'                          incoming event XML.
'
' The QBDataEventSubscriber is the simplest of all three components,
' we have a simple form with two buttons, one to subscribe and one to
' unsubscribe from events using the ProcessSubscription call to the RequestProcessor2

Private Sub SubscribeBtn_Click()
    ' Subscribe to events...
    Dim ConnOpen As Integer
    ConnOpen = 0
    On Error GoTo cleanup
    
    ' Get the RequestProcessor and open a connection.
    Dim RP As New QBXMLRP2Lib.RequestProcessor2
    RP.OpenConnection "", "IDN DataEventSample"
    ConnOpen = 1
    

    ' Note that there is no need to BeginSession in order to send a
    ' subscription request, the QuickBooks administrator will be asked
    ' to grant your application permission to receive events the next time
    ' QuickBooks starts...
    
    ' Create the outer subscription request XML "envelope"
    Dim doc As New DOMDocument40
    Dim QBXML As IXMLDOMElement
    Set QBXML = doc.createElement("QBXML")
    doc.appendChild QBXML
    Dim SubReq As IXMLDOMElement
    ' Note, for a subscription our envelope has a QBXMLSubscriptionMsgsRq, not
    ' a QBXMLMsgsRq
    Set SubReq = doc.createElement("QBXMLSubscriptionMsgsRq")
    QBXML.appendChild SubReq
    
    ' Now create the Data Event subscription to find out about customer changes
    ' Note that we will get notifications indirectly if the customer balance changes, etc.
    ' when an invoice is created...
    Dim DataSubReq As IXMLDOMElement
    Set DataSubReq = doc.createElement("DataEventSubscriptionAddRq")
    SubReq.appendChild DataSubReq
    Dim DataSubAdd As IXMLDOMElement
    Set DataSubAdd = doc.createElement("DataEventSubscriptionAdd")
    DataSubReq.appendChild DataSubAdd
    
    ' Note that our SubscriberID is the same as the DataEventApp uses for
    ' the OwnerID of data extensions.  this isn't necessary but it is convenient...
    AddSimpleElement doc, DataSubAdd, "SubscriberID", "{2B6C9DB4-EBE2-45E7-A14F-4E1C49C965F7}"
    
    ' Set up the COMCallback information
    Dim COMCallback As IXMLDOMElement
    Set COMCallback = doc.createElement("COMCallbackInfo")
    DataSubAdd.appendChild COMCallback
    AddSimpleElement doc, COMCallback, "AppName", "DataEventSample"
    
    ' We supply the ProgID, this is the most convenient way to supply callback
    ' information for a COM class written in visual basic which does it's best to
    ' hide information like CLSIDs from the programmer.
    '
    ' Note!  The ProgID will be converted to a CLSID internally by QuickBooks so it
    ' is very important that you set the properties of the project implementing the
    ' callback class to "binary compatibility" mode on the "Components" tab of the
    ' project properties dialog, otherwise VB will be foolish enough to change the
    ' CLSID with each recompile...
    AddSimpleElement doc, COMCallback, "ProgID", "QBDataEventManager.QBEventHandler"
    
    '
    ' We set to DeliverAlways just to show how we can have a fine granularity of
    ' whether to process events or not by implementing an application-specific event
    ' queue.  This could easily be changed to DeliverOnlyIfRunning and then as soon
    ' as the QBDataEventApp creates an instance of the QBDateEventHandler class to
    ' start tracking events QuickBooks will start delivering the events...
    '
    AddSimpleElement doc, DataSubAdd, "DeliveryPolicy", "DeliverAlways"
    
    '
    ' With QuickBooks 2005 we can avoid getting our own events, so lets do that if we can...
    '
    Dim strXMLVersionsArray() As String
    strXMLVersionsArray = RP.QBXMLVersionsForSubscription
    Dim useVersion As String
    Dim supportsXML40 As Boolean
    useVersion = "3.0"
    supportsXML40 = False
    For i = LBound(strXMLVersionsArray) To UBound(strXMLVersionsArray)
        If (strXMLVersionsArray(i) = "4.0" Or strXMLVersionsArray(i) = "CA4.0") Then
            supportsXML40 = True
            useVersion = strXMLVersionsArray(i)
        End If
    Next i
    If (supportsXML40) Then
        AddSimpleElement doc, DataSubAdd, "DeliverOwnEvents", "false"
    End If
    
    '
    ' Now we tell QuickBooks what events we are interested in, we'll subscribe
    ' to ListEvents that affect Customer objects, and we care about all types of
    ' events (add, modify, delete, and merge).
    '
    Dim ListEventSub As IXMLDOMElement
    Set ListEventSub = doc.createElement("ListEventSubscription")
    DataSubAdd.appendChild ListEventSub
    AddSimpleElement doc, ListEventSub, "ListEventType", "Customer"
    AddSimpleElement doc, ListEventSub, "ListEventOperation", "Add"
    AddSimpleElement doc, ListEventSub, "ListEventOperation", "Modify"
    AddSimpleElement doc, ListEventSub, "ListEventOperation", "Delete"
    AddSimpleElement doc, ListEventSub, "ListEventOperation", "Merge"
    
    'Now, because our EventManager has a UI that we want to dismiss when
    'we're all done we are going to subscribe to the UIEvents for company
    'close so we can shut down when QuickBooks shuts down.  Similarly, we want
    'our application (which might be holding a pointer to the RequestProcessor) to
    'drop those pointers as quickly as possible when we get a close event, otherwise
    'QuickBooks won't be able to close while our main app is running.  Note also
    'we STILL want the handler to close when it gets the close event because if the
    'EventApp has shut down then we want the handler to shut down completely...
    Dim UISubReq As IXMLDOMElement
    Set UISubReq = doc.createElement("UIEventSubscriptionAddRq")
    SubReq.appendChild UISubReq
    Dim UISubAdd As IXMLDOMElement
    Set UISubAdd = doc.createElement("UIEventSubscriptionAdd")
    UISubReq.appendChild UISubAdd
    ' Same subscriber ID as for DataEvents
    AddSimpleElement doc, UISubAdd, "SubscriberID", "{2B6C9DB4-EBE2-45E7-A14F-4E1C49C965F7}"
    
    ' Same Callback info as for DataEvents, and same delivery policy
    Set COMCallback = doc.createElement("COMCallbackInfo")
    UISubAdd.appendChild COMCallback
    AddSimpleElement doc, COMCallback, "AppName", "DataEventSample"
    AddSimpleElement doc, COMCallback, "ProgID", "QBDataEventManager.QBEventHandler"
    AddSimpleElement doc, UISubAdd, "DeliveryPolicy", "DeliverAlways"
    Dim UIEventSub As IXMLDOMElement
    Set UIEventSub = doc.createElement("CompanyFileEventSubscription")
    UISubAdd.appendChild UIEventSub
    ' We care only about the Close event, not the open event.
    AddSimpleElement doc, UIEventSub, "CompanyFileEventOperation", "Close"
    
    
    'Finally, send the subscription request to QuickBooks
    Dim subXML As String
    subXML = "<?xml version=""1.0""?>" & vbCrLf & "<?qbxml version=""" & useVersion & """?>" & vbCrLf & doc.xml
    saveXMLStream subXML
    Dim resp As String
    resp = RP.ProcessSubscription(subXML)
    ' We'll just show the response.  Note that if we try to subscribe again we'll get
    ' an error.  We're not doing anything else with the response...
    MsgBox Prompt:=resp, Title:="Subscribe Complete"
    
    ' And finally close our connection to QuickBooks
cleanup:
    If (Err.Number) Then
        MsgBox ("Unexpected error from QuickBooks: " & Err.Number & ": " & Err.Description)
    End If
    If (ConnOpen) Then
        RP.CloseConnection
    End If
End Sub


Private Sub Unsubscribe_Click()
    'Simply delete the subscriptions we set up in Subscribe_click
    Dim ConnOpen As Integer
    ConnOpen = 0
    On Error GoTo cleanup
    Dim RP As New QBXMLRP2Lib.RequestProcessor2
    RP.OpenConnection "", "IDN DataEventSample"
    ConnOpen = 1
    
    ' Create the outer subscription request XML "envelope"
    Dim doc As New DOMDocument40
    Dim QBXML As IXMLDOMElement
    Set QBXML = doc.createElement("QBXML")
    doc.appendChild QBXML
    Dim SubReq As IXMLDOMElement
    Set SubReq = doc.createElement("QBXMLSubscriptionMsgsRq")
    QBXML.appendChild SubReq
    
    ' Now create a subscription delete request for the Data event subscription
    Dim DataSubDel As IXMLDOMElement
    Set DataSubDel = doc.createElement("SubscriptionDelRq")
    SubReq.appendChild DataSubDel
    AddSimpleElement doc, DataSubDel, "SubscriberID", "{2B6C9DB4-EBE2-45E7-A14F-4E1C49C965F7}"
    AddSimpleElement doc, DataSubDel, "SubscriptionType", "Data"
    
    'and a subscription delete request for the UIEvent for getting the company close
    Dim UISubDel As IXMLDOMElement
    Set UISubDel = doc.createElement("SubscriptionDelRq")
    SubReq.appendChild UISubDel
    AddSimpleElement doc, UISubDel, "SubscriberID", "{2B6C9DB4-EBE2-45E7-A14F-4E1C49C965F7}"
    AddSimpleElement doc, UISubDel, "SubscriptionType", "UI"
        
    'Finally, send the subscription delete requests to QuickBooks
    Dim subXML As String
    subXML = "<?xml version=""1.0""?>" & vbCrLf & "<?qbxml version=""3.0""?>" & vbCrLf & doc.xml
    saveXMLStream subXML
    Dim resp As String
    resp = RP.ProcessSubscription(subXML)
    
    ' Show the response we got, note that if there was no subscription we'll get an
    ' error but we don't care...
    MsgBox Prompt:=resp, Title:="Subscriptions Removed"
    
    ' Finally close the connection to QuickBooks.
cleanup:
    If (Err.Number) Then
        MsgBox ("Unexpected error from QuickBooks: " & Err.Number & ": " & Err.Description)
    End If
    If (ConnOpen) Then
        RP.CloseConnection
    End If

End Sub

'
' A little utility routine for adding basic XML elements (i.e. <foo>Value</foo>) to
' a DOM hierarchy...
Private Sub AddSimpleElement(doc As DOMDocument40, parent As IXMLDOMElement, name As String, value As String)
    Dim newElem As IXMLDOMElement
    Set newElem = doc.createElement(name)
    newElem.Text = value
    parent.appendChild newElem
End Sub

'
' A little utility routine for saving XML data to disk for debugging using the
' qbXML validator...
'
Private Sub saveXMLStream(xml As String)
    Dim fname As String
    fname = App.Path & "\subReq.qbxml"
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    Set ts = fso.CreateTextFile(fname, True)
    ts.Write (xml)
End Sub
