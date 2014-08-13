VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QBEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Note!  The ProgID supplied in the subscription request will be converted to
' a CLSID internally by QuickBooks so it ' is very important that you set the
' properties of this project implementing the callback class to "binary
' compatibility" mode on the "Components" tab of the project properties dialog,
' otherwise VB will be foolish enough to change the CLSID with each recompile...
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
' Strictly speaking, we could probably do this in 2 apps (one to subscribe and one to
' do everything else) but to ensure that we process events from QuickBooks as quickly
' as possible and our main UI does not interfere with our event processing we divide it
' up, with the event manager here simply acting as a "broker" to the main application,
' serving events to it as requested.
'
' The QBDataEventManager is made up of one primary class that provides a COM interface:
' QBEventHandler -- Receives events from QuickBooks
'
' And one private utility class for managing a Queue of events
' QBEventQueue   -- Stores events for processing by the QBDataEventApp
'
' And one Form, just to show events as they are received by the handler and to
' serve as a kind of "static class variable" holder for variables we want to share
' between instances of the QBEventHandler COM class.  (See below for more info).
' QBDataEventManagerDisplay -- displays and counts events as they come in.
'

'
' We need to make sure we can bring our window to the front the only way to do this
' is via the Win32 function since the VB zOrder property is only the zOrder within our
' application, not all windows.
'
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private EventCounter As Integer

'
' One thing worth noting, the member variable for holding the EventQueue is a member
' of our FORM class (QBDataEventManagerDisplay) because this is a COM class which has
' two instances, the one used by QuickBooks (to call IQBEventCallback::inform) and the
' one used by QBDataEventApp.  If we were to make the QBEventQueue variable a part of
' this handler class then there would be two EventQueues, one belonging to the COM
' instance held by QuickBooks (which would keep having events added to it) and one
' belonging to the COM instance held by the DataEventApp (which would never have
' events added to it.  Since there is no way in Visual Basic that I could find to
' declare a CLASS variable I punted and took advantage of the fact that there is
' never more than one instance of a FORM in a MultiUse COM class.  So the FORM has
' a member variable for the EventQueue.
'
' For the same reason, the Boolean (Tracking) we use to determine whether to
' Queue an event or not is stored in the FORM as well.

'
' This is the routine called by QuickBooks, implementing the IQBEventCallBack interface
'
Implements QBSDKEVENTLib.IQBEventCallback
Public Sub IQBEventCallback_inform(ByVal eventXML As String)
' We treat this basically like interrupt handling in an OS, we want to process the
' interrupt as quickly as possible and take care of the details later if we can.
' hence we do not attempt to parse the XML here, just save it in a queue
' (if the DataEventApp has asked for it) for the DataEventApp to handle later...
'
    
    ' Display the event we got
    Dim tmpXML As String
    tmpXML = eventXML
    EventCounter = EventCounter + 1
    QBDataEventManagerDisplay.eventXML.Text = Replace(tmpXML, vbLf, vbCrLf, 1, -1, vbTextCompare)
    QBDataEventManagerDisplay.eventLabel.Caption = "Received Event #" & EventCounter
    QBDataEventManagerDisplay.Show
    SetForegroundWindow QBDataEventManagerDisplay.hwnd
    
    'Now check if it is a company close event or a data event
    If (InStr(1, eventXML, "CompanyFileEventOperation>Close<", vbTextCompare) > 0) Then
        'close is a priority event
        QBDataEventManagerDisplay.Debug.Text = "!!!COMPANY CLOSE!!!"
        QBDataEventManagerDisplay.EventQueue.PriorityEnQueue eventXML
    Else
        'regardless of the event, queue it up if we are supposed to be tracking events.
        'The EventApp will need to get the close event to know that it needs to shut
        'itself down. We simply Enqueue to the EventQueue held by the form.
        If (QBDataEventManagerDisplay.Tracking) Then
            Dim listIDpos As Integer
            listIDpos = InStr(1, eventXML, "<ListID>")
            Dim shouldQueue As Boolean
            shouldQueue = True
            If (listIDpos > 0) Then
                listIDpos = listIDpos + 8 ' pass the <ListID>
                Dim closeListID As Integer
                closeListID = InStr(listIDpos, eventXML, "</ListID>")
                If (closeListID > 0) Then
                    Dim length As Integer
                    length = closeListID - listIDpos
                    Dim ListID As String
                    ListID = Mid$(eventXML, listIDpos, length)
                    If (QBDataEventManagerDisplay.BlockEvent(ListID)) Then
                        QBDataEventManagerDisplay.RemoveBlock ListID
                        QBDataEventManagerDisplay.Debug.Text = "Blocking event #" & EventCounter
                        shouldQueue = False
                    End If
                End If
            End If
            If (shouldQueue) Then
                QBDataEventManagerDisplay.Debug.Text = "Queing event #" & EventCounter
                QBDataEventManagerDisplay.EventQueue.EnQueue (eventXML)
            End If
        End If
    End If
End Sub

'
' This is the primary routine used by our main application to check if there
' are any events for it to process, we simply deQueue from the EventQueue held
' by the form
'
Public Function GetEvent() As String
    GetEvent = QBDataEventManagerDisplay.EventQueue.DeQueue
End Function

' If we implemented the main UI in this app we could dispense with these
' and just set the DeliveryPolicy in our subscriber to only get events when
' our application is running.  Since our main app and our event handler are
' separate objects, we implement these routines to let the main app decide
' whether events should be dropped on the floor or queued.
Public Sub StartEventTracking()
    'Simply set our Tracking Boolean to True, the inform routine handles
    'deciding whether to Queue or not based on this variable
    QBDataEventManagerDisplay.Tracking = True
    QBDataEventManagerDisplay.Debug.Text = "Tracking resumed"
End Sub

Public Sub StopEventTracking()
    'Simply set our tracking Boolean to False, the inform routine handles
    'deciding whether to Queue or not based on this variable
    QBDataEventManagerDisplay.Tracking = False
    QBDataEventManagerDisplay.Debug.Text = "Tracking stopped"
End Sub

'
' Since we really don't want to process our own events, we specify an ID we are
' about to modify (if we were adding objects and wanted to filter our events we'd
' do something similar, basically holding onto events until we get the response from
' our add, then filtering the "on hold" events by the ID we get in the response.
'
' Our app only needs to filter its own modifications, so we already know the ID
' to filter.
'
' It should be noted that it *can* happen that we filter a mod event from some other
' app, but that's OK, because it means the mod we were filtering for will fail (the
' other app "won" and so our <EditSequence> is out of date) and we'll have to re-try
' the mod after doing a refresh Query to update our EditSequence.
Public Sub AddFilter(ID As String)
    If (QBDataEventManagerDisplay.Tracking) Then
        QBDataEventManagerDisplay.AddBlock ID
    End If
End Sub

Public Sub Shutdown()
    QBDataEventManagerDisplay.Hide
    Unload QBDataEventManagerDisplay
End Sub

