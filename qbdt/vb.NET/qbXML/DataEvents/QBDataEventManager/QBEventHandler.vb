Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports QBSDKEVENTLib
<ComClass(QBEventHandler.ClassId, QBEventHandler.InterfaceId,
          QBEventHandler.EventsId), ComVisible(True)>
Public Class QBEventHandler
    Inherits ReferenceCountedObject
    Implements QBSDKEVENTLib.IQBEventCallback
    Public Const ClassId As String = "A665E117-5CE3-4D8B-9E28-7F47B6D1F691"
    Public Const InterfaceId As String = "9B2D2CE9-AD77-4B8C-8B34-201129103E68"
    Public Const EventsId As String = "D6F3090D-C9D3-4EC5-AA54-3BDAE69E1D5A"

    Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Integer) As Integer
    Private EventCounter As Short

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
    Public Sub inform(eventXML As String) Implements IQBEventCallback.inform

        ' We treat this basically like interrupt handling in an OS, we want to process the
        ' interrupt as quickly as possible and take care of the details later if we can.
        ' hence we do not attempt to parse the XML here, just save it in a queue
        ' (if the DataEventApp has asked for it) for the DataEventApp to handle later...
        '

        ' Display the event we got

        Dim tmpXML As String
        tmpXML = eventXML
        EventCounter = EventCounter + 1
        If frm.InvokeRequired Then

            frm.Invoke(Sub()

                           frm.eventXML.Text = Replace(tmpXML, vbLf, vbCrLf, 1, -1, CompareMethod.Text)
                           frm.eventLabel.Text = "Received Event #" & EventCounter
                           frm.Show()
                           SetForegroundWindow(frm.Handle.ToInt32)

                           'Now check if it is a company close event or a data event
                           Dim listIDpos As Short
                           Dim shouldQueue As Boolean
                           Dim closeListID As Short
                           Dim length As Short
                           Dim ListID As String
                           If (InStr(1, eventXML, "CompanyFileEventOperation>Close<", CompareMethod.Text) > 0) Then
                               'close is a priority event
                               frm.Debug.Text = "!!!COMPANY CLOSE!!!"
                               frm.EventQueue.PriorityEnQueue(eventXML)
                           Else
                               'regardless of the event, queue it up if we are supposed to be tracking events.
                               'The EventApp will need to get the close event to know that it needs to shut
                               'itself down. We simply Enqueue to the EventQueue held by the form.
                               If (frm.Tracking) Then
                                   listIDpos = InStr(1, eventXML, "<ListID>")
                                   shouldQueue = True
                                   If (listIDpos > 0) Then
                                       listIDpos = listIDpos + 8 ' pass the <ListID>
                                       closeListID = InStr(listIDpos, eventXML, "</ListID>")
                                       If (closeListID > 0) Then
                                           length = closeListID - listIDpos
                                           ListID = Mid(eventXML, listIDpos, length)
                                           If (frm.BlockEvent(ListID)) Then
                                               frm.RemoveBlock(ListID)
                                               frm.Debug.Text = "Blocking event #" & EventCounter
                                               shouldQueue = False
                                           End If
                                       End If
                                   End If
                                   If (shouldQueue) Then
                                       frm.Debug.Text = "Queing event #" & EventCounter
                                       frm.EventQueue.EnQueue((eventXML))
                                   End If
                               End If
                           End If
                       End Sub)
        Else
            frm.eventXML.Text = Replace(tmpXML, vbLf, vbCrLf, 1, -1, CompareMethod.Text)
            frm.eventLabel.Text = "Received Event #" & EventCounter
            frm.Show()
            SetForegroundWindow(frm.Handle.ToInt32)

            'Now check if it is a company close event or a data event
            Dim listIDpos As Short
            Dim shouldQueue As Boolean
            Dim closeListID As Short
            Dim length As Short
            Dim ListID As String
            If (InStr(1, eventXML, "CompanyFileEventOperation>Close<", CompareMethod.Text) > 0) Then
                'close is a priority event
                frm.Debug.Text = "!!!COMPANY CLOSE!!!"
                frm.EventQueue.PriorityEnQueue(eventXML)
            Else
                'regardless of the event, queue it up if we are supposed to be tracking events.
                'The EventApp will need to get the close event to know that it needs to shut
                'itself down. We simply Enqueue to the EventQueue held by the form.
                If (frm.Tracking) Then
                    listIDpos = InStr(1, eventXML, "<ListID>")
                    shouldQueue = True
                    If (listIDpos > 0) Then
                        listIDpos = listIDpos + 8 ' pass the <ListID>
                        closeListID = InStr(listIDpos, eventXML, "</ListID>")
                        If (closeListID > 0) Then
                            length = closeListID - listIDpos
                            ListID = Mid(eventXML, listIDpos, length)
                            If (frm.BlockEvent(ListID)) Then
                                frm.RemoveBlock(ListID)
                                frm.Debug.Text = "Blocking event #" & EventCounter
                                shouldQueue = False
                            End If
                        End If
                    End If
                    If (shouldQueue) Then
                        frm.Debug.Text = "Queing event #" & EventCounter
                        frm.EventQueue.EnQueue((eventXML))
                    End If
                End If
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
    '
    ' This is the primary routine used by our main application to check if there
    ' are any events for it to process, we simply deQueue from the EventQueue held
    ' by the form
    '
    Public Function GetEvent() As String

        Dim eventStr As String
        eventStr = ""
        If frm.InvokeRequired Then
            frm.Invoke(Sub()
                           eventStr = frm.EventQueue.DeQueue
                           GetEvent = eventStr
                       End Sub)
        Else
            eventStr = frm.EventQueue.DeQueue
            GetEvent = eventStr
        End If
    End Function

    ' If we implemented the main UI in this app we could dispense with these
    ' and just set the DeliveryPolicy in our subscriber to only get events when
    ' our application is running.  Since our main app and our event handler are
    ' separate objects, we implement these routines to let the main app decide
    ' whether events should be dropped on the floor or queued.
    Public Sub StartEventTracking()
        'Simply set our Tracking Boolean to True, the inform routine handles
        'deciding whether to Queue or not based on this variable
        If frm.InvokeRequired Then
            frm.Invoke(Sub()
                           frm.Tracking = True
                           frm.Debug.Text = "Tracking resumed"
                       End Sub)
        Else
            frm.Tracking = True
            frm.Debug.Text = "Tracking resumed"
        End If
    End Sub

    Public Sub StopEventTracking()
        'Simply set our tracking Boolean to False, the inform routine handles
        'deciding whether to Queue or not based on this variable
        If frm.InvokeRequired Then
            frm.Invoke(Sub()
                           frm.Tracking = False
                           frm.Debug.Text = "Tracking stopped"
                       End Sub)
        Else
            frm.Tracking = False
            frm.Debug.Text = "Tracking stopped"
        End If

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
    Public Sub AddFilter(ByRef ID As String)
        Dim temp As String
        temp = ID
        If frm.InvokeRequired Then
            frm.Invoke(Sub()
                           If (frm.Tracking) Then
                               frm.AddBlock(temp)
                           End If
                       End Sub)
        Else
            If (frm.Tracking) Then
                frm.AddBlock(ID)
            End If
        End If
    End Sub

    Public Sub Shutdown()
        If frm.InvokeRequired Then
            frm.Invoke(Sub()
                           frm.Hide()
                           frm.Close()
                       End Sub)
        Else
            frm.Hide()
            frm.Close()
        End If
    End Sub
End Class

