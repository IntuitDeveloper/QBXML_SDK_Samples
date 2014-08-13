VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QBEventQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Note!  The ProgID supplied in the subscription request will be converted to
' a CLSID internally by QuickBooks so it ' is very important that you set the
' properties of this project implementing the callback class to "binary
' compatibility" mode on the "Components" tab of the project properties dialog,
' otherwise VB will be foolish enough to change the CLSID with each recompile...
'
'
' this is just a trivial little Queue class which uses a Collection to
' manage a queue.  We Add to the end to Enqueue, grab item 1 and remove it from
' the collection to DeQueue.
'
Private EventQueue As Collection

'
' The alert reader will note that in our EnQueue and DeQueue routines we blindly
' use the Collection.  This is because this class is entirely enveloped by the
' QBDataEventManagerDisplay and so we can count on it having been initialized before
' it is used.  Part of the reason for this approach is that EnQueue and DeQueue are
' called from the event handler and we want it to be as efficient as possible, so we
' don't want the EventQueue to be an auto-instancing variable (which VB checks if it
' needs to create "behind the scenes" on each use of the variable and is therefore
' inefficient) and we don't want EnQueue and DeQueue to have to check if it needs to
' create the Collection each time they are called for the same reason!
'
' This Init function therefore creates the collection if we need it.
Public Sub Init()
    If (EventQueue Is Nothing) Then
        Set EventQueue = New Collection
    End If
End Sub

Public Sub EnQueue(eventXML As String)
    EventQueue.Add eventXML
End Sub

'
' need a way to put priority events to the front of the queue
'
Public Sub PriorityEnQueue(eventXML As String)
    If (EventQueue.Count = 0) Then
        EventQueue.Add eventXML
    Else
        EventQueue.Add eventXML, , 1
    End If
End Sub
'
' DeQueue is intended to be bulletproof, just return an empty string
' if there's nothing on the queue.
'
Public Function DeQueue() As String
    If (EventQueue.Count > 0) Then
        DeQueue = EventQueue.Item(1)
        EventQueue.Remove (1)
    Else
        DeQueue = ""
    End If
End Function

