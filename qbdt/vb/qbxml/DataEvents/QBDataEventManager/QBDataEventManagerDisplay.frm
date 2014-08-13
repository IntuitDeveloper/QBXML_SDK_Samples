VERSION 5.00
Begin VB.Form QBDataEventManagerDisplay 
   Caption         =   "Form1"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   ScaleHeight     =   7335
   ScaleWidth      =   8910
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Debug 
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   6615
   End
   Begin VB.TextBox eventXML 
      CausesValidation=   0   'False
      Height          =   6615
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "QBDataEventManagerDisplay.frx":0000
      Top             =   480
      Width           =   8415
   End
   Begin VB.Label eventLabel 
      Caption         =   "Received Event #"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "QBDataEventManagerDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Note!  The ProgID supplied in the subscription request will be converted to
' a CLSID internally by QuickBooks so it ' is very important that you set the
' properties of this project implementing the callback class to "binary
' compatibility" mode on the "Components" tab of the project properties dialog,
' otherwise VB will be foolish enough to change the CLSID with each recompile...
'
'
' since there is no user-interactive part of this form (it is just used to
' display information to the user) there is nothing here for the form.  Just
' stuff to support variables that we want to have only one instance of for all
' instances of the COM class.
'
' This one we do right, with private variables and Let/Get property methods because
' we need to initialize the EventQueue properly the first time it is accessed.
'
Private m_EventQueue As QBEventQueue
Private m_BlockEvent As Collection
Private m_Tracking As Boolean

Property Get EventQueue() As QBEventQueue
    If (m_EventQueue Is Nothing) Then
        Set m_EventQueue = New QBEventQueue
        m_EventQueue.Init
    End If
    Set EventQueue = m_EventQueue
End Property

Property Get Tracking() As Boolean
    Tracking = m_Tracking
End Property

Property Let Tracking(ByVal newVal As Boolean)
    m_Tracking = newVal
End Property

Public Function BlockEvent(ListID As String) As Boolean
    If (m_BlockEvent Is Nothing) Then
        Set m_BlockEvent = New Collection
    End If
    On Error Resume Next
    Dim tmp As Variant
    Set tmp = m_BlockEvent.Item(ListID)
    BlockEvent = (Err <> 5)
End Function

Public Sub AddBlock(ListID As String)
    If (m_BlockEvent Is Nothing) Then
        Set m_BlockEvent = New Collection
    End If
    m_BlockEvent.Add ListID, ListID
End Sub

Public Sub RemoveBlock(ListID As String)
    On Error Resume Next
    If (Not m_BlockEvent Is Nothing) Then
        m_BlockEvent.Remove ListID
    End If
End Sub
