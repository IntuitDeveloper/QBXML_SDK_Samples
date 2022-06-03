Option Strict Off
Option Explicit On
Friend Class QBDataEventManagerDisplay
    Inherits System.Windows.Forms.Form
    Private m_EventQueue As QBEventQueue
    Private m_BlockEvent As Collection
    Private m_Tracking As Boolean

    ReadOnly Property EventQueue() As QBEventQueue
        Get
            If (m_EventQueue Is Nothing) Then
                m_EventQueue = New QBEventQueue
                m_EventQueue.Init()
            End If
            EventQueue = m_EventQueue
        End Get
    End Property


    Property Tracking() As Boolean
        Get
            Tracking = m_Tracking
        End Get
        Set(ByVal Value As Boolean)
            m_Tracking = Value
        End Set
    End Property

    Public Function BlockEvent(ByRef ListID As String) As Boolean
        If (m_BlockEvent Is Nothing) Then
            m_BlockEvent = New Collection
        End If
        On Error Resume Next
        Dim tmp As Object
        tmp = m_BlockEvent.Item(ListID)
        BlockEvent = (Err.Number <> 5)
    End Function

    Public Sub AddBlock(ByRef ListID As String)
        If (m_BlockEvent Is Nothing) Then
            m_BlockEvent = New Collection
        End If
        m_BlockEvent.Add(ListID, ListID)
    End Sub

    Public Sub RemoveBlock(ByRef ListID As String)
        On Error Resume Next
        If (Not m_BlockEvent Is Nothing) Then
            m_BlockEvent.Remove(ListID)
        End If
    End Sub

End Class