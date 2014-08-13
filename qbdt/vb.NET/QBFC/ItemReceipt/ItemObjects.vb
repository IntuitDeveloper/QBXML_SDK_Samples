Public Class VendorItem
    Inherits Object

    Private myName As String
    Private myListID As String

    ReadOnly Property Name() As String
        Get
            Name = myName
        End Get
    End Property

    ReadOnly Property ListID() As String
        Get
            ListID = myListID
        End Get
    End Property

    Public Overrides Function ToString() As String
        ToString = myName
    End Function

    Public Sub New(ByVal ID As String, ByVal Nm As String)
        myListID = ID
        myName = Nm
    End Sub
End Class

Public Class POItem
    Inherits Object

    Private myTxnID As String
    Private myTxnDate As String
    Private myTxnNumber As String
    Private myRefNumber As String

    Public Sub New(ByVal ID As String, ByVal txnDate As String, ByVal refNum As String, ByVal txnNum As String)
        myTxnID = ID
        myTxnDate = txnDate
        myRefNumber = refNum
        myTxnNumber = txnNum
    End Sub

    Public Overrides Function ToString() As String
        ToString = myTxnDate & " : " & myRefNumber & " : " & myTxnNumber
    End Function

    ReadOnly Property TxnID() As String
        Get
            TxnID = myTxnID
        End Get
    End Property
End Class

Public Class POLineItem
    Inherits Object

    Private myTxnID As String
    Private myTxnLineID As String
    Private myNewRcv As Integer
    Private myDirty As Boolean

    Public Sub New(ByVal ID As String, ByVal lineID As String)
        myTxnID = ID
        myTxnLineID = lineID
        myNewRcv = 0
        myDirty = False
    End Sub

    Public Overrides Function ToString() As String
        ToString = myNewRcv
    End Function

    ReadOnly Property TxnID() As String
        Get
            TxnID = myTxnID
        End Get
    End Property

    ReadOnly Property TxnLineID() As String
        Get
            TxnLineID = myTxnLineID
        End Get
    End Property

    ReadOnly Property IsDirty() As Boolean
        Get
            IsDirty = myDirty
        End Get
    End Property

    Property Received() As Integer
        Get
            Received = myNewRcv
        End Get
        Set(ByVal Value As Integer)
            myNewRcv = Value
            myDirty = True
        End Set
    End Property
End Class