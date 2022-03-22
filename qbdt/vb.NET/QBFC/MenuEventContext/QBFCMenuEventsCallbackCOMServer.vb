
Imports System.Threading

Friend NotInheritable Class QBFCMenuEventsCallbackCOMServer

    Private Sub New()
    End Sub

    Private Shared _instance As QBFCMenuEventsCallbackCOMServer = New QBFCMenuEventsCallbackCOMServer
    Public Shared ReadOnly Property Instance() As QBFCMenuEventsCallbackCOMServer
        Get
            Return QBFCMenuEventsCallbackCOMServer._instance
        End Get
    End Property


    Private syncRoot As Object = New Object
    Private _bRunning As Boolean = False


    Private threadId As UInt32 = 0

    Private lockCount As Integer = 0

    Private timer As Timer

    Private Shared Sub GarbageCollect(ByVal stateInfo As Object)
        GC.Collect()
    End Sub

    Private msg As UInt32

    Private Sub PreMessageLoop()


        Dim clsid As New Guid(QBMenuListenerClass.ClassId)


        Dim hResult As Integer = ComInit.CoRegisterClassObject(
        clsid, New QBFCMenuEventsCallbackClassFactory, CLSCTX.LOCAL_SERVER,
        REGCLS.SUSPENDED Or REGCLS.MULTIPLEUSE, Me.msg)
        If (hResult <> 0) Then
            Throw New ApplicationException(
            "CoRegisterClassObject failed w/err 0x" & hResult.ToString("X"))
        End If
        hResult = ComInit.CoResumeClassObjects
        If (hResult <> 0) Then

            If (Me.msg <> 0) Then
                ComInit.CoRevokeClassObject(Me.msg)
            End If

            Throw New ApplicationException(
            "CoResumeClassObjects failed w/err 0x" & hResult.ToString("X"))
        End If


        Me.threadId = NativeHelper.GetCurrentThreadId

        Me.lockCount = 0

        Me.timer = New Timer(
        New TimerCallback(AddressOf QBFCMenuEventsCallbackCOMServer.GarbageCollect), Nothing,
        6000, 6000)

    End Sub

    Private Sub RunMessageLoop()
        Dim msg As MSG
        Do While NativeHelper.GetMessage(msg, IntPtr.Zero, 0, 0)
            NativeHelper.TranslateMessage((msg))
            NativeHelper.DispatchMessage((msg))
        Loop
    End Sub

    Private Sub PostMessageLoop()

        If (Me.msg <> 0) Then
            ComInit.CoRevokeClassObject(Me.msg)
        End If

        If (Not Me.timer Is Nothing) Then
            Me.timer.Dispose()
        End If

        ' Wait for any threads to finish.
        Thread.Sleep(1000)

    End Sub

    Public Sub Run()
        SyncLock Me.syncRoot
            If Me._bRunning Then
                Return
            End If

            Me._bRunning = True
        End SyncLock

        Try

            Me.PreMessageLoop()

            Me.RunMessageLoop()

            Me.PostMessageLoop()
        Finally
            Me._bRunning = False
        End Try
    End Sub


    Public Function Lock() As Integer
        Return Interlocked.Increment(Me.lockCount)
    End Function


    Public Function Unlock() As Integer
        Dim nRet As Integer = Interlocked.Decrement(Me.lockCount)

        If (nRet = 0) Then

            NativeHelper.PostThreadMessage(
            threadId, NativeHelper.WM_QUIT, UIntPtr.Zero, IntPtr.Zero)
        End If
        Return nRet
    End Function


    Public Function GetLockCount() As Integer
        Return Me.lockCount
    End Function

End Class