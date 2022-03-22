Imports System.Runtime.InteropServices

<ComImport(), ComVisible(False),
InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
Guid("00000001-0000-0000-C000-000000000046")>
Friend Interface IQBFCMenuEventsCallbackClassFactory
    <PreserveSig()>
    Function CreateInstance(ByVal pUnkOuter As IntPtr, ByRef riid As Guid,
                            <Out()> ByRef ppvObject As IntPtr) As Integer

    <PreserveSig()>
    Function LockServer(ByVal fLock As Boolean) As Integer

End Interface

Friend Class QBFCMenuEventsCallbackClassFactory
    Implements IQBFCMenuEventsCallbackClassFactory

    Public Function CreateInstance(ByVal pUnkOuter As IntPtr, ByRef riid As Guid,
                                   <Out()> ByRef ppvObject As IntPtr) As Integer _
                                   Implements IQBFCMenuEventsCallbackClassFactory.CreateInstance
        ppvObject = IntPtr.Zero

        If (pUnkOuter <> IntPtr.Zero) Then
            Marshal.ThrowExceptionForHR(ComInit.CLASS_E_NOAGGREGATION)
        End If

        If ((riid = New Guid(QBMenuListenerClass.ClassId)) OrElse
            (riid = New Guid(ComInit.IID_IDispatch)) OrElse
            (riid = New Guid(ComInit.IID_IUnknown))) Then

            ppvObject = Marshal.GetComInterfaceForObject(
            New QBMenuListenerClass, GetType(QBMenuListenerClass).GetInterface("_QBMenuListenerClass"))
        Else
            Marshal.ThrowExceptionForHR(ComInit.E_NOINTERFACE)
        End If

        Return 0
    End Function


    Public Function LockServer(ByVal fLock As Boolean) As Integer _
    Implements IQBFCMenuEventsCallbackClassFactory.LockServer
        Return 0
    End Function

End Class


''' <summary>
''' Reference counted object base.
''' </summary>
''' <remarks></remarks>
<ComVisible(False)>
Public Class ReferenceCountedObject

    Public Sub New()
        QBFCMenuEventsCallbackCOMServer.Instance.Lock()
    End Sub

    Protected Overrides Sub Finalize()
        Try
            QBFCMenuEventsCallbackCOMServer.Instance.Unlock()
        Finally
            MyBase.Finalize()
        End Try
    End Sub

End Class
