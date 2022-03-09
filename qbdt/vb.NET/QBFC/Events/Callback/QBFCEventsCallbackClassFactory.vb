Imports System.Runtime.InteropServices

<ComImport(), ComVisible(False),
InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
Guid("00000001-0000-0000-C000-000000000046")>
Friend Interface IQBFCEventsCallbackClassFactory
    <PreserveSig()>
    Function CreateInstance(ByVal pUnkOuter As IntPtr, ByRef riid As Guid,
                            <Out()> ByRef ppvObject As IntPtr) As Integer

    <PreserveSig()>
    Function LockServer(ByVal fLock As Boolean) As Integer

End Interface

Friend Class QBFCEventsCallbackClassFactory
    Implements IQBFCEventsCallbackClassFactory

    Public Function CreateInstance(ByVal pUnkOuter As IntPtr, ByRef riid As Guid,
                                   <Out()> ByRef ppvObject As IntPtr) As Integer _
                                   Implements IQBFCEventsCallbackClassFactory.CreateInstance
        ppvObject = IntPtr.Zero

        If (pUnkOuter <> IntPtr.Zero) Then
            Marshal.ThrowExceptionForHR(ComInit.CLASS_E_NOAGGREGATION)
        End If

        If ((riid = New Guid(QBFCEventsCallbackClass.ClassId)) OrElse
            (riid = New Guid(ComInit.IID_IDispatch)) OrElse
            (riid = New Guid(ComInit.IID_IUnknown))) Then

            ppvObject = Marshal.GetComInterfaceForObject(
            New QBFCEventsCallbackClass, GetType(QBFCEventsCallbackClass).GetInterface("_QBFCEventsCallbackClass"))
        Else
            Marshal.ThrowExceptionForHR(ComInit.E_NOINTERFACE)
        End If

        Return 0
    End Function


    Public Function LockServer(ByVal fLock As Boolean) As Integer _
    Implements IQBFCEventsCallbackClassFactory.LockServer
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
        QBFCEventsCallbackCOMServer.Instance.Lock()
    End Sub

    Protected Overrides Sub Finalize()
        Try
            QBFCEventsCallbackCOMServer.Instance.Unlock()
        Finally
            MyBase.Finalize()
        End Try
    End Sub

End Class
