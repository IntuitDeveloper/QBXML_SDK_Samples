Option Strict Off
Option Explicit On
Imports System.ComponentModel
Imports System.Runtime.InteropServices

<ComClass(QBMenuListenerClass.ClassId, QBMenuListenerClass.InterfaceId,
          QBMenuListenerClass.EventsId), ComVisible(True)>
Public Class QBMenuListenerClass
    Inherits ReferenceCountedObject
    Implements QBSDKEVENTLib.IQBEventCallback
    Public Const ClassId As String = "EBCBCD28-F409-4D8B-B3B0-D48D8F7325F1"
    Public Const InterfaceId As String = "FA9D8CC5-4ED0-4AEC-91CD-53D5F4B98830"
    Public Const EventsId As String = "350E19DC-718D-45FF-98D9-B7407CC3B6D8"
    '
    ' This is our little COM service class to handle the callback from
    ' QuickBooks if our menu item is selected from the QuickBooks File menu.
    '
    '
    ' finally, implement the QuickBooks callback.  We'll crack the
    '
    Public Sub IQBEventCallback_inform(ByVal eventXML As String) Implements QBSDKEVENTLib.IQBEventCallback.inform
        
        Dim sessMgr As QBFC15Lib.QBSessionManager
        sessMgr = New QBFC15Lib.QBSessionManager
        Dim eventSet As QBFC15Lib.IEventsMsgSet
        eventSet = sessMgr.ToEventsMsgSet(eventXML, 4, 0)
        ' UIExtensionEvent should be the only type we get, since that's all we subscribe to
        Dim displayStr As String
        If (Not eventSet.OREvent.UIExtensionEvent Is Nothing) Then
            With eventSet.OREvent.UIExtensionEvent
                displayStr = "Company File: " & .CompanyFilePath.GetValue & vbCrLf
                displayStr = displayStr & "Menu Tag: " & .EventTag.GetValue & vbCrLf
                displayStr = displayStr & "QB Info: " & .HostInfo.ProductName.GetValue & vbCrLf
                If (Not .CurrentWindow Is Nothing) Then
                    displayStr = displayStr & "Current Window: "
                    With .CurrentWindow.ORCurrentWindow
                        If (Not .ListTypeID Is Nothing) Then
                            displayStr = displayStr & .ListTypeID.ListType.GetAsString
                            If (Not .ListTypeID.ListID Is Nothing) Then
                                displayStr = displayStr & " - ListID: " & .ListTypeID.ListID.GetValue
                            End If
                        End If
                        If (Not .TxnTypeID Is Nothing) Then
                            displayStr = displayStr & .TxnTypeID.TxnType.GetAsString
                            If (Not .TxnTypeID.TxnID Is Nothing) Then
                                
                                displayStr = displayStr & " - TxnID: " & .TxnTypeID.TxnID.GetValue
                            End If
                        End If
                    End With
                End If
            End With
            If frm.InvokeRequired Then
                frm.Invoke(Sub()
                               frm.EventData.Text = displayStr & vbCrLf
                           End Sub)
            Else
                frm.EventData.Text = displayStr & vbCrLf
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
End Class