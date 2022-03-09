VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QBMenuListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
' This is our little COM service class to handle the callback from
' QuickBooks if our menu item is selected from the QuickBooks File menu.
'
Implements QBSDKEVENTLib.IQBEventCallback

' Make sure our GUI is loaded when the class is initialized.
'
Private Sub Class_Initialize()
    Load MenuEventSample
End Sub

' If the class is terminated then QuickBooks has shut down and the user has
' dismissed us interactively, so hide our UI and unload it.
Private Sub Class_Terminate()
    MenuEventSample.Hide
    Unload MenuEventSample
End Sub

'
' finally, implement the QuickBooks callback.  We'll crack the
'
Public Sub IQBEventCallback_inform(ByVal eventXML As String)
    Dim sessMgr As QBSessionManager
    Set sessMgr = New QBSessionManager
    Dim eventSet As IEventsMsgSet
    Set eventSet = sessMgr.ToEventsMsgSet(eventXML, 4, 0)
    ' UIExtensionEvent should be the only type we get, since that's all we subscribe to
    If (Not eventSet.OREvent.UIExtensionEvent Is Nothing) Then
        With eventSet.OREvent.UIExtensionEvent
            Dim displayStr As String
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
                            displayStr = dispalyStr & " - TxnID: " & .TxnTypeID.TxnID.GetValue
                        End If
                    End If
                End With
            End If
        End With
        MenuEventSample.EventData.Text = displayStr & vbCrLf
    End If
    MenuEventSample.Show
End Sub

