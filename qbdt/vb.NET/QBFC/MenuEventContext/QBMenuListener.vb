Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("QBMenuListener_NET.QBMenuListener")> Public Class QBMenuListener
	Implements QBSDKEVENTLib.IQBEventCallback
	'
	' This is our little COM service class to handle the callback from
	' QuickBooks if our menu item is selected from the QuickBooks File menu.
	'
	
	' Make sure our GUI is loaded when the class is initialized.
	'
	
	Private Sub Class_Initialize_Renamed()
        'Load(MenuEventSample)
        Dim Form As New MenuEventSample
        Form.ShowDialog()
    End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	' If the class is terminated then QuickBooks has shut down and the user has
	' dismissed us interactively, so hide our UI and unload it.
	
	Private Sub Class_Terminate_Renamed()
		MenuEventSample.Hide()
		MenuEventSample.Close()
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	'
	' finally, implement the QuickBooks callback.  We'll crack the
	'
	Public Sub IQBEventCallback_inform(ByVal eventXML As String) Implements QBSDKEVENTLib.IQBEventCallback.inform
		Dim dispalyStr As Object
		
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
								
								displayStr = dispalyStr & " - TxnID: " & .TxnTypeID.TxnID.GetValue
							End If
						End If
					End With
				End If
			End With
			MenuEventSample.EventData.Text = displayStr & vbCrLf
		End If
		MenuEventSample.Show()
	End Sub
End Class