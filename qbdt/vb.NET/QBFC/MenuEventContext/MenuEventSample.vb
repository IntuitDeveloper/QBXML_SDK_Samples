Option Strict Off
Option Explicit On
Friend Class MenuEventSample
	Inherits System.Windows.Forms.Form
	'
	' This application is very simple, it simply subscribes to a UIExtension event with
	' QuickBooks, creating a new menu item (with 5 cascade sub-menu items) in the File
	' menu of QuickBooks.
	'
	' In order to receive the events, this is an ActiveX VB project, and the project
	' properties are set to make the startup object be the Main subroutine in
	' MainMod  (though the application will get started automatically by QuickBooks if
	' necessary when our menu items are used from QuickBooks).  Sub Main handles ensuring
	' our form is loaded and displayed.
	'
	
	Private appName As String
	Private appGUID As String
	
	
	Private Sub AddMenu_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles AddMenu.Click
		Dim i As Object
		On Error GoTo handleError
		'
		' Note, we don't dim as new QBSessionManager because that causes VB to compile
		' every reference to sessMgr to include code to check if it is nothing and
		' to instantiate it dynamically if it is nothing -- this is incredibly inefficient
		'
		'UPGRADE_WARNING: Arrays in structure sessMgr may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim sessMgr As QBFC15Lib.QBSessionManager
		sessMgr = New QBFC15Lib.QBSessionManager
		
		'
		' For a subscription request we only need an OpenConnection, no session...
		sessMgr.OpenConnection("", appName)
		
		'
		' Create a Subscription request
		Dim subRq As QBFC15Lib.ISubscriptionMsgSetRequest
		subRq = sessMgr.CreateSubscriptionMsgSetRequest(4, 0)
		
		'
		' Add a UIExtension subscription to our request
		Dim subAdd As QBFC15Lib.IUIExtensionSubscriptionAdd
		subAdd = subRq.AppendUIExtensionSubscriptionAddRq
		
		'
		' set up the subscription request with the required information, we're adding to
		' the file menu in this case, and just for fun, we're making it a cascading menu
		subAdd.SubscriberID.SetValue(appGUID)
		subAdd.COMCallbackInfo.appName.SetValue(appName)
        subAdd.COMCallbackInfo.ORProgCLSID.ProgID.SetValue("MenuEventContext.QBMenuListenerClass")
        subAdd.MenuExtensionSubscription.AddToMenu.SetValue(QBFC15Lib.ENAddToMenu.atmFile)
		
		'
		' For the cascade fun, we're just going to add items to the cascade menu...
		Dim subMenu As QBFC15Lib.IMenuItem
		For i = 1 To 5
			subMenu = subAdd.MenuExtensionSubscription.ORMenuSubmenu.subMenu.MenuItemList.Append
			'
			' this is the text that the user will see in QuickBooks:
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			subMenu.MenuText.SetValue("Sub Item " & i)
			'
			' this is the tag we'll get in our event handler to know which menu item was
			' selected:
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			subMenu.EventTag.SetValue("SubMenu" & i)
		Next i
		
		'
		' Send the request and get the response, since we're sending only one request there
		' will be only one response in the response list
		Dim subRs As QBFC15Lib.ISubscriptionMsgSetResponse
		subRs = sessMgr.DoSubscriptionRequests(subRq)
		Dim resp As QBFC15Lib.IResponse
		
		'
		' Check the response and display an appropriate message to the user.
		resp = subRs.ResponseList.GetAt(0)
		If (resp.StatusCode = 0) Then
			MsgBox("Successfully added to QuickBooks File menu, restart QuickBooks to see results")
		Else
			MsgBox("Could not add to QuickBooks menu: " & resp.StatusMessage)
		End If
		sessMgr.CloseConnection()
		'UPGRADE_NOTE: Object sessMgr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		sessMgr = Nothing
		Exit Sub
handleError:
        sessMgr.CloseConnection()
        MsgBox("Encountered error subscribing: " & Err.Description)
	End Sub
	
	Private Sub MenuEventSample_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		appName = "IDN Menu Event Context Sample"
		appGUID = "{C0082D6F-0D97-44a8-98F4-5153E8805E44}"
	End Sub
	
	Private Sub RemoveMenu_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles RemoveMenu.Click
		' we don't get handed a session manager, so we need to set up the session
		' with QuickBooks
		'UPGRADE_WARNING: Arrays in structure sessMgr may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim sessMgr As QBFC15Lib.QBSessionManager
		sessMgr = New QBFC15Lib.QBSessionManager
		' Again, we're dealing with subscriptions, which are independent of the company
		' so there is no need to BeginSession, just open the connection.
		sessMgr.OpenConnection("", appName)
		
		' Set up the SubscriptionDel request
		Dim submsg As QBFC15Lib.ISubscriptionMsgSetRequest
		submsg = sessMgr.CreateSubscriptionMsgSetRequest(4, 0)
		Dim uiextend As QBFC15Lib.ISubscriptionDel
		uiextend = submsg.AppendSubscriptionDelRq
		uiextend.SubscriberID.SetValue(appGUID)
		uiextend.SubscriptionType.SetValue(QBFC15Lib.ENSubscriptionType.stUIExtension)
		
		' Send the request
		Dim subresp As QBFC15Lib.ISubscriptionMsgSetResponse
		subresp = sessMgr.DoSubscriptionRequests(submsg)
		Dim resp As QBFC15Lib.IResponse
		resp = subresp.ResponseList.GetAt(0)
		
		' Check the result and display an appropriate message to the user
		If (resp.StatusCode = 0) Then
			MsgBox("Successfully removed from QuickBooks File menu, restart QuickBooks to see results")
		Else
			MsgBox("Could not remove from QuickBooks menu: " & resp.StatusMessage)
		End If
		
		' Close the connection with QuickBooks, we didn't Begin a session so there is
		' no need to EndSession.
		sessMgr.CloseConnection()
		'UPGRADE_NOTE: Object sessMgr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		sessMgr = Nothing
	End Sub
End Class