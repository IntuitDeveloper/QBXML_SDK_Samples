VERSION 5.00
Begin VB.Form MenuEventSample 
   Caption         =   "Menu Event Sample"
   ClientHeight    =   4140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   6405
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox EventData 
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "MenuEventSample.frx":0000
      Top             =   1800
      Width           =   5535
   End
   Begin VB.CommandButton RemoveMenu 
      Caption         =   "Remove Menu From QuickBooks"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2775
   End
   Begin VB.CommandButton AddMenu 
      Caption         =   "Add Menu To QuickBooks"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Menu Event Data"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   3015
   End
End
Attribute VB_Name = "MenuEventSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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


Private Sub AddMenu_Click()
    On Error GoTo handleError
    '
    ' Note, we don't dim as new QBSessionManager because that causes VB to compile
    ' every reference to sessMgr to include code to check if it is nothing and
    ' to instantiate it dynamically if it is nothing -- this is incredibly inefficient
    '
    Dim sessMgr As QBSessionManager
    Set sessMgr = New QBSessionManager
    
    '
    ' For a subscription request we only need an OpenConnection, no session...
    sessMgr.OpenConnection "", appName
    
    '
    ' Create a Subscription request
    Dim subRq As ISubscriptionMsgSetRequest
    Set subRq = sessMgr.CreateSubscriptionMsgSetRequest(4, 0)
    
    '
    ' Add a UIExtension subscription to our request
    Dim subAdd As IUIExtensionSubscriptionAdd
    Set subAdd = subRq.AppendUIExtensionSubscriptionAddRq
    
    '
    ' set up the subscription request with the required information, we're adding to
    ' the file menu in this case, and just for fun, we're making it a cascading menu
    subAdd.SubscriberID.SetValue appGUID
    subAdd.COMCallbackInfo.appName.SetValue appName
    subAdd.COMCallbackInfo.ORProgCLSID.ProgId.SetValue "MenuEventContext.QBMenuListener"
    subAdd.MenuExtensionSubscription.AddToMenu.SetValue atmFile
    
    '
    ' For the cascade fun, we're just going to add items to the cascade menu...
    Dim subMenu As IMenuItem
    For i = 1 To 5
        Set subMenu = subAdd.MenuExtensionSubscription.ORMenuSubmenu.subMenu.MenuItemList.Append
        '
        ' this is the text that the user will see in QuickBooks:
        subMenu.MenuText.SetValue "Sub Item " & i
        '
        ' this is the tag we'll get in our event handler to know which menu item was
        ' selected:
        subMenu.EventTag.SetValue "SubMenu" & i
    Next i
    
    '
    ' Send the request and get the response, since we're sending only one request there
    ' will be only one response in the response list
    Dim subRs As ISubscriptionMsgSetResponse
    Set subRs = sessMgr.DoSubscriptionRequests(subRq)
    Dim resp As IResponse
    
    '
    ' Check the response and display an appropriate message to the user.
    Set resp = subRs.ResponseList.GetAt(0)
    If (resp.StatusCode = 0) Then
        MsgBox "Successfully added to QuickBooks File menu, restart QuickBooks to see results"
    Else
        MsgBox "Could not add to QuickBooks menu: " & resp.StatusMessage
    End If
    sessMgr.CloseConnection
    Set sessMgr = Nothing
    Exit Sub
handleError:
    sessMgr.CloseConnection
    MsgBox "Encountered error subscribing: " & Err.Description
End Sub

Private Sub Form_Load()
    appName = "IDN Menu Event Context Sample"
    appGUID = "{C0082D6F-0D97-44a8-98F4-5153E8805E44}"
End Sub

Private Sub RemoveMenu_Click()
    ' we don't get handed a session manager, so we need to set up the session
    ' with QuickBooks
    Dim sessMgr As QBSessionManager
    Set sessMgr = New QBSessionManager
    ' Again, we're dealing with subscriptions, which are independent of the company
    ' so there is no need to BeginSession, just open the connection.
    sessMgr.OpenConnection "", appName
    
    ' Set up the SubscriptionDel request
    Dim submsg As ISubscriptionMsgSetRequest
    Set submsg = sessMgr.CreateSubscriptionMsgSetRequest(4, 0)
    Dim uiextend As ISubscriptionDel
    Set uiextend = submsg.AppendSubscriptionDelRq
    uiextend.SubscriberID.SetValue appGUID
    uiextend.SubscriptionType.SetValue stUIExtension
    
    ' Send the request
    Dim subresp As ISubscriptionMsgSetResponse
    Set subresp = sessMgr.DoSubscriptionRequests(submsg)
    Dim resp As IResponse
    Set resp = subresp.ResponseList.GetAt(0)
    
    ' Check the result and display an appropriate message to the user
    If (resp.StatusCode = 0) Then
        MsgBox "Successfully removed from QuickBooks File menu, restart QuickBooks to see results"
    Else
        MsgBox "Could not remove from QuickBooks menu: " & resp.StatusMessage
    End If
    
    ' Close the connection with QuickBooks, we didn't Begin a session so there is
    ' no need to EndSession.
    sessMgr.CloseConnection
    Set sessMgr = Nothing
End Sub
