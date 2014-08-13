VERSION 5.00
Begin VB.Form CommPrefsDlg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   5835
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   Begin VB.Frame W 
      Caption         =   "Communicate with QuickBooks"
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   4935
      Begin VB.OptionButton QBOE 
         Caption         =   "Online Edition"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   3135
      End
      Begin VB.OptionButton QBDT 
         Caption         =   "On My Desktop"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5520
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   5520
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame DTComm 
      Caption         =   "Setup Desktop Connection Preferences"
      Height          =   3975
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   6375
      Begin VB.Frame CompanyFileFrame 
         Caption         =   "Company File Info"
         Height          =   1455
         Left            =   240
         TabIndex        =   15
         Top             =   2400
         Width           =   5895
         Begin VB.TextBox CompanyFile 
            Height          =   375
            Left            =   120
            TabIndex        =   18
            Top             =   960
            Width           =   5655
         End
         Begin VB.OptionButton Interactive 
            Caption         =   "Use whatever file is open"
            Height          =   255
            Left            =   3000
            TabIndex        =   17
            Top             =   480
            Width           =   2415
         End
         Begin VB.OptionButton UseCurrentCompany 
            Caption         =   "Use this company file (QuickBooks can be closed):"
            Height          =   375
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Width           =   2415
         End
      End
      Begin VB.CommandButton QBConnect 
         Caption         =   "Connect To QuickBooks!"
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Top             =   1440
         Width           =   3375
      End
      Begin VB.Label Label4 
         Caption         =   $"CommPrefsDlg.frx":0000
         Height          =   855
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   6015
      End
   End
   Begin VB.Frame OEComm 
      Caption         =   "Set up Online Edition communications"
      Height          =   3975
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   6375
      Begin VB.TextBox ConnectionTicket 
         Height          =   375
         Left            =   2280
         TabIndex        =   12
         Top             =   3480
         Width           =   3855
      End
      Begin VB.CommandButton Connect 
         Caption         =   "Set up connection!"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   2880
         Width           =   2055
      End
      Begin VB.CommandButton GoToOE 
         Caption         =   "Let me set up a company!"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Enter connection ticket here:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   3600
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   $"CommPrefsDlg.frx":013C
         Height          =   975
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   6135
      End
      Begin VB.Label Label1 
         Caption         =   $"CommPrefsDlg.frx":0295
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   6015
      End
   End
End
Attribute VB_Name = "CommPrefsDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' In addition to being a dialog class, we're going to use this to hold everything we know about
' how to connect to QuickBooks and let the client class ask us for information.
'
' The intent here is for this Communication Preferences dialog to be *completely* reusable by any
' application.  The Client application simply provides the necessary information (see the Set/Get
' routines below) for QB Online Edition (application ID, application login, etc.) and QB Desktop
' editions and this class handles everything about setting up the connection and remembering the
' preferences for the next time the application is run, etc.
'
' The only time a client application will need to care about whether the connection is with QBOE or
' QBDT is if the application is using functionality that differs between OE or DT, in which case it
' can always check CommPrefsDlg.UseOnlineEdition to determine if it is communicating with QBOE or not.
'

Option Explicit

'
' Constants for communication setup with QuickBooks Online Edition
'
Const QBOELoginServer = "https://login.quickbooks.com"
Const QBOEConnectionURL = "/j/qbn/sdkapp/confirm?serviceid=2004"
Const QBOELoginURL = "/j/qbn/sdkapp/sessionauth2?serviceid=2004"
Const QBOEAuthURL = "/j/qbn/sdkapp/connauth?serviceid=2004"

'
' Private class variables
'
Private myAppName As String 'Used to identify our key in the registry for remembering preferences

'
' Declarations to support building a GUID (see GetGUID at the bottom of this file)
'
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
Private Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long

'
' Public properties
'
Public isOK As Boolean 'If "OK" was clicked last time the dialog was displayed.

'
' Load the form and display the correct group box for the online edition or the desktop
'
Private Sub Form_Load()
    If (Me.UseOnlineEdition(myAppName)) Then
        QBOE.Value = True
    Else
        QBDT.Value = True
    End If
    ConnectionTicket.Text = Me.GetConnTkt(myAppName)
    CompanyFile.Text = Me.GetDTCompanyFile(myAppName)
    If ("" <> CompanyFile.Text) Then
        UseCurrentCompany.Value = True
    Else
        Interactive.Value = True
    End If
End Sub
Private Sub CancelButton_Click()
    isOK = False
    CommPrefsDlg.Hide
End Sub

'
' The "Connect to QuickBooks Online" button was clicked, open a browser to QBOE with a Query string
' indicating the desire to set up a connection between the application and QBOE
'
Private Sub Connect_Click()
    Dim appID As String
    appID = Me.GetAppID(myAppName)
    If ("" = appID) Then
        MsgBox "Must call CommPrefsDlg.SetAppID before displaying the CommPrefsDlg", vbCritical, "Error"
    Else
        showURL (QBOELoginServer & QBOEConnectionURL & "&appid=" & appID)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload SessionTicketDlg
End Sub

'
' The "Let me set up a company" button was clicked.
'
Private Sub GoToOE_Click()
    showURL ("http://oe.quickbooks.com")
End Sub

Private Sub Interactive_Click()
    CompanyFile.Text = ""
End Sub

'
' The OK button was clicked, verify that we have the information we need and save the
' communication preferences to the registry if we do.  If we do not, show an error and
' don't close the dialog.
'
Private Sub OKButton_Click()
    Dim Good As Boolean 'Do we have all the information we need
    Good = True         'assume we do.
    If (QBOE.Value) Then
        SaveSetting myAppName, "CommunicationPreferences", "UseOE", "Yes"
        If ("" = ConnectionTicket.Text) Then
            MsgBox "You must specify the connection ticket when communicating with QuickBooks Online Edition", vbExclamation
            Good = False
        End If
        SaveSetting myAppName, "CommunicationPreferences", "ConnTkt", ConnectionTicket.Text
    Else
        SaveSetting myAppName, "CommunicationPreferences", "UseOE", "No"
        '
        ' If the user wants us to save the company file in use, do so.
        '
        If (UseCurrentCompany.Value) Then
            SaveSetting myAppName, "CommunicationPreferences", "CompanyFile", CompanyFile.Text
        Else
            CompanyFile.Text = ""
            SaveSetting myAppName, "CommunicationPreferences", "CompanyFile", ""
        End If
    End If
    isOK = True
    If (Good) Then
        CommPrefsDlg.Hide
    End If
End Sub

'
' The Connect button was clicked in the desktop preferences group,
' open a session long enough to find out what the open company file is so
' we can display it (and save it for later use).
'
Private Sub QBConnect_Click()
    On Error GoTo ErrHandler
    '
    ' Set up a traditional desktop session manager
    '
    Dim SessionManager As New QBSessionManager
    SessionManager.OpenConnection "", myAppName
    SessionManager.BeginSession "", omDontCare
    
    '
    ' Get the company file name and display it and set the default to use it
    ' again later.
    '
    Dim CompanyFileName As String
    CompanyFileName = SessionManager.GetCurrentCompanyFileName()
    CompanyFile.Text = CompanyFileName
    UseCurrentCompany.Value = True
    Interactive.Value = False
    
    '
    ' Close down the current session
    '
    SessionManager.EndSession
    SessionManager.CloseConnection
    Exit Sub
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
End Sub

'
' Selected the desktop communication radio button, show the desktop related
' preferences group and hide the online edition group.
'
Private Sub QBDT_Click()
    OEComm.Visible = False
    DTComm.Visible = True
End Sub

'
' Selected the online edition communication radio button, show the online
' related preferences group and hide the desktop group
'
Private Sub QBOE_Click()
    OEComm.Visible = True
    DTComm.Visible = False
End Sub

'
' Return True if the user has requested the online edition
'
Public Function UseOnlineEdition(appName As String) As Boolean
    myAppName = appName
    Dim OE As String
    OE = GetSetting(appName, "CommunicationPreferences", "UseOE", "")
    If ("Yes" = OE) Then
        UseOnlineEdition = True
    Else
        UseOnlineEdition = False
    End If
End Function

'
' Return the installation ID that is saved in the registry, if we
' haven't saved one yet, generate a GUID and save it, then return it
'
' The online edition uses the Installation ID to keep the different
' application users of an online company separate for the purposes of
' supporting an error recovery list for each unique installation ID.
'
Public Function GetInstallID(appName As String) As String
    Dim InstallID As String
    myAppName = appName
    InstallID = GetSetting(appName, "CommunicationPreferences", "InstallID", "")
    If ("" = InstallID) Then
        InstallID = GetGUID()
        SaveSetting appName, "CommunicationPreferences", "InstallID", InstallID
    End If
    GetInstallID = InstallID
End Function

'
' Now we just have a bunch of Get/Set functions to allow the client to look up
' and (in some cases) set the value of, the information needed to set up a
' connection to QuickBooks
'

'
' Get (but no set, the set comes from this dialog) the connection ticket for
' QuickBooks Online Edition
'
Public Function GetConnTkt(appName As String) As String
    myAppName = appName
    GetConnTkt = GetSetting(appName, "CommunicationPreferences", "ConnTkt", "")
End Function

'
' Get, (but no set, the set comes from this dialog) the company file name for
' QuickBooks Desktop editions.
'
Public Function GetDTCompanyFile(appName As String) As String 'Company file for the desktop editions
    myAppName = appName
    GetDTCompanyFile = GetSetting(appName, "CommunicationPreferences", "CompanyFile", "")
End Function

'
' Get and set the application ID used by QuickBooks Online Edition
'
Public Function GetAppID(appName As String) As String
    myAppName = appName
    GetAppID = GetSetting(appName, "CommunicationPreferences", "AppID", "")
End Function

Public Sub SetAppID(appName As String, appID As String)
    myAppName = appName
    SaveSetting appName, "CommunicationPreferences", "AppID", appID
End Sub

'
' Get and set the application login ID used by QuickBooks Online Edition
'
Public Function GetAppLogin(appName As String) As String
    myAppName = appName
    GetAppLogin = GetSetting(appName, "CommunicationPreferences", "AppLogin", "")
End Function

Public Sub SetAppLogin(appName As String, appLogin As String)
    myAppName = appName
    SaveSetting appName, "CommunicationPreferences", "AppLogin", appLogin
End Sub

'
' Get and set the Language used by QuickBooks Online Edition, the default
' is "English".
'
Public Function GetLanguage(appName As String) As String
    myAppName = appName
    GetLanguage = GetSetting(appName, "CommunicationPreferences", "Language", "English")
End Function

Public Sub SetLanguage(appName As String, language As String)
    myAppName = appName
    SaveSetting appName, "CommunicationPreferences", "Language", language
End Sub

'
' Get and set the application version used by QuickBooks Online Edition, the
' default is "1.0".
'
Public Function GetAppVersion(appName As String) As String
    myAppName = appName
    GetAppVersion = GetSetting(appName, "CommunicationPreferences", "AppVer", "1.0")
End Function

Public Sub SetAppVersion(appName As String, version As String)
    myAppName = appName
    SaveSetting appName, "CommunicationPreferences", "AppVer", version
End Sub

'
' Set up the connection to QuickBooks and return the resulting SessionManager, note that
' the QBOESessionManager object implements the QBSessionManager interface so we return a
' QBSessionManager object even if the object we created is a QBOESessionManager.  Once
' the connection parameters are set up (in this routine) the app doesn't treat the SessionManager
' any differently for OE or desktop.
'
Public Function ConnectToQuickBooks(appName As String) As QBSessionManager
    myAppName = appName
    On Error GoTo ErrHandler
    If (Me.UseOnlineEdition(appName)) Then
        '
        ' User wants the Online Edition, so set up the QBOESessionManager and tell it
        ' all the necessary information for the Online Edition
        '
        Dim OESessionManager As QBOESessionManager
        Set OESessionManager = New QBOESessionManager
        OESessionManager.ConnectionTicket.setValue Me.GetConnTkt(appName)
        OESessionManager.ApplicationLogin.setValue Me.GetAppLogin(appName)
        OESessionManager.InstallationID.setValue Me.GetInstallID(appName)
        OESessionManager.appID.setValue Me.GetAppID(appName)
        OESessionManager.AppVer.setValue Me.GetAppVersion(appName)
        OESessionManager.language.setValue Me.GetLanguage(appName)
        
        '
        ' These don't really do anything for the Online Edition right now but it is
        ' worth keeping them in for symmetry :-)
        '
        OESessionManager.OpenConnection "", appName
        OESessionManager.BeginSession "", omDontCare
            
        '
        ' Now, things get interesting.  The Online Edition supports two types of connections,
        ' one requires that a user login to the Online Edition in order to obtain a "session ticket"
        ' before the application can actually use the connection ticket it has.  The session ticket
        ' represents the user that logged in on behalf of the application and the application's
        ' permissions are restricted based on the logged-in user's permissions.  This is the type
        ' of connection the business owner is *strongly recommended* to use, especially for desktop
        ' applications like this sample.  The other type of connection does not require the user to
        ' login, the connection ticket is sufficient to allow the application to access the company
        ' data.
        '
        ' The only way we can know which type of connection we have is to send a request without a
        ' session ticket and find out if we know the session ticket when the request is done (in
        ' which case we don't need the user to login to QBOE and provide us with a session ticket)
        '
        ' So, we send a trivial request just to see if we need to manually get a session ticket
        '
        Dim CompanyRequestSet As IMsgSetRequest
        '
        ' First supported version of qbXML for Online edition is 2.0 so use it, no need to check
        ' latest supported version here as we do elsewhere.
        '
        Set CompanyRequestSet = OESessionManager.CreateMsgSetRequest("", 2, 0)
        CompanyRequestSet.AppendCompanyQueryRq
        Dim CompanyResponseSet As IMsgSetResponse
        Set CompanyResponseSet = OESessionManager.DoRequests(CompanyRequestSet)
        If (Not OESessionManager.SessionTicket.IsSet) Then
            '
            ' If there is no session ticket after we get the response then we need to manually
            ' get a session ticket by directing the user to login to QuickBooks Online with
            ' a special URL so present the user with the situation and let them do the login.
            '
            Dim sessionURL As String
            Dim authURL As String
            sessionURL = QBOELoginServer & QBOELoginURL & "&appid=" & Me.GetAppID(appName)
            authURL = QBOELoginServer & QBOEAuthURL & "&appid=" & Me.GetAppID(appName) & "&conntkt=" & Me.GetConnTkt(appName)
            SessionTicketDlg.GetSessionTicket OESessionManager, sessionURL, authURL
            If (Not OESessionManager.SessionTicket.IsSet) Then
                MsgBox "Cannot proceed without a session ticket", vbExclamation, "Need Session Ticket"
            End If
        End If
                
        '
        ' All done!  Return the QBOESessionManager as a QBSessionManager object.
        '
        Set ConnectToQuickBooks = OESessionManager
    Else
        '
        ' User wants to talk to the desktop
        '
        Dim SessionManager As New QBSessionManager
        SessionManager.OpenConnection "", appName
        
        '
        ' Check if there is a company file we should be using and use it if we should
        '
        Dim DTCompanyFile As String
        DTCompanyFile = CommPrefsDlg.GetDTCompanyFile(appName)
        SessionManager.BeginSession DTCompanyFile, omDontCare
        
        '
        ' Return the QBSessionManager as a QBSessionManager
        Set ConnectToQuickBooks = SessionManager
    End If
    Exit Function
ErrHandler:
    '
    ' In the event of a problem, display the error and return Nothing for the
    ' SessionManager because we know something went wrong.
    '
    MsgBox Err.Description, vbExclamation, "Error"
    Set ConnectToQuickBooks = Nothing
End Function

'
' Disconnect from QuickBooks, no magic here, just EndSession and CloseConnection
'
Public Sub DisconnectFromQuickBooks(SessionManager As QBSessionManager)
    On Error GoTo ErrHandler
    SessionManager.EndSession
    SessionManager.CloseConnection
    Exit Sub
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
End Sub

'
' Open an IE browser to a given URL to support Online Edition connection
' and session setup
'
Public Sub showURL(iFname As String)
    Dim IE1 As New InternetExplorer
    IE1.Visible = True
    IE1.Navigate iFname
End Sub

'
' Just a simple function to create a GUID for our InstallIDs.
'
Public Function GetGUID() As String

    Dim udtGUID As GUID

    If (CoCreateGuid(udtGUID) = 0) Then
        GetGUID = _
        String(8 - Len(Hex$(udtGUID.Data1)), "0") & Hex$(udtGUID.Data1) & _
        "-" & _
        String(4 - Len(Hex$(udtGUID.Data2)), "0") & Hex$(udtGUID.Data2) & _
        "-" & _
        String(4 - Len(Hex$(udtGUID.Data3)), "0") & Hex$(udtGUID.Data3) & _
        "-" & _
        IIf((udtGUID.Data4(0) < &H10), "0", "") & Hex$(udtGUID.Data4(0)) & _
        IIf((udtGUID.Data4(1) < &H10), "0", "") & Hex$(udtGUID.Data4(1)) & _
        "-" & _
        IIf((udtGUID.Data4(2) < &H10), "0", "") & Hex$(udtGUID.Data4(2)) & _
        IIf((udtGUID.Data4(3) < &H10), "0", "") & Hex$(udtGUID.Data4(3)) & _
        IIf((udtGUID.Data4(4) < &H10), "0", "") & Hex$(udtGUID.Data4(4)) & _
        IIf((udtGUID.Data4(5) < &H10), "0", "") & Hex$(udtGUID.Data4(5)) & _
        IIf((udtGUID.Data4(6) < &H10), "0", "") & Hex$(udtGUID.Data4(6)) & _
        IIf((udtGUID.Data4(7) < &H10), "0", "") & Hex$(udtGUID.Data4(7))
    End If
End Function

