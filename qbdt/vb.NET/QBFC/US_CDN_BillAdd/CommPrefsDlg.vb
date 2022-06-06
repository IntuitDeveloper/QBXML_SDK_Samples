Option Strict Off
Option Explicit On
Friend Class CommPrefsDlg
	Inherits System.Windows.Forms.Form
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
	
	
	'
	' Constants for communication setup with QuickBooks Online Edition
	'
	Const QBOELoginServer As String = "https://login.quickbooks.com"
	Const QBOEConnectionURL As String = "/j/qbn/sdkapp/confirm?serviceid=2004"
	Const QBOELoginURL As String = "/j/qbn/sdkapp/sessionauth2?serviceid=2004"
	Const QBOEAuthURL As String = "/j/qbn/sdkapp/connauth?serviceid=2004"
	
	'
	' Private class variables
	'
	Private myAppName As String 'Used to identify our key in the registry for remembering preferences
	
	'
	' Declarations to support building a GUID (see GetGUID at the bottom of this file)
	'
	Private Structure GUID
		Dim Data1 As Integer
		Dim Data2 As Short
		Dim Data3 As Short
		<VBFixedArray(7)> Dim Data4() As Byte
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			ReDim Data4(7)
		End Sub
	End Structure
	'UPGRADE_WARNING: Structure GUID may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function CoCreateGuid Lib "OLE32.DLL" (ByRef pGuid As GUID) As Integer
	
	'
	' Public properties
	'
	Public isOK As Boolean 'If "OK" was clicked last time the dialog was displayed.
	
	'
	' Load the form and display the correct group box for the online edition or the desktop
	'
	Private Sub CommPrefsDlg_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		If (Me.UseOnlineEdition(myAppName)) Then
			QBOE.Checked = True
		Else
			QBDT.Checked = True
		End If
		ConnectionTicket.Text = Me.GetConnTkt(myAppName)
		CompanyFile.Text = Me.GetDTCompanyFile(myAppName)
		If ("" <> CompanyFile.Text) Then
			UseCurrentCompany.Checked = True
		Else
			Interactive.Checked = True
		End If
	End Sub
	Private Sub CancelButton_Renamed_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CancelButton_Renamed.Click
		isOK = False
		Me.Hide()
	End Sub
	
	'
	' The "Connect to QuickBooks Online" button was clicked, open a browser to QBOE with a Query string
	' indicating the desire to set up a connection between the application and QBOE
	'
	Private Sub Connect_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Connect.Click
		Dim appID As String
		appID = Me.GetAppID(myAppName)
		If ("" = appID) Then
			MsgBox("Must call CommPrefsDlg.SetAppID before displaying the CommPrefsDlg", MsgBoxStyle.Critical, "Error")
		Else
			showURL((QBOELoginServer & QBOEConnectionURL & "&appid=" & appID))
		End If
	End Sub
	
	Private Sub CommPrefsDlg_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		SessionTicketDlg.Close()
	End Sub
	
	'
	' The "Let me set up a company" button was clicked.
	'
	Private Sub GoToOE_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles GoToOE.Click
		showURL(("http://oe.quickbooks.com"))
	End Sub
	
	'UPGRADE_WARNING: Event Interactive.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Interactive_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Interactive.CheckedChanged
		If eventSender.Checked Then
			CompanyFile.Text = ""
		End If
	End Sub
	
	'
	' The OK button was clicked, verify that we have the information we need and save the
	' communication preferences to the registry if we do.  If we do not, show an error and
	' don't close the dialog.
	'
	Private Sub OKButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OKButton.Click
		Dim Good As Boolean 'Do we have all the information we need
		Good = True 'assume we do.
		If (QBOE.Checked) Then
			SaveSetting(myAppName, "CommunicationPreferences", "UseOE", "Yes")
			If ("" = ConnectionTicket.Text) Then
				MsgBox("You must specify the connection ticket when communicating with QuickBooks Online Edition", MsgBoxStyle.Exclamation)
				Good = False
			End If
			SaveSetting(myAppName, "CommunicationPreferences", "ConnTkt", ConnectionTicket.Text)
		Else
			SaveSetting(myAppName, "CommunicationPreferences", "UseOE", "No")
			'
			' If the user wants us to save the company file in use, do so.
			'
			If (UseCurrentCompany.Checked) Then
				SaveSetting(myAppName, "CommunicationPreferences", "CompanyFile", CompanyFile.Text)
			Else
				CompanyFile.Text = ""
				SaveSetting(myAppName, "CommunicationPreferences", "CompanyFile", "")
			End If
		End If
		isOK = True
		If (Good) Then
			Me.Hide()
		End If
	End Sub
	
	'
	' The Connect button was clicked in the desktop preferences group,
	' open a session long enough to find out what the open company file is so
	' we can display it (and save it for later use).
	'
	Private Sub QBConnect_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles QBConnect.Click
		On Error GoTo ErrHandler
		'
		' Set up a traditional desktop session manager
		'
		Dim SessionManager As New QBFC15Lib.QBSessionManager
		SessionManager.OpenConnection("", myAppName)
		SessionManager.BeginSession("", QBFC15Lib.ENOpenMode.omDontCare)
		
		'
		' Get the company file name and display it and set the default to use it
		' again later.
		'
		Dim CompanyFileName As String
		CompanyFileName = SessionManager.GetCurrentCompanyFileName()
		CompanyFile.Text = CompanyFileName
		UseCurrentCompany.Checked = True
		Interactive.Checked = False
		
		'
		' Close down the current session
		'
		SessionManager.EndSession()
		SessionManager.CloseConnection()
		Exit Sub
ErrHandler: 
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
	End Sub
	
	'
	' Selected the desktop communication radio button, show the desktop related
	' preferences group and hide the online edition group.
	'
	'UPGRADE_WARNING: Event QBDT.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub QBDT_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles QBDT.CheckedChanged
		If eventSender.Checked Then
			OEComm.Visible = False
			DTComm.Visible = True
		End If
	End Sub
	
	'
	' Selected the online edition communication radio button, show the online
	' related preferences group and hide the desktop group
	'
	'UPGRADE_WARNING: Event QBOE.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub QBOE_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles QBOE.CheckedChanged
		If eventSender.Checked Then
			OEComm.Visible = True
			DTComm.Visible = False
		End If
	End Sub
	
	'
	' Return True if the user has requested the online edition
	'
	Public Function UseOnlineEdition(ByRef appName As String) As Boolean
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
	Public Function GetInstallID(ByRef appName As String) As String
		Dim InstallID As String
		myAppName = appName
		InstallID = GetSetting(appName, "CommunicationPreferences", "InstallID", "")
		If ("" = InstallID) Then
			InstallID = GetGUID()
			SaveSetting(appName, "CommunicationPreferences", "InstallID", InstallID)
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
	Public Function GetConnTkt(ByRef appName As String) As String
		myAppName = appName
		GetConnTkt = GetSetting(appName, "CommunicationPreferences", "ConnTkt", "")
	End Function
	
	'
	' Get, (but no set, the set comes from this dialog) the company file name for
	' QuickBooks Desktop editions.
	'
	Public Function GetDTCompanyFile(ByRef appName As String) As String 'Company file for the desktop editions
		myAppName = appName
		GetDTCompanyFile = GetSetting(appName, "CommunicationPreferences", "CompanyFile", "")
	End Function
	
	'
	' Get and set the application ID used by QuickBooks Online Edition
	'
	Public Function GetAppID(ByRef appName As String) As String
		myAppName = appName
		GetAppID = GetSetting(appName, "CommunicationPreferences", "AppID", "")
	End Function
	
	Public Sub SetAppID(ByRef appName As String, ByRef appID As String)
		myAppName = appName
		SaveSetting(appName, "CommunicationPreferences", "AppID", appID)
	End Sub
	
	'
	' Get and set the application login ID used by QuickBooks Online Edition
	'
	Public Function GetAppLogin(ByRef appName As String) As String
		myAppName = appName
		GetAppLogin = GetSetting(appName, "CommunicationPreferences", "AppLogin", "")
	End Function
	
	Public Sub SetAppLogin(ByRef appName As String, ByRef appLogin As String)
		myAppName = appName
		SaveSetting(appName, "CommunicationPreferences", "AppLogin", appLogin)
	End Sub
	
	'
	' Get and set the Language used by QuickBooks Online Edition, the default
	' is "English".
	'
	Public Function GetLanguage(ByRef appName As String) As String
		myAppName = appName
		GetLanguage = GetSetting(appName, "CommunicationPreferences", "Language", "English")
	End Function
	
	Public Sub SetLanguage(ByRef appName As String, ByRef language As String)
		myAppName = appName
		SaveSetting(appName, "CommunicationPreferences", "Language", language)
	End Sub
	
	'
	' Get and set the application version used by QuickBooks Online Edition, the
	' default is "1.0".
	'
	Public Function GetAppVersion(ByRef appName As String) As String
		myAppName = appName
		GetAppVersion = GetSetting(appName, "CommunicationPreferences", "AppVer", "1.0")
	End Function
	
	Public Sub SetAppVersion(ByRef appName As String, ByRef version As String)
		myAppName = appName
		SaveSetting(appName, "CommunicationPreferences", "AppVer", version)
	End Sub
	
	'
	' Set up the connection to QuickBooks and return the resulting SessionManager, note that
	' the QBOESessionManager object implements the QBSessionManager interface so we return a
	' QBSessionManager object even if the object we created is a QBOESessionManager.  Once
	' the connection parameters are set up (in this routine) the app doesn't treat the SessionManager
	' any differently for OE or desktop.
	'
	Public Function ConnectToQuickBooks(ByRef appName As String) As QBFC15Lib.QBSessionManager
		myAppName = appName
		On Error GoTo ErrHandler
		Dim CompanyRequestSet As QBFC15Lib.IMsgSetRequest
		Dim CompanyResponseSet As QBFC15Lib.IMsgSetResponse
		Dim sessionURL As String
		Dim authURL As String
		Dim SessionManager As New QBFC15Lib.QBSessionManager
		Dim DTCompanyFile As String
		'UPGRADE_WARNING: Arrays in structure OESessionManager may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim OESessionManager As QBFC15Lib.QBOESessionManager
		If (Me.UseOnlineEdition(appName)) Then
			'
			' User wants the Online Edition, so set up the QBOESessionManager and tell it
			' all the necessary information for the Online Edition
			'
			OESessionManager = New QBFC15Lib.QBOESessionManager
			OESessionManager.ConnectionTicket.setValue(Me.GetConnTkt(appName))
			OESessionManager.ApplicationLogin.setValue(Me.GetAppLogin(appName))
			OESessionManager.InstallationID.setValue(Me.GetInstallID(appName))
			OESessionManager.appID.setValue(Me.GetAppID(appName))
			OESessionManager.AppVer.setValue(Me.GetAppVersion(appName))
			OESessionManager.language.setValue(Me.GetLanguage(appName))
			
			'
			' These don't really do anything for the Online Edition right now but it is
			' worth keeping them in for symmetry :-)
			'
			OESessionManager.OpenConnection("", appName)
			OESessionManager.BeginSession("", QBFC15Lib.ENOpenMode.omDontCare)
			
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
			'
			' First supported version of qbXML for Online edition is 2.0 so use it, no need to check
			' latest supported version here as we do elsewhere.
			'
			CompanyRequestSet = OESessionManager.CreateMsgSetRequest("", 2, 0)
			CompanyRequestSet.AppendCompanyQueryRq()
			CompanyResponseSet = OESessionManager.DoRequests(CompanyRequestSet)
			If (Not OESessionManager.SessionTicket.IsSet) Then
				'
				' If there is no session ticket after we get the response then we need to manually
				' get a session ticket by directing the user to login to QuickBooks Online with
				' a special URL so present the user with the situation and let them do the login.
				'
				sessionURL = QBOELoginServer & QBOELoginURL & "&appid=" & Me.GetAppID(appName)
				authURL = QBOELoginServer & QBOEAuthURL & "&appid=" & Me.GetAppID(appName) & "&conntkt=" & Me.GetConnTkt(appName)
				SessionTicketDlg.GetSessionTicket(OESessionManager, sessionURL, authURL)
				If (Not OESessionManager.SessionTicket.IsSet) Then
					MsgBox("Cannot proceed without a session ticket", MsgBoxStyle.Exclamation, "Need Session Ticket")
				End If
			End If
			
			'
			' All done!  Return the QBOESessionManager as a QBSessionManager object.
			'
			ConnectToQuickBooks = OESessionManager
		Else
			'
			' User wants to talk to the desktop
			'
			SessionManager.OpenConnection("", appName)
			
			'
			' Check if there is a company file we should be using and use it if we should
			'
			DTCompanyFile = Me.GetDTCompanyFile(appName)
			SessionManager.BeginSession(DTCompanyFile, QBFC15Lib.ENOpenMode.omDontCare)
			
			'
			' Return the QBSessionManager as a QBSessionManager
			ConnectToQuickBooks = SessionManager
		End If
		Exit Function
ErrHandler: 
		'
		' In the event of a problem, display the error and return Nothing for the
		' SessionManager because we know something went wrong.
		'
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
		'UPGRADE_NOTE: Object ConnectToQuickBooks may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		ConnectToQuickBooks = Nothing
	End Function
	
	'
	' Disconnect from QuickBooks, no magic here, just EndSession and CloseConnection
	'
	Public Sub DisconnectFromQuickBooks(ByRef SessionManager As QBFC15Lib.QBSessionManager)
		On Error GoTo ErrHandler
		SessionManager.EndSession()
		SessionManager.CloseConnection()
		Exit Sub
ErrHandler: 
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
	End Sub

    Private Shared Function GetDefaultBrowserPath() As String
        ' get the name of default browser first. Its in user regKey 'key' below with key progId
        Dim appExecPath As String
        Dim key As String = "Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.html\UserChoice"
        Dim registryKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(key, False)
        Dim appName As String = DirectCast(registryKey.GetValue("ProgId"), String)

        ' Now get the path of the exec from root reg keys.
        Dim appExec As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(String.Concat(appName, "\shell\open\command"), False)
        appExecPath = DirectCast(appExec.GetValue(Nothing, Nothing), String).Split(""""c)(1)
        Return appExecPath
    End Function

    '
    ' Open an IE browser to a given URL to support Online Edition connection
    ' and session setup
    ' 
    ' Sreehari, 22/2/2022 - change to open in default web browser instead of ie.
    '
    Public Sub showURL(ByRef iFname As String)
        Try
            Process.Start(GetDefaultBrowserPath(), """" + iFname + """")
        Catch ex As Exception
            Dim IE1 As New System.Windows.Forms.WebBrowser
            IE1.Visible = True
            IE1.Navigate(New System.Uri(iFname))
        End Try
    End Sub

    '
    ' Just a simple function to create a GUID for our InstallIDs.
    '
    Public Function GetGUID() As String
		
		'UPGRADE_WARNING: Arrays in structure udtGUID may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim udtGUID As GUID
		
		If (CoCreateGuid(udtGUID) = 0) Then
			GetGUID = New String("0", 8 - Len(Hex(udtGUID.Data1))) & Hex(udtGUID.Data1) & "-" & New String("0", 4 - Len(Hex(udtGUID.Data2))) & Hex(udtGUID.Data2) & "-" & New String("0", 4 - Len(Hex(udtGUID.Data3))) & Hex(udtGUID.Data3) & "-" & IIf((udtGUID.Data4(0) < &H10), "0", "") & Hex(udtGUID.Data4(0)) & IIf((udtGUID.Data4(1) < &H10), "0", "") & Hex(udtGUID.Data4(1)) & "-" & IIf((udtGUID.Data4(2) < &H10), "0", "") & Hex(udtGUID.Data4(2)) & IIf((udtGUID.Data4(3) < &H10), "0", "") & Hex(udtGUID.Data4(3)) & IIf((udtGUID.Data4(4) < &H10), "0", "") & Hex(udtGUID.Data4(4)) & IIf((udtGUID.Data4(5) < &H10), "0", "") & Hex(udtGUID.Data4(5)) & IIf((udtGUID.Data4(6) < &H10), "0", "") & Hex(udtGUID.Data4(6)) & IIf((udtGUID.Data4(7) < &H10), "0", "") & Hex(udtGUID.Data4(7))
		End If
	End Function
End Class