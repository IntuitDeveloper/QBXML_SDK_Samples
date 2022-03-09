Option Strict Off
Option Explicit On
Friend Class SessionTicketDlg
	Inherits System.Windows.Forms.Form
	'
	' This is just a simple dialog that explains that the user must login to
	' QuickBooks Online Edition and obtain a session ticket which they paste into
	' this dialog so that we can tell the QBOESessionManager the session ticket that
	' was obtained manually
	'
	' Note that this dialog and the CommPrefsDlg work hand-in-hand
	'
	
	'
	' Private variables
	'
	Private DidOK As Boolean 'Was OK pressed?
	Private myLoginURL As String 'Passed in when dialog is invoked to tell us the QBOE URL to get a ticket
	Private myAuthURL As String 'Passed in when dialog is invoked to tell us our connection ticket
	
	'
	' The "main" function of this dialog, this is the only way this dialog should be invoked
	' (i.e. the program should not call SessionTicketDlg.Show itself, this function does it)
	'
	Public Sub GetSessionTicket(ByRef OESessionManager As QBFC15Lib.QBOESessionManager, ByRef loginURL As String, ByRef authURL As String)
		'
		' Save the URL for later
		'
		myLoginURL = loginURL
		myAuthURL = authURL
		'
		' Show the dialog
		'
		Me.ShowDialog()
		'
		' If the user clicked OK store the session ticket we got in the OESessionManager
		'
		If (DidOK) Then
			myAuthURL = myAuthURL & "&sessiontkt=" & Me.SessionTicketText.Text
			OESessionManager.SessionTicket.setValue(GetRealSessionTicket)
		End If
	End Sub
	
	Private Sub CancelButton_Renamed_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CancelButton_Renamed.Click
		DidOK = False
		Me.Hide()
	End Sub
	
	Private Sub OKButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OKButton.Click
		DidOK = True
		Me.Hide()
	End Sub
	
	'
	' Open a browser to QuickBooks Online so the user can get the session ticket for us.
	'
	Private Sub QBLogin_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles QBLogin.Click
		showURL(myLoginURL)
	End Sub
	
	Private Function GetRealSessionTicket() As String
        Dim http As New MSXML2.XMLHTTP60
        http.open("GET", myAuthURL, False)
		http.send()
		Dim resp As String
		resp = http.responseText
		resp = Mid(resp, 4)
		GetRealSessionTicket = resp
	End Function

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
    Public Sub showURL(ByRef iFname As String)
        Try
            Process.Start(GetDefaultBrowserPath(), """" + iFname + """")
        Catch ex As Exception
            Dim IE1 As New System.Windows.Forms.WebBrowser
            IE1.Visible = True
            IE1.Navigate(New System.Uri(iFname))
        End Try
    End Sub
End Class