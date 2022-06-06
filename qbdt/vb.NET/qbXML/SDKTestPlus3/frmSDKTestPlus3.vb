Option Strict Off
Option Explicit On
Imports System.Windows.Controls
Imports Microsoft.Win32
Imports SHDocVw

Friend Class frmSDKTestPlus3
    Inherits System.Windows.Forms.Form
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' (c) 2003 Intuit Inc. All Rights Reserved                          '
    ' Use is subject to IP Rights Notice and Restrictions available at: '
    ' http://developer.quickbooks.com/legal/IPRNotice_021201.html       '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim strCompanyFilename As String
    Dim strInXMLFile As String
    Dim strOutXMLFile As String
    Dim strError As String
    Dim strAppName As String
    Dim strOldDrive As String
    Public AuthPrefsDirty As Boolean
    Public ConnPrefsDirty As Boolean

    Public DisplayFile As String

    Private Sub FileError(ByRef FileType As String)
        MsgBox("You must select " & FileType & " file.", MsgBoxStyle.Exclamation, "Beta qbXML Test")
    End Sub
    Private Sub Clear_Status_Area()
        lblStatus.Text = ""
        lblStatus.Refresh()
    End Sub

    Private Sub EnableAuthFlagsSettings(ByRef EnableOrDisable As Boolean)

        AuthFlagsFrame.Enabled = EnableOrDisable
        qbEnterprise.Enabled = EnableOrDisable
        qbPremier.Enabled = EnableOrDisable
        qbPro.Enabled = EnableOrDisable
        qbSimple.Enabled = EnableOrDisable
        qbForceAuthDialog.Enabled = EnableOrDisable
    End Sub

    Private Shared Function GetDefaultBrowserPath() As String
        ' get the name of default browser first. Its in user regKey 'key' below with key progId
        Dim appExecPath As String
        Dim key As String = "Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.html\UserChoice"
        Dim registryKey As RegistryKey = Registry.CurrentUser.OpenSubKey(key, False)
        Dim appName As String = DirectCast(registryKey.GetValue("ProgId"), String)

        ' Now get the path of the exec from root reg keys.
        Dim appExec As RegistryKey = Registry.ClassesRoot.OpenSubKey(String.Concat(appName, "\shell\open\command"), False)
        appExecPath = DirectCast(appExec.GetValue(Nothing, Nothing), String).Split(""""c)(1)
        Return appExecPath
    End Function

    Public Sub showURL(ByRef iFname As String)
        ' changed from IE to default web browser. 
        ' Sreehari - 18/2/2022
        Try
            Process.Start(GetDefaultBrowserPath(), """" + iFname + """")
        Catch ex As Exception
            ' Fallback incase we are unable to fetch the default web app.
            Dim IE1 As Object ' to show the response in IE
            IE1 = CreateObject("InternetExplorer.Application")
            IE1.Visible = True
            IE1.Navigate(iFname)
        End Try
    End Sub

    Private Sub cmdBeginSession_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdBeginSession.Click

        Dim strFileMode As String

        If optSingleUser.Checked Then
            strFileMode = "Single"
        ElseIf optMultiUser.Checked Then
            strFileMode = "Multi"
        Else
            strFileMode = "DoNotCare"
        End If

        strError = BeginSession((CompanyFile.Text), strFileMode, lblStatus)

        If IsNothing(strError) Then
            cmdBeginSession.Enabled = False
            cmdCloseConnection.Enabled = False
            cmdEndSession.Enabled = True
            sessionPrefsControl((False))
            cmdSendXML.Enabled = (Not RequestFile.Text = "")
        End If
    End Sub

    Private Sub cmdCloseConnection_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCloseConnection.Click
        CloseConnection(lblStatus)
        cmdBeginSession.Enabled = False
        cmdOpenConnection.Enabled = True
        cmdCloseConnection.Enabled = False
        cmdProcessSubscription.Enabled = False
        txtApplicationName.Enabled = True
    End Sub

    Private Sub cmdEndSession_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdEndSession.Click
        EndSession(lblStatus)
        cmdBeginSession.Enabled = True
        cmdCloseConnection.Enabled = True
        cmdEndSession.Enabled = False
        cmdSendXML.Enabled = False
        sessionPrefsControl((True))
    End Sub

    Private Sub cmdExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdExit.Click
        End
    End Sub

    Private Sub cmdOpenConnection_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOpenConnection.Click
        Dim strFileMode As String

        If IsNothing(txtApplicationName.Text) Then
            MsgBox("You must supply an application name before opening a connection")
            Exit Sub
        End If


        strAppName = txtApplicationName.Text
        strError = OpenConnection("", strAppName, lblStatus)

        If IsNothing(strError) Then
            cmdBeginSession.Enabled = True
            cmdCloseConnection.Enabled = True
            cmdOpenConnection.Enabled = False
            txtApplicationName.Enabled = False
            cmdProcessSubscription.Enabled = (Not RequestFile.Text = "")
            sessionPrefsControl((True))
            lblStatus.Text = "OpenConnection call successful"
            lblStatus.Refresh()
        End If
    End Sub


    Private Sub sessionPrefsControl(ByRef state As Boolean)
        SessionPrefsFrame.Enabled = state
        Unattended.Enabled = state
        ReadOnly_Renamed.Enabled = state
        pdRequired.Enabled = state
        pdOptional.Enabled = state
        pdNotNeeded.Enabled = state
        optDontCare.Enabled = state
        optMultiUser.Enabled = state
        optSingleUser.Enabled = state
    End Sub
    Private Sub cmdProcessSubscription_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdProcessSubscription.Click

        strOutXMLFile = CurDir() & "\SubscriptionResponse.xml"

        strError = SendXMLFile("SubscriptionProcessor", (RequestFile.Text), strOutXMLFile, lblStatus)
        If IsNothing(strError) Then
            cmdViewOutput.Enabled = True
        End If
    End Sub

    Private Sub cmdSendXML_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSendXML.Click

        strOutXMLFile = CurDir() & "\QBResponse.xml"

        strError = SendXMLFile("RequestProcessor", (RequestFile.Text), strOutXMLFile, lblStatus)
        If IsNothing(strError) Then
            cmdViewOutput.Enabled = True
        End If
    End Sub

    Private Sub cmdViewInput_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdViewInput.Click
        showURL((RequestFile.Text))
    End Sub

    Private Sub cmdViewOutput_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdViewOutput.Click
        showURL(strOutXMLFile)
    End Sub


    Private Sub CompanyBrowse_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CompanyBrowse.Click
        OpenFileDialogOpen.Filter = "QuickBooks Files|*.qbw"
        OpenFileDialogOpen.InitialDirectory = "c:\program files\Intuit"
        OpenFileDialogOpen.ShowDialog()
        CompanyFile.Text = OpenFileDialogOpen.FileName
        ' CompanyFile.Text = "C:\\Users\\Public\\Documents\\Intuit\\QuickBooks\\Company Files\\Phoenix_Dummy_CF1.qbw"
    End Sub

    Private Sub optUseCurrentFile_Click()
        Dim lstCompanyFile As Object
        lstCompanyFile.Refresh()
    End Sub

    Private Sub ConnLocal_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ConnLocal.CheckedChanged
        If eventSender.Checked Then
            ConnPrefsDirty = True
            EnableAuthFlagsSettings((True))
        End If
    End Sub

    Private Sub ConnLocalUI_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ConnLocalUI.CheckedChanged
        If eventSender.Checked Then
            ConnPrefsDirty = True
            EnableAuthFlagsSettings((True))
        End If
    End Sub

    Private Sub ConnNone_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ConnNone.CheckedChanged
        If eventSender.Checked Then
            ConnPrefsDirty = True
            EnableAuthFlagsSettings((True))
        End If
    End Sub

    Private Sub ConnQBOE_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ConnQBOE.CheckedChanged
        If eventSender.Checked Then
            ConnPrefsDirty = True
            EnableAuthFlagsSettings((False))
        End If
    End Sub

    Private Sub ConnRemote_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ConnRemote.CheckedChanged
        If eventSender.Checked Then
            ConnPrefsDirty = True
            EnableAuthFlagsSettings((False))
        End If
    End Sub

    Private Sub frmSDKTestPlus3_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        If (qbEnterprise.CheckState = 1 Or qbPremier.CheckState = 1 Or qbPro.CheckState = 1 Or qbSimple.CheckState = 1) Then
            AuthPrefsDirty = True
        Else
            AuthPrefsDirty = False
        End If
        ConnPrefsDirty = False
    End Sub

    Private Sub pdNotNeeded_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles pdNotNeeded.CheckedChanged
        If eventSender.Checked Then
            AuthPrefsDirty = True
        End If
    End Sub

    Private Sub pdOptional_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles pdOptional.CheckedChanged
        If eventSender.Checked Then
            AuthPrefsDirty = True

        End If
    End Sub

    Private Sub pdRequired_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles pdRequired.CheckedChanged
        If eventSender.Checked Then
            AuthPrefsDirty = True

        End If
    End Sub


    Private Sub ReadOnly_Renamed_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ReadOnly_Renamed.CheckStateChanged
        AuthPrefsDirty = True
    End Sub

    Private Sub RequestBrowse_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles RequestBrowse.Click
        OpenFileDialogOpen.Filter = "XML Files|*.xml;*.qbxml"
        OpenFileDialogOpen.InitialDirectory = "../../../../xmlfiles"
        OpenFileDialogOpen.ShowDialog()
        RequestFile.Text = OpenFileDialogOpen.FileName
    End Sub

    Private Sub RequestFile_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles RequestFile.TextChanged
        cmdSendXML.Enabled = (Not RequestFile.Text = "")
        cmdProcessSubscription.Enabled = (Not RequestFile.Text = "")
        cmdViewInput.Enabled = (Not RequestFile.Text = "")
    End Sub

    Private Sub Unattended_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Unattended.CheckStateChanged
        AuthPrefsDirty = True
    End Sub

    Private Sub frmSDKTestPlus3_Paint(sender As Object, e As PaintEventArgs) Handles MyBase.Paint
        Dim myGraphics As Graphics = Me.CreateGraphics
        Dim myPen As Pen
        myPen = New Pen(Brushes.Black, 6)
        e.Graphics.DrawLine(myPen, 80, 295, 80, 353)
        myPen.EndCap = Drawing2D.LineCap.ArrowAnchor
        e.Graphics.DrawLine(myPen, 122, 295, 185, 295)
        e.Graphics.DrawLine(myPen, 300, 295, 416, 295)
        e.Graphics.DrawLine(myPen, 500, 295, 551, 295)
        e.Graphics.DrawLine(myPen, 305, 350, 545, 350)
        e.Graphics.DrawLine(myPen, 80, 350, 185, 350)
        myPen.StartCap = Drawing2D.LineCap.ArrowAnchor
        e.Graphics.DrawLine(myPen, 590, 300, 590, 335)
    End Sub

End Class