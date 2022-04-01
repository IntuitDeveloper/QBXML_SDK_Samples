Option Strict Off
Option Explicit On

Imports Microsoft.Win32

Module Log
	'
	' Copyright © 2002-2020 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	' $Id: Log.bas,v 1.1 2002/08/05 12:55:39 giza Exp $
	'
	'
	' Write to streaming log-message window.
	'
	Public Sub out(ByRef msg As String)
		UIForm.logBox.Text = msg & vbCrLf & UIForm.logBox.Text
	End Sub
    '
    ' Write to streaming log-message window.
    '
    Public Sub clear()
        UIForm.logBox.Text = ""
    End Sub
    Private Function GetDefaultBrowserPath() As String
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
    '
    ' Display file in browser.
    '
    Public Sub showURL(ByRef iFname As String)
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
    '
    ' Save xml to a file and then display in browser.
    '
    Public Sub showXMLStream(ByRef xml As String)
		Dim fname As String
		Dim fnum As Short
		Dim isOpen As Boolean
		On Error GoTo problem
		fname = My.Application.Info.DirectoryPath & "\LastResponse.xml"
		fnum = FreeFile
		FileOpen(fnum, fname, OpenMode.Output)
		isOpen = True
		PrintLine(fnum, xml)
		FileClose(fnum)
		showURL((fname))
		Exit Sub
problem: 
		If isOpen Then FileClose(fnum)
		If Err.Number Then Err.Raise(Err.Number,  , Err.Description)
	End Sub
	
	Public Sub clearXML()
		Dim fname As String
		fname = My.Application.Info.DirectoryPath & "\LastResponse.xml"
		On Error Resume Next
		Kill(fname)
	End Sub
End Module