Option Strict Off
Option Explicit On
Module Log
	'
	' Copyright © 2021-2022 Intuit Inc. All rights reserved.
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
	'
	' Display file in browser.
	'
	Public Sub showURL(ByRef iFname As String)
        Dim IE1 As Object ' to show the response in IE
        IE1 = CreateObject("InternetExplorer.Application")
        IE1.Visible = True
        IE1.navigate(iFname)
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