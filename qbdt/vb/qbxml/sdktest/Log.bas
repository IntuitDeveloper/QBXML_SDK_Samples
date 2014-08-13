Attribute VB_Name = "Log"
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
' $Id: Log.bas,v 1.1 2002/08/05 12:55:39 giza Exp $
'
Option Explicit
'
' Write to streaming log-message window.
'
Public Sub out(msg As String)
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
Public Sub showURL(iFname As String)
    Dim IE1 As New InternetExplorer
    IE1.Visible = True
    IE1.Navigate (iFname)
End Sub
'
' Save xml to a file and then display in browser.
'
Public Sub showXMLStream(xml As String)
    Dim fname As String
    Dim fnum As Integer
    Dim isOpen As Boolean
    On Error GoTo problem
    fname = App.Path & "\LastResponse.xml"
    fnum = FreeFile()
    Open fname For Output As #fnum
    isOpen = True
    Print #fnum, xml
    Close #fnum
    showURL (fname)
    Exit Sub
problem:
    If isOpen Then Close #fnum
    If Err Then Err.Raise Err.Number, , Err.Description
End Sub

Public Sub clearXML()
    Dim fname As String
    fname = App.Path & "\LastResponse.xml"
    On Error Resume Next
    Kill fname
End Sub

