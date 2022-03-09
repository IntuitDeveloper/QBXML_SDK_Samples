Option Strict Off
Option Explicit On
Friend Class BigDisplayForm
	Inherits System.Windows.Forms.Form
	' BigDisplayForm.frm
	' Created July, 2002
	'
	' This form contains an embedded browser which is
	' used to display the report we generate to the user.
	'
	' Copyright © 2021-2022 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	'
	
	
	
	'
	' Display the html page at url in the embedded browser.
	'
	Public Sub displayResults(ByRef url As String)
		On Error GoTo ErrorHandler
		
		Me.Show()
        browser.Navigate(New System.Uri(url))
        Exit Sub
		
ErrorHandler: 
		MsgBox("Error displaying results")
		Me.Close()
		Exit Sub
	End Sub
	
	
	
	'
	' Close the form when the user clicks the close button.
	'
	Private Sub closeButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles closeButton.Click
		Me.Close()
	End Sub
End Class