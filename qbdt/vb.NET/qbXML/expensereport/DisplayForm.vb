Option Strict Off
Option Explicit On
Friend Class DisplayForm
	Inherits System.Windows.Forms.Form
	' DisplayForm.frm
	' Created July, 2002
	'
	' This form can be used to display the request and response
	' qbxml to the user for illustrational purposes.
	'
	' Copyright © 2021-2022 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	'
	'
	
	
	Private XMLString As String
	
	
	'
	' Set the xml string to be displayed later.
	'
	Public Sub setXML(ByRef xmlToShow As String)
		XMLString = xmlToShow
	End Sub
	
	
	
	'
	' Display xml to the user.
	'
	Public Sub showXML(ByRef title As String)
		Me.Text = title
		displayText.Text = XMLString
		Me.Show()
	End Sub
	
	
	
	'
	' Close the form.
	'
	Private Sub closeButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles closeButton.Click
		Me.Close()
	End Sub
End Class