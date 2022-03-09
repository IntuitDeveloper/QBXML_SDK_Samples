Option Strict Off
Option Explicit On
Friend Class DisplayXML
	Inherits System.Windows.Forms.Form
	'-----------------------------------------------------------
	' Form Module: DisplayFrm
	'
	' Description: this form is to display qbXML request/response.
	'
	' Created On: 09/15/2001
	' Updated to SDK 2.0 On: 07/30/2002
	'
	' Unpublished work. INTUIT CONFIDENTIAL.
	' Copyright © 2021-2022 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	'----------------------------------------------------------
	
	
	Private Sub Comm_Exit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Comm_Exit.Click
		Me.Close()
	End Sub
End Class