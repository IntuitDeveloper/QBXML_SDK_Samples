Option Strict Off
Option Explicit On
Friend Class Display
	Inherits System.Windows.Forms.Form
	'-----------------------------------------------------------
	' Form Module: DisplayFrm
	'
	' Description:  Form to display qbXML request/response.
	'
	' Created On: 11/09/2001
	' Updated to SDK 2.0: 08/05/2002
	'
	' Copyright © 2021-2022 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	'----------------------------------------------------------
	
	
	Private Sub Comm_Exit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Comm_Exit.Click
		Me.Close()
	End Sub
End Class