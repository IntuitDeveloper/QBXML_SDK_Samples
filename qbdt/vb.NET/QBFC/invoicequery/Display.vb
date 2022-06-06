Option Strict Off
Option Explicit On
Friend Class Display
	Inherits System.Windows.Forms.Form
	'-----------------------------------------------------------
	' Form Module: Display
	'
	' Description:  Display invoice list
	'
	' Last Updated: 08/2002
	'
	' Copyright © 2002-2020 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	'----------------------------------------------------------
	
	
	Private Sub Comm_Exit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Comm_Exit.Click
		Me.Close()
	End Sub
End Class