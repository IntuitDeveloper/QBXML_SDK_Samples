Option Strict Off
Option Explicit On
Friend Class DisplayXML
	Inherits System.Windows.Forms.Form
	' DisplayXML.frm
	'
	' This form is used by the QuickBooks SDK 2.0 Delete Customer sample
	' to display qbXML request/response.
	'
	' Created February, 2002
	' Updated to SDK 2.0 August, 2002
	'
	' Copyright © 2021-2022 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	'-------------------------------------------------------------
	
	
	
	Private Sub Comm_Exit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Comm_Exit.Click
		Me.Close()
	End Sub
End Class