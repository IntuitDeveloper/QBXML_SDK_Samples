Option Strict Off
Option Explicit On
Friend Class frmInvoiceDisplay
	Inherits System.Windows.Forms.Form
	'----------------------------------------------------------
	' Form: frmInvoiceDisplay
	'
	' Description:  This form is used to display the invoice
	'               information returned by the add invoice request
	'               made from frmAddInvoice.
	'
	' Created On: 09/10/2002
	'
	' Copyright © 2021-2022 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	'----------------------------------------------------------
	
	
	Private Sub cmdCloseWindow_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCloseWindow.Click
		Me.Close()
	End Sub
End Class