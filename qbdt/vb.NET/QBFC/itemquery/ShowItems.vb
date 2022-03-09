Option Strict Off
Option Explicit On
Friend Class ShowItems
	Inherits System.Windows.Forms.Form
	'----------------------------------------------------------
	' ShowItems.frm
	'
	' This form is simply used to display items which are
	' returned from QuickBooks.  All interaction with QB
	' takes place in the Main Module.
	'
	' Last Updated: 08/2002
	'
	' Copyright © 2021-2022 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	'----------------------------------------------------------
	
	Private Sub Done_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Done.Click
		Me.Close()
	End Sub
End Class