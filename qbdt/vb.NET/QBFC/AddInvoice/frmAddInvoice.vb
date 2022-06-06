Option Strict Off
Option Explicit On
Friend Class frmAddInvoice
	Inherits System.Windows.Forms.Form
	'----------------------------------------------------------
	' Form: frmAddInvoice
	'
	' Description:  This sample demonstrates the use of qbXML 2.0 or
	'               QBFC, by adding an invoice within.
	'               It includes examples of the following:
	'                   - Constructing an invoice add request including
	'                     - Using a different shipping address
	'                     - Adding multiple invoice lines including two
	'                       item lines and one item group line
	'
	'               All non form code is found in modAddInvoice
	'
	' Created On: 09/10/2002
	'
	' Copyright © 2002-2020 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	'----------------------------------------------------------
	
	
	Private Sub cmdAddInvoice_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAddInvoice.Click
		If optqbXMLRP.Checked Then
			qbXML_AddInvoice()
		Else
			QBFC_AddInvoice()
		End If
	End Sub
	
	Private Sub cmdExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdExit.Click
		End
	End Sub
End Class