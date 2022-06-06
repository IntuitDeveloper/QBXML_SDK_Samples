Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmAddDataExtension
	Inherits System.Windows.Forms.Form
	'----------------------------------------------------------
	' Form: frmAddDataExtension
	'
	' Description: Allows the user to select a select a customer causing
	'              the form to list the unused data extensions for that
	'              customer.  Then the user may highlight one of the
	'              unused data extensions, provide a value for it and
	'              add it to the customer record
	'
	' Copyright © 2021-2022 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	'----------------------------------------------------------
	Private Sub cmdAddValue_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAddValue.Click
		
		If lstCustomers.SelectedIndex < 0 Then
			MsgBox("You must select a customer first")
			Exit Sub
		End If
		
		If lstUnusedDataExtDefs.SelectedIndex < 0 Then
			MsgBox("You must select a Data Extension to add first")
			Exit Sub
		End If
		
		If txtDataExtValue.Text = "" Then
			MsgBox("You must supply a value for the Data Extension first")
			Exit Sub
		End If
		
		If AddDataExt(VB6.GetItemString(lstCustomers, lstCustomers.SelectedIndex), VB.Left(VB6.GetItemString(lstUnusedDataExtDefs, lstUnusedDataExtDefs.SelectedIndex), InStr(VB6.GetItemString(lstUnusedDataExtDefs, lstUnusedDataExtDefs.SelectedIndex), "  |") - 1), (txtDataExtValue.Text)) Then
			lstUnusedDataExtDefs.Items.RemoveAt((lstUnusedDataExtDefs.SelectedIndex))
			txtDataExtValue.Text = ""
			txtDataExtValue.Refresh()
		End If
		
		cmdShowRequest.Enabled = True
		cmdShowResponse.Enabled = True
	End Sub
	
	Private Sub cmdCloseWindow_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCloseWindow.Click
		Me.Close()
	End Sub
	
	Private Sub cmdShowRequest_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdShowRequest.Click
		frmqbXMLDisplay.txtqbXML.Text = PrettyXMLString((modDataExtSample.strAddRequest))
        frmqbXMLDisplay.ShowDialog()
    End Sub

    Private Sub cmdShowResponse_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdShowResponse.Click
        frmqbXMLDisplay.txtqbXML.Text = PrettyXMLString((modDataExtSample.strAddResponse))
        frmqbXMLDisplay.ShowDialog()
    End Sub

    Private Sub frmAddDataExtension_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		GetCustomers(lstCustomers)
	End Sub

    Private Sub lstCustomers_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstCustomers.SelectedIndexChanged
        lstUnusedDataExtDefs.Items.Clear()
        FillUnusedDataExts(VB6.GetItemString(lstCustomers, lstCustomers.SelectedIndex), lstUnusedDataExtDefs)
    End Sub
End Class