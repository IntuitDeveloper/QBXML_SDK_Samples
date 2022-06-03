Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmModDataExtension
	Inherits System.Windows.Forms.Form
	'----------------------------------------------------------
	' Form: frmModDataExtension
	'
	' Description: Allows the user to highlight a customer which brings
	'              up a list of used data extensions for that customer.
	'              The user may then highlight a data extension and modify
	'              the value given for that customer.
	'
	' Copyright © 2021-2022 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	'----------------------------------------------------------
	Private Sub cmdCloseWindow_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCloseWindow.Click
		Me.Close()
	End Sub
	
	Private Sub cmdModRequest_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdModRequest.Click
		frmqbXMLDisplay.txtqbXML.Text = PrettyXMLString((modDataExtSample.strModRequest))
        frmqbXMLDisplay.ShowDialog()
    End Sub
	
	Private Sub cmdModResponse_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdModResponse.Click
		frmqbXMLDisplay.txtqbXML.Text = PrettyXMLString((modDataExtSample.strModResponse))
        frmqbXMLDisplay.ShowDialog()
    End Sub
	
	Private Sub cmdModValue_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdModValue.Click
		If lstUsedDataExts.SelectedIndex < 0 Then
			MsgBox("You must select a data extension to modify")
			Exit Sub
		End If
		
		If txtDataExtValue.Text = "" Then
			MsgBox("You must select a data extension to modify")
			Exit Sub
		End If
		
		If ModDataExt(VB6.GetItemString(lstCustomers, lstCustomers.SelectedIndex), VB.Left(VB6.GetItemString(lstUsedDataExts, lstUsedDataExts.SelectedIndex), InStr(VB6.GetItemString(lstUsedDataExts, lstUsedDataExts.SelectedIndex), "  ") - 1), (txtDataExtValue.Text)) Then
			
			GetUsedCustomerDataExts(VB6.GetItemString(lstCustomers, lstCustomers.SelectedIndex), lstUsedDataExts, True)
		End If
		
		cmdModRequest.Enabled = True
		cmdModResponse.Enabled = True
	End Sub
	
	Private Sub frmModDataExtension_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		GetCustomers(lstCustomers)
	End Sub

    Private Sub lstCustomers_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstCustomers.SelectedIndexChanged
        lstUsedDataExts.Items.Clear()
        lstUsedDataExts.Refresh()

        GetUsedCustomerDataExts(VB6.GetItemString(lstCustomers, lstCustomers.SelectedIndex), lstUsedDataExts, True)

        If lstUsedDataExts.Items.Count = 0 Then
            lstUsedDataExts.Items.Add("This customer has no data extensions added to their record")
        End If
    End Sub
End Class