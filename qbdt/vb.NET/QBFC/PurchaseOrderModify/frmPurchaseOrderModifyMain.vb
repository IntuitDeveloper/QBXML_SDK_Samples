Option Strict Off
Option Explicit On
Friend Class frmPurchaseOrderModifyMain
	Inherits System.Windows.Forms.Form
	'----------------------------------------------------------
	' Copyright © 2003-2013 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	'----------------------------------------------------------
	
	
	Private Sub cmdQBFC_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdQBFC.Click
		If Not DatesOK Then
			Exit Sub
		End If
		
		SetDates(CDate(txtFromDate.Text), CDate(txtToDate.Text))
		
		SetImplementation("QBFC")
		If QuickBooksVersionOK Then
            'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
            Dim form As New frmPurchaseOrderSelect
            form.Show()
            Me.Hide()
        End If
	End Sub
	
	Private Sub cmdQBXMLRP_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdQBXMLRP.Click
		If Not DatesOK Then
			Exit Sub
		End If
		
		SetDates(CDate(txtFromDate.Text), CDate(txtToDate.Text))
		
		SetImplementation("QBXMLRP")
		If QuickBooksVersionOK Then
            'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
            Dim form As New frmPurchaseOrderSelect
            form.Show()
            Me.Hide()
        End If
	End Sub
	
	Private Sub cmdQuit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdQuit.Click
		End
	End Sub
	
	Private Function DatesOK() As Boolean
		If Not IsDate(txtFromDate.Text) Then
			MsgBox("The From date you've supplied is illegal")
			DatesOK = False
			Exit Function
		End If
		
		If Not IsDate(txtToDate.Text) Then
			MsgBox("The To date you've supplied is illegal")
			DatesOK = False
			Exit Function
		End If
		
		If CDate(txtFromDate.Text) > CDate(txtToDate.Text) Then
			MsgBox("The To date you've supplied is earlier than the From date")
			DatesOK = False
			Exit Function
		End If
		
		If CDate(txtFromDate.Text) < CDate("1/1/1970") Or CDate(txtFromDate.Text) > CDate("1/18/2038") Or CDate(txtToDate.Text) > CDate("1/18/2038") Then
			MsgBox("The dates you supply must be between Jan 1, 1970 and Jan 18, 2038")
			DatesOK = False
			Exit Function
		End If
		
		DatesOK = True
	End Function
End Class