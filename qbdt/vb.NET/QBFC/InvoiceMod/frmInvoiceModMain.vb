Option Strict Off
Option Explicit On
Friend Class frmInvoiceModMain
	Inherits System.Windows.Forms.Form
	'----------------------------------------------------------
	' Copyright © 2003-2013 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/rdmgr/?ID=100
	'
	'----------------------------------------------------------
	
	
	Private Sub cmdQBFC_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdQBFC.Click
		SetImplementation("QBFC")
		If QuickBooksVersionOK Then
            'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
            frmInvoiceQuery.ShowDialog()
            Me.Hide()
		End If
	End Sub
	
	Private Sub cmdQBXMLRP_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdQBXMLRP.Click
		SetImplementation("QBXMLRP")
		If QuickBooksVersionOK Then
            'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
            frmInvoiceQuery.ShowDialog()
            Me.Hide()
		End If
	End Sub
	
	Private Sub cmdQuit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdQuit.Click
		End
	End Sub
	
	Private Sub frmInvoiceModMain_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
        frmPatience.ShowDialog()
        'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
        frmModifying.ShowDialog()
    End Sub
End Class