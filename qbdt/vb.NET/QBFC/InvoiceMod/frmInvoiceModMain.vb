Option Strict Off
Option Explicit On
Friend Class frmInvoiceModMain
	Inherits System.Windows.Forms.Form
	'----------------------------------------------------------
	' Copyright © 2021-2022 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/rdmgr/?ID=100
	'
	'----------------------------------------------------------
	
	
	Private Sub cmdQBFC_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdQBFC.Click
		SetImplementation("QBFC")
		If QuickBooksVersionOK Then
            
            frmInvoiceQuery.ShowDialog()
            Me.Hide()
		End If
	End Sub
	
	Private Sub cmdQBXMLRP_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdQBXMLRP.Click
		SetImplementation("QBXMLRP")
		If QuickBooksVersionOK Then
            
            frmInvoiceQuery.ShowDialog()
            Me.Hide()
		End If
	End Sub
	
	Private Sub cmdQuit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdQuit.Click
		End
	End Sub
	
	Private Sub frmInvoiceModMain_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        
        frmPatience.ShowDialog()
        
        frmModifying.ShowDialog()
    End Sub
End Class