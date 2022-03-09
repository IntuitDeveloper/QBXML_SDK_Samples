Option Strict Off
Option Explicit On
Friend Class frmDepositAdd
	Inherits System.Windows.Forms.Form
	'-----------------------------------------------------------
	' Form Module: frmDepositAdd
	'
	' Description: this form is to display payments available for deposit
	'              and allow the user to deposit them.
	'
	' Copyright © 2021-2022 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	'----------------------------------------------------------
	
	Dim booFundsSelected As Boolean
	
	Private Sub cmdDepositFunds_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDepositFunds.Click
		If Not booFundsSelected Then
			MsgBox("You must select funds to deposit before attempting to deposit them.")
			Exit Sub
		End If
		
		DepositFunds(VB6.GetItemString(lstFundsForDeposit, lstFundsForDeposit.SelectedIndex))
		GetFundsForDeposit(lstFundsForDeposit)
	End Sub
	
	Private Sub cmdExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdExit.Click
		EndSessionCloseConnection()
		End
	End Sub
	
	Private Sub frmDepositAdd_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		booFundsSelected = False
		Connect()
		GetFundsForDeposit(lstFundsForDeposit)
	End Sub

    Private Sub lstFundsForDeposit_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstFundsForDeposit.SelectedIndexChanged
        booFundsSelected = True
    End Sub
End Class