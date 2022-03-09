Option Strict Off
Option Explicit On
Friend Class Macro
	Inherits System.Windows.Forms.Form
	'----------------------------------------------------------
	' Form: Macro
	'
	' Description: This sample program will show how to use the
	'              "UseMacro" and "DefMacro".  The "Send One
	'              Request" Button will build the InvoiceAdd
	'              and ReceivePayment transactions in one
	'              request.  The ReceivePayment uses the
	'              UseMacro to refer to the InvoiceAdd
	'              transaction.  The "Send Two Requests" button
	'              will show that the DefMacro value is
	'              persistent over separate requests.
	'              Information from an Estimate transaction
	'              is used to build the InvoiceAdd request.
	'
	' Copyright © 2021-2022 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	'----------------------------------------------------------
	
	Private Sub CloseWindow_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CloseWindow.Click
		EndSessionCloseConnection()
		End
	End Sub
	
	Private Sub Macro_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim booConnectionOpened As Object
		Dim booSessionBegun As Object
		
		booSessionBegun = False
		
		booConnectionOpened = False
		OpenConnectionBeginSession()
		GetEstimates(estimatesList)
	End Sub
	
	Private Sub SendOneRequest_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles SendOneRequest.Click
		Dim index As Short = SendOneRequest.GetIndex(eventSender)
		Dim i As Short
		i = estimatesList.SelectedIndex
		If (i <> -1) Then
			SendOneQBRequest(i)
		Else
			MsgBox("An Estimate must be selected in the List Box.")
		End If
	End Sub
	
	Private Sub SendTwoRequests_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles SendTwoRequests.Click
		Dim index As Short = SendTwoRequests.GetIndex(eventSender)
		Dim i As Short
		i = estimatesList.SelectedIndex
		If (i <> -1) Then
			SendTwoQBRequests(i)
		Else
			MsgBox("An Estimate must be selected in the List Box.")
		End If
	End Sub
End Class