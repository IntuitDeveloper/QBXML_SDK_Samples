Option Strict Off
Option Explicit On
Friend Class frmPurchaseOrderDetail
	Inherits System.Windows.Forms.Form
	'----------------------------------------------------------
	' Copyright © 2003-2013 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	'----------------------------------------------------------
	
	
	Dim strPOLines(30) As String
	Dim strTxnID As String
	Dim strEditSequence As String
	Dim strRefNumber As String
	Dim strTxnDate As String
	Dim strVendor As String
	
	Dim intPOLines As Short
	Dim intHighlightedPOLineRow As Short
	Dim intActualHighlightedPOLine As Short
	Dim intTopDisplayedPOLine As Short
	
	
	Private Sub frmPurchaseOrderDetail_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		ClearstrPOLines()
		GetPOInfo(strTxnID, strEditSequence, strRefNumber, strTxnDate, strVendor, strPOLines)
		
		'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If IsNothing(strPOLines(1)) Then
			MsgBox("Error encountered getting PO lines, exiting")
			End
		End If
		
		FillPODisplay()
		intHighlightedPOLineRow = -1
		intActualHighlightedPOLine = -1
	End Sub
	
	
	Private Sub cmdReturn_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdReturn.Click
		frmPurchaseOrderSelect.RefreshPOList(False)
		Me.Close()
	End Sub
	
	
	Private Sub cmdExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdExit.Click
		EndSessionCloseConnection()
		End
	End Sub
	
	
	Private Sub cmdReceiveAll_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdReceiveAll.Click
		SetQuantitiesAndBillForRemainingItems(strTxnID, strEditSequence, strVendor, strRefNumber, strTxnDate, strPOLines, 0)
		frmPurchaseOrderSelect.RefreshPOList(False)
		Me.Close()
	End Sub
	
	
	Private Sub cmdReceiveOne_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdReceiveOne.Click
		
		If intActualHighlightedPOLine = -1 Then
			MsgBox("You must select a PO line before using this button")
			Exit Sub
		End If
		
		If LineType(strPOLines(intActualHighlightedPOLine)) = "GroupItem" Then
			MsgBox("You cannot receive a Group Item line" & vbCrLf & vbCrLf & "You can only receive the Item lines in a Group Item")
			Exit Sub
		End If
		
		SetQuantitiesAndBillForRemainingItems(strTxnID, strEditSequence, strVendor, strRefNumber, strTxnDate, strPOLines, intActualHighlightedPOLine)
		frmPurchaseOrderSelect.RefreshPOList(False)
		Me.Close()
	End Sub
	
	
	Private Sub cmdCloseLine_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCloseLine.Click
		
		If intActualHighlightedPOLine = -1 Then
			MsgBox("You must select a PO line before using this button")
			Exit Sub
		End If
		
		If LineType(strPOLines(intActualHighlightedPOLine)) = "GroupItem" Then
			MsgBox("You cannot close a Group Item line" & vbCrLf & vbCrLf & "You can only close the Item lines in a Group Item")
			Exit Sub
		End If
		
		If POLineClosed(strPOLines(intActualHighlightedPOLine)) Then
			MsgBox("You cannot close a line that's already closed")
			Exit Sub
		End If
		
		ChangePOLine("Close", strTxnID, strEditSequence, strPOLines, intActualHighlightedPOLine)
		UnHighlightPOLineRow(intHighlightedPOLineRow)
		frmPurchaseOrderSelect.RefreshPOList(False)
		Me.Close()
	End Sub
	
	
	Private Sub cmdCloseAll_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCloseAll.Click
		ClosePO((txtTxnID.Text), (txtEditSequence.Text))
		frmPurchaseOrderSelect.RefreshPOList(False)
		Me.Close()
	End Sub
	
	
	Private Sub ClearstrPOLines()
		
		Dim i As Short
		For i = 1 To UBound(strPOLines)
			strPOLines(i) = CStr(Nothing)
		Next i
	End Sub
	
	
	Private Sub FillPODisplay()
		
		txtTxnID.Text = strTxnID
		txtEditSequence.Text = strEditSequence
		txtRefNumber.Text = strRefNumber
		txtTxnDate.Text = strTxnDate
		txtVendor.Text = strVendor
		
		Dim strSplits() As String
		Dim i As Short
		Dim booDone As Boolean
		i = 1
		booDone = False
		
		intTopDisplayedPOLine = 1
		Do 
			If i < 11 Then
				strSplits = Split(strPOLines(i), "<spliter>")
				txtItem(i - 1).Text = strSplits(0)
				txtDescription(i - 1).Text = strSplits(1)
				txtQuantity(i - 1).Text = strSplits(2)
				txtRate(i - 1).Text = strSplits(3)
				txtAmount(i - 1).Text = strSplits(5)
				txtReceived(i - 1).Text = strSplits(6)
				txtClosed(i - 1).Text = strSplits(7)
			End If
			
			If i = 30 Then
				booDone = True
				'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			ElseIf IsNothing(strPOLines(i + 1)) Then 
				booDone = True
			Else
				i = i + 1
			End If
		Loop Until booDone
		
		intPOLines = i
		
		If intPOLines < 11 Then
			vscPOLineScroll.Enabled = False
			vscPOLineScroll.Visible = False
			vscPOLineScroll.Minimum = 1
			vscPOLineScroll.Value = 1
			Exit Sub
		End If
		
		vscPOLineScroll.Minimum = 1
		vscPOLineScroll.Maximum = (intPOLines - 9 + vscPOLineScroll.LargeChange - 1)
		vscPOLineScroll.Enabled = True
		vscPOLineScroll.Visible = True
	End Sub
	
	
	Private Sub HighlightPOLineRow(ByRef intRowNum As Short)
		If intHighlightedPOLineRow >= 0 Then
			UnHighlightPOLineRow(intHighlightedPOLineRow)
		End If
		
		txtItem(intRowNum).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFF00)
		txtSpace1(intRowNum).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFF00)
		txtDescription(intRowNum).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFF00)
		txtSpace2(intRowNum).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFF00)
		txtQuantity(intRowNum).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFF00)
		txtSpace3(intRowNum).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFF00)
		txtRate(intRowNum).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFF00)
		txtSpace4(intRowNum).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFF00)
		txtAmount(intRowNum).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFF00)
		txtReceived(intRowNum).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFF00)
		txtClosed(intRowNum).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFF00)
		
		intHighlightedPOLineRow = intRowNum
		intActualHighlightedPOLine = vscPOLineScroll.Value + intRowNum
	End Sub
	
	
	Private Sub UnHighlightPOLineRow(ByRef intRowNum As Short)
		If intHighlightedPOLineRow >= 0 Then
			txtItem(intRowNum).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			txtSpace1(intRowNum).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			txtDescription(intRowNum).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			txtSpace2(intRowNum).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			txtQuantity(intRowNum).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			txtSpace3(intRowNum).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			txtRate(intRowNum).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			txtSpace4(intRowNum).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			txtAmount(intRowNum).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			txtReceived(intRowNum).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			txtClosed(intRowNum).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
		End If
		
		intHighlightedPOLineRow = -1
		intActualHighlightedPOLine = -1
	End Sub
	
	Private Sub txtAmount_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAmount.Click
		Dim Index As Short = txtAmount.GetIndex(eventSender)
		If Index < intPOLines Then
			HighlightPOLineRow((Index))
		End If
	End Sub
	
	Private Sub txtClosed_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtClosed.Click
		Dim Index As Short = txtClosed.GetIndex(eventSender)
		If Index < intPOLines Then
			HighlightPOLineRow((Index))
		End If
	End Sub
	
	Private Sub txtDescription_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescription.Click
		Dim Index As Short = txtDescription.GetIndex(eventSender)
		If Index < intPOLines Then
			HighlightPOLineRow((Index))
		End If
	End Sub
	
	Private Sub txtItem_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtItem.Click
		Dim Index As Short = txtItem.GetIndex(eventSender)
		If Index < intPOLines Then
			HighlightPOLineRow((Index))
		End If
	End Sub
	
	Private Sub txtQuantity_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtQuantity.Click
		Dim Index As Short = txtQuantity.GetIndex(eventSender)
		If Index < intPOLines Then
			HighlightPOLineRow((Index))
		End If
	End Sub
	
	Private Sub txtRate_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRate.Click
		Dim Index As Short = txtRate.GetIndex(eventSender)
		If Index < intPOLines Then
			HighlightPOLineRow((Index))
		End If
	End Sub
	
	Private Sub txtReceived_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtReceived.Click
		Dim Index As Short = txtReceived.GetIndex(eventSender)
		If Index < intPOLines Then
			HighlightPOLineRow((Index))
		End If
	End Sub
	
	Private Sub txtSpace1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSpace1.Click
		Dim Index As Short = txtSpace1.GetIndex(eventSender)
		If Index < intPOLines Then
			HighlightPOLineRow((Index))
		End If
	End Sub
	
	Private Sub txtSpace2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSpace2.Click
		Dim Index As Short = txtSpace2.GetIndex(eventSender)
		If Index < intPOLines Then
			HighlightPOLineRow((Index))
		End If
	End Sub
	
	Private Sub txtSpace3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSpace3.Click
		Dim Index As Short = txtSpace3.GetIndex(eventSender)
		If Index < intPOLines Then
			HighlightPOLineRow((Index))
		End If
	End Sub
	
	Private Sub txtSpace4_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSpace4.Click
		Dim Index As Short = txtSpace4.GetIndex(eventSender)
		If Index < intPOLines Then
			HighlightPOLineRow((Index))
		End If
	End Sub
	
	'UPGRADE_NOTE: vscPOLineScroll.Change was changed from an event to a procedure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="4E2DC008-5EDA-4547-8317-C9316952674F"'
	'UPGRADE_WARNING: VScrollBar event vscPOLineScroll.Change has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub vscPOLineScroll_Change(ByVal newScrollValue As Integer)
		UnHighlightPOLineRow(intHighlightedPOLineRow)
		
		Dim strSplits() As String
		Dim i As Short
		intTopDisplayedPOLine = newScrollValue
		For i = 0 To 9
			'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If Not IsNothing(strPOLines(newScrollValue + i)) Then
				strSplits = Split(strPOLines(newScrollValue + i), "<spliter>")
				txtItem(i).Text = strSplits(0)
				txtDescription(i).Text = strSplits(1)
				txtQuantity(i).Text = strSplits(2)
				txtRate(i).Text = strSplits(3)
				txtAmount(i).Text = strSplits(5)
				txtReceived(i).Text = strSplits(6)
				txtClosed(i).Text = strSplits(7)
			End If
		Next 
		
		If intActualHighlightedPOLine >= newScrollValue And intActualHighlightedPOLine <= newScrollValue + 9 Then
			HighlightPOLineRow(intActualHighlightedPOLine - newScrollValue)
		End If
		
		If Me.Visible Then
			txtItem(0).Focus()
		End If
		
	End Sub
	
	
	Private Function POLineClosed(ByRef strLine As String) As Boolean
		
		Dim strSplits() As String
		strSplits = Split(strLine, "<spliter>")
		POLineClosed = (strSplits(7) = "X")
	End Function
	Private Sub vscPOLineScroll_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.ScrollEventArgs) Handles vscPOLineScroll.Scroll
		Select Case eventArgs.type
			Case System.Windows.Forms.ScrollEventType.EndScroll
				vscPOLineScroll_Change(eventArgs.newValue)
		End Select
	End Sub
End Class