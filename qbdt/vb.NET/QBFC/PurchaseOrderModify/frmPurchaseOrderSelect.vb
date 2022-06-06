Option Strict Off
Option Explicit On
Friend Class frmPurchaseOrderSelect
	Inherits System.Windows.Forms.Form
	'----------------------------------------------------------
	' Copyright © 2003-2013 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	'----------------------------------------------------------
	
	
	Dim strPOInfo(30) As String
	Dim intPOCount As Short
	Dim intHighlightedPORow As Short
	Dim intActualHighlightedPO As Short
	Dim intTopDisplayedPO As Short
	
	Private Sub frmPurchaseOrderSelect_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		intHighlightedPORow = -1
		intActualHighlightedPO = -1
		
		RefreshPOList(True)
	End Sub
	
	
	Private Sub cmdExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdExit.Click
		EndSessionCloseConnection()
		End
	End Sub


    Private Sub cmdPODetails_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPODetails.Click
        If intActualHighlightedPO = -1 Then
            MsgBox("You must select a Purchase Order before you can view it's details")
            Exit Sub
        End If

        SetSelectedPOInfo(strPOInfo(intActualHighlightedPO))
        'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
        Dim form As New frmPurchaseOrderDetail
        form.Show()
    End Sub


    Private Sub FillPOInfoDisplay()
		
		intPOCount = 0
		
		If InStr(1, strPOInfo(1), "There were no open purchase orders returned") > 0 Then
			txtPOVendor(0).Text = "There were no open purchase orders returned"
			vscPOScroll.Enabled = False
			vscPOScroll.Visible = False
		End If
		
		intTopDisplayedPO = 1
		Dim strPOInfoSplit() As String
		Dim booDone As Boolean
		booDone = False
		Do 
			intPOCount = intPOCount + 1
			If intPOCount < 11 Then
				strPOInfoSplit = Split(strPOInfo(intPOCount), "<spliter>")
				txtPONumber(intPOCount - 1).Text = strPOInfoSplit(0)
				txtPODate(intPOCount - 1).Text = strPOInfoSplit(1)
				txtPOVendor(intPOCount - 1).Text = strPOInfoSplit(2)
				txtPODeliveryDate(intPOCount - 1).Text = strPOInfoSplit(3)
				txtPOAmount(intPOCount - 1).Text = strPOInfoSplit(4)
			End If
			If intPOCount = UBound(strPOInfo) Then
				booDone = True
				'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			ElseIf IsNothing(strPOInfo(intPOCount + 1)) Then 
				booDone = True
			End If
		Loop Until booDone
		
		vscPOScroll.Value = 1
		If intPOCount < 11 Then
			vscPOScroll.Enabled = False
			vscPOScroll.Visible = False
			Exit Sub
		End If
		
		vscPOScroll.Minimum = 1
		vscPOScroll.Maximum = (intPOCount - 9 + vscPOScroll.LargeChange - 1)
		vscPOScroll.Enabled = True
		vscPOScroll.Visible = True
		
	End Sub
	
	
	Private Sub ClearPOInfo()
		Dim i As Short
		For i = 1 To 30
			strPOInfo(i) = CStr(Nothing)
		Next 
		
		For i = 0 To 9
			txtPONumber(i).Text = CStr(Nothing)
			txtPODate(i).Text = CStr(Nothing)
			txtPOVendor(i).Text = CStr(Nothing)
			txtPODeliveryDate(i).Text = CStr(Nothing)
			txtPOAmount(i).Text = CStr(Nothing)
		Next i
	End Sub
	
	Private Sub HighlightPORow(ByRef intRowNum As Short)
		If intHighlightedPORow >= 0 Then
			UnHighlightPORow(intHighlightedPORow)
		End If
		
		txtPONumber(intRowNum).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFF00)
		txtSpace1(intRowNum).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFF00)
		txtPODate(intRowNum).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFF00)
		txtSpace2(intRowNum).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFF00)
		txtPOVendor(intRowNum).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFF00)
		txtPODeliveryDate(intRowNum).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFF00)
		txtPOAmount(intRowNum).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFF00)
		
		intHighlightedPORow = intRowNum
		intActualHighlightedPO = vscPOScroll.Value + intRowNum
	End Sub
	
	
	Private Sub UnHighlightPORow(ByRef intRowNum As Short)
		If intHighlightedPORow >= 0 Then
			txtPONumber(intRowNum).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			txtSpace1(intRowNum).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			txtPODate(intRowNum).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			txtSpace2(intRowNum).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			txtPOVendor(intRowNum).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			txtPODeliveryDate(intRowNum).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			txtPOAmount(intRowNum).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
		End If
		
		intHighlightedPORow = -1
	End Sub
	
	Private Sub txtPOAmount_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPOAmount.Click
		Dim Index As Short = txtPOAmount.GetIndex(eventSender)
		If Index < intPOCount Then
			HighlightPORow((Index))
		End If
	End Sub
	
	Private Sub txtPODate_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPODate.Click
		Dim Index As Short = txtPODate.GetIndex(eventSender)
		If Index < intPOCount Then
			HighlightPORow((Index))
		End If
	End Sub
	
	Private Sub txtPODeliveryDate_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPODeliveryDate.Click
		Dim Index As Short = txtPODeliveryDate.GetIndex(eventSender)
		If Index < intPOCount Then
			HighlightPORow((Index))
		End If
	End Sub
	
	Private Sub txtPONumber_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPONumber.Click
		Dim Index As Short = txtPONumber.GetIndex(eventSender)
		If Index < intPOCount Then
			HighlightPORow((Index))
		End If
	End Sub
	
	Private Sub txtPOVendor_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPOVendor.Click
		Dim Index As Short = txtPOVendor.GetIndex(eventSender)
		If Index < intPOCount Then
			HighlightPORow((Index))
		End If
	End Sub
	
	Private Sub txtSpace1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSpace1.Click
		Dim Index As Short = txtSpace1.GetIndex(eventSender)
		If Index < intPOCount Then
			HighlightPORow((Index))
		End If
	End Sub
	
	Private Sub txtSpace2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSpace2.Click
		Dim Index As Short = txtSpace2.GetIndex(eventSender)
		If Index < intPOCount Then
			HighlightPORow((Index))
		End If
	End Sub
	
	'UPGRADE_NOTE: vscPOScroll.Change was changed from an event to a procedure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="4E2DC008-5EDA-4547-8317-C9316952674F"'
	'UPGRADE_WARNING: VScrollBar event vscPOScroll.Change has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub vscPOScroll_Change(ByVal newScrollValue As Integer)
		UnHighlightPORow(intHighlightedPORow)
		
		Dim strPOInfoSplit() As String
		Dim i As Short
		intTopDisplayedPO = newScrollValue
		For i = 0 To 9
			'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If Not IsNothing(strPOInfo(newScrollValue + i)) Then
				strPOInfoSplit = Split(strPOInfo(newScrollValue + i), "<spliter>")
				txtPONumber(i).Text = strPOInfoSplit(0)
				txtPODate(i).Text = strPOInfoSplit(1)
				txtPOVendor(i).Text = strPOInfoSplit(2)
				txtPODeliveryDate(i).Text = strPOInfoSplit(3)
				txtPOAmount(i).Text = strPOInfoSplit(4)
			End If
		Next 
		
		If intActualHighlightedPO >= newScrollValue And intActualHighlightedPO <= newScrollValue + 9 Then
			HighlightPORow(intActualHighlightedPO - newScrollValue)
		End If
		
		If Me.Visible Then
			txtPONumber(0).Focus()
		End If
	End Sub
	
	Public Sub RefreshPOList(ByRef booGiveWarning As Boolean)
		
		If intHighlightedPORow >= 0 Then
			UnHighlightPORow(intHighlightedPORow)
		End If
		
		ClearPOInfo()
		LoadPOInfoArray(strPOInfo, booGiveWarning)
		FillPOInfoDisplay()
		
		intHighlightedPORow = -1
		intActualHighlightedPO = -1
	End Sub
	Private Sub vscPOScroll_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.ScrollEventArgs) Handles vscPOScroll.Scroll
		Select Case eventArgs.type
			Case System.Windows.Forms.ScrollEventType.EndScroll
				vscPOScroll_Change(eventArgs.newValue)
		End Select
	End Sub
End Class