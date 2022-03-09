Option Strict Off
Option Explicit On
Friend Class InventoryAdjust
	Inherits System.Windows.Forms.Form
	'----------------------------------------------------------
	' Form: InventoryAdjust
	'
	' Description: This form allows the user to select the ItemInventory item,
	'              Customer, Class and Account, define changes to the Inventory
	'              Item and then add it to the currently open QuickBooks company
	'              file.
	'
	' Copyright © 2021-2022 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	'----------------------------------------------------------
	Dim bClassListFilled As Boolean
	Dim bCustomerListFilled As Boolean
	
	Private Sub AddBtn_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles AddBtn.Click
		On Error GoTo Errs
		
		Dim item As String
		Dim bQuantity As Boolean
		Dim bRelative As Boolean
		Dim li As System.Windows.Forms.ListViewItem
		bQuantity = QuantityOption.Checked
		item = ItemList.Text
		
		' item must be selected
		If (item = "") Then
			MsgBox("An Item must be selected.")
			Exit Sub
		End If
		
		li = LineItems.Items.Add(item, item, "")
		If (bQuantity) Then
			
			If li.SubItems.Count > 1 Then
				li.SubItems(1).Text = "Quantity"
			Else
				li.SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "Quantity"))
			End If
			bRelative = DiffCheck.CheckState
			If bRelative Then
				
				If li.SubItems.Count > 2 Then
					li.SubItems(2).Text = "Relative"
				Else
					li.SubItems.Insert(2, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "Relative"))
				End If
			Else
				
				If li.SubItems.Count > 2 Then
					li.SubItems(2).Text = "Absolute"
				Else
					li.SubItems.Insert(2, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "Absolute"))
				End If
			End If
			
			If li.SubItems.Count > 3 Then
				li.SubItems(3).Text = QuantityAdjust.Text
			Else
				li.SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, QuantityAdjust.Text))
			End If
		Else
			
			If li.SubItems.Count > 1 Then
				li.SubItems(1).Text = "Value"
			Else
				li.SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "Value"))
			End If
			
			If li.SubItems.Count > 2 Then
				li.SubItems(2).Text = ValueAdjust.Text
			Else
				li.SubItems.Insert(2, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, ValueAdjust.Text))
			End If
			
			If li.SubItems.Count > 3 Then
				li.SubItems(3).Text = QuantityAdjust.Text
			Else
				li.SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, QuantityAdjust.Text))
			End If
		End If
		Exit Sub
Errs: 
		If (Err.Number = 35602) Then
			MsgBox("You cannot use the same item in two different line items in this transaction")
		Else
			MsgBox("Error in AddBtn_Click" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)
		End If
	End Sub
	
	Private Sub ClassList_DropDown(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ClassList.DropDown
		If (bClassListFilled) Then
			Exit Sub
		End If
		fillClassList(ClassList)
		bClassListFilled = True
	End Sub
	
	Private Sub CustomerList_DropDown(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CustomerList.DropDown
		If (bCustomerListFilled) Then
			Exit Sub
		End If
		fillCustomerList(CustomerList)
		bCustomerListFilled = True
	End Sub
	
	Private Sub InventoryAdjust_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim booConnectionOpen As Object
		Dim booSessionBegun As Object
		
		booSessionBegun = False
		
		booConnectionOpen = False
		qbConnect()
		If ClassesEnabled() Then
			ClassList.Enabled = True
			ClassLabel.Enabled = True
		End If
		fillAccountList(AccountList)
		fillItemList(ItemList)
		bClassListFilled = False
		bCustomerListFilled = False
	End Sub
	
	Private Sub InventoryAdjust_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		EndSessionCloseConnection()
		End
	End Sub
	
	
	Private Sub QuantityOption_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles QuantityOption.CheckedChanged
		If eventSender.Checked Then
			ValueLabel.Visible = False
			ValueAdjust.Visible = False
			DiffCheck.Visible = True
		End If
	End Sub
	
	Private Sub SendBtn_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles SendBtn.Click
		Dim acct As String
		Dim cust As String
		
		Dim class_Renamed As String
		Dim thememo As String
		Dim resp As String
		acct = AccountList.Text
		cust = CustomerList.Text
		class_Renamed = ClassList.Text
		thememo = Memo.Text
		
		' the Account must be specified
		If (acct = "") Then
			MsgBox("An Account must be selected.")
			Exit Sub
		End If
		
		' there must be at least one line item
		If (LineItems.Items.Count <= 0) Then
			MsgBox("At least One Line Item must be specified.")
			Exit Sub
		End If
		
		resp = AdjustInventory(acct, cust, class_Renamed, thememo, LineItems)
		MsgBox(resp)
	End Sub
	
	
	Private Sub ValueOption_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ValueOption.CheckedChanged
		If eventSender.Checked Then
			DiffCheck.Visible = False
			ValueLabel.Visible = True
			ValueAdjust.Visible = True
		End If
	End Sub
End Class