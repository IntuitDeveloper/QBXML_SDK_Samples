Public Class MainForm
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents CloseWindowButton As System.Windows.Forms.Button
    Friend WithEvents ModItemButton As System.Windows.Forms.Button
    Friend WithEvents ItemInventoryListBox As System.Windows.Forms.ListBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents NameTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents IsActiveCheckBox As System.Windows.Forms.CheckBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents ReorderPointTextBox As System.Windows.Forms.TextBox
    Friend WithEvents PurchaseCostTextBox As System.Windows.Forms.TextBox
    Friend WithEvents PurchaseDescTextBox As System.Windows.Forms.TextBox
    Friend WithEvents SalesPriceTextBox As System.Windows.Forms.TextBox
    Friend WithEvents SalesDescTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents EditSequenceTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ParentComboBox As System.Windows.Forms.ComboBox
    Friend WithEvents SalesTaxCodeComboBox As System.Windows.Forms.ComboBox
    Friend WithEvents COGSAccountComboBox As System.Windows.Forms.ComboBox
    Friend WithEvents PrefVendorComboBox As System.Windows.Forms.ComboBox
    Friend WithEvents AssetAccountComboBox As System.Windows.Forms.ComboBox
    Friend WithEvents ClearParent As System.Windows.Forms.Button
    Friend WithEvents ClearSalesTaxCode As System.Windows.Forms.Button
    Friend WithEvents ClearPrefVendor As System.Windows.Forms.Button
    Friend WithEvents ReFillLists As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.ItemInventoryListBox = New System.Windows.Forms.ListBox()
        Me.CloseWindowButton = New System.Windows.Forms.Button()
        Me.ModItemButton = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.ClearPrefVendor = New System.Windows.Forms.Button()
        Me.ClearSalesTaxCode = New System.Windows.Forms.Button()
        Me.ClearParent = New System.Windows.Forms.Button()
        Me.AssetAccountComboBox = New System.Windows.Forms.ComboBox()
        Me.PrefVendorComboBox = New System.Windows.Forms.ComboBox()
        Me.COGSAccountComboBox = New System.Windows.Forms.ComboBox()
        Me.SalesTaxCodeComboBox = New System.Windows.Forms.ComboBox()
        Me.ParentComboBox = New System.Windows.Forms.ComboBox()
        Me.ReorderPointTextBox = New System.Windows.Forms.TextBox()
        Me.PurchaseCostTextBox = New System.Windows.Forms.TextBox()
        Me.PurchaseDescTextBox = New System.Windows.Forms.TextBox()
        Me.SalesPriceTextBox = New System.Windows.Forms.TextBox()
        Me.SalesDescTextBox = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.IsActiveCheckBox = New System.Windows.Forms.CheckBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.NameTextBox = New System.Windows.Forms.TextBox()
        Me.EditSequenceTextBox = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.ReFillLists = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ItemInventoryListBox
        '
        Me.ItemInventoryListBox.Location = New System.Drawing.Point(24, 40)
        Me.ItemInventoryListBox.Name = "ItemInventoryListBox"
        Me.ItemInventoryListBox.Size = New System.Drawing.Size(424, 69)
        Me.ItemInventoryListBox.TabIndex = 0
        '
        'CloseWindowButton
        '
        Me.CloseWindowButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.CloseWindowButton.Location = New System.Drawing.Point(322, 472)
        Me.CloseWindowButton.Name = "CloseWindowButton"
        Me.CloseWindowButton.Size = New System.Drawing.Size(88, 23)
        Me.CloseWindowButton.TabIndex = 1
        Me.CloseWindowButton.Text = "CloseWindow"
        '
        'ModItemButton
        '
        Me.ModItemButton.Location = New System.Drawing.Point(62, 472)
        Me.ModItemButton.Name = "ModItemButton"
        Me.ModItemButton.Size = New System.Drawing.Size(88, 23)
        Me.ModItemButton.TabIndex = 2
        Me.ModItemButton.Text = "Modify Item"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Item Inventory List:"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.ClearPrefVendor, Me.ClearSalesTaxCode, Me.ClearParent, Me.AssetAccountComboBox, Me.PrefVendorComboBox, Me.COGSAccountComboBox, Me.SalesTaxCodeComboBox, Me.ParentComboBox, Me.ReorderPointTextBox, Me.PurchaseCostTextBox, Me.PurchaseDescTextBox, Me.SalesPriceTextBox, Me.SalesDescTextBox, Me.Label13, Me.Label12, Me.Label11, Me.Label10, Me.Label9, Me.Label8, Me.Label7, Me.Label6, Me.Label5, Me.Label4, Me.IsActiveCheckBox, Me.Label2, Me.NameTextBox})
        Me.GroupBox1.Location = New System.Drawing.Point(24, 144)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(432, 320)
        Me.GroupBox1.TabIndex = 28
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "You can modify the following fields:"
        '
        'ClearPrefVendor
        '
        Me.ClearPrefVendor.Location = New System.Drawing.Point(312, 240)
        Me.ClearPrefVendor.Name = "ClearPrefVendor"
        Me.ClearPrefVendor.Size = New System.Drawing.Size(96, 23)
        Me.ClearPrefVendor.TabIndex = 57
        Me.ClearPrefVendor.Text = "Clear Selection"
        '
        'ClearSalesTaxCode
        '
        Me.ClearSalesTaxCode.Location = New System.Drawing.Point(312, 96)
        Me.ClearSalesTaxCode.Name = "ClearSalesTaxCode"
        Me.ClearSalesTaxCode.Size = New System.Drawing.Size(96, 23)
        Me.ClearSalesTaxCode.TabIndex = 55
        Me.ClearSalesTaxCode.Text = "Clear Selection"
        '
        'ClearParent
        '
        Me.ClearParent.Location = New System.Drawing.Point(312, 72)
        Me.ClearParent.Name = "ClearParent"
        Me.ClearParent.Size = New System.Drawing.Size(96, 23)
        Me.ClearParent.TabIndex = 54
        Me.ClearParent.Text = "Clear Selection"
        '
        'AssetAccountComboBox
        '
        Me.AssetAccountComboBox.AllowDrop = True
        Me.AssetAccountComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.AssetAccountComboBox.Location = New System.Drawing.Point(112, 264)
        Me.AssetAccountComboBox.Name = "AssetAccountComboBox"
        Me.AssetAccountComboBox.Size = New System.Drawing.Size(192, 21)
        Me.AssetAccountComboBox.TabIndex = 53
        '
        'PrefVendorComboBox
        '
        Me.PrefVendorComboBox.AllowDrop = True
        Me.PrefVendorComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.PrefVendorComboBox.Location = New System.Drawing.Point(112, 240)
        Me.PrefVendorComboBox.Name = "PrefVendorComboBox"
        Me.PrefVendorComboBox.Size = New System.Drawing.Size(192, 21)
        Me.PrefVendorComboBox.TabIndex = 52
        '
        'COGSAccountComboBox
        '
        Me.COGSAccountComboBox.AllowDrop = True
        Me.COGSAccountComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.COGSAccountComboBox.Location = New System.Drawing.Point(112, 216)
        Me.COGSAccountComboBox.Name = "COGSAccountComboBox"
        Me.COGSAccountComboBox.Size = New System.Drawing.Size(192, 21)
        Me.COGSAccountComboBox.TabIndex = 51
        '
        'SalesTaxCodeComboBox
        '
        Me.SalesTaxCodeComboBox.AllowDrop = True
        Me.SalesTaxCodeComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.SalesTaxCodeComboBox.Location = New System.Drawing.Point(112, 96)
        Me.SalesTaxCodeComboBox.Name = "SalesTaxCodeComboBox"
        Me.SalesTaxCodeComboBox.Size = New System.Drawing.Size(192, 21)
        Me.SalesTaxCodeComboBox.TabIndex = 50
        '
        'ParentComboBox
        '
        Me.ParentComboBox.AllowDrop = True
        Me.ParentComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ParentComboBox.Location = New System.Drawing.Point(112, 72)
        Me.ParentComboBox.Name = "ParentComboBox"
        Me.ParentComboBox.Size = New System.Drawing.Size(192, 21)
        Me.ParentComboBox.TabIndex = 49
        '
        'ReorderPointTextBox
        '
        Me.ReorderPointTextBox.Location = New System.Drawing.Point(112, 288)
        Me.ReorderPointTextBox.Name = "ReorderPointTextBox"
        Me.ReorderPointTextBox.Size = New System.Drawing.Size(192, 20)
        Me.ReorderPointTextBox.TabIndex = 48
        Me.ReorderPointTextBox.Text = ""
        '
        'PurchaseCostTextBox
        '
        Me.PurchaseCostTextBox.Location = New System.Drawing.Point(112, 192)
        Me.PurchaseCostTextBox.Name = "PurchaseCostTextBox"
        Me.PurchaseCostTextBox.Size = New System.Drawing.Size(192, 20)
        Me.PurchaseCostTextBox.TabIndex = 44
        Me.PurchaseCostTextBox.Text = ""
        '
        'PurchaseDescTextBox
        '
        Me.PurchaseDescTextBox.Location = New System.Drawing.Point(112, 168)
        Me.PurchaseDescTextBox.Name = "PurchaseDescTextBox"
        Me.PurchaseDescTextBox.Size = New System.Drawing.Size(192, 20)
        Me.PurchaseDescTextBox.TabIndex = 43
        Me.PurchaseDescTextBox.Text = ""
        '
        'SalesPriceTextBox
        '
        Me.SalesPriceTextBox.Location = New System.Drawing.Point(112, 144)
        Me.SalesPriceTextBox.Name = "SalesPriceTextBox"
        Me.SalesPriceTextBox.Size = New System.Drawing.Size(192, 20)
        Me.SalesPriceTextBox.TabIndex = 42
        Me.SalesPriceTextBox.Text = ""
        '
        'SalesDescTextBox
        '
        Me.SalesDescTextBox.Location = New System.Drawing.Point(112, 120)
        Me.SalesDescTextBox.Name = "SalesDescTextBox"
        Me.SalesDescTextBox.Size = New System.Drawing.Size(192, 20)
        Me.SalesDescTextBox.TabIndex = 41
        Me.SalesDescTextBox.Text = ""
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(8, 288)
        Me.Label13.Name = "Label13"
        Me.Label13.TabIndex = 38
        Me.Label13.Text = "ReorderPoint:"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(8, 264)
        Me.Label12.Name = "Label12"
        Me.Label12.TabIndex = 37
        Me.Label12.Text = "AssetAccount:"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(8, 240)
        Me.Label11.Name = "Label11"
        Me.Label11.TabIndex = 36
        Me.Label11.Text = "PrefVendor:"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(8, 216)
        Me.Label10.Name = "Label10"
        Me.Label10.TabIndex = 35
        Me.Label10.Text = "COGSAccount:"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(8, 192)
        Me.Label9.Name = "Label9"
        Me.Label9.TabIndex = 34
        Me.Label9.Text = "PurchaseCost:"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(8, 168)
        Me.Label8.Name = "Label8"
        Me.Label8.TabIndex = 33
        Me.Label8.Text = "PurchaseDesc:"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(8, 144)
        Me.Label7.Name = "Label7"
        Me.Label7.TabIndex = 32
        Me.Label7.Text = "SalesPrice:"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(8, 120)
        Me.Label6.Name = "Label6"
        Me.Label6.TabIndex = 31
        Me.Label6.Text = "SalesDesc:"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(8, 96)
        Me.Label5.Name = "Label5"
        Me.Label5.TabIndex = 30
        Me.Label5.Text = "SalesTaxCode:"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(8, 72)
        Me.Label4.Name = "Label4"
        Me.Label4.TabIndex = 29
        Me.Label4.Text = "Parent:"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'IsActiveCheckBox
        '
        Me.IsActiveCheckBox.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.IsActiveCheckBox.Location = New System.Drawing.Point(24, 48)
        Me.IsActiveCheckBox.Name = "IsActiveCheckBox"
        Me.IsActiveCheckBox.TabIndex = 28
        Me.IsActiveCheckBox.Text = "IsActive:      "
        Me.IsActiveCheckBox.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 18
        Me.Label2.Text = "Name:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'NameTextBox
        '
        Me.NameTextBox.Location = New System.Drawing.Point(112, 24)
        Me.NameTextBox.Name = "NameTextBox"
        Me.NameTextBox.Size = New System.Drawing.Size(192, 20)
        Me.NameTextBox.TabIndex = 17
        Me.NameTextBox.Text = ""
        '
        'EditSequenceTextBox
        '
        Me.EditSequenceTextBox.Location = New System.Drawing.Point(168, 8)
        Me.EditSequenceTextBox.Name = "EditSequenceTextBox"
        Me.EditSequenceTextBox.Size = New System.Drawing.Size(192, 20)
        Me.EditSequenceTextBox.TabIndex = 49
        Me.EditSequenceTextBox.Text = ""
        Me.EditSequenceTextBox.Visible = False
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(24, 112)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(344, 32)
        Me.Label3.TabIndex = 50
        Me.Label3.Text = "Click on an Item Inventory above to display the current values for the fields bel" & _
        "ow."
        '
        'ReFillLists
        '
        Me.ReFillLists.Location = New System.Drawing.Point(192, 472)
        Me.ReFillLists.Name = "ReFillLists"
        Me.ReFillLists.Size = New System.Drawing.Size(88, 23)
        Me.ReFillLists.TabIndex = 51
        Me.ReFillLists.Text = "Refill Lists"
        '
        'MainForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.CloseWindowButton
        Me.ClientSize = New System.Drawing.Size(472, 501)
        Me.ControlBox = False
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.ReFillLists, Me.Label3, Me.GroupBox1, Me.Label1, Me.ModItemButton, Me.CloseWindowButton, Me.ItemInventoryListBox, Me.EditSequenceTextBox})
        Me.Name = "MainForm"
        Me.Text = "Item Inventory Modification"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    '----------------------------------------------------------
    ' Form: mainForm
    '
    ' Description: This the main form and entry point for this sample
    '              program.  It displays the list of Item Inventory
    '              FullNames.  Then when an Item Inventory is selected,
    '              the field values are filled in the controls in the form.
    '              The user can change the values in the controls.  Then
    '              press the modify button to change the values of those
    '              fields for the selected Item Inventory.  The list box
    '              of Item Inventory FullNames are refreshed from QB 
    '
    '              The form calls OpenConnectionBeginSession to make sure
    '              a company file is open.
    '
    ' Copyright © 2002-2013 Intuit Inc. All rights reserved.
    ' Use is subject to the terms specified at:
    '      http://developer.intuit.com/legal/devsite_tos.html
    '
    '----------------------------------------------------------
    Private Sub MainForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        ' OpenConnection and BeginSession to QB
        OpenConnectionBeginSession()

        ReFillLists_Click(sender, e)
    End Sub

    Private Sub CloseWindowButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CloseWindowButton.Click

        ' EndSession and CloseConnection with QB
        EndSessionCloseConnection()

        End
    End Sub

    Private Sub ModItemButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ModItemButton.Click

        ' make sure an Item Inventory FullName is selected in the list box
        If ItemInventoryListBox.SelectedIndex < 0 Then
            MsgBox("You must select an Inventory Item first")
            Exit Sub
        End If

        Dim itemInfo, itemInfo2 As FullNameListIDClass
        Dim selectedListID As String

        ' get the information for the Inventory Item selected
        itemInfo = ItemInventoryListBox.SelectedItem()
        If (Not itemInfo Is Nothing) Then
            selectedListID = itemInfo.ListID

            ' modify the field values in QB for this Item Inventory
            ModifyItemInventoryFields(itemInfo, Me)

            ' have to refill the Item Inventory List Box because this list box
            ' has names of the items which could have changed if the 
            ' Name or Parent fields were modified.
            FillItemInventoryListBox(ItemInventoryListBox, ParentComboBox)

            ' reselect the list item that was modified
            Dim index As Short
            Dim count As Short
            Dim bItemFound As Boolean = False
            count = ItemInventoryListBox.Items.Count
            For index = 0 To count - 1
                itemInfo2 = ItemInventoryListBox.Items.Item(index)
                If (itemInfo2.ListID = selectedListID) Then
                    ItemInventoryListBox.SetSelected(index, True)
                    bItemFound = True
                    Exit For
                End If
            Next
            If (Not bItemFound) Then
                MsgBox("Could not find the modified item in the list box.  We will reset the field values.")
                ClearControls()
            End If
        End If
    End Sub

    Private Sub ItemInventoryListBox_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ItemInventoryListBox.SelectedIndexChanged

        Dim itemInfo As FullNameListIDClass

        ' determine which item is selected in the list box
        itemInfo = ItemInventoryListBox.SelectedItem()
        If (Not itemInfo Is Nothing) Then

            ' clear the values in the fields in the form
            ClearControls()

            ' reset the fields with the current values of the Item selected in the list box
            QueryItemInventoryFields(itemInfo, Me)
        End If
    End Sub

    Public Sub ClearControls()
        ReorderPointTextBox.Clear()
        AssetAccountComboBox.SelectedIndex = -1
        PrefVendorComboBox.SelectedIndex = -1
        COGSAccountComboBox.SelectedIndex = -1
        PurchaseCostTextBox.Clear()
        PurchaseDescTextBox.Clear()
        SalesPriceTextBox.Clear()
        SalesDescTextBox.Clear()
        SalesTaxCodeComboBox.SelectedIndex = -1
        ParentComboBox.SelectedIndex = -1
        NameTextBox.Clear()
        EditSequenceTextBox.Clear()
        IsActiveCheckBox.Checked = False

    End Sub

    Private Sub ClearParent_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ClearParent.Click
        ParentComboBox.SelectedIndex = -1
    End Sub

    Private Sub ClearSalesTaxCode_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ClearSalesTaxCode.Click
        SalesTaxCodeComboBox.SelectedIndex = -1
    End Sub

    Private Sub ClearPrefVendor_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ClearPrefVendor.Click
        PrefVendorComboBox.SelectedIndex = -1
    End Sub

    Private Sub ReFillLists_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ReFillLists.Click

        ' fill the list box with Item Inventory and Parent with FullNames
        FillItemInventoryListBox(ItemInventoryListBox, ParentComboBox)

        ' fill the SalesTaxCode Combo Box with the SalesTaxCode list of FullNames
        FillSalesTaxCodeList(SalesTaxCodeComboBox)

        ' fill the COGSAccount Combo Box with the COGSAccount names
        FillCOGSAccountList(COGSAccountComboBox)

        ' fill in the Preferred Vendor Combo Box with the Vendor names
        FillPrefVendorList(PrefVendorComboBox)

        ' fill in the AssetAccount Combo box with the Asset Account names
        FillAssetAccountList(AssetAccountComboBox)

        ' clear the values in the fields in the form
        ClearControls()

    End Sub

End Class
