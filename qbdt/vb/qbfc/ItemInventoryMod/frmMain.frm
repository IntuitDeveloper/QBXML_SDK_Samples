VERSION 5.00
Begin VB.Form MainForm 
   Caption         =   "Item Inventory Modification"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   6240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CloseWindowButton 
      Cancel          =   -1  'True
      Caption         =   "Close Window"
      Height          =   495
      Left            =   4373
      TabIndex        =   33
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton ReFillLists 
      Caption         =   "Refill Lists"
      Height          =   495
      Left            =   2393
      TabIndex        =   32
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton ModItemButton 
      Caption         =   "Modify Item"
      Height          =   495
      Left            =   413
      TabIndex        =   31
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "You can modify the following fields:"
      Height          =   4695
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   5775
      Begin VB.CheckBox IsActiveCheckBox 
         Caption         =   "Check1"
         Height          =   255
         Left            =   1560
         TabIndex        =   34
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton ClearPrefVendor 
         Caption         =   "Clear Selection"
         Height          =   375
         Left            =   4200
         TabIndex        =   30
         Top             =   3600
         Width           =   1335
      End
      Begin VB.CommandButton ClearSalesTaxCode 
         Caption         =   "Clear Selection"
         Height          =   375
         Left            =   4200
         TabIndex        =   29
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton ClearParent 
         Caption         =   "Clear Selection"
         Height          =   375
         Left            =   4200
         TabIndex        =   28
         Top             =   1080
         Width           =   1335
      End
      Begin VB.ComboBox AssetAccountComboBox 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   3960
         Width           =   2415
      End
      Begin VB.ComboBox PrefVendorComboBox 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   3600
         Width           =   2415
      End
      Begin VB.ComboBox COGSAccountComboBox 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   3240
         Width           =   2415
      End
      Begin VB.ComboBox SalesTaxCodeComboBox 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   1440
         Width           =   2415
      End
      Begin VB.ComboBox ParentComboBox 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox ReorderPointTextBox 
         Height          =   285
         Left            =   1560
         TabIndex        =   22
         Top             =   4320
         Width           =   2415
      End
      Begin VB.TextBox PurchaseCostTextBox 
         Height          =   285
         Left            =   1560
         TabIndex        =   21
         Top             =   2880
         Width           =   2415
      End
      Begin VB.TextBox PurchaseDescTextBox 
         Height          =   285
         Left            =   1560
         TabIndex        =   20
         Top             =   2520
         Width           =   2415
      End
      Begin VB.TextBox SalesPriceTextBox 
         Height          =   285
         Left            =   1560
         TabIndex        =   19
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox SalesDescTextBox 
         Height          =   285
         Left            =   1560
         TabIndex        =   18
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox NameTextBox 
         Height          =   285
         Left            =   1560
         TabIndex        =   17
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Parent:"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "SalesTaxCode:"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "SalesDesc:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "SalesPrice:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "PurchaseDesc:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "PurchaseCost:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "COGSAccount:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "PrefVendor:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "AssetAccount:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "ReorderPoint:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "IsActive:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.ListBox ItemInventoryListBox 
      Height          =   1035
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   5775
   End
   Begin VB.TextBox EditSequenceTextBox 
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "Click on an Item Inventory above to display the current values for the fields below."
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   5895
   End
   Begin VB.Label Label1 
      Caption         =   "Item Inventory List:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub ClearParent_Click()
        ParentComboBox.listIndex = -1
End Sub

Private Sub ClearPrefVendor_Click()
        PrefVendorComboBox.listIndex = -1
End Sub

Private Sub ClearSalesTaxCode_Click()
        SalesTaxCodeComboBox.listIndex = -1
End Sub

Private Sub CloseWindowButton_Click()

        ' EndSession and CloseConnection with QB
        EndSessionCloseConnection

        End

End Sub

Private Sub Form_Load()

        booSessionBegun = False
        booConnectionOpened = False
        
        ' OpenConnection and BeginSession to QB
        OpenConnectionBeginSession

        ReFillLists_Click
        
End Sub

    Public Sub ClearControls()
        ReorderPointTextBox.Text = ""
        AssetAccountComboBox.listIndex = -1
        PrefVendorComboBox.listIndex = -1
        COGSAccountComboBox.listIndex = -1
        PurchaseCostTextBox.Text = ""
        PurchaseDescTextBox.Text = ""
        SalesPriceTextBox.Text = ""
        SalesDescTextBox.Text = ""
        SalesTaxCodeComboBox.listIndex = -1
        ParentComboBox.listIndex = -1
        NameTextBox.Text = ""
        EditSequenceTextBox.Text = ""
        IsActiveCheckBox.value = 0

    End Sub

Private Sub ItemInventoryListBox_Click()

        ' determine which item is selected in the list box
        Dim index As Integer
        index = ItemInventoryListBox.listIndex
        If (index <> -1) Then
            ' clear the values in the fields in the form
            ClearControls

            ' reset the fields with the current values of the Item selected in the list box
            QueryItemInventoryFields index, Me
        End If
End Sub

Private Sub ModItemButton_Click()

        Dim index As Integer
        index = ItemInventoryListBox.listIndex
        
        ' make sure an Item Inventory FullName is selected in the list box
        If index < 0 Then
            MsgBox ("You must select an Inventory Item first")
            Exit Sub
        End If

        Dim itemInfo As FullNameListIDClass

        ' get the information for the Inventory Item selected
        Set itemInfo = itemCollection.Item(index + 1)
        If (Not itemInfo Is Nothing) Then

            ' modify the field values in QB for this Item Inventory
            Dim fullName As String
            fullName = ModifyItemInventoryFields(itemInfo, Me)

            ' have to refill the Item Inventory List Box because this list box
            ' has names of the items which could have changed if the
            ' Name or Parent fields were modified.
            FillItemInventoryListBox ItemInventoryListBox, ParentComboBox

            ' reselect the list item that was modified
            Dim i As Integer
            Dim count As Integer
            Dim bItemFound As Boolean
            bItemFound = False
            count = ItemInventoryListBox.ListCount
            For i = 0 To count - 1
                If (ItemInventoryListBox.List(i) = fullName) Then
                    ItemInventoryListBox.listIndex = i
                    bItemFound = True
                    Exit For
                End If
            Next
            If (Not bItemFound) Then
                If (Not fullName = "") Then
                    MsgBox ("Could not find the modified item in the list box.  We will reset the field values.")
                End If
                ClearControls
            End If
        End If

End Sub

Private Sub ReFillLists_Click()
        ' fill the list box with Item Inventory and Parent with FullNames
        FillItemInventoryListBox ItemInventoryListBox, ParentComboBox

        ' fill the SalesTaxCode Combo Box with the SalesTaxCode list of FullNames
        FillSalesTaxCodeList SalesTaxCodeComboBox

        ' fill the COGSAccount Combo Box with the COGSAccount names
        FillCOGSAccountList COGSAccountComboBox

        ' fill in the Preferred Vendor Combo Box with the Vendor names
        FillPrefVendorList PrefVendorComboBox

        ' fill in the AssetAccount Combo box with the Asset Account names
        FillAssetAccountList AssetAccountComboBox

        ' clear the values in the fields in the form
        ClearControls

End Sub
