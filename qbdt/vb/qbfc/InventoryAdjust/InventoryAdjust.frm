VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form InventoryAdjust 
   Caption         =   "InventoryAdjust"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   8625
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox WarningMsg 
      Height          =   735
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   21
      Text            =   "InventoryAdjust.frx":0000
      Top             =   0
      Width           =   8295
   End
   Begin VB.CommandButton SendBtn 
      Caption         =   "Send to QuickBooks!"
      Height          =   375
      Left            =   3105
      TabIndex        =   17
      Top             =   8040
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Line Item"
      Height          =   1935
      Left            =   360
      TabIndex        =   5
      Top             =   3000
      Width           =   7815
      Begin VB.CommandButton AddBtn 
         Caption         =   "Add Line Item"
         Height          =   375
         Left            =   5400
         TabIndex        =   15
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox ValueAdjust 
         Height          =   375
         Left            =   2760
         TabIndex        =   12
         Text            =   "0"
         Top             =   1440
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox QuantityAdjust 
         Height          =   375
         Left            =   1440
         TabIndex        =   10
         Text            =   "0"
         Top             =   1440
         Width           =   855
      End
      Begin VB.ComboBox ItemList 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   480
         Width           =   5895
      End
      Begin VB.OptionButton ValueOption 
         Caption         =   "Value"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton QuantityOption 
         Caption         =   "Quantity"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.CheckBox DiffCheck 
         Caption         =   "Difference?"
         Height          =   375
         Left            =   2640
         TabIndex        =   14
         Top             =   1440
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.Label ValueLabel 
         Caption         =   "Adjust Value:"
         Height          =   495
         Left            =   2760
         TabIndex        =   13
         Top             =   960
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Adjust Quantity:"
         Height          =   495
         Left            =   1440
         TabIndex        =   11
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Select Item:"
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSComctlLib.ListView LineItems 
      Height          =   2295
      Left            =   360
      TabIndex        =   4
      Top             =   5520
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   4048
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      Enabled         =   0   'False
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Value/Quantity"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Type/Value"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Quantity"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox Memo 
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Text            =   "Memo"
      Top             =   2400
      Width           =   7815
   End
   Begin VB.ComboBox ClassList 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1920
      Width           =   5415
   End
   Begin VB.ComboBox CustomerList 
      Height          =   315
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1440
      Width           =   5415
   End
   Begin VB.ComboBox AccountList 
      Height          =   315
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   960
      Width           =   5415
   End
   Begin VB.Label ClassLabel 
      Caption         =   "Select Class[Optional]:"
      Enabled         =   0   'False
      Height          =   315
      Left            =   960
      TabIndex        =   20
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Select Customer [Optional]:"
      Height          =   315
      Left            =   720
      TabIndex        =   19
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Select Account:"
      Height          =   315
      Left            =   1440
      TabIndex        =   18
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Line Items:"
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   5280
      Width           =   1815
   End
End
Attribute VB_Name = "InventoryAdjust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------
' Form: InventoryAdjust
'
' Description: This form allows the user to select the ItemInventory item,
'              Customer, Class and Account, define changes to the Inventory
'              Item and then add it to the currently open QuickBooks company
'              file.
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'----------------------------------------------------------
Dim bClassListFilled As Boolean
Dim bCustomerListFilled As Boolean

Private Sub AddBtn_Click()
    On Error GoTo Errs
    
    Dim item As String
    Dim bQuantity As Boolean
    Dim bRelative As Boolean
    Dim li As ListItem
    bQuantity = QuantityOption.value
    item = ItemList.Text
    
    ' item must be selected
    If (item = "") Then
        MsgBox "An Item must be selected."
        Exit Sub
    End If
    
    Set li = LineItems.ListItems.Add(, item, item)
    If (bQuantity) Then
        li.SubItems(1) = "Quantity"
        bRelative = DiffCheck.value
        If bRelative Then
            li.SubItems(2) = "Relative"
        Else
            li.SubItems(2) = "Absolute"
        End If
        li.SubItems(3) = QuantityAdjust.Text
    Else
        li.SubItems(1) = "Value"
        li.SubItems(2) = ValueAdjust.Text
        li.SubItems(3) = QuantityAdjust.Text
    End If
    Exit Sub
Errs:
    If (Err.Number = 35602) Then
        MsgBox "You cannot use the same item in two different line items in this transaction"
    Else
        MsgBox "Error in AddBtn_Click" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description
    End If
End Sub

Private Sub ClassList_DropDown()
    If (bClassListFilled) Then
        Exit Sub
    End If
    fillClassList ClassList
    bClassListFilled = True
End Sub

Private Sub CustomerList_DropDown()
    If (bCustomerListFilled) Then
        Exit Sub
    End If
    fillCustomerList CustomerList
    bCustomerListFilled = True
End Sub

Private Sub Form_Load()
    booSessionBegun = False
    booConnectionOpen = False
    qbConnect
    If ClassesEnabled() Then
        ClassList.Enabled = True
        ClassLabel.Enabled = True
    End If
    fillAccountList AccountList
    fillItemList ItemList
    bClassListFilled = False
    bCustomerListFilled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    EndSessionCloseConnection
    End
End Sub

Private Sub QuantityOption_Click()
    ValueLabel.Visible = False
    ValueAdjust.Visible = False
    DiffCheck.Visible = True
End Sub

Private Sub SendBtn_Click()
    Dim acct As String
    Dim cust As String
    Dim class As String
    Dim thememo As String
    Dim resp As String
    acct = AccountList.Text
    cust = CustomerList.Text
    class = ClassList.Text
    thememo = Memo.Text
    
    ' the Account must be specified
    If (acct = "") Then
        MsgBox "An Account must be selected."
        Exit Sub
    End If
    
    ' there must be at least one line item
    If (LineItems.ListItems.count <= 0) Then
        MsgBox "At least One Line Item must be specified."
        Exit Sub
    End If
    
    resp = AdjustInventory(acct, cust, class, thememo, LineItems)
    MsgBox resp
End Sub

Private Sub ValueOption_Click()
    DiffCheck.Visible = False
    ValueLabel.Visible = True
    ValueAdjust.Visible = True
End Sub

