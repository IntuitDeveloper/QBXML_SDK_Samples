VERSION 5.00
Begin VB.Form frmAddDataExtDef 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Data Extension Definition"
   ClientHeight    =   8790
   ClientLeft      =   3930
   ClientTop       =   1860
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   6465
   Begin VB.TextBox Text1 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   36
      TabStop         =   0   'False
      Text            =   "frmAddDataExtDef.frx":0000
      Top             =   600
      Width           =   6015
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "Clear All Checkboxes"
      Height          =   375
      Left            =   2280
      TabIndex        =   35
      Top             =   7200
      Width           =   1815
   End
   Begin VB.CommandButton cmdAllTransactions 
      Caption         =   "All Transactions"
      Height          =   375
      Left            =   4560
      TabIndex        =   34
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton cmdAllEntities 
      Caption         =   "All Entities"
      Height          =   375
      Left            =   3000
      TabIndex        =   33
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "All"
      Height          =   375
      Left            =   2160
      TabIndex        =   32
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Close Window"
      Height          =   615
      Left            =   4080
      TabIndex        =   31
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Data Extension Definition"
      Height          =   615
      Left            =   840
      TabIndex        =   30
      Top             =   7920
      Width           =   2895
   End
   Begin VB.CheckBox chkSalesTaxPaymentCheck 
      Caption         =   "Sales Tax Payment Check"
      Height          =   255
      Left            =   2160
      TabIndex        =   29
      Top             =   6600
      Width           =   2175
   End
   Begin VB.CheckBox chkSalesReceipt 
      Caption         =   "Sales Receipt"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   6600
      Width           =   1815
   End
   Begin VB.CheckBox chkReceivePayment 
      Caption         =   "Receive Payment"
      Height          =   255
      Left            =   4200
      TabIndex        =   27
      Top             =   6240
      Width           =   2055
   End
   Begin VB.CheckBox chkPurchaseOrder 
      Caption         =   "Purchase Order"
      Height          =   255
      Left            =   2160
      TabIndex        =   26
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CheckBox chkInvoice 
      Caption         =   "Invoice"
      Height          =   255
      Left            =   4200
      TabIndex        =   25
      Top             =   5880
      Width           =   2055
   End
   Begin VB.CheckBox chkJournalEntry 
      Caption         =   "Journal Entry"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CheckBox chkInventoryAdjustment 
      Caption         =   "Inventory Adjustment"
      Height          =   255
      Left            =   2160
      TabIndex        =   23
      Top             =   5880
      Width           =   1935
   End
   Begin VB.CheckBox chkEstimate 
      Caption         =   "Estimate"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CheckBox chkDeposit 
      Caption         =   "Deposit"
      Height          =   255
      Left            =   4200
      TabIndex        =   21
      Top             =   5520
      Width           =   2055
   End
   Begin VB.CheckBox chkCreditMemo 
      Caption         =   "Credit Memo"
      Height          =   255
      Left            =   2160
      TabIndex        =   20
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CheckBox chkCreditCardCredit 
      Caption         =   "Credit Card Credit"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CheckBox chkCreditCardCharge 
      Caption         =   "Credit Card Charge"
      Height          =   255
      Left            =   4200
      TabIndex        =   18
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CheckBox chkCheck 
      Caption         =   "Check"
      Height          =   255
      Left            =   2160
      TabIndex        =   17
      Top             =   5160
      Width           =   1935
   End
   Begin VB.CheckBox chkCharge 
      Caption         =   "Charge"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CheckBox chkBillPaymentCreditCard 
      Caption         =   "Bill Payment Credit Card"
      Height          =   255
      Left            =   4200
      TabIndex        =   15
      Top             =   4800
      Width           =   2055
   End
   Begin VB.CheckBox chkItem 
      Caption         =   "Item"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   4200
      Width           =   615
   End
   Begin VB.CheckBox chkBill 
      Caption         =   "Bill"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   4800
      Width           =   615
   End
   Begin VB.CheckBox chkAccount 
      Caption         =   "Account"
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CheckBox chkBillPaymentCheck 
      Caption         =   "Bill Payment Check"
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CheckBox chkOtherName 
      Caption         =   "Other Name"
      Height          =   255
      Left            =   3360
      TabIndex        =   10
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CheckBox chkEmployee 
      Caption         =   "Employee"
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CheckBox chkVendor 
      Caption         =   "Vendor"
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   3600
      Width           =   855
   End
   Begin VB.CheckBox chkVendorCredit 
      Caption         =   "Vendor Credit"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   6960
      Width           =   2055
   End
   Begin VB.CheckBox chkCustomer 
      Caption         =   "Customer"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Width           =   1095
   End
   Begin VB.ListBox lstDataExtType 
      Height          =   840
      ItemData        =   "frmAddDataExtDef.frx":00D6
      Left            =   4560
      List            =   "frmAddDataExtDef.frx":00D8
      TabIndex        =   4
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox txtDataExtName 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Apply Data Extension To:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Extension Type"
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Extension Name"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "New Data Extensions for OwnerID {E09C86CF-9D6E-4EF2-BCBE-4D66B6B0F754}"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "frmAddDataExtDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------
' Form: frmAddDataExtDef
'
' Description: This form allows the user to type in the name of a
'              new data extension definition, select the type for it,
'              select the items and transactions to associate it with
'              and the add it to the currently open QuickBooks company
'              file.
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'----------------------------------------------------------

Private Sub cmdAdd_Click()
  
  Dim strObjects As String
  
  If txtDataExtName.Text = "" Then
    MsgBox "You must enter a Data Extension Name"
    Exit Sub
  End If
  
  If lstDataExtType.ListIndex < 0 Then
    MsgBox "You must select a Data Extension Type"
    Exit Sub
  End If
  
  strObjects = ""
  If chkCustomer.Value = 1 Then strObjects = strObjects & " Customer"
  If chkVendor.Value = 1 Then strObjects = strObjects & " Vendor"
  If chkEmployee.Value = 1 Then strObjects = strObjects & " Employee"
  If chkOtherName.Value = 1 Then strObjects = strObjects & " OtherName"
  If chkItem.Value = 1 Then strObjects = strObjects & " Item"
  If chkAccount.Value = 1 Then strObjects = strObjects & " Account"
  If chkBill.Value = 1 Then strObjects = strObjects & " Bill"
  If chkBillPaymentCheck.Value = 1 Then strObjects = strObjects & " BillPaymentCheck"
  If chkBillPaymentCreditCard.Value = 1 Then strObjects = strObjects & " BillPaymentCreditCard"
  If chkCharge.Value = 1 Then strObjects = strObjects & " Charge"
  If chkCheck.Value = 1 Then strObjects = strObjects & " Check"
  If chkCreditCardCharge.Value = 1 Then strObjects = strObjects & " CreditCardCharge"
  If chkCreditCardCredit.Value = 1 Then strObjects = strObjects & " CreditCardCredit"
  If chkCreditMemo.Value = 1 Then strObjects = strObjects & " CreditMemo"
  If chkDeposit.Value = 1 Then strObjects = strObjects & " Deposit"
  If chkEstimate.Value = 1 Then strObjects = strObjects & " Estimate"
  If chkInventoryAdjustment.Value = 1 Then strObjects = strObjects & " InventoryAdjustment"
  If chkInvoice.Value = 1 Then strObjects = strObjects & " Invoice"
  If chkJournalEntry.Value = 1 Then strObjects = strObjects & " JournalEntry"
  If chkPurchaseOrder.Value = 1 Then strObjects = strObjects & " PurchaseOrder"
  If chkReceivePayment.Value = 1 Then strObjects = strObjects & " ReceivePayment"
  If chkSalesReceipt.Value = 1 Then strObjects = strObjects & " SalesReceipt"
  If chkSalesTaxPaymentCheck.Value = 1 Then strObjects = strObjects & " SalesTaxPaymentCheck"
  If chkVendorCredit.Value = 1 Then strObjects = strObjects & " VendorCredit"
  
  If strObjects = "" Then
    MsgBox "You need to pick at least one target for the data extension"
    Exit Sub
  End If
  
  'Remove the leading space from strObjects
  strObjects = Right(strObjects, Len(strObjects) - 1)
  
  AddDataExtDef _
    txtDataExtName.Text, _
    lstDataExtType.List(lstDataExtType.ListIndex), _
    strObjects

  Unload frmAddDataExtDef
End Sub

Private Sub cmdAll_Click()
  chkCustomer.Value = 1
  chkVendor.Value = 1
  chkEmployee.Value = 1
  chkOtherName.Value = 1
  chkItem.Value = 1
  chkAccount.Value = 1
  chkBill.Value = 1
  chkBillPaymentCheck.Value = 1
  chkBillPaymentCreditCard.Value = 1
  chkCharge.Value = 1
  chkCheck.Value = 1
  chkCreditCardCharge.Value = 1
  chkCreditCardCredit.Value = 1
  chkCreditMemo.Value = 1
  chkDeposit.Value = 1
  chkEstimate.Value = 1
  chkInventoryAdjustment.Value = 1
  chkInvoice.Value = 1
  chkJournalEntry.Value = 1
  chkPurchaseOrder.Value = 1
  chkReceivePayment.Value = 1
  chkSalesReceipt.Value = 1
  chkSalesTaxPaymentCheck.Value = 1
  chkVendorCredit.Value = 1
End Sub

Private Sub cmdAllEntities_Click()
  chkCustomer.Value = 1
  chkVendor.Value = 1
  chkEmployee.Value = 1
  chkOtherName.Value = 1
End Sub

Private Sub cmdAllTransactions_Click()
  chkBill.Value = 1
  chkBillPaymentCheck.Value = 1
  chkBillPaymentCreditCard.Value = 1
  chkCharge.Value = 1
  chkCheck.Value = 1
  chkCreditCardCharge.Value = 1
  chkCreditCardCredit.Value = 1
  chkCreditMemo.Value = 1
  chkDeposit.Value = 1
  chkEstimate.Value = 1
  chkInventoryAdjustment.Value = 1
  chkInvoice.Value = 1
  chkJournalEntry.Value = 1
  chkPurchaseOrder.Value = 1
  chkReceivePayment.Value = 1
  chkSalesReceipt.Value = 1
  chkSalesTaxPaymentCheck.Value = 1
  chkVendorCredit.Value = 1
End Sub

Private Sub cmdClearAll_Click()
  chkCustomer.Value = 0
  chkVendor.Value = 0
  chkEmployee.Value = 0
  chkOtherName.Value = 0
  chkItem.Value = 0
  chkAccount.Value = 0
  chkBill.Value = 0
  chkBillPaymentCheck.Value = 0
  chkBillPaymentCreditCard.Value = 0
  chkCharge.Value = 0
  chkCheck.Value = 0
  chkCreditCardCharge.Value = 0
  chkCreditCardCredit.Value = 0
  chkCreditMemo.Value = 0
  chkDeposit.Value = 0
  chkEstimate.Value = 0
  chkInventoryAdjustment.Value = 0
  chkInvoice.Value = 0
  chkJournalEntry.Value = 0
  chkPurchaseOrder.Value = 0
  chkReceivePayment.Value = 0
  chkSalesReceipt.Value = 0
  chkSalesTaxPaymentCheck.Value = 0
  chkVendorCredit.Value = 0
End Sub

Private Sub cmdQuit_Click()
  Unload frmAddDataExtDef
End Sub

Private Sub Form_Load()
  'Set up the data extension type list box with the accepted types
  lstDataExtType.Clear
  lstDataExtType.AddItem "INTTYPE"
  lstDataExtType.AddItem "AMTTYPE"
  lstDataExtType.AddItem "PRICETYPE"
  lstDataExtType.AddItem "QUANTYPE"
  lstDataExtType.AddItem "PERCENTTYPE"
  lstDataExtType.AddItem "DATETIMETYPE"
  lstDataExtType.AddItem "STR255TYPE"
  lstDataExtType.AddItem "STR1024TYPE"
End Sub
