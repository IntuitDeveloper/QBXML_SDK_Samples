VERSION 5.00
Begin VB.Form frmEditInvoiceLine 
   Caption         =   "Edit Invoice Line"
   ClientHeight    =   5745
   ClientLeft      =   3645
   ClientTop       =   3450
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   8655
   Begin VB.TextBox txtSavedDesc 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   855
      Left            =   7560
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   25
      Text            =   "frmEditInvoiceLine.frx":0000
      Top             =   3840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox chkDisable 
      Caption         =   "Disable modify warnings (don't seek approval to clear fields based on changes to other fields)"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   4320
      Width           =   6975
   End
   Begin VB.CommandButton cmdFinish 
      Caption         =   "Finish without saving changes"
      CausesValidation=   0   'False
      Height          =   735
      Left            =   5880
      TabIndex        =   23
      Top             =   4800
      Width           =   2655
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Invoice Line"
      Height          =   735
      Left            =   3240
      TabIndex        =   22
      Top             =   4800
      Width           =   2415
   End
   Begin VB.CommandButton cmdUndo 
      Caption         =   "Undo Changes / Get Original Values"
      CausesValidation=   0   'False
      Height          =   735
      Left            =   120
      TabIndex        =   21
      Top             =   4800
      Width           =   2895
   End
   Begin VB.ComboBox cmbOverrideItemAccount 
      Height          =   315
      Left            =   1800
      TabIndex        =   20
      Top             =   3840
      Width           =   5070
   End
   Begin VB.ComboBox cmbSalesTaxCode 
      Height          =   315
      Left            =   7680
      TabIndex        =   18
      Top             =   3360
      Width           =   870
   End
   Begin VB.TextBox txtServiceDate 
      Height          =   285
      Left            =   4920
      TabIndex        =   16
      Top             =   3360
      Width           =   975
   End
   Begin VB.ComboBox cmbClass 
      Height          =   315
      Left            =   600
      TabIndex        =   14
      Top             =   3360
      Width           =   3015
   End
   Begin VB.TextBox txtAmount 
      Height          =   285
      Left            =   6840
      TabIndex        =   12
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox txtRate 
      Height          =   285
      Left            =   4080
      TabIndex        =   10
      Top             =   2880
      Width           =   1215
   End
   Begin VB.OptionButton optRatePercent 
      Caption         =   "Rate Percent"
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   2880
      Width           =   1335
   End
   Begin VB.OptionButton optRate 
      Caption         =   "Rate"
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox txtQuantity 
      Height          =   285
      Left            =   840
      TabIndex        =   7
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox txtDesc 
      Height          =   1575
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   1080
      Width           =   7935
   End
   Begin VB.ComboBox cmbItem 
      Height          =   315
      Left            =   600
      TabIndex        =   3
      Top             =   600
      Width           =   7935
   End
   Begin VB.TextBox txtTxnLineID 
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lblOverrideItemAccount 
      Caption         =   "Override Item Account"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label lblSalesTaxCode 
      Caption         =   "Sales Tax Code"
      Height          =   255
      Left            =   6360
      TabIndex        =   17
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label lblServiceDate 
      Caption         =   "Service Date"
      Height          =   255
      Left            =   3840
      TabIndex        =   15
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label lblClass 
      Caption         =   "Class"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label lblAmount 
      Caption         =   "Amount"
      Height          =   255
      Left            =   6120
      TabIndex        =   11
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Quantity"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label lblDesc 
      Caption         =   "Desc"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Item"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "TxnLineID"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmEditInvoiceLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------
' Copyright © 2003-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/rdmgr/?ID=100
'
'----------------------------------------------------------

Option Explicit

'Variables to track changes made on the form
Dim booItemChanged As Boolean
Dim booDescChanged As Boolean
Dim booQuantityChanged As Boolean
Dim booRatePercentChanged As Boolean
Dim booRateChanged As Boolean
Dim booAmountChanged As Boolean
Dim booClassChanged As Boolean
Dim booServiceDateChanged As Boolean
Dim booSalesTaxCodeChanged As Boolean
Dim booOverrideItemAccountChanged As Boolean

Dim booGroupLine As Boolean

'Variables to keep the type and state of the line we're editing
Dim strLineType As String
Dim strLineState As String

'Variable to track the type of line the content of the items combo box
'is currently set up for
Dim strItemListType As String

Private Sub Form_Load()
  ClearForm
  
  frmPatience.Show
  FillComboBox cmbClass, "Class", "FullName", "", False
  FillComboBox cmbItem, _
    "ItemService,ItemInventory,ItemInventoryAssembly,ItemNonInventory,ItemOtherCharge,ItemSubtotal,ItemGroup,ItemDiscount,ItemPayment,ItemSalesTax", _
    "FullName,FullName,FullName,FullName,FullName,Name,Name,FullName,Name,Name", _
    "", True
  strItemListType = "Item"
  
  FillComboBox cmbSalesTaxCode, "SalesTaxCode", "Name", "", False
  FillComboBox cmbOverrideItemAccount, "Account", "FullName", _
    "<AccountType>Income</AccountType>", False
  frmPatience.Hide
  
  DisplayLineInfo
End Sub


Private Sub Form_Activate()
  ClearForm
  
  DisplayLineInfo
End Sub


Private Sub cmdFinish_Click()
  'Don't save any changes, just unload the form and return to the invoice
  'modify form
  frmEditInvoiceLine.Hide
End Sub


Private Sub cmdSave_Click()
  If AnyValueChanged Then SaveLineInfo
  frmEditInvoiceLine.Hide
End Sub


Private Sub cmdUndo_Click()
  ClearForm
  DisplayLineInfo
End Sub


Private Sub cmbClass_Change()
  booClassChanged = True
End Sub

Private Sub cmbItem_Click()
  booItemChanged = True
  
  If Not booGroupLine Then
    If InStr(1, cmbItem.Text, " - Group Item") > 0 Then
      SetFormForItemLine False
    ElseIf txtDesc.Locked = True Then
      SetFormForItemLine True
    End If
  End If
  
  Dim strDesc As String
  Dim strRate As String
  Dim strRatePercent As String
  Dim strSalesTaxCode As String
  
  GetItemInfo _
    Replace(cmbItem.Text, " - Group Item", ""), _
    strDesc, strRate, strRatePercent, strSalesTaxCode

  Dim strDescSplits() As String
  Dim i As Integer
  
  strDescSplits = Split(strDesc, Chr(10))
  
  If UBound(strDescSplits) > 0 Then
    txtDesc.Text = strDescSplits(0)
    For i = 1 To UBound(strDescSplits)
      txtDesc.Text = txtDesc.Text & vbCrLf & strDescSplits(i)
    Next i
  Else
    txtDesc.Text = strDesc
  End If
  txtSavedDesc.Text = txtDesc.Text

  txtRate.Text = strRate
  cmbSalesTaxCode.Text = strSalesTaxCode
  If strRatePercent = "Rate" Then
    optRate.Value = True
    optRatePercent.Value = False
  Else
    optRate.Value = False
    optRatePercent.Value = True
  End If
End Sub

Private Sub cmbOverrideItemAccount_Change()
  booOverrideItemAccountChanged = True
End Sub

Private Sub cmbSalesTaxCode_Change()
  booSalesTaxCodeChanged = True
End Sub

Private Sub optRate_Click()
  booRatePercentChanged = True
End Sub

Private Sub optRatePercent_Click()
  booRatePercentChanged = True
End Sub

Private Sub txtAmount_Validate(Cancel As Boolean)

  txtAmount.Text = Trim(txtAmount.Text)
  
  Dim strSplits() As String
  strSplits = Split(GetInvoiceLineInfo, "<spliter>")
  
  Dim strOldAmount As String
  
  If UBound(strSplits) = 0 Then
    strOldAmount = ""
  Else
    strOldAmount = strSplits(5)
  End If
  If Trim(txtAmount.Text) = Empty Then
    If Trim(strOldAmount) <> Empty Then booAmountChanged = True
    Exit Sub
  End If
  
  If (Not IsNumeric(txtAmount.Text)) And (Trim(txtAmount.Text) = Empty) Then
    MsgBox "Amount must be numeric, please re-enter Amount"
    txtAmount.Text = strOldAmount
    txtAmount.SetFocus
    booAmountChanged = False
    Exit Sub
  ElseIf Not InStr(1, txtAmount.Text, ".") > 0 Then
    txtAmount.Text = txtAmount.Text & ".00"
  ElseIf Len(txtAmount.Text) - InStr(1, txtAmount.Text, ".") = 1 Then
    txtAmount.Text = txtAmount.Text & "0"
  ElseIf Len(txtAmount.Text) - InStr(1, txtAmount.Text, ".") > 2 Then
    MsgBox "Amount can only have two decimal places"
    txtAmount.Text = _
      Left(txtAmount.Text, _
           Len(txtAmount.Text) + 3 - InStr(1, txtAmount.Text, "."))
  End If
  
  booAmountChanged = Trim(txtAmount.Text) <> Trim(strOldAmount)
  
  If Trim(txtRate.Text) = Empty Then Exit Sub
  
  If CDbl(txtRate.Text) = 0 Then
    txtRate.Text = Empty
    Exit Sub
  End If
  
  Dim objMsgBoxAnswer As VbMsgBoxResult
  If Trim(txtRate.Text) <> Empty And chkDisable.Value = 0 And _
     booAmountChanged Then
    objMsgBoxAnswer = _
      MsgBox("Changing Amount requires Rate to be cleared" & vbCrLf & _
             vbCrLf & "Do you still wish to change Amount and clear Rate?", _
             vbYesNo)
  End If
  
  If (objMsgBoxAnswer = vbYes Or chkDisable.Value <> 0) And _
      booAmountChanged Then
    txtRate.Text = Empty
    booRateChanged = True
    optRate.Value = 1
    optRatePercent.Value = 0
  Else
    txtAmount.Text = Trim(strOldAmount)
  End If
End Sub

Private Sub txtDesc_Change()
  booDescChanged = True
End Sub

Private Sub txtQuantity_Validate(Cancel As Boolean)

  Dim strOldQuantity As String
  
  txtQuantity.Text = Trim(txtQuantity.Text)
  
  Dim strSplits() As String
  strSplits = Split(GetInvoiceLineInfo, "<spliter>")
  
  If UBound(strSplits) = 0 Then
    strOldQuantity = ""
  Else
    strOldQuantity = strSplits(1)
  End If
  
  If Trim(txtQuantity.Text) = Empty Then
    If Trim(strOldQuantity) <> Empty Then booQuantityChanged = True
    Exit Sub
  End If
  
  If (Not IsNumeric(txtQuantity.Text)) And (Trim(txtQuantity.Text) <> Empty) Then
    MsgBox "Quantity must be numeric, please re-enter Quantity"
    txtQuantity.Text = strOldQuantity
    txtQuantity.SetFocus
    booQuantityChanged = False
    Exit Sub
  End If
  
  booQuantityChanged = txtQuantity.Text <> Trim(strOldQuantity)
  
  If Trim(txtAmount.Text) = Empty Then Exit Sub
  
  Dim objMsgBoxAnswer As VbMsgBoxResult
  If Trim(txtRate.Text) <> Empty And chkDisable.Value = 0 And _
     booQuantityChanged Then
    objMsgBoxAnswer = _
      MsgBox("Changing Quantity requires Amount to be cleared" & vbCrLf & _
             vbCrLf & "Do you still wish to change Quantity and clear Amount?", _
             vbYesNo)
  End If
  
  If (objMsgBoxAnswer = vbYes Or chkDisable.Value <> 0) And _
      booQuantityChanged Then
    txtAmount.Text = Empty
    booAmountChanged = True
  Else
    txtQuantity.Text = Trim(strOldQuantity)
  End If
End Sub

Private Sub txtRate_Validate(Cancel As Boolean)
  txtRate.Text = Trim(txtRate.Text)
  
  Dim strSplits() As String
  strSplits = Split(GetInvoiceLineInfo, "<spliter>")
  
  Dim strOldRate As String
  
  If UBound(strSplits) = 0 Then
    strOldRate = ""
  Else
    strOldRate = strSplits(4)
  End If
  
  If Trim(txtRate.Text) = Empty Then
    If Trim(strOldRate) <> Empty Then booRateChanged = True
    Exit Sub
  End If
  
  If Not IsNumeric(txtRate.Text) Then
    MsgBox "Rate must be numeric, please re-enter Rate"
    txtRate.Text = strOldRate
    txtRate.SetFocus
    booRateChanged = False
    Exit Sub
  End If
  
  booRateChanged = txtRate.Text <> Trim(strOldRate)
  
  If Trim(txtAmount.Text) = Empty Then Exit Sub
  
  Dim objMsgBoxAnswer As VbMsgBoxResult
  If Trim(txtRate.Text) <> Empty And Trim(txtAmount.Text) <> Empty And _
     chkDisable.Value = 0 And booRateChanged Then
    objMsgBoxAnswer = _
      MsgBox("Changing Rate requires Amount to be cleared" & vbCrLf & _
             vbCrLf & "Do you still wish to change Rate and clear Amount?", _
             vbYesNo)
  End If
  
  If (objMsgBoxAnswer = vbYes Or chkDisable.Value <> 0) And _
      booRateChanged Then
    txtAmount.Text = Empty
    booAmountChanged = True
  Else
    txtRate.Text = Trim(strOldRate)
  End If
End Sub

Private Sub txtServiceDate_Change()
  booServiceDateChanged = True
End Sub

Private Sub txtDesc_Click()
  If txtDesc.Locked Then MsgBox "You cannot change the description of a group item"
End Sub


Private Sub ClearForm()
  txtTxnLineID.Text = Empty
  cmbItem.Text = Empty
  txtDesc.Text = Empty
  txtQuantity.Text = Empty
  optRate.Value = False
  optRatePercent.Value = False
  txtRate.Text = Empty
  txtAmount.Text = Empty
  cmbClass.Text = Empty
  txtServiceDate.Text = Empty
  cmbSalesTaxCode.Text = Empty
  cmbOverrideItemAccount.Text = Empty
  
  booItemChanged = False
  booDescChanged = False
  booQuantityChanged = False
  booRatePercentChanged = False
  booRateChanged = False
  booAmountChanged = False
  booClassChanged = False
  booServiceDateChanged = False
  booSalesTaxCodeChanged = False
  booOverrideItemAccountChanged = False
  
  SetFormForItemLine True
End Sub


Private Sub DisplayLineInfo()

  Dim strPassedLine As String
  Dim strSplits() As String
  
  'We could be editing a new line in which case the invoice line info will
  'be empty and we don't have anything to display except for a -1 for the
  'TxnLineID
  strPassedLine = GetInvoiceLineInfo
  If Right(strPassedLine, 4) = ",New" Then
    strLineType = Left(strPassedLine, InStr(1, strPassedLine, ",") - 1)
    strLineState = "New"
    txtTxnLineID.Text = "-1 (New " & strLineType & ")"
    booGroupLine = False
    
    If strLineType = "SubItem" And strItemListType <> "SubItem" Then
      FillComboBox cmbItem, _
        "ItemService,ItemInventory,ItemInventoryAssembly,ItemNonInventory,ItemOtherCharge,ItemSubtotal,ItemDiscount,ItemPayment,ItemSalesTax", _
        "FullName,FullName,FullName,FullName,FullName,Name,FullName,Name,Name", _
        "", False
      strItemListType = "SubItem"
      cmbItem.Refresh
    End If
    Exit Sub
  End If
  
  strSplits = Split(strPassedLine, "<spliter>")
  
  If strSplits(11) = "SubItem" And strItemListType <> "SubItem" Then
    FillComboBox cmbItem, _
      "ItemService,ItemInventory,ItemInventoryAssembly,ItemNonInventory,ItemOtherCharge,ItemSubtotal,ItemDiscount,ItemPayment,ItemSalesTax", _
      "FullName,FullName,FullName,FullName,FullName,Name,FullName,Name,Name", _
      "", False
    strItemListType = "SubItem"
    cmbItem.Refresh
    booGroupLine = False
  ElseIf strSplits(11) = "Group" And strItemListType <> "Group" Then
    FillComboBox cmbItem, "ItemGroup", "Name", "", False
    strItemListType = "Group"
    cmbItem.Refresh
    booGroupLine = True
  ElseIf strSplits(11) = "Item" And strItemListType <> "Item" Then
    FillComboBox cmbItem, _
      "ItemService,ItemInventory,ItemInventoryAssembly,ItemNonInventory,ItemOtherCharge,ItemSubtotal,ItemGroup,ItemDiscount,ItemPayment,ItemSalesTax", _
      "FullName,FullName,FullName,FullName,FullName,Name,Name,FullName,Name,Name", _
      "", True
    strItemListType = "Item"
    cmbItem.Refresh
    booGroupLine = False
  End If
  
  SetFormForItemLine (strSplits(11) <> "Group")
  
  txtTxnLineID.Text = strSplits(0)
  cmbItem.Text = strSplits(2)
  
  Dim strDescSplits() As String
  Dim i As Integer
  
  strDescSplits = Split(strSplits(3), Chr(10))
  
  If UBound(strDescSplits) > 0 Then
    txtDesc.Text = strDescSplits(0)
    For i = 1 To UBound(strDescSplits)
      txtDesc.Text = txtDesc.Text & vbCrLf & strDescSplits(i)
    Next i
  Else
    txtDesc.Text = strSplits(3)
  End If
  txtSavedDesc.Text = txtDesc.Text
  
  txtQuantity.Text = strSplits(1)
  
  If strSplits(9) = "Rate" Then
    optRate.Value = True
  ElseIf strSplits(9) = "RatePercent" Then
    optRatePercent.Value = True
  End If
  
  txtRate.Text = strSplits(4)
  txtAmount.Text = strSplits(5)
  cmbClass.Text = strSplits(6)
  txtServiceDate.Text = strSplits(7)
  cmbSalesTaxCode.Text = strSplits(8)
  
  strLineType = strSplits(11)
  strLineState = strSplits(12)
  If strLineState <> "New" Then strLineState = "Changed"
End Sub


Private Function AnyValueChanged() As Boolean
  
  If booItemChanged Or _
     booDescChanged Or _
     booQuantityChanged Or _
     booRatePercentChanged Or _
     booRateChanged Or _
     booAmountChanged Or _
     booClassChanged Or _
     booServiceDateChanged Or _
     booSalesTaxCodeChanged Or _
     booOverrideItemAccountChanged Then

    If Right(GetInvoiceLineInfo, 3) = "New" Then
      If Trim(txtTxnLineID.Text) <> Empty Or _
         Trim(txtQuantity.Text) <> Empty Or _
         Trim(cmbItem.Text) <> Empty Or _
         Trim(txtDesc.Text) <> Empty Or _
         Trim(txtRate.Text) <> Empty Or _
         Trim(txtAmount.Text) <> Empty Or _
         Trim(cmbClass.Text) <> Empty Or _
         Trim(txtServiceDate.Text) <> Empty Or _
         Trim(cmbSalesTaxCode.Text) <> Empty Or _
         optRate.Value = True Or _
         optRatePercent.Value = True Or _
         cmbOverrideItemAccount.Text <> Empty Then
        AnyValueChanged = True
      Else
        AnyValueChanged = False
      End If
      
      Exit Function
    End If
    
    Dim strSplits() As String
    strSplits = Split(GetInvoiceLineInfo, "<spliter>")
    
    If Trim(txtTxnLineID.Text) <> Trim(strSplits(0)) Or _
       Trim(txtQuantity.Text) <> Trim(strSplits(1)) Or _
       Trim(cmbItem.Text) <> Trim(strSplits(2)) Or _
       Trim(txtDesc.Text) <> Trim(txtSavedDesc.Text) Or _
       Trim(txtRate.Text) <> Trim(strSplits(4)) Or _
       Trim(txtAmount.Text) <> Trim(strSplits(5)) Or _
       Trim(cmbClass.Text) <> Trim(strSplits(6)) Or _
       Trim(txtServiceDate.Text) <> Trim(strSplits(7)) Or _
       Trim(cmbSalesTaxCode.Text) <> Trim(strSplits(8)) Or _
       (strSplits(9) = "Rate" And optRate.Value <> True) Or _
       (strSplits(9) = "RatePercent" And optRatePercent.Value <> True) Or _
       cmbOverrideItemAccount.Text <> Empty Then
       AnyValueChanged = True
    Else
      AnyValueChanged = False
    End If
  Else
    AnyValueChanged = False
  End If
End Function


Private Sub SaveLineInfo()

  Dim strInvoiceLineInfo As String
  
  strInvoiceLineInfo = _
    txtTxnLineID.Text & "<spliter>" & _
    txtQuantity.Text & "<spliter>" & _
    cmbItem.Text & "<spliter>" & _
    Replace(txtDesc.Text, vbCrLf, vbCr) & "<spliter>" & _
    txtRate.Text & "<spliter>" & _
    txtAmount.Text & "<spliter>" & _
    cmbClass.Text & "<spliter>" & _
    txtServiceDate.Text & "<spliter>" & _
    cmbSalesTaxCode.Text & "<spliter>"

  
  If optRate.Value = True Then
    strInvoiceLineInfo = strInvoiceLineInfo & "Rate<spliter>"
  ElseIf optRatePercent.Value = True Then
    strInvoiceLineInfo = strInvoiceLineInfo & "RatePercent<spliter>"
  Else
    strInvoiceLineInfo = strInvoiceLineInfo & "Neither<spliter>"
  End If
  
  strInvoiceLineInfo = _
    strInvoiceLineInfo & cmbOverrideItemAccount.Text & "<spliter>" & _
    strLineType & "<spliter>" & strLineState

  SetInvoiceLineInfo strInvoiceLineInfo
End Sub


Public Sub SetFormForItemLine(Value As Boolean)
  txtDesc.Locked = Not Value
  optRate.Enabled = Value
  optRatePercent.Enabled = Value
  txtRate.Enabled = Value
  txtAmount.Enabled = Value
  cmbClass.Enabled = Value
  txtServiceDate.Enabled = Value
  cmbSalesTaxCode.Enabled = Value
  cmbOverrideItemAccount.Enabled = Value
  lblAmount.Enabled = Value
  lblClass.Enabled = Value
  lblServiceDate.Enabled = Value
  lblSalesTaxCode.Enabled = Value
  lblOverrideItemAccount.Enabled = Value
End Sub
