VERSION 5.00
Begin VB.Form frmInvoiceQuery 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Invoice for modification"
   ClientHeight    =   8865
   ClientLeft      =   4065
   ClientTop       =   1665
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkShow 
      Caption         =   "Show invoice query"
      Height          =   255
      Left            =   240
      TabIndex        =   58
      Top             =   8280
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "Query for Invoices"
      Height          =   615
      Left            =   2160
      TabIndex        =   27
      Top             =   8160
      Width           =   1575
   End
   Begin VB.ComboBox cmbAccounts 
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   16
      Top             =   3000
      Width           =   4815
   End
   Begin VB.ComboBox cmbCustomers 
      Height          =   315
      Left            =   1440
      TabIndex        =   14
      Top             =   2520
      Width           =   4815
   End
   Begin VB.Frame frameRefNumberArea 
      Height          =   1095
      Left            =   120
      TabIndex        =   47
      Top             =   3360
      Width           =   7335
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   855
         Left            =   120
         TabIndex        =   49
         Top             =   120
         Width           =   5655
         Begin VB.ComboBox cmbRefNumberCriteria 
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmInvoiceQuery.frx":0000
            Left            =   1080
            List            =   "frmInvoiceQuery.frx":000D
            TabIndex        =   22
            Text            =   "Starts With"
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtToRefNumber 
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   285
            Left            =   3960
            TabIndex        =   20
            Top             =   120
            Width           =   1575
         End
         Begin VB.TextBox txtFromRefNumber 
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   285
            Left            =   2040
            TabIndex        =   19
            Top             =   120
            Width           =   1575
         End
         Begin VB.TextBox txtRefNumberPart 
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   285
            Left            =   2640
            TabIndex        =   23
            Top             =   480
            Width           =   1575
         End
         Begin VB.OptionButton optRefNumberFilter 
            Caption         =   "Invoice #"
            Height          =   255
            Left            =   0
            TabIndex        =   21
            Top             =   480
            Width           =   975
         End
         Begin VB.OptionButton optRefNumberRange 
            Caption         =   "Invoice # Range"
            Height          =   255
            Left            =   0
            TabIndex        =   18
            Top             =   120
            Width           =   1575
         End
         Begin VB.Label lblToRefNumber 
            Caption         =   "To"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3720
            TabIndex        =   51
            Top             =   120
            Width           =   255
         End
         Begin VB.Label lblFromRefNumber 
            Caption         =   "From"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1560
            TabIndex        =   50
            Top             =   120
            Width           =   375
         End
      End
      Begin VB.CommandButton cmdClearRefNumber 
         Caption         =   "Clear Invoice # Filtering"
         Height          =   735
         Left            =   6120
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1935
      Left            =   120
      TabIndex        =   38
      Top             =   480
      Width           =   7335
      Begin VB.TextBox txtToTime 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3480
         TabIndex        =   55
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtFromTime 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3480
         TabIndex        =   54
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtToDate 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   53
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtFromDate 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   52
         Top             =   600
         Width           =   1095
      End
      Begin VB.Frame frameFromAMPM 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   4200
         TabIndex        =   44
         Top             =   480
         Width           =   1335
         Begin VB.OptionButton optFromAM 
            Caption         =   "AM"
            Enabled         =   0   'False
            Height          =   255
            Left            =   0
            TabIndex        =   4
            Top             =   120
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton optFromPM 
            Caption         =   "PM"
            Enabled         =   0   'False
            Height          =   255
            Left            =   600
            TabIndex        =   5
            Top             =   120
            Width           =   615
         End
      End
      Begin VB.Frame frameToAMPM 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   4200
         TabIndex        =   43
         Top             =   840
         Width           =   1335
         Begin VB.OptionButton optToPM 
            Caption         =   "PM"
            Enabled         =   0   'False
            Height          =   255
            Left            =   600
            TabIndex        =   7
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton optToAM 
            Caption         =   "AM"
            Enabled         =   0   'False
            Height          =   255
            Left            =   0
            TabIndex        =   6
            Top             =   120
            Value           =   -1  'True
            Width           =   615
         End
      End
      Begin VB.Frame frameDateCriteria 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   495
         Left            =   120
         TabIndex        =   42
         Top             =   120
         Width           =   5775
         Begin VB.OptionButton optModified 
            Caption         =   "Created or Modified"
            Height          =   255
            Left            =   0
            TabIndex        =   1
            Top             =   120
            Width           =   1695
         End
         Begin VB.OptionButton optTxnDate 
            Caption         =   "Transaction Date"
            Height          =   255
            Left            =   1920
            TabIndex        =   2
            Top             =   120
            Width           =   1695
         End
         Begin VB.OptionButton optMacroDate 
            Caption         =   "Transaction Date Macro"
            Height          =   255
            Left            =   3720
            TabIndex        =   3
            Top             =   120
            Width           =   2055
         End
      End
      Begin VB.CommandButton cmdClearDateSelection 
         Caption         =   "Clear Date Filtering"
         Height          =   495
         Left            =   6120
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
      Begin VB.Frame frmThisLast 
         Height          =   495
         Left            =   120
         TabIndex        =   40
         Top             =   1320
         Width           =   2175
         Begin VB.OptionButton optThis 
            Caption         =   "This (to date)"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   120
            Width           =   1335
         End
         Begin VB.OptionButton optLast 
            Caption         =   "Last"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1440
            TabIndex        =   9
            Top             =   120
            Width           =   615
         End
      End
      Begin VB.Frame frameMacroPeriod 
         Height          =   495
         Left            =   2400
         TabIndex        =   39
         Top             =   1320
         Width           =   4815
         Begin VB.OptionButton optWeek 
            Caption         =   "Week"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   120
            Width           =   735
         End
         Begin VB.OptionButton optMonth 
            Caption         =   "Month"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1080
            TabIndex        =   11
            Top             =   120
            Width           =   855
         End
         Begin VB.OptionButton optFiscalQuarter 
            Caption         =   "Fiscal Quarter"
            Enabled         =   0   'False
            Height          =   255
            Left            =   2040
            TabIndex        =   12
            Top             =   120
            Width           =   1335
         End
         Begin VB.OptionButton optFiscalYear 
            Caption         =   "Fiscal Year"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3600
            TabIndex        =   13
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Label lblToTime 
         Caption         =   "hh:mm"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2880
         TabIndex        =   57
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblFromTime 
         Caption         =   "hh:mm"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2880
         TabIndex        =   56
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblAfter 
         Caption         =   "From (yyyy-mm-dd)"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblBefore 
         Caption         =   "To    (yyyy-mm-dd)"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   960
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   840
      TabIndex        =   35
      Top             =   4440
      Width           =   2775
      Begin VB.OptionButton optPaidOnly 
         Caption         =   "only paid"
         Height          =   255
         Left            =   1800
         TabIndex        =   26
         Top             =   120
         Width           =   975
      End
      Begin VB.OptionButton optUnpaidOnly 
         Caption         =   "only unpaid"
         Height          =   255
         Left            =   600
         TabIndex        =   25
         Top             =   120
         Width           =   1215
      End
      Begin VB.OptionButton optAll 
         Caption         =   "all"
         Height          =   255
         Left            =   0
         TabIndex        =   24
         Top             =   120
         Value           =   -1  'True
         Width           =   615
      End
   End
   Begin VB.CheckBox chkAccountChildren 
      Caption         =   "With Children"
      Height          =   255
      Left            =   6360
      TabIndex        =   17
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CheckBox chkCustomerChildren 
      Caption         =   "With Children"
      Height          =   255
      Left            =   6360
      TabIndex        =   15
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox txtInvoiceNumber 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   6000
      TabIndex        =   30
      Top             =   8160
      Width           =   1575
   End
   Begin VB.CommandButton cmdModifyInvoice 
      Caption         =   "Modify Invoice"
      Enabled         =   0   'False
      Height          =   615
      Left            =   4080
      TabIndex        =   29
      Top             =   8160
      Width           =   1575
   End
   Begin VB.ListBox lstInvoices 
      Height          =   2985
      ItemData        =   "frmInvoiceQuery.frx":0033
      Left            =   240
      List            =   "frmInvoiceQuery.frx":0035
      TabIndex        =   28
      Top             =   5040
      Width           =   7335
   End
   Begin VB.Label Label8 
      Caption         =   "Invoices"
      Height          =   255
      Left            =   3720
      TabIndex        =   37
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "Return"
      Height          =   255
      Left            =   240
      TabIndex        =   36
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Account"
      Height          =   255
      Left            =   240
      TabIndex        =   34
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Customer/Job"
      Height          =   255
      Left            =   240
      TabIndex        =   33
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "or find up to 30 invoices that meet the below criteria"
      Height          =   255
      Left            =   2640
      TabIndex        =   32
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Find Invoice #"
      Height          =   255
      Left            =   240
      TabIndex        =   31
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmInvoiceQuery"
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

'This variable determines whether the filter controls are enabled
Dim booFiltersEnabled As Boolean

'Saves the last command from the form for use in Form_Activate
Dim strLastCommand As String

'Save the state of the Invoice Modify form
Dim booInvoiceModifyLoaded As Boolean

'These are the variables which will be set to do the invoice query
Dim strRefNumber As String
Dim strFromDateTime As String
Dim strToDateTime As String
Dim strDateQueryType As String
Dim strDateMacro As String
Dim strCustomerJob As String
Dim booCustomerWithChildren As Boolean
Dim strAccount As String
Dim booAccountWithChildren As Boolean
Dim strFromRefNumberRange As String
Dim strToRefNumberRange As String
Dim strRefNumberPiece As String
Dim strRefNumberCriteria As String
Dim strPaidStatus As String


Private Sub Form_Load()
  strLastCommand = Empty
  booInvoiceModifyLoaded = False
  
  booFiltersEnabled = True
  If Not SupportsModify Then
    cmdModifyInvoice.Caption = "Invoice Details"
  End If
  
  frmPatience.Show
  FillComboBox cmbCustomers, "Customer", "FullName", "", False
  
  FillComboBox cmbAccounts, "Account", "FullName", "<AccountType>AccountsReceivable</AccountType>", False
  frmPatience.Hide
End Sub


Private Sub Form_Activate()
  If strLastCommand <> Empty And strLastCommand <> "Query" Then
    lstInvoices.Clear
    ClearQueryVariables
    If Not SetQueryVariables Then
      Exit Sub
    End If
  
    FillInvoiceList lstInvoices, strRefNumber, strFromDateTime, _
      strToDateTime, strDateQueryType, strDateMacro, strCustomerJob, _
      booCustomerWithChildren, strAccount, booAccountWithChildren, _
      strFromRefNumberRange, strToRefNumberRange, strRefNumberPiece, _
      strRefNumberCriteria, strPaidStatus
  End If
End Sub


Private Sub cmdModifyInvoice_Click()

  strLastCommand = "ModifyInvoice"
  
  If lstInvoices.ListIndex < 0 Then
    MsgBox "You must select an invoice to modify"
    Exit Sub
  End If
  
  If lstInvoices.List(lstInvoices.ListIndex) = _
       "No invoices match the query filter used" Or _
     lstInvoices.ListCount = 0 Then
    MsgBox "You must query for a valid invoice before you can modify one"
    Exit Sub
  End If
  
  GetInvoice Right(lstInvoices.List(lstInvoices.ListIndex), Len(lstInvoices.List(lstInvoices.ListIndex)) - InStrRev(lstInvoices.List(lstInvoices.ListIndex), " "))
  
  If Not booInvoiceModifyLoaded Then
    Load frmInvoiceModify
    booInvoiceModifyLoaded = True
  End If
  
  frmInvoiceModify.Show
End Sub


Private Sub cmdClearDateSelection_Click()
  optFiscalQuarter.Value = False
  optFiscalYear.Value = False
  optLast.Value = False
  optMacroDate.Value = False
  optModified.Value = False
  optMonth.Value = False
  optThis.Value = False
  optTxnDate.Value = False
  optWeek.Value = False
  SetToFromDatesEnabled False
  SetMacroDatesEnabled False
End Sub


Private Sub cmdClearRefNumber_Click()
  optRefNumberFilter.Value = False
  optRefNumberRange.Value = False
  txtFromRefNumber.Enabled = False
  txtToRefNumber.Enabled = False
  lblFromRefNumber.Enabled = False
  lblToRefNumber.Enabled = False
  cmbRefNumberCriteria.Enabled = False
  txtRefNumberPart.Enabled = False
  
  txtFromRefNumber.BackColor = &H80000004
  txtToRefNumber.BackColor = &H80000004
  cmbRefNumberCriteria.BackColor = &H80000004
  txtRefNumberPart.BackColor = &H80000004
End Sub


Private Sub cmdQuery_Click()
  strLastCommand = "Query"
  
  lstInvoices.Clear
  ClearQueryVariables
  If Not SetQueryVariables Then
    Exit Sub
  End If
  
  FillInvoiceList lstInvoices, strRefNumber, strFromDateTime, _
    strToDateTime, strDateQueryType, strDateMacro, strCustomerJob, _
    booCustomerWithChildren, strAccount, booAccountWithChildren, _
    strFromRefNumberRange, strToRefNumberRange, strRefNumberPiece, _
    strRefNumberCriteria, strPaidStatus

  If chkShow.Value = 1 Then
    Load frmShowRequest
    frmShowRequest.Show
  End If
End Sub


Private Sub cmdQuit_Click()
  EndSessionCloseConnection
  End
End Sub


Private Sub lstInvoices_Click()
  cmdModifyInvoice.Enabled = True
End Sub

Private Sub optMacroDate_Click()
  SetToFromDatesEnabled False
  SetMacroDatesEnabled True
End Sub


Private Sub optModified_Click()
  SetToFromDatesEnabled True
  SetMacroDatesEnabled False
End Sub


Private Sub optTxnDate_Click()
  SetToFromDatesEnabled True
  SetMacroDatesEnabled False
End Sub


Private Sub optRefNumberFilter_Click()
  cmbRefNumberCriteria.Enabled = True
  txtRefNumberPart.Enabled = True
  cmbRefNumberCriteria.BackColor = &H80000005
  txtRefNumberPart.BackColor = &H80000005
  
  lblFromRefNumber.Enabled = False
  lblToRefNumber.Enabled = False
  txtFromRefNumber.Enabled = False
  txtToRefNumber.Enabled = False
  txtFromRefNumber.BackColor = &H80000004
  txtToRefNumber.BackColor = &H80000004
End Sub


Private Sub optRefNumberRange_Click()
  lblFromRefNumber.Enabled = True
  lblToRefNumber.Enabled = True
  txtFromRefNumber.Enabled = True
  txtToRefNumber.Enabled = True
  txtFromRefNumber.BackColor = &H80000005
  txtToRefNumber.BackColor = &H80000005
  
  cmbRefNumberCriteria.Enabled = False
  txtRefNumberPart.Enabled = False
  cmbRefNumberCriteria.BackColor = &H80000004
  txtRefNumberPart.BackColor = &H80000004
End Sub


Private Sub txtInvoiceNumber_Change()
  If booFiltersEnabled <> (Trim(txtInvoiceNumber.Text) = "") Then
    SetFiltersEnabled (Trim(txtInvoiceNumber.Text) = "")
    booFiltersEnabled = (Trim(txtInvoiceNumber.Text) = "")
  End If
End Sub


Private Sub SetToFromDatesEnabled(Value As Boolean)
  lblAfter.Enabled = Value
  txtFromDate.Enabled = Value
  
  lblBefore.Enabled = Value
  txtToDate.Enabled = Value
  
  If Value And SupportsDateTime And optModified.Value Then
    txtFromTime.Enabled = Value
    lblFromTime.Enabled = Value
    optFromAM.Enabled = Value
    optFromPM.Enabled = Value
    
    txtToTime.Enabled = Value
    lblToTime.Enabled = Value
    optToAM.Enabled = Value
    optToPM.Enabled = Value
  ElseIf Not Value Then
    optFromAM.Enabled = Value
    optFromPM.Enabled = Value
    optToAM.Enabled = Value
    optToPM.Enabled = Value
  End If
  
  If optTxnDate.Value Then
    txtFromTime.Enabled = False
    lblFromTime.Enabled = False
    optFromAM.Enabled = False
    optFromPM.Enabled = False
    
    txtToTime.Enabled = False
    lblToTime.Enabled = False
    optToAM.Enabled = False
    optToPM.Enabled = False
  
    txtFromTime.BackColor = &H80000004
    txtToTime.BackColor = &H80000004
  End If
  
  If Value Then
    txtFromDate.BackColor = &H80000005
    
    txtToDate.BackColor = &H80000005
    
    If SupportsDateTime And optModified.Value Then
      txtFromTime.BackColor = &H80000005
      txtToTime.BackColor = &H80000005
    End If
  Else
    txtFromDate.BackColor = &H80000004
    txtFromTime.BackColor = &H80000004
    
    txtToDate.BackColor = &H80000004
    txtToTime.BackColor = &H80000004
  End If
End Sub


Private Sub SetMacroDatesEnabled(Value As Boolean)
  optThis.Enabled = Value
  optLast.Enabled = Value
  optWeek.Enabled = Value
  optMonth.Enabled = Value
  optFiscalQuarter.Enabled = Value
  optFiscalYear.Enabled = Value
End Sub


Private Sub SetFiltersEnabled(Value As Boolean)
  Label2.Enabled = Value
  
  optModified.Enabled = Value
  optTxnDate.Enabled = Value
  optMacroDate.Enabled = Value
  
  If Value Then
    If (optModified.Value Or optTxnDate.Value) Then
      SetToFromDatesEnabled True
    ElseIf optMacroDate.Value Then
      SetMacroDatesEnabled True
    End If
  Else
    SetToFromDatesEnabled False
    SetMacroDatesEnabled False
  End If
  cmdClearDateSelection.Enabled = Value
  
  Label3.Enabled = Value
  cmbCustomers.Enabled = Value
  chkCustomerChildren.Enabled = Value
  
  Label4.Enabled = Value
  cmbAccounts.Enabled = Value
  chkAccountChildren.Enabled = Value
  
  optRefNumberRange.Enabled = Value
  lblFromRefNumber.Enabled = Value
  txtFromRefNumber.Enabled = Value
  lblToRefNumber.Enabled = Value
  txtToRefNumber.Enabled = Value
  
  optRefNumberFilter.Enabled = Value
  cmbRefNumberCriteria.Enabled = Value
  txtRefNumberPart.Enabled = Value
  
  cmdClearRefNumber.Enabled = Value
  
  Label7.Enabled = Value
  optAll.Enabled = Value
  optUnpaidOnly.Enabled = Value
  optPaidOnly.Enabled = Value
  Label8.Enabled = Value
  
  If Value Then
    cmbCustomers.BackColor = &H80000005
    cmbAccounts.BackColor = &H80000005
    txtFromRefNumber.BackColor = &H80000005
    txtToRefNumber.BackColor = &H80000005
    cmbRefNumberCriteria.BackColor = &H80000005
    txtRefNumberPart.BackColor = &H80000005
  Else
    cmbCustomers.BackColor = &H80000004
    cmbAccounts.BackColor = &H80000004
    txtFromRefNumber.BackColor = &H80000004
    txtToRefNumber.BackColor = &H80000004
    cmbRefNumberCriteria.BackColor = &H80000004
    txtRefNumberPart.BackColor = &H80000004
  End If
End Sub


Private Sub ClearQueryVariables()
  strRefNumber = ""
  strFromDateTime = ""
  strToDateTime = ""
  strDateQueryType = ""
  strDateMacro = ""
  strCustomerJob = ""
  booCustomerWithChildren = False
  strAccount = ""
  booAccountWithChildren = False
  strFromRefNumberRange = ""
  strToRefNumberRange = ""
  strRefNumberPiece = ""
  strRefNumberCriteria = ""
  strPaidStatus = ""
End Sub


Private Function SetQueryVariables() As Boolean

  Dim strHours As String
  Dim strMinutes As String
  
  If Trim(txtInvoiceNumber.Text) <> Empty Then
    strRefNumber = Trim(txtInvoiceNumber.Text)
    'If we have an invoice number that's all we need
    SetQueryVariables = True
    Exit Function
  End If
  
  If optModified.Value Then
    If DateValid(txtFromDate.Text, "From") Then
      If TimeValid(txtFromTime.Text, "From") Then
        If DateValid(txtToDate.Text, "To") Then
          If TimeValid(txtToTime.Text, "To") Then
    
            If Trim(txtFromDate.Text) <> Empty Then
              strFromDateTime = _
                DateTimeString(txtFromDate.Text, txtFromTime.Text, _
                               optFromAM.Value, SupportsDateTime)
            End If
    
            If Trim(txtToDate.Text) <> Empty Then
              strToDateTime = _
                DateTimeString(txtToDate.Text, txtToTime.Text, _
                               optToAM.Value, SupportsDateTime)
            End If
            
            If strFromDateTime <> Empty And strToDateTime <> Empty Then
              If CDate(Replace(strFromDateTime, "T", " ")) > _
                 CDate(Replace(strToDateTime, "T", " ")) Then
                MsgBox "The from date must be before the to date"
                SetQueryVariables = False
                Exit Function
              End If
            End If
          Else
            SetQueryVariables = False
            Exit Function
          End If
        Else
          SetQueryVariables = False
          Exit Function
        End If
      Else
        SetQueryVariables = False
        Exit Function
      End If
    Else
      SetQueryVariables = False
      Exit Function
    End If
    
    strDateQueryType = "ModifiedDateRangeFilter"
  ElseIf optTxnDate.Value Then
    If DateValid(txtFromDate.Text, "From") Then
      If DateValid(txtToDate.Text, "To") Then
      
        If Trim(txtFromDate.Text) <> Empty Then
          strFromDateTime = _
            DateTimeString(txtFromDate.Text, txtFromTime.Text, _
                           optFromAM.Value, False)
        End If
    
        If Trim(txtToDate.Text) <> Empty Then
          strToDateTime = _
            DateTimeString(txtToDate.Text, txtToTime.Text, _
                           optToAM.Value, False)
        End If
      
        If strFromDateTime <> Empty And strToDateTime <> Empty Then
          If CDate(strFromDateTime) > CDate(strToDateTime) Then
            MsgBox "The from date must be before the to date"
            SetQueryVariables = False
            Exit Function
          End If
        End If
        strDateQueryType = "TxnDateRangeFilter"
      Else
        SetQueryVariables = False
        Exit Function
      End If
    Else
      SetQueryVariables = False
      Exit Function
    End If
  ElseIf optMacroDate.Value Then
    If optThis.Value Then
      strDateMacro = "This"
    ElseIf optLast.Value Then
      strDateMacro = "Last"
    Else
      MsgBox "You must select either This (to date) or Last when using the Date Macro filter"
      SetQueryVariables = False
      strDateMacro = ""
      Exit Function
    End If
    
    If optWeek.Value Then
      strDateMacro = strDateMacro & "Week"
    ElseIf optMonth.Value Then
      strDateMacro = strDateMacro & "Month"
    ElseIf optFiscalQuarter.Value Then
      strDateMacro = strDateMacro & "FiscalQuarter"
    ElseIf optFiscalYear.Value Then
      strDateMacro = strDateMacro & "FiscalYear"
    Else
      MsgBox "You much choose, Week, Month, Fiscal Quarter or Fiscal Year when using the Date Macro filter"
      SetQueryVariables = False
      strDateMacro = ""
      Exit Function
    End If
  
    If optThis.Value Then
      strDateMacro = strDateMacro & "ToDate"
    End If
  End If

  If Trim(cmbCustomers.Text) <> Empty Then
    strCustomerJob = Trim(cmbCustomers.Text)
    booCustomerWithChildren = (chkCustomerChildren.Value = 1)
  End If
  
  If Trim(cmbAccounts.Text) <> Empty Then
    strAccount = Trim(cmbAccounts.Text)
    booAccountWithChildren = (chkAccountChildren.Value = 1)
  End If
  
  If optRefNumberRange.Value Then
    strFromRefNumberRange = Trim(txtFromRefNumber.Text)
    strToRefNumberRange = Trim(txtToRefNumber.Text)
  ElseIf optRefNumberFilter.Value Then
    strRefNumberCriteria = cmbRefNumberCriteria.Text
    strRefNumberCriteria = Replace(strRefNumberCriteria, " ", "")
    strRefNumberPiece = Trim(txtRefNumberPart.Text)
  End If
  
  If optAll.Value Then
    strPaidStatus = "All"
  ElseIf optPaidOnly.Value Then
    strPaidStatus = "PaidOnly"
  Else
    strPaidStatus = "NotPaidOnly"
  End If
  
  SetQueryVariables = True
End Function



