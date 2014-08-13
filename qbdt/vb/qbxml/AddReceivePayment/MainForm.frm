VERSION 5.00
Begin VB.Form MainForm 
   Caption         =   "Enter a Payment"
   ClientHeight    =   9150
   ClientLeft      =   4230
   ClientTop       =   1935
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   ScaleHeight     =   9150
   ScaleWidth      =   8865
   Begin VB.Frame Frame3 
      Caption         =   "Step 2: Select Payment Options"
      Height          =   1935
      Left            =   120
      TabIndex        =   13
      Top             =   2880
      Width           =   8595
      Begin VB.ComboBox Combo_Credit_Memos 
         Height          =   315
         Left            =   2940
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1260
         Width           =   5415
      End
      Begin VB.ComboBox Combo_Invoices 
         Height          =   315
         ItemData        =   "MainForm.frx":0000
         Left            =   1680
         List            =   "MainForm.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   6675
      End
      Begin VB.Label Label_Credit_Memo 
         Caption         =   "Choose a Credit Memo (Optional):"
         Height          =   315
         Left            =   180
         TabIndex        =   19
         Top             =   1260
         Width           =   2595
      End
      Begin VB.Label Label4 
         Caption         =   "Choose an Invoice:"
         Height          =   375
         Left            =   180
         TabIndex        =   18
         Top             =   780
         Width           =   1935
      End
   End
   Begin VB.CommandButton Command_Apply_Payment 
      Caption         =   "Apply Payment"
      Height          =   375
      Left            =   6240
      TabIndex        =   10
      Top             =   6960
      Width           =   2175
   End
   Begin VB.TextBox Text_Ref_Number 
      Height          =   285
      Left            =   1080
      TabIndex        =   5
      Top             =   7080
      Width           =   2500
   End
   Begin VB.ComboBox Combo_Customer 
      Height          =   315
      ItemData        =   "MainForm.frx":0004
      Left            =   1680
      List            =   "MainForm.frx":0006
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   2040
      Width           =   2505
   End
   Begin VB.ComboBox Combo_Pay_Method 
      Height          =   315
      ItemData        =   "MainForm.frx":0008
      Left            =   1080
      List            =   "MainForm.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   6360
      Width           =   2500
   End
   Begin VB.Frame Frame2 
      Caption         =   "Step 1: Choose a Customer"
      Height          =   1815
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   8595
      Begin VB.CommandButton Command_Display_Invoices 
         Caption         =   "Fill In Customer Info Below"
         Height          =   435
         Left            =   4320
         TabIndex        =   1
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label8 
         Caption         =   $"MainForm.frx":000C
         Height          =   615
         Left            =   1320
         TabIndex        =   21
         Top             =   540
         Width           =   6495
      End
      Begin VB.OLE OLE1 
         AutoActivate    =   3  'Automatic
         AutoVerbMenu    =   0   'False
         CausesValidation=   0   'False
         Class           =   "Paint.Picture"
         Enabled         =   0   'False
         Height          =   555
         Left            =   660
         OleObjectBlob   =   "MainForm.frx":00C3
         SourceDoc       =   "C:\Jenn\qbXML vb\sdk1.1\ReceivePayment\customers.bmp"
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Step 3: Apply Payment"
      Height          =   3075
      Left            =   120
      TabIndex        =   14
      Top             =   5160
      Width           =   8595
      Begin VB.TextBox Text_Pay_Date 
         Height          =   285
         Left            =   1020
         TabIndex        =   4
         Top             =   600
         Width           =   2500
      End
      Begin VB.TextBox Text_Amt_Paid 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   7
         Top             =   2520
         Width           =   1200
      End
      Begin VB.TextBox Text_Credit_Memo 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Left            =   7080
         TabIndex        =   8
         Top             =   600
         Width           =   1200
      End
      Begin VB.TextBox Text_Discount 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Left            =   7080
         TabIndex        =   9
         Top             =   1080
         Width           =   1200
      End
      Begin VB.Label Label3 
         Caption         =   "Payment Date (MM/DD/YYYY):"
         Height          =   315
         Left            =   420
         TabIndex        =   24
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "Ref/Check Number:"
         Height          =   375
         Left            =   360
         TabIndex        =   23
         Top             =   1680
         Width           =   2115
      End
      Begin VB.Label Label1 
         Caption         =   "Payment Method:"
         Height          =   315
         Left            =   360
         TabIndex        =   22
         Top             =   960
         Width           =   1755
      End
      Begin VB.Label Label6 
         Caption         =   "Amount of Payment:   $"
         Height          =   315
         Left            =   480
         TabIndex        =   17
         Top             =   2520
         Width           =   2235
      End
      Begin VB.Label Label_Credit_Amt 
         Alignment       =   1  'Right Justify
         Caption         =   "Credit Memo (If Selected Above):   $"
         Height          =   375
         Left            =   4320
         TabIndex        =   16
         Top             =   600
         Width           =   2595
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Discount (Optional):   $"
         Height          =   315
         Left            =   4320
         TabIndex        =   15
         Top             =   1080
         Width           =   2535
      End
   End
   Begin VB.OLE OLE2 
      Class           =   "Paint.Picture"
      Enabled         =   0   'False
      Height          =   525
      Left            =   1140
      OleObjectBlob   =   "MainForm.frx":12DB
      SizeMode        =   1  'Stretch
      SourceDoc       =   "C:\Jenn\qbXML vb\sdk1.1\ReceivePayment\qbbanner.bmp"
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   8520
      Width           =   6435
   End
   Begin VB.Label Label7 
      Caption         =   " Receive Payment"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   8115
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' MainForm.frm
'
' This is the form file for the sample receive payment code.
' The user is first asked to choose a customer.  Then lists
' are displayed of the customer's open invoices and credit
' memos as retrieved from QuickBooks.  The user can choose
' an invoice and optionally a credit memo, fill in all of the
' payment information, and then send the payment to QuickBooks.
'
' This application works with a few selected customers. A sample
' QuickBooks company file: ReceivePaymentsCompanyFile.qbw has
' been provided with this application, which has the relevant
' information for invoices and credit memos set up for this
' purpose.
'
'
'
' Created February, 2002
' Updated to SDK 2.0 July, 2002
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'-------------------------------------------------------------

Option Explicit

' Global variable holds the customer information so it will not
' be inadvertantly modified between when "Display Invoices"
' is clicked and when "Apply Payment" is clicked
Dim customer As String
Public connType As QBXMLRPConnectionType

' Hold info about all of the current customer's open invoices
Dim invoice_TxnIds() As String
Dim invoice_AppliedAmts() As Currency
Dim invoice_Balances() As Currency
Dim invoice_RefNumbers() As String

' Hold info about all of the current customer's credit memos
Dim credit_TxnIds() As String
Dim credit_TotalAmts() As Currency
Dim credit_AmtLeft() As Currency
'

' After the user has filled in all appropriate information,
' we can collect this info and finally make a payment on the
' appropriate invoice.
' Most of this subroutine is concerned with error checking.
' The user can not specify a credit memo amount if no credit
' memo is selected.  The user can't make a payment larger than
' the total amount of the invoice.  There are a number of
' required fields as well.
'
Private Sub Command_Apply_Payment_Click()

    Dim oldDate As String
    Dim payDate As String
    Dim refNum As String
    Dim payMethod As String
    Dim payAmt As Currency
    Dim memoAmt As Currency
    Dim discountAmt As Currency
    Dim invoice As String
    Dim credMemo As String

    ' We need to find the invoice and credit memo data
    ' for the correct invoice and credit memo in the
    ' global arrays
    Dim invNumber As Integer
    Dim credNumber As Integer
    Dim invTxnID As String
    Dim credTxnID As String

    Dim invNumString() As String
    Dim credNumString() As String

    ' Make sure an invoice is selected
    If "" = Combo_Invoices.Text Then
        MsgBox "You must select an invoice."
        Exit Sub
    End If
        
    ' Figure out which invoice the user has selected
    invoice = Combo_Invoices.Text
    invNumString = Split(invoice, ">>")
    
    ' InvNumber is the array index of this invoice in
    ' the four invoice related global variables
    invNumber = CInt(invNumString(0)) - 1
    
    ' See if the user has selected a credit memo, if
    ' so, figure out which it is
    If "" = Combo_Credit_Memos.Text Then
        credMemo = "" ' No credit memo selected
    Else
        ' Same way we figured out which invoice was selected
        credMemo = Combo_Credit_Memos.Text
        credNumString = Split(credMemo, ">>")
        credNumber = CInt(credNumString(0)) - 1
    End If
    
    
    ' Get the amount that the customer wants to apply
    ' to the credit memo and make sure that it is not larger
    ' than the amount of the credit memo the customer is using
    If Not ("" = Text_Credit_Memo.Text) Then
        memoAmt = Text_Credit_Memo.Text
        
        If ("" = credMemo) Then
            ' If user has tried to specify a credit amount without
            ' choosing a credit memo, go back to form
            MsgBox "You may not specify a credit memo amount unless you " _
                & "choose a credit memo.  Please fix this error and " _
                & "hit Apply Payment again."
            Exit Sub
        
        Else
            
            If credit_AmtLeft(credNumber) = 0 Then
                ' The remaining balance for the credit memo is $0.
                ' Another credit memo will need to be selected.
                MsgBox "There is no remaining balance available on this credit memo to apply a payment. " & _
                "Please select another credit memo."
                Exit Sub

            End If
            
            If memoAmt < 0 Then
                MsgBox "Please specify a positive credit memo amount."
                Exit Sub
            End If
            
            
            If memoAmt > credit_AmtLeft(credNumber) Then
                ' Can not use more credit than they have!
                ' Go back to form if user tries to do this
                MsgBox "In order to use this credit memo, you must " & _
                    "select a credit amount less than the remaining balance: $" & _
                    credit_AmtLeft(credNumber)
                Exit Sub

            End If
        End If
    
    Else
        ' If the customer is not applying a credit memo..
        memoAmt = 0
    End If
        
        
    ' Get the discount amount if there is one filled in
    If Not ("" = Text_Discount.Text) Then
        discountAmt = Text_Discount.Text
        If discountAmt < 0 Then
            MsgBox "Please specify a positive amount for discount."
            Exit Sub
        End If
    Else
        discountAmt = 0
    End If


    ' Get the amount of the cash/credit/check/etc payment
    If Not ("" = Text_Amt_Paid.Text) Then
        payAmt = Text_Amt_Paid.Text
        If payAmt < 0 Then
            MsgBox "Please specify a positive payment amount."
            Exit Sub
        End If
    Else
        payAmt = 0
    End If


    ' Make sure that the total of the memo amount, the discount
    ' amount, and the regular payment amount is no larger than
    ' the total amount of the invoice -- Return to the form
    ' if it is.
    If ((memoAmt + discountAmt + payAmt) > _
        invoice_Balances(invNumber)) Then
        MsgBox "You may not pay more than the remaining balance amount " & _
            "of the selected invoice.  Please make sure the total of " & _
            "your payment amounts is less than $" & _
            invoice_Balances(invNumber) & " and hit Apply again."
        Exit Sub
    End If
        
        
    ' Get date info from the form and convert it to a date
    ' that QuickBooks can read.  Return to form if date not valid.
    oldDate = Text_Pay_Date.Text
    payDate = toQBDate(oldDate)
    If ("error" = payDate) Then
        MsgBox "The date you entered is not valid."
        Exit Sub
    End If
    
    ' Get the reference number and payment method from the form
    refNum = Text_Ref_Number.Text
    payMethod = Combo_Pay_Method.Text

    ' Force all mandatory fields to be entered -- The invoice
    ' field is not listed here because it should have already
    ' been checked above
    If ("" = payDate) Or _
        ("" = refNum) Or ("" = payMethod) _
        Or ("" = customer) Or (0 = payAmt) Then
        
        ' Go back to form if not all required fields are filled in
        MsgBox "The following fields are required before you apply a payment to an invoice:" _
            & vbCrLf & "Payment Date" & vbCrLf & "Ref/Check Number" & _
            vbCrLf & "Payment Method" & vbCrLf & "Payment Amount" & _
            vbCrLf & "Invoice"
        Exit Sub
    
    Else
        ' Send the payment to QuickBooks!
        
        If (memoAmt > 0) Then
    
            SendPaymentToQB customer, payDate, refNum, payMethod, payAmt, _
                memoAmt, discountAmt, invoice_TxnIds(invNumber), _
                credit_TxnIds(credNumber)
    
        Else
            
            ' Avoid out of range error if there is no credit memo
            ' being used
            SendPaymentToQB customer, payDate, refNum, payMethod, payAmt, _
                memoAmt, discountAmt, invoice_TxnIds(invNumber), ""
    
        End If
    
        ' Now that we have applied the payment, clear all of the fields
        ' so that the data will not be stale
        Combo_Invoices.Clear
        Combo_Credit_Memos.Clear
        Text_Pay_Date.Text = Date
        Text_Ref_Number.Text = ""
        Text_Amt_Paid.Text = ""
        Text_Credit_Memo.Text = ""
        Text_Discount.Text = ""
    
    End If

End Sub


' Once a customer has been selected and the user clicks the button,
' we need to fetch all of the users open invoices and credit memos
' and display them in the appropriate list boxes.
'
Private Sub Command_Display_Invoices_Click()
    
    ' Set global variable to make sure that its value
    ' will remain the same even if the data is somehow
    ' deleted or changed before "Apply Payment" is clicked
    customer = Combo_Customer.Text

    ' Force mandatory field to be entered
    If ("" = customer) Then
        MsgBox "You must specify a customer before invoices can be displayed."
    Else
        
        ' Before we can display the new lists, we have to clear
        ' the old ones in case there were choices left over from
        ' another customer
        Combo_Invoices.Clear
        Combo_Credit_Memos.Clear
        
        ' Clear all fields that appear below on the form also
        Text_Pay_Date.Text = Date
        Text_Ref_Number.Text = ""
        Text_Amt_Paid.Text = ""
        Text_Credit_Memo.Text = ""
        Text_Discount.Text = ""
        
        ' This is the first time we communicate with QuickBooks.  A
        ' connection is opened and we ask for a list of invoices for
        ' the currently selected customer.  These invoices are then
        ' displayed on the form.  The same is done for credit memos.
        displayLists (customer)
    
    End If

End Sub


' Add original lists of customers and payment methods.
' Also set the date field to the current date.
'
Private Sub Form_Load()
    Text_Pay_Date.Text = Date
    connType = localQBDLaunchUI
    If (MsgBox("Do you want to communicate with QuickBooks Online Edition?", vbYesNo, "Use QuickBooks Online Edition") = vbYes) Then
        connType = remoteQBOE
    End If
    OpenConnection
    fillCustomerList
    
    
    
    Combo_Pay_Method.AddItem "American Express"
    Combo_Pay_Method.AddItem "Barter"
    Combo_Pay_Method.AddItem "Cash"
    Combo_Pay_Method.AddItem "Check"
    Combo_Pay_Method.AddItem "Discover"
    Combo_Pay_Method.AddItem "Master Card"
    Combo_Pay_Method.AddItem "VISA"
     
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If blnIsOpenConnection Then
        'Call CloseConnection
        CloseConnection
    End If
    
End Sub

' Need a public sub that allows qbooks.bas to set the size
' of the invoice related arrays based on
' how many open invoices are found
'
Public Sub SetInvoiceArrayLengths(numInvoices As Integer)

    ReDim invoice_TxnIds(numInvoices) As String
    ReDim invoice_AppliedAmts(numInvoices) As Currency
    ReDim invoice_Balances(numInvoices) As Currency
    ReDim invoice_SuggDiscounts(numInvoices) As Currency
    ReDim invoice_RefNumbers(numInvoices) As String

End Sub




' Need a public sub that allows qbooks.bas to set the size
' of the credit memo related arrays based on
' how many open credit memos are found
'
Public Sub SetCreditMemoArrayLengths(numCreditMemos As Integer)

    ReDim credit_TxnIds(numCreditMemos) As String
    ReDim credit_TotalAmts(numCreditMemos) As Currency
    ReDim credit_AmtLeft(numCreditMemos) As Currency
    
End Sub

Public Sub AddCustomerToList(FullName As String)
    Combo_Customer.AddItem (FullName)
End Sub

' Need a public sub so that qbooks.bas can add
' invoices to the list when getInvoiceList() is called
'
Public Sub AddInvoiceToList(txnId As String, _
    appliedAmt As Currency, balanceRemaining As Currency, _
    suggDiscount As Currency, invRefNum As String, number As Integer)
    
    Dim invoiceData As String
    Dim index As Integer
    
    ' Start numbering the list of invoices at 1 instead of 0
    index = number + 1
    
    ' Store info in all of the global invoice arrays so that
    ' we will have it later once the user chooses an invoice
    invoice_TxnIds(number) = txnId
    invoice_AppliedAmts(number) = appliedAmt
    invoice_Balances(number) = balanceRemaining
    invoice_RefNumbers(number) = invRefNum
    
    invoiceData = index & ">> " & _
        " Ref Num: " & invRefNum & _
        "  /  Remaining Balance: $" & currToString(balanceRemaining) & _
        "  /  Applied: $" & currToString(appliedAmt)
    
    ' Add invoice to the list of choices on the form
    Combo_Invoices.AddItem invoiceData
    
    ' Want the first invoice to be selected initially
    If (1 = index) Then
        Combo_Invoices.Text = invoiceData
    End If
End Sub


' Need a public sub so that qbooks.bas can add
' credit memos to the list when getCreditMemoList() is called
'
Public Sub AddCreditMemoToList(txnId As String, _
    TotalAmount As Currency, amountLeft As Currency, num As Integer)
    
    Dim cmData As String
    Dim index As Integer
    
    ' Want to start numbering the list of credit memos that is
    ' displayed at 1 instead of 0
    index = num + 1
    
    ' Add the information about these credit memos to the global
    ' variables so that it will be available when the user
    ' chooses a memo later.
    credit_TxnIds(num) = txnId
    credit_TotalAmts(num) = TotalAmount
    credit_AmtLeft(num) = amountLeft
    
    cmData = index & ">> Total Amount: $" & currToString(TotalAmount) & _
            "  /  Remaining Balance: $" & currToString(amountLeft)
    
    ' Add credit memo to the list of choices on the form
    Combo_Credit_Memos.AddItem cmData
    
    ' Want the first credit memo to be selected initially
    If (1 = index) Then
        Combo_Credit_Memos.Text = cmData
    End If
End Sub


' Adds credit memo boxes to form if they have been removed --
' this is used when a customer who has credit memos open is
' selected after a previous one who did not have open credit memos
'
Public Sub TurnOnCreditMemos()

    Combo_Credit_Memos.Visible = True
    Text_Credit_Memo.Visible = True
    Label_Credit_Memo.Visible = True
    Label_Credit_Amt.Visible = True

End Sub

' Removes text boxes and labels dealing with credit memos --
' this is used when a customer who does not have any open
' credit memos is selected
'
Public Sub TurnOffCreditMemos()

    Combo_Credit_Memos.Visible = False
    Text_Credit_Memo.Visible = False
    Label_Credit_Memo.Visible = False
    Label_Credit_Amt.Visible = False

End Sub

