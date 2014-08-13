VERSION 5.00
Begin VB.Form Invoice 
   Caption         =   "Add Invoice Sample"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11670
   ScaleHeight     =   6990
   ScaleWidth      =   11670
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command_Save 
      Caption         =   "Save"
      Height          =   495
      Left            =   4080
      TabIndex        =   25
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton Command_Close 
      Caption         =   "Close"
      Height          =   495
      Left            =   7920
      TabIndex        =   27
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton Command_Clear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   6000
      TabIndex        =   26
      Top             =   6240
      Width           =   1575
   End
   Begin VB.TextBox Text_ExRate 
      Height          =   375
      Left            =   1800
      TabIndex        =   21
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox Text_Total 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8280
      TabIndex        =   24
      Top             =   5400
      Width           =   1455
   End
   Begin VB.TextBox Text_PST 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   4105
         SubFormatType   =   2
      EndProperty
      Enabled         =   0   'False
      Height          =   285
      Left            =   8880
      TabIndex        =   23
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox Text_GST 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   4105
         SubFormatType   =   2
      EndProperty
      Enabled         =   0   'False
      Height          =   285
      Left            =   8880
      TabIndex        =   22
      Top             =   4560
      Width           =   855
   End
   Begin VB.Frame Frame6 
      Height          =   495
      Left            =   9840
      TabIndex        =   41
      Top             =   1560
      Width           =   1335
      Begin VB.Label Label11 
         Caption         =   "Tax Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame5 
      Height          =   495
      Left            =   8280
      TabIndex        =   40
      Top             =   1560
      Width           =   1575
      Begin VB.Label Label10 
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   46
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Height          =   495
      Left            =   7200
      TabIndex        =   39
      Top             =   1560
      Width           =   1095
      Begin VB.Label Label9 
         Caption         =   "Rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Height          =   495
      Left            =   6360
      TabIndex        =   38
      Top             =   1560
      Width           =   855
      Begin VB.Label Label8 
         Caption         =   "Qty"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   2640
      TabIndex        =   37
      Top             =   1560
      Width           =   3735
      Begin VB.Label Label6 
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   43
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   480
      TabIndex        =   36
      Top             =   1560
      Width           =   2175
      Begin VB.Label Label7 
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   42
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.ComboBox Combo_Second_TaxCode 
      Height          =   315
      Left            =   9840
      TabIndex        =   14
      Top             =   3000
      Width           =   1095
   End
   Begin VB.ComboBox Combo_Third_TaxCode 
      Height          =   315
      Left            =   9840
      TabIndex        =   20
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox Text_Second_Amount 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   4105
         SubFormatType   =   2
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   13
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox Text_Third_Amount 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   4105
         SubFormatType   =   2
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   19
      Top             =   3600
      Width           =   1575
   End
   Begin VB.ComboBox Combo_Second_Item 
      Height          =   315
      Left            =   600
      TabIndex        =   9
      Top             =   3000
      Width           =   2055
   End
   Begin VB.ComboBox Combo_Third_Item 
      Height          =   315
      Left            =   600
      TabIndex        =   15
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox Text_Third_Rate 
      Height          =   375
      Left            =   7200
      TabIndex        =   18
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox Text_Third_QTY 
      Height          =   375
      Left            =   6360
      TabIndex        =   17
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox Text_Desc_Third_Item 
      Height          =   375
      Left            =   2640
      TabIndex        =   16
      Top             =   3600
      Width           =   3735
   End
   Begin VB.TextBox Text_Second_Rate 
      Height          =   375
      Left            =   7200
      TabIndex        =   12
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox Text_Second_QTY 
      Height          =   375
      Left            =   6360
      TabIndex        =   11
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox Text_Desc_Second_Item 
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   3000
      Width           =   3735
   End
   Begin VB.ComboBox Combo_First_TaxCode 
      Height          =   315
      Left            =   9840
      TabIndex        =   8
      Top             =   2400
      Width           =   1095
   End
   Begin VB.ComboBox Combo_First_Item 
      Height          =   315
      Left            =   600
      TabIndex        =   3
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Frame Frame 
      Height          =   2415
      Left            =   480
      TabIndex        =   35
      Top             =   1920
      Width           =   10695
      Begin VB.TextBox Text_First_Amount 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   4105
            SubFormatType   =   2
         EndProperty
         Height          =   375
         Left            =   7800
         TabIndex        =   7
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox Text_First_Rate 
         Height          =   375
         Left            =   6720
         TabIndex        =   6
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Text_First_QTY 
         Height          =   375
         Left            =   5880
         TabIndex        =   5
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Text_Desc_First_Item 
         Height          =   375
         Left            =   2160
         TabIndex        =   4
         Top             =   480
         Width           =   3735
      End
      Begin VB.Line Line5 
         X1              =   7800
         X2              =   7800
         Y1              =   120
         Y2              =   2400
      End
      Begin VB.Line Line4 
         X1              =   6720
         X2              =   6720
         Y1              =   120
         Y2              =   2400
      End
      Begin VB.Line Line3 
         X1              =   9360
         X2              =   9360
         Y1              =   120
         Y2              =   2400
      End
      Begin VB.Line Line2 
         X1              =   5880
         X2              =   5880
         Y1              =   120
         Y2              =   2400
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000007&
         X1              =   2160
         X2              =   2160
         Y1              =   120
         Y2              =   2400
      End
   End
   Begin VB.TextBox Text_PO_Number 
      Height          =   285
      Left            =   9720
      TabIndex        =   33
      Top             =   3120
      Width           =   1335
   End
   Begin VB.ComboBox Combo_Terms 
      Height          =   315
      Left            =   8160
      TabIndex        =   31
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox Text_Date 
      Height          =   285
      Left            =   6840
      TabIndex        =   2
      Top             =   600
      Width           =   1695
   End
   Begin VB.ComboBox Combo_ArAccount 
      Height          =   315
      Left            =   3600
      TabIndex        =   1
      Top             =   600
      Width           =   2535
   End
   Begin VB.ComboBox Combo_Customer 
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   2415
   End
   Begin VB.Image Image_QBBANNER 
      Height          =   615
      Left            =   600
      Top             =   6240
      Width           =   3015
   End
   Begin VB.Line Line6 
      X1              =   11880
      X2              =   120
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Label Label20 
      Caption         =   "Exchange Rate:"
      Height          =   255
      Left            =   360
      TabIndex        =   51
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label Label14 
      Caption         =   "Total"
      Height          =   375
      Left            =   7680
      TabIndex        =   50
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label13 
      Caption         =   "PST"
      Height          =   255
      Left            =   8280
      TabIndex        =   49
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label Label12 
      Caption         =   "GST"
      Height          =   255
      Left            =   8280
      TabIndex        =   48
      Top             =   4560
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "P.O. No."
      Height          =   255
      Left            =   9720
      TabIndex        =   34
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Terms"
      Height          =   255
      Left            =   8160
      TabIndex        =   32
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Date"
      Height          =   255
      Left            =   6840
      TabIndex        =   30
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Account"
      Height          =   255
      Left            =   3600
      TabIndex        =   29
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Customer:Job"
      Height          =   255
      Left            =   480
      TabIndex        =   28
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Invoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Invoice.frm
' This form is part of the Invoice sample program
' for the QuickBooks SDK Version CA2.0.
'
' Created September, 2002
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'-------------------------------------------------------------
Private Sub Combo_ArAccount_Change()
'If not in the list, we clear the combo box
    If Combo_ArAccount.ListIndex = -1 Then
        Combo_ArAccount = ""
    End If
End Sub

Private Sub Combo_ArAccount_Click()
'Once the account is selected, We must make sure that the currency is the same as the customer selected (If multi-currency is turned on in the preferences)
    
    If Combo_ArAccount <> "" Then
        Dim strArAccountArray() As String 'I am using the name in the combo box to retrieve the info about
                                                                        'This ArAccount from the collection
        strArAccountArray = colArAccount.Item(Combo_ArAccount)
        If strArAccountArray(1) <> "" Then 'If multicurrency is turned on in the preference, The currencyref is not ""
    
            Dim strCustomerArray() As String
            strCustomerArray = colCustomerCurrencyList.Item(Combo_Customer) 'I am using the name in the combo box to retrieve the info about
                                                                            'This customer from the collection
            
            If strCustomerArray(1) <> strArAccountArray(1) Then  ' Make sure that both account and customer use the same currency by comparing the currencyref
            MsgBox "Please select an account that has the same currency as the customer", , "Change account"
            Combo_ArAccount = ""
            Combo_ArAccount.SetFocus
        
            End If
        End If
    End If
    

End Sub

Private Sub Combo_Customer_Change()
'If not in the list, we clear the combo box
    If Combo_Customer.ListIndex = -1 Then
        Combo_Customer = ""
    End If
End Sub

Private Sub Combo_Customer_Click()
' Once the customer is selected, We need to find the Exchange rate if multi-currency
' To do so, the use the currency used by the customer and find the exchange rate associated with this currency

'Clear the form
    Call FormClear
    If Combo_Customer <> "" Then
'Get the customer info from the collection
        Dim strCustomerArray() As String
        strCustomerArray = colCustomerCurrencyList.Item(Combo_Customer) 'I am using the name in the combo box to retrieve the info about
                                                                        'This customer from the collection
        
        If strCustomerArray(1) <> "" Then 'If multicurrency is turned on in the preference, The currencyref is not ""
            Text_ExRate = GetCurExRate(strCustomerArray(1)) ' Call the GetCurExRate that will use the currency ListID to find the Exchange Rate
        End If
    End If
End Sub

Private Sub Combo_First_Item_Change()
'If not in the list, we clear the combo box
    If Combo_First_Item.ListIndex = -1 Then
        Combo_First_Item = ""
    End If
End Sub

Private Sub Combo_First_Item_Click()
'Once the Item is selected, Populate the Description, rate and tax rate fields

    If Combo_First_Item <> "" Then
'Get the Item info from the collection
        Dim strItemArray() As String
        strItemArray = colItemList.Item(Combo_First_Item) 'I am using the name in the combo box to retrieve the info about
                                                                        'This item from the collection
        'strItemArray(0) is the name of the item
        'strItemArray(1) is the taxcode associated to this currency
        'strItemArray(2) is Description associated to this item
        'strItemArray(3) is the price associated to this item
        'strItemArray(4) tell us if it's a sale or a purchase desc and price
        Text_Desc_First_Item = strItemArray(2)
        Text_First_Rate = strItemArray(3)
        Combo_First_TaxCode.Text = strItemArray(1)
    End If
End Sub

Private Sub Combo_First_TaxCode_Change()
'If not in the list, we clear the combo box
    If Combo_First_TaxCode.ListIndex = -1 Then
        Combo_First_TaxCode = ""
    End If
    CalculateTaxAndTotal
End Sub

Private Sub Combo_First_TaxCode_Click()
'Recalculate the total and taxes
    CalculateTaxAndTotal
End Sub
Private Sub Combo_Second_Item_Change()
'If not in the list, we clear the combo box
    If Combo_Second_Item.ListIndex = -1 Then
        Combo_Second_Item = ""
    End If
    CalculateTaxAndTotal
End Sub

Private Sub Combo_Second_Item_Click()
'Once the Item is selected, Populate the Description, rate and tax rate fields

    If Combo_Second_Item <> "" Then
'Get the Item info from the collection
        Dim strItemArray() As String
        strItemArray = colItemList.Item(Combo_Second_Item) 'I am using the name in the combo box to retrieve the info about
                                                                        'This item from the collection
        'strItemArray(0) is the name of the item
        'strItemArray(1) is the taxcode associated to this currency
        'strItemArray(2) is Description associated to this item
        'strItemArray(3) is the price associated to this item
        'strItemArray(4) tell us if it's a sale or a purchase desc and price
        Text_Desc_Second_Item = strItemArray(2)
        Text_Second_Rate = strItemArray(3)
        Combo_Second_TaxCode.Text = strItemArray(1)
    End If
End Sub

Private Sub Combo_Second_TaxCode_Change()
'If not in the list, we clear the combo box
    If Combo_Second_TaxCode.ListIndex = -1 Then
        Combo_Second_TaxCode = ""
    End If
    CalculateTaxAndTotal

End Sub

Private Sub Combo_Second_TaxCode_Click()
'Recalculate the total and taxes
    CalculateTaxAndTotal
End Sub

Private Sub Combo_Terms_Change()
'If not in the list, we clear the combo box
    If Combo_Terms.ListIndex = -1 Then
        Combo_Terms = ""
    End If
End Sub

Private Sub Combo_Third_Item_Change()
'If not in the list, we clear the combo box
    If Combo_Third_Item.ListIndex = -1 Then
        Combo_Third_Item = ""
    End If
End Sub

Private Sub Combo_Third_Item_Click()
'Once the Item is selected, Populate the Description, rate and tax rate fields

    If Combo_Third_Item <> "" Then
'Get the Item info from the collection
        Dim strItemArray() As String
        strItemArray = colItemList.Item(Combo_Third_Item) 'I am using the name in the combo box to retrieve the info about
                                                                        'This item from the collection
        'strItemArray(0) is the name of the item
        'strItemArray(1) is the taxcode associated to this currency
        'strItemArray(2) is Description associated to this item
        'strItemArray(3) is the price associated to this item
        'strItemArray(4) tell us if it's a sale or a purchase desc and price
        Text_Desc_Third_Item = strItemArray(2)
        Text_Third_Rate = strItemArray(3)
        Combo_Third_TaxCode.Text = strItemArray(1)
    End If
End Sub

Private Sub Combo_Third_TaxCode_Change()
'If not in the list, we clear the combo box
    If Combo_Third_TaxCode.ListIndex = -1 Then
        Combo_Third_TaxCode = ""
    End If
    CalculateTaxAndTotal
End Sub

Private Sub Combo_Third_TaxCode_Click()
'Recalculate the total and taxes
    CalculateTaxAndTotal

End Sub

Private Sub Command_Clear_Click()
    Combo_Customer = ""
'I call the FormClear sub
    Call FormClear
End Sub

Private Sub Command_Close_Click()
    CloseConnection ' Make sure that all connections are closed with QuickBooks
    Unload Me ' close the window
End Sub

Private Sub Command_Save_Click()
'This function is used to save the new invoice

    If Combo_Customer = "" Then 'Need to have a customer selected
        MsgBox "Please select a customer", , "Customer missing"
        Combo_Customer.SetFocus
    ElseIf Combo_First_Item = "" And Combo_Second_Item = "" And Combo_Third_Item = "" Then
        MsgBox "Please select at least an item", , "Item missing"
        Combo_First_Item.SetFocus
    ElseIf "" <> Text_Date And Not IsDate(Text_Date) Then 'Verify the date format
        MsgBox "The Date Format is wrong" & vbCrLf & "The format of the date should be:" & vbCrLf & _
            "DD/MM/YY" & vbCrLf & "DD/MM/YYYY" & vbCrLf & "DD-MM-YYYY" & vbCrLf & vbCrLf & _
            "Examples: 14/08/02 14/08/2002 14-08-02 14-08-2002"
        Text_Date.SetFocus
    
    Else
        Dim blnCurrencySame As Boolean
        blnCurrencySame = True
        If Text_Date <> "" Then 'make sure that the date is in the right format for QuickBooks
           Text_Date = toQBDate(Text_Date)
        End If
        

        Dim strXMLRequest As String
        
        strXMLRequest = GenerateInvoiceXMLRequest 'Create the XML request using the information on the form
        'Save the invoice in QuickBooks
        FuncTionSaveInvoice (strXMLRequest) 'Call this function to send the request to QuickBooks
    End If
End Sub

Private Sub Form_Load()
' The FormLoad function is used to populate the different combo boxes on the form.
' It stores the information received from QuickBooks in order to use them later.


    Dim appPath
    appPath = App.Path
    Image_QBBANNER.Picture = LoadPicture(appPath & "/qbbanner.bmp")
    
    Text_Date = toQBDate(Date)
' Populating the Customer combo box
    Dim intIndex As Integer
    
    
    intIndex = 0
        Set colCustomerCurrencyList = New Collection
        'Call the GetCustomerList procedure that will populate the colCustomerCurrencyList collection
        GetCustomerList
    If blnIsOpenConnection = True Then ' If we could not connect with the first function, I do not continue

        intIndex = 1
        Dim strCustomerArray() As String
        'We loop through the customer arrays in the colCustomerCurrencyList collection
        While intIndex <= colCustomerCurrencyList.Count
            'Adding the customer's name to the combo box
            strCustomerArray = colCustomerCurrencyList.Item(intIndex)
            Combo_Customer.AddItem strCustomerArray(0)
            intIndex = intIndex + 1 ' Go to the next customer
        Wend
    
    

    '' Populating the Item combo box
        Set colItemList = New Collection
        'Call the GetItemList procedure that will populate the colItemList collection
        GetItemList

        intIndex = 1
        Dim strItemArray() As String
        'We loop through the item arrays in the colItemList collection
        While intIndex <= colItemList.Count
            'Adding the Item name to all three item combo boxes
            strItemArray = colItemList.Item(intIndex)
            Combo_First_Item.AddItem strItemArray(0)
            Combo_Second_Item.AddItem strItemArray(0)
            Combo_Third_Item.AddItem strItemArray(0)
            intIndex = intIndex + 1 ' Go to the next item
        Wend
    
    
    
    
    ' Populating the Tax Code combo boxes
    
        Set colTaxCodeList = New Collection
        'Call the GetTaxCodeList procedure that will populate the colItemList collection
        GetTaxCodeList
        intIndex = 1
        Dim strTaxCodeArray() As String
        'We loop through the taxcode arrays in the colTaxCodeList collection
        While intIndex <= colTaxCodeList.Count
            'Adding the TaxCode name to all three TaxCode combo boxes
            strTaxCodeArray = colTaxCodeList.Item(intIndex)
            Combo_First_TaxCode.AddItem strTaxCodeArray(0)
            Combo_Second_TaxCode.AddItem strTaxCodeArray(0)
            Combo_Third_TaxCode.AddItem strTaxCodeArray(0)
            intIndex = intIndex + 1 ' Go to the next TaxCode
        Wend
        
        
        
        
    '
    ' Populating the ArAccount combo box
        Dim intArAccountCount As Integer
        Dim StringArAccountArray() As String 'This String array will contain the list of item names returned from QuickBooks
        intArAccountCount = 0
        intIndex = 1
    
        'Call the function to retrieve the list of accounts
        Set colArAccount = New Collection
        'Call the GetArAccountList procedure that will populate the colArAccount collection
        GetArAccountList
    
        Dim strArAccountArray() As String
        'We loop through the ArAccount arrays in the colArAccount collection
        While intIndex <= colArAccount.Count
            'Adding the ArAccount name to the ArAccount combo box
            strArAccountArray = colArAccount.Item(intIndex)
            Combo_ArAccount.AddItem strArAccountArray(0)
            intIndex = intIndex + 1 ' Go to the next ArAccount
        Wend

    End If
End Sub

Private Sub Text_First_Amount_LostFocus()
'Process the entry of a new amount
    If Text_First_Amount <> "" Then
        If Not IsNumeric(Text_First_QTY) Then
            MsgBox "Please enter a numeric value", , "Enter an amount"
            Text_First_Amount = ""
        Else
            If Text_First_QTY <> "" Then
            
                Text_First_Rate = Text_First_Amount / Text_First_QTY
                Text_First_Rate = FormatNumber(Text_First_Rate)
            ElseIf Text_First_QTY = "" Then
                Text_First_Rate = FormatNumber(Text_First_Amount)
            End If
            Text_First_Amount = FormatNumber(Text_First_Amount)
        End If
        Call CalculateTaxAndTotal 'Call the function that calculated the taxes and total

    End If

End Sub

Private Sub Text_First_QTY_LostFocus()
'Process the entry of a new Quantity
    If Text_First_QTY <> "" Then
        If Not IsNumeric(Text_First_QTY) Then
            MsgBox "Please enter a numeric value", , "Enter a qty"
            Text_First_QTY = ""
        Else
            If Text_First_Rate <> "" Then
                Text_First_Amount = FormatNumber(Text_First_Rate * Text_First_QTY)
            End If
            
        End If
        Call CalculateTaxAndTotal 'Call the function that calculated the taxes and total
    End If
End Sub


Private Sub Text_First_Rate_LostFocus()
'Process the entry of a new Rate
    If Text_First_Rate <> "" Then
        If Not IsNumeric(Text_First_Rate) Then
            MsgBox "Please enter a numeric value", , "Enter a Rate"
            Text_First_Rate = ""
        Else
            If Text_First_Amount <> "" And Text_First_QTY = "" Then
                Text_First_Amount = FormatNumber(Text_First_Rate)
            ElseIf Text_First_QTY <> "" Then
                Text_First_Amount = FormatNumber(Text_First_Rate * Text_First_QTY)
            End If
           Text_First_Rate = FormatNumber(Text_First_Rate)
        End If
        Call CalculateTaxAndTotal 'Call the function that calculated the taxes and total
        
    End If

End Sub



Private Sub Text_Second_Amount_LostFocus()
'Process the entry of a new Amount
    If Text_Second_Amount <> "" Then
        If Not IsNumeric(Text_Second_QTY) Then
            MsgBox "Please enter a numeric value", , "Enter an amount"
            Text_Second_Amount = ""
        Else
            If Text_Second_QTY <> "" Then
            
                Text_Second_Rate = Text_Second_Amount / Text_Second_QTY
                Text_Second_Rate = FormatNumber(Text_Second_Rate)
            ElseIf Text_Second_QTY = "" Then
                Text_Second_Rate = FormatNumber(Text_Second_Amount)
            End If
            Text_Second_Amount = FormatNumber(Text_Second_Amount)
        End If
        Call CalculateTaxAndTotal 'Call the function that calculated the taxes and total

    End If


End Sub

Private Sub Text_Second_QTY_LostFocus()
'Process the entry of a new Quantity
    If Text_Second_QTY <> "" Then
        If Not IsNumeric(Text_Second_QTY) Then
            MsgBox "Please enter a numeric value", , "Enter a qty"
            Text_Second_QTY = ""
        Else
            If Text_Second_Rate <> "" Then
                Text_Second_Amount = FormatNumber(Text_Second_Rate * Text_Second_QTY)
            End If
            
        End If
        Call CalculateTaxAndTotal 'Call the function that calculated the taxes and total
    End If
End Sub
Private Sub Text_Second_Rate_LostFocus()
'Process the entry of a new Rate
    If Text_Second_Rate <> "" Then
        If Not IsNumeric(Text_Second_Rate) Then
            MsgBox "Please enter a numeric value", , "Enter a Rate"
            Text_Second_Rate = ""
        Else
            If Text_Second_Amount <> "" And Text_Second_QTY = "" Then
                Text_Second_Amount = FormatNumber(Text_Second_Rate)
            ElseIf Text_Second_QTY <> "" Then
                Text_Second_Amount = FormatNumber(Text_Second_Rate * Text_Second_QTY)
            End If
           Text_Second_Rate = FormatNumber(Text_Second_Rate)
        End If
        Call CalculateTaxAndTotal 'Call the function that calculated the taxes and total
    End If
End Sub
Private Sub Text_Third_Amount_LostFocus()
'Process the entry of a new Amount
    If Text_Third_Amount <> "" Then
        If Not IsNumeric(Text_Third_QTY) Then
            MsgBox "Please enter a numeric value", , "Enter an amount"
            Text_Third_Amount = ""
        Else
            If Text_Third_QTY <> "" Then
            
                Text_Third_Rate = Text_Third_Amount / Text_Third_QTY
                Text_Third_Rate = FormatNumber(Text_Third_Rate)
            ElseIf Text_Third_QTY = "" Then
                Text_Third_Rate = FormatNumber(Text_Third_Amount)
            End If
            Text_Third_Amount = FormatNumber(Text_Third_Amount)
        End If
        Call CalculateTaxAndTotal 'Call the function that calculated the taxes and total

    End If

End Sub

Private Sub Text_Third_QTY_LostFocus()
'Process the entry of a new Quantity
    If Text_Third_QTY <> "" Then
        If Not IsNumeric(Text_Third_QTY) Then
            MsgBox "Please enter a numeric value", , "Enter a qty"
            Text_Third_QTY = ""
        Else
            If Text_Third_Rate <> "" Then
                Text_Third_Amount = FormatNumber(Text_Third_Rate * Text_Third_QTY)
            End If
            
        End If
        Call CalculateTaxAndTotal 'Call the function that calculated the taxes and total
    End If

End Sub


Private Sub Text_Third_Rate_LostFocus()
'Process the entry of a new Rate
    If Text_Third_Rate <> "" Then
        If Not IsNumeric(Text_Third_Rate) Then
            MsgBox "Please enter a numeric value", , "Enter a Rate"
            Text_Third_Rate = ""
        Else
            If Text_Third_Amount <> "" And Text_Third_QTY = "" Then
                Text_Third_Amount = FormatNumber(Text_Third_Rate)
            ElseIf Text_Third_QTY <> "" Then
                Text_Third_Amount = FormatNumber(Text_Third_Rate * Text_Third_QTY)
            End If
           Text_Third_Rate = FormatNumber(Text_Third_Rate)
        End If
        Call CalculateTaxAndTotal 'Call the function that calculated the taxes and total
        
    End If

End Sub


Public Function toQBDate(theDate As String) As String
'Make sure that the date is in a format that QuickBooks accepts
On Error GoTo ErrHandler
    toQBDate = Format(theDate, "yyyy-mm-dd")
    
    If toQBDate <> "" And Not IsDate(toQBDate) Then
        toQBDate = "error"
        Exit Function
    End If
    
    Exit Function

ErrHandler:
    toQBDate = "error formating the date"
End Function
Sub SetNumberInCurrency()
'Format the rate and amount fields into a field with 2 decimal point
    If Text_First_Rate <> "" Then
        Text_First_Rate = FormatNumber(Text_First_Rate)
    End If
    If Text_First_Amount <> "" Then
        Text_First_Amount = FormatNumber(Text_First_Amount)
    End If
    If Text_Second_Rate <> "" Then
        Text_Second_Rate = FormatNumber(Text_Second_Rate)
    End If
    If Text_Second_Amount <> "" Then
        Text_Second_Amount = FormatNumber(Text_Second_Amount)
    End If
    If Text_Third_Rate <> "" Then
        Text_Third_Rate = FormatNumber(Text_Third_Rate)
    End If
    If Text_Third_Amount <> "" Then
        Text_Third_Amount = FormatNumber(Text_Third_Amount)
    End If
End Sub

Sub CalculateTaxAndTotal()
'Calculating the Taxes and total
    Text_GST = ""
    Text_PST = ""
    
    Dim strTaxCodeArray() As String 'I am using the name in the combo box to retrieve the info about
                                                                    'This ArAccount from the collection

    'strTaxCodeArray(0) is TaxCode Name
    'strTaxCodeArray (1) is GST (TaxRate1)
    'strTaxCodeArray(2) is PST (TaxRate2)




    Dim strTempTaxRate As String

    If Combo_First_TaxCode <> "" Then
        'If tax rate selected for the first item...
        Dim strFirstComboTaxRate1 As String
        Dim strFirstComboTaxRate2 As String
        strTaxCodeArray = colTaxCodeList.Item(Combo_First_TaxCode) 'We  retrieve the taxcode array corresponding to the taxcode selected in the combo box

        If strTaxCodeArray(1) <> "" Then
            strFirstComboTaxRate1 = strTaxCodeArray(1) / 100
        End If
        If strTaxCodeArray(2) <> "" Then
            strFirstComboTaxRate2 = strTaxCodeArray(2) / 100
        Else
            strFirstComboTaxRate2 = "0"
        End If
        
    End If

    If Combo_Second_TaxCode <> "" Then
        'If tax rate selected for the second item...
        Dim strSecondComboTaxRate1 As String
        Dim strSecondComboTaxRate2 As String
        strTaxCodeArray = colTaxCodeList.Item(Combo_Second_TaxCode) 'We  retrieve the taxcode array corresponding to the taxcode selected in the combo box

        If strTaxCodeArray(1) <> "" Then
            strSecondComboTaxRate1 = strTaxCodeArray(1) / 100
        End If
        If strTaxCodeArray(2) <> "" Then
            strSecondComboTaxRate2 = strTaxCodeArray(2) / 100
        Else
            strSecondComboTaxRate2 = "0"
        End If

    End If

    If Combo_Third_TaxCode <> "" Then
        'If tax rate selected for the third item...
        Dim strThirdComboTaxRate1 As String
        Dim strThirdComboTaxRate2 As String
        strTaxCodeArray = colTaxCodeList.Item(Combo_Third_TaxCode) 'We  retrieve the taxcode array corresponding to the taxcode selected in the combo box

        If strTaxCodeArray(1) <> "" Then
            strThirdComboTaxRate1 = strTaxCodeArray(1) / 100
        End If
        If strTaxCodeArray(2) <> "" Then
            strThirdComboTaxRate2 = strTaxCodeArray(2) / 100
        Else
            strThirdComboTaxRate2 = "0"
            
        End If

    End If
'Calculate the Total GST and PST
    If Text_First_Amount <> "" And Combo_First_TaxCode <> "" Then
        Text_GST = Text_First_Amount * strFirstComboTaxRate1
        Text_PST = Text_First_Amount * strFirstComboTaxRate2
    End If

    If Text_Second_Amount <> "" And Combo_Second_TaxCode <> "" Then
        If Text_GST = "" Then '
            Text_GST = Text_Second_Amount * strSecondComboTaxRate1
        Else
            Text_GST = Text_GST + Text_Second_Amount * strSecondComboTaxRate1
        End If
        If Text_PST = "" Then
            Text_PST = Text_Second_Amount * strSecondComboTaxRate2
        Else
            Text_PST = Text_PST + Text_Second_Amount * strSecondComboTaxRate2
        End If
    End If

    If Text_Third_Amount <> "" And Combo_Third_TaxCode <> "" Then
        If Text_GST = "" Then
            Text_GST = Text_Third_Amount * strThirdComboTaxRate1
        Else
           Text_GST = Text_GST + Text_Third_Amount * strThirdComboTaxRate1

        End If

        If Text_PST = "" Then
            Text_PST = Text_Third_Amount * strThirdComboTaxRate2
        Else
            Text_PST = Text_PST + Text_Third_Amount * strThirdComboTaxRate2
        End If
    End If
    If Text_GST <> "" Then
    'Format the GST amount with 2 decimal
        Text_GST = FormatNumber(Text_GST)
    End If
    If Text_PST <> "" Then
    'Format the PST amount with 2 decimal
        Text_PST = FormatNumber(Text_PST)
    End If


'Calculate the total
    Text_Total = 0
    If Text_First_Amount <> "" Then
        Text_Total = Text_Total + CDbl(Text_First_Amount)
    End If
    If Text_Second_Amount <> "" Then
        Text_Total = Text_Total + CDbl(Text_Second_Amount)
    End If
    If Text_Third_Amount <> "" Then
        Text_Total = Text_Total + CDbl(Text_Third_Amount)
    End If
    If Text_GST <> "" Then
        Text_Total = Text_Total + CDbl(Text_GST)
    End If
    If Text_PST <> "" Then
        Text_Total = Text_Total + CDbl(Text_PST)
    End If
    If Text_Total <> "" Then
        Text_Total = FormatNumber(Text_Total)
    End If
End Sub
Function GenerateInvoiceXMLRequest() As String
    ' We generate the request qbXML in order to get the ArAccount list with filter to have only account receivables.
   Dim requestXML As String
    requestXML = "<?xml version=""1.0"" ?>" & _
    "<?qbxml version=""CA2.0""?>" & _
    "<QBXML><QBXMLMsgsRq onError=""continueOnError"">" & _
    "<InvoiceAddRq requestID=""1"">" & _
    "<InvoiceAdd>" & _
    "<CustomerRef>" & _
       "<FullName>" & Combo_Customer & "</FullName>" & _
    "</CustomerRef>"
    If Combo_ArAccount <> "" Then
        requestXML = requestXML + "<ARAccountRef>" & _
           "<FullName>" & Combo_ArAccount & "</FullName>" & _
        "</ARAccountRef>"
    End If
    
    If Text_Date <> "" Then
        requestXML = requestXML + "<TxnDate>" & Text_Date & "</TxnDate>"
    End If
'If the first item is selected...
    If Combo_First_Item <> "" Then
        requestXML = requestXML + "<InvoiceLineAdd>" & _
        "<ItemRef>" & _
           "<FullName>" & Combo_First_Item & "</FullName>" & _
        "</ItemRef>"
        If Text_Desc_First_Item <> "" Then
            requestXML = requestXML + "<Desc>" & Text_Desc_First_Item & "</Desc>"
        End If
        
        If Text_First_QTY <> "" Then
            requestXML = requestXML + "<Quantity>" & Text_First_QTY & "</Quantity>"
        End If
    
        If Text_First_Rate <> "" Then
            requestXML = requestXML + "<Rate>" & Text_First_Rate & "</Rate>"
        End If
   
        If Combo_First_TaxCode <> "" Then
            requestXML = requestXML + "<TaxCodeRef>" & _
               "<FullName>" & Combo_First_TaxCode & "</FullName>" & _
               "</TaxCodeRef>"
        End If
        requestXML = requestXML & "</InvoiceLineAdd>"
    End If
'If the second item is selected...
    If Combo_Second_Item <> "" Then
        requestXML = requestXML + "<InvoiceLineAdd>" & _
        "<ItemRef>" & _
           "<FullName>" & Combo_Second_Item & "</FullName>" & _
        "</ItemRef>"
        If Text_Desc_Second_Item <> "" Then
            requestXML = requestXML + "<Desc>" & Text_Desc_Second_Item & "</Desc>"
        End If
        
        If Text_Second_QTY <> "" Then
            requestXML = requestXML + "<Quantity>" & Text_Second_QTY & "</Quantity>"
        End If
    
        If Text_Second_Rate <> "" Then
            requestXML = requestXML + "<Rate>" & Text_Second_Rate & "</Rate>"
        End If
    
         If Combo_Second_TaxCode <> "" Then
            requestXML = requestXML + "<TaxCodeRef>" & _
               "<FullName>" & Combo_Second_TaxCode & "</FullName>" & _
               "</TaxCodeRef>"
        End If
        requestXML = requestXML & "</InvoiceLineAdd>"
        
        
    End If
'If the third item is selected...
    If Combo_Third_Item <> "" Then
        requestXML = requestXML + "<InvoiceLineAdd>" & _
        "<ItemRef>" & _
           "<FullName>" & Combo_Third_Item & "</FullName>" & _
        "</ItemRef>"
        If Text_Desc_Third_Item <> "" Then
            requestXML = requestXML + "<Desc>" & Text_Desc_Third_Item & "</Desc>"
        End If
        
        If Text_Third_QTY <> "" Then
            requestXML = requestXML + "<Quantity>" & Text_Third_QTY & "</Quantity>"
        End If
    
        If Text_Third_Rate <> "" Then
            requestXML = requestXML + "<Rate>" & Text_Third_Rate & "</Rate>"
        End If
    
    
         If Combo_Third_TaxCode <> "" Then
            requestXML = requestXML + "<TaxCodeRef>" & _
               "<FullName>" & Combo_Third_TaxCode & "</FullName>" & _
               "</TaxCodeRef>"
        End If
        requestXML = requestXML & "</InvoiceLineAdd>"

    End If
'Add the exchange rate to the request
    If Text_ExRate <> "" Then
        requestXML = requestXML + "<ExchangeRate>" & Text_ExRate & "</ExchangeRate>"
    End If
    
    requestXML = requestXML + "</InvoiceAdd>" & _
    "</InvoiceAddRq>" & _
    "</QBXMLMsgsRq></QBXML>"
    
    GenerateInvoiceXMLRequest = requestXML 'return the XML request


End Function
Sub FormClear()
'Clear the form
    Combo_ArAccount = ""
    Combo_First_Item = ""
    Text_Desc_First_Item = ""
    Text_First_QTY = ""
    Text_First_Rate = ""
    Text_First_Amount = ""
    Combo_First_TaxCode = ""

    Combo_Second_Item = ""
    Text_Desc_Second_Item = ""
    Text_Second_QTY = ""
    Text_Second_Rate = ""
    Text_Second_Amount = ""
    Combo_Second_TaxCode = ""

    Combo_Third_Item = ""
    Text_Desc_Third_Item = ""
    Text_Third_QTY = ""
    Text_Third_Rate = ""
    Text_Third_Amount = ""
    Combo_Third_TaxCode = ""
    
    Text_ExRate = ""
    Text_GST = ""
    Text_PST = ""
    Text_Total = ""
End Sub
