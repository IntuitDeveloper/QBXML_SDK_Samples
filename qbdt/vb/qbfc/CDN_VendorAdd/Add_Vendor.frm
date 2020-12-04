VERSION 5.00
Begin VB.Form AddVendor 
   Caption         =   "Add a vendor"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command_Clear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   8040
      TabIndex        =   36
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox Text_Note_MultiCurrency 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5400
      TabIndex        =   34
      Text            =   "This data file doesn't have multicurrency turned on"
      Top             =   4440
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.ComboBox Combo_Currency 
      Height          =   315
      Left            =   5400
      TabIndex        =   15
      Top             =   3840
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.ComboBox Combo_Vendor_Type 
      Height          =   315
      Left            =   5400
      TabIndex        =   14
      Top             =   3240
      Width           =   2055
   End
   Begin VB.ComboBox Combo_State_Prov 
      Height          =   315
      Left            =   1560
      TabIndex        =   7
      Top             =   4440
      Width           =   855
   End
   Begin VB.TextBox Text_Postal_Code 
      Height          =   285
      Left            =   2640
      TabIndex        =   8
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox Text_EMail 
      Height          =   285
      Left            =   5400
      TabIndex        =   11
      Top             =   1710
      Width           =   2055
   End
   Begin VB.TextBox Text_Alt_Phone 
      Height          =   285
      Left            =   5400
      TabIndex        =   10
      Top             =   1260
      Width           =   2055
   End
   Begin VB.TextBox Text_Phone 
      Height          =   285
      Left            =   5400
      TabIndex        =   9
      Top             =   810
      Width           =   2055
   End
   Begin VB.TextBox Text_Name_On_Check 
      Height          =   285
      Left            =   5400
      TabIndex        =   12
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox Text_First_Name 
      Height          =   288
      Left            =   1560
      TabIndex        =   1
      Top             =   1365
      Width           =   1215
   End
   Begin VB.TextBox Text_Last_Name 
      Height          =   288
      Left            =   1560
      TabIndex        =   3
      Top             =   1905
      Width           =   2055
   End
   Begin VB.TextBox Text_Salutation 
      Height          =   288
      Left            =   1560
      TabIndex        =   0
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text_Address_1 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   2670
      Width           =   2055
   End
   Begin VB.TextBox Text_Address_2 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   3195
      Width           =   2055
   End
   Begin VB.TextBox Text_City 
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Top             =   3720
      Width           =   2055
   End
   Begin VB.TextBox Text_Account_Number 
      Height          =   285
      Left            =   5400
      TabIndex        =   13
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CommandButton Command_Add_Vendor 
      Caption         =   "Add Vendor"
      Height          =   495
      Left            =   8040
      TabIndex        =   17
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton Command_Exit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   8040
      TabIndex        =   16
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox Text_Middle_Name 
      Height          =   285
      Left            =   3240
      TabIndex        =   2
      Top             =   1380
      Width           =   375
   End
   Begin VB.Label Label_Note 
      Caption         =   "Note on data file"
      Height          =   375
      Left            =   4320
      TabIndex        =   35
      Top             =   4440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label_Multicurrency 
      Caption         =   "Currency"
      Height          =   375
      Left            =   4320
      TabIndex        =   33
      Top             =   3840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Vendor Type"
      Height          =   375
      Left            =   4320
      TabIndex        =   32
      Top             =   3240
      Width           =   855
   End
   Begin VB.Image Image_QBBANNER 
      Height          =   615
      Left            =   1440
      Top             =   5640
      Width           =   3015
   End
   Begin VB.Label Label7 
      Caption         =   "Postal Code"
      Height          =   255
      Left            =   2640
      TabIndex        =   31
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "State/Province"
      Height          =   375
      Left            =   1440
      TabIndex        =   30
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label16 
      Caption         =   "Add Vendor"
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
      Left            =   600
      TabIndex        =   29
      Top             =   240
      Width           =   4635
   End
   Begin VB.Label Label1 
      Caption         =   "First Name"
      Height          =   255
      Left            =   600
      TabIndex        =   28
      Top             =   1380
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Last Name"
      Height          =   255
      Left            =   600
      TabIndex        =   27
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Salutation"
      Height          =   255
      Left            =   600
      TabIndex        =   26
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Address"
      Height          =   375
      Left            =   600
      TabIndex        =   25
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "City"
      Height          =   375
      Left            =   600
      TabIndex        =   24
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Phone"
      Height          =   375
      Left            =   4320
      TabIndex        =   23
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "E-mail"
      Height          =   375
      Left            =   4320
      TabIndex        =   22
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "Alt. Ph."
      Height          =   375
      Left            =   4320
      TabIndex        =   21
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label13 
      Caption         =   "Name on check"
      Height          =   375
      Left            =   4320
      TabIndex        =   20
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "Account Number"
      Height          =   375
      Left            =   4320
      TabIndex        =   19
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label17 
      Caption         =   "M.I."
      Height          =   255
      Left            =   2880
      TabIndex        =   18
      Top             =   1440
      Width           =   495
   End
End
Attribute VB_Name = "AddVendor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Add_Vendor.frm
' This form is part of the Add Vendor sample program
' for the Canadian QuickBooks SDK Version CA2.0 and US QuickBooks SDK version 2.0.
'
' Created October, 2002
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/rdmgr/?ID=100
'
'-------------------------------------------------------------


Private Sub Combo_Currency_Change()
'If not in the list, we clear the combo box
    If Combo_Currency.ListIndex = -1 Then
        Combo_Currency = ""
    End If

End Sub

Private Sub Combo_State_Prov_Change()
'If not in the list, we clear the combo box
    If Combo_State_Prov.ListIndex = -1 Then
        Combo_State_Prov = ""
    End If

End Sub

Private Sub Combo_Vendor_Type_Change()
'If not in the list, we clear the combo box
    If Combo_Vendor_Type.ListIndex = -1 Then
        Combo_Vendor_Type = ""
    End If

End Sub

Private Sub Command_Add_Vendor_Click()
'Verify the required fields
    If Text_First_Name = "" And Text_Last_Name = "" Then
        MsgBox "You need at least a first name or a last name", , "Missing name"
        Text_First_Name.SetFocus
    Else
        If strSDKVersion = "US" Then 'Save using the US SDK
            SaveWithUSSDK
        Else 'Save using the Canadian SDK
            SaveWithCanadaSDK
        End If
    End If

End Sub

Private Sub Command_Clear_Click()
'Clear the form
    Text_Salutation = ""
    Text_First_Name = ""
    Text_Last_Name = ""
    Text_Middle_Name = ""
    Text_Address_1 = ""
    Text_Address_2 = ""
    Text_City = ""
    Combo_State_Prov = ""
    Text_Postal_Code = ""
    Text_Phone = ""
    Text_Alt_Phone = ""
    Text_EMail = ""
    Text_Name_On_Check = ""
    Text_Account_Number = ""
    Combo_Vendor_Type = ""
    Combo_Currency = ""

End Sub

Private Sub Command_Exit_Click()
    If blnConnectionIsOpen = True Then 'If the connection is still open
        CloseQBFCConnection  'The procedure to close the connection is called
    End If
    Unload Me ' close the app
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler

    Dim appPath
    appPath = App.Path
    Image_QBBANNER.Picture = LoadPicture(appPath & "/qbbanner.bmp")
    
    Dim blnIsOpenConnection As Boolean
'Find Which SDK version QuickBooks supports.  When I call the OpenConnection, the SDKVersion public
'variable is set and tell us if the app is connected to a Canadian QuickBooks or a US QuickBooks

    blnIsOpenConnection = OpenConnection

    If blnIsOpenConnection = True Then 'If the connection was established, I populate the combo boxex
        
        'Populate the combo boxes
        If strSDKVersion = "Canada" Then 'List of province abbreviations
            Combo_State_Prov.AddItem "AB"
            Combo_State_Prov.AddItem "BC"
            Combo_State_Prov.AddItem "MB"
            Combo_State_Prov.AddItem "NB"
            Combo_State_Prov.AddItem "NF"
            Combo_State_Prov.AddItem "NS"
            Combo_State_Prov.AddItem "NT"
            Combo_State_Prov.AddItem "NU"
            Combo_State_Prov.AddItem "ON"
            Combo_State_Prov.AddItem "PE"
            Combo_State_Prov.AddItem "QC"
            Combo_State_Prov.AddItem "SK"
            Combo_State_Prov.AddItem "YT"
           
            
        Else 'Partial list of the US states
            Combo_State_Prov.AddItem "AL"
            Combo_State_Prov.AddItem "AK"
            Combo_State_Prov.AddItem "AZ"
            Combo_State_Prov.AddItem "AR"
            Combo_State_Prov.AddItem "CA"
            Combo_State_Prov.AddItem "CO"
            Combo_State_Prov.AddItem "DC"
            Combo_State_Prov.AddItem "FL"
            Combo_State_Prov.AddItem "GA"
            
            
        End If
            Dim strVendorTypeString As String
            Dim VendorTypeStringArray() As String 'This String array will contain the list of Vendor Type returned from QuickBooks
            Dim nIndex As Integer
            Dim intVendorTypeCount As Integer
    
'Populate the Vendor Type combo box
            ' The GetVendorTypeListCA and GetVendorTypeListUS function returns a list a vendor type that are separated by "***", We use the
            'Split function to separate the customer names in order to populate the vendor type combo box
            intIndex = 0
            
            If strSDKVersion = "US" Then
                strVendorTypeString = GetVendorTypeListUS
            Else
                strVendorTypeString = GetVendorTypeListCA
            End If
            VendorTypeStringArray = Split(strVendorTypeString, "***") 'Split the vendor type list string into an array
                
            intVendorTypeCount = UBound(VendorTypeStringArray)
                
                
            If intVendorTypeCount = 0 Then ' If can't find at least one vendor type, give an error message
                MsgBox "Could not find any Vendor Type in QuickBooks", vbInformation, "No vendor type found"
            Else
                While intIndex < intVendorTypeCount
                    'Adding the customer's name to the combo box
                    Combo_Vendor_Type.AddItem VendorTypeStringArray(intIndex)
                    intIndex = intIndex + 1 'Go to the next customer
                Wend
            End If
    
'Populate the currency list if multicurrency feature is selected
            
            intIndex = 0
            
            If strSDKVersion = "Canada" Then 'Verify first that the multicurrency feature has been turned on.  The multicurrency
                                                'is set in the preferences.  We query the preferences to find if multicurrency is turned on or not.
                Dim blnMultiCurrencyOn As Boolean
                Dim strCurrencyString As String
                Dim CurrencyStringArray() As String 'This String array will contain the list of currency returned from QuickBooks
                Dim intCurrencyCount As Integer

                blnMultiCurrencyOn = FindIfMulticurrencyOn
'                blnMultiCurrencyOn = True
                If blnMultiCurrencyOn = True Then
                'I make the currency combo box visible and populate it
                    Combo_Currency.Visible = True
                    Label_Multicurrency.Visible = True
                    
                    strCurrencyString = GetCurrencyListCA
                    CurrencyStringArray = Split(strCurrencyString, "***") 'Split the vendor type list string into an array
                        
                    intCurrencyCount = UBound(CurrencyStringArray)
                        
                        
                    If intCurrencyCount = 0 Then ' If can't find at least one currency, give an error message
                        MsgBox "Could not find any currency in QuickBooks", vbInformation, "No currency found"
                    Else
                        While intIndex < intCurrencyCount
                            'Adding the customer's name to the combo box
                            Combo_Currency.AddItem CurrencyStringArray(intIndex)
                            intIndex = intIndex + 1 'Go to the next customer
                        Wend
                    End If
                Else
                    'Making the multicurrency note field visible
                    Text_Note_MultiCurrency.Visible = True
                    Label_Note.Visible = True
                    
                End If
            End If
    
    Else 'Either the connection could not be established, the connection was refused by user or one or more dlls are missing.  The error message has already been given to the user
        'we just need to shut down the app.
        Unload Me
    End If
    Exit Sub

ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
End Sub
Sub SaveWithUSSDK()
On Error GoTo ErrHandler

    'Create the message set request object
    Dim requestMsgSet As IMsgSetRequest
    Set requestMsgSet = sessionManagerUS.CreateMsgSetRequest("US", cQBXMLMajorVersion, cQBXMLMinorVersion)
    'Initialize the request's attributes
    requestMsgSet.Attributes.OnError = roeContinue
    
    'Add the request to the message set request object
    Dim VendorAdd As IVendorAdd
    Set VendorAdd = requestMsgSet.AppendVendorAddRq
    
    If Text_Salutation <> "" Then
        VendorAdd.Salutation.SetValue Text_Salutation
    End If
    
    If Text_First_Name <> "" And Text_Last_Name <> "" Then
        VendorAdd.Name.SetValue Text_First_Name & " " & Text_Last_Name
        VendorAdd.FirstName.SetValue Text_First_Name
        VendorAdd.LastName.SetValue Text_Last_Name
    ElseIf Text_First_Name <> "" Then
        VendorAdd.Name.SetValue Text_First_Name
        VendorAdd.FirstName.SetValue Text_First_Name
    ElseIf Text_Last_Name <> "" Then
        VendorAdd.Name.SetValue Text_Last_Name
        VendorAdd.LastName.SetValue Text_Last_Name
    End If
    
    If Text_Middle_Name <> "" Then
        VendorAdd.MiddleName.SetValue Text_Middle_Name
    End If
    
    If Text_Address_1 <> "" Then
        VendorAdd.VendorAddress.Addr1.SetValue Text_Address_1
    End If
    
    If Text_Address_2 <> "" Then
        VendorAdd.VendorAddress.Addr2.SetValue Text_Address_2
    End If
    If Text_City <> "" Then
        VendorAdd.VendorAddress.City.SetValue Text_City
    End If
    If Combo_State_Prov <> "" Then
        VendorAdd.VendorAddress.State.SetValue Combo_State_Prov
    End If
    
    If Text_Postal_Code <> "" Then
        VendorAdd.VendorAddress.PostalCode.SetValue Text_Postal_Code
    End If
    
    If Text_Phone <> "" Then
        VendorAdd.Phone.SetValue Text_Phone
    End If
    
    If Text_Alt_Phone <> "" Then
        VendorAdd.AltPhone.SetValue Text_Alt_Phone
    End If
    
    If Text_EMail <> "" Then
        VendorAdd.Email.SetValue Text_EMail
    End If
   
    If Text_Name_On_Check <> "" Then
        VendorAdd.NameOnCheck.SetValue Text_Name_On_Check
    End If
    
    If Text_Account_Number <> "" Then
        VendorAdd.AccountNumber.SetValue Text_Account_Number
    End If
    
    If Combo_Vendor_Type <> "" Then
        VendorAdd.VendorTypeRef.FullName.SetValue Combo_Vendor_Type
    End If
    
    
    'Perform the request
    Dim responseMsgSet As IMsgSetResponse
    Set responseMsgSet = sessionManagerUS.DoRequests(requestMsgSet)
    
    'Interpret the response
    Dim response As IResponse
   
    'The response list contains only one response,
    'which corresponds to our single request
    Set response = responseMsgSet.ResponseList.GetAt(0)
    
    If (response.statusCode <> 0) Then
        MsgBox "Status: Code = " & CStr(response.statusCode) & _
          ", Severity = " & response.statusSeverity & _
          ", Message = " & response.statusMessage
    End If
        
    'The response detail for Add and Mod requests is a 'Ret' object
    'In our case, it's IVendorRet
    Dim VendorRet As IVendorRet
    Set VendorRet = response.Detail
    If (Not (VendorRet Is Nothing)) Then
        MsgBox "A new vendor has been created with ListID = " & VendorRet.ListID.GetValue, , "Vendor added"
    End If
   'Clear the form
    Command_Clear_Click
   
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"


End Sub
Sub SaveWithCanadaSDK()
On Error GoTo ErrHandler

    'Create the message set request object
    Dim requestMsgSet As IMsgSetRequest
    Set requestMsgSet = sessionManagerCA.CreateMsgSetRequest("CA", cQBXMLMajorVersion, cQBXMLMinorVersion)
    'Initialize the request's attributes
    requestMsgSet.Attributes.OnError = roeContinue
    
    'Add the request to the message set request object
    Dim VendorAdd As IVendorAdd
    Set VendorAdd = requestMsgSet.AppendVendorAddRq

    If Text_Salutation <> "" Then
        VendorAdd.Salutation.SetValue Text_Salutation
    End If
    
    If Text_First_Name <> "" And Text_Last_Name <> "" Then
        VendorAdd.Name.SetValue Text_First_Name & " " & Text_Last_Name
        VendorAdd.FirstName.SetValue Text_First_Name
        VendorAdd.LastName.SetValue Text_Last_Name
    ElseIf Text_First_Name <> "" Then
        VendorAdd.Name.SetValue Text_First_Name
        VendorAdd.FirstName.SetValue Text_First_Name
    ElseIf Text_Last_Name <> "" Then
        VendorAdd.Name.SetValue Text_Last_Name
        VendorAdd.LastName.SetValue Text_Last_Name
    End If
    
    If Text_Middle_Name <> "" Then
        VendorAdd.MiddleName.SetValue Text_Middle_Name
    End If
    
    If Text_Address_1 <> "" Then
        VendorAdd.VendorAddress.Addr1.SetValue Text_Address_1
    End If
    
    If Text_Address_2 <> "" Then
        VendorAdd.VendorAddress.Addr2.SetValue Text_Address_2
    End If
    If Text_City <> "" Then
        VendorAdd.VendorAddress.City.SetValue Text_City
    End If
    If Combo_State_Prov <> "" Then
        VendorAdd.VendorAddress.Province.SetValue Combo_State_Prov
    End If
    
    If Text_Postal_Code <> "" Then
        VendorAdd.VendorAddress.PostalCode.SetValue Text_Postal_Code
    End If
    
    If Text_Phone <> "" Then
        VendorAdd.Phone.SetValue Text_Phone
    End If
    
    If Text_Alt_Phone <> "" Then
        VendorAdd.AltPhone.SetValue Text_Alt_Phone
    End If
    
    If Text_EMail <> "" Then
        VendorAdd.Email.SetValue Text_EMail
    End If
   
    If Text_Name_On_Check <> "" Then
        VendorAdd.NameOnCheck.SetValue Text_Name_On_Check
    End If
    
    If Text_Account_Number <> "" Then
        VendorAdd.AccountNumber.SetValue Text_Account_Number
    End If
    
    If Combo_Vendor_Type <> "" Then
        VendorAdd.VendorTypeRef.FullName.SetValue Combo_Vendor_Type
    End If
    
    If Combo_Currency <> "" Then
        VendorAdd.CurrencyRef.FullName.SetValue Combo_Currency
    End If
    
    'Perform the request
    Dim responseMsgSet As IMsgSetResponse
    Set responseMsgSet = sessionManagerCA.DoRequests(requestMsgSet)
    
    'Interpret the response
    Dim response As IResponse
   
    'The response list contains only one response,
    'which corresponds to our single request
    Set response = responseMsgSet.ResponseList.GetAt(0)
    
    If (response.statusCode <> 0) Then
        MsgBox "Status: Code = " & CStr(response.statusCode) & _
          ", Severity = " & response.statusSeverity & _
          ", Message = " & response.statusMessage
    End If
        
    'The response detail for Add and Mod requests is a 'Ret' object
    'In our case, it's IVendorRet
    Dim VendorRet As IVendorRet
    Set VendorRet = response.Detail
    If (Not (VendorRet Is Nothing)) Then
        MsgBox "A new vendor has been created with ListID = " & VendorRet.ListID.GetValue, , "Vendor added"
    End If
   'Clear the form
    Command_Clear_Click
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"


End Sub


