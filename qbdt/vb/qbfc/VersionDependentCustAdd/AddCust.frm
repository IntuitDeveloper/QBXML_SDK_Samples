VERSION 5.00
Begin VB.Form AddCust 
   Caption         =   "QBFC Sample: Create Requests dependent on the qbxml version"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6570
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   6570
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox UseQBOE 
      Caption         =   "Communicate with QuickBooks Online Edition"
      Height          =   375
      Left            =   240
      TabIndex        =   35
      Top             =   120
      Width           =   3615
   End
   Begin VB.Frame Frame3 
      Caption         =   "Data Extension Information (optional)"
      Height          =   1095
      Left            =   240
      TabIndex        =   26
      Top             =   7200
      Width           =   6135
      Begin VB.TextBox Text_DataExtValue 
         Height          =   285
         Left            =   2160
         TabIndex        =   30
         Top             =   720
         Width           =   3735
      End
      Begin VB.TextBox Text_DataExtName 
         Height          =   285
         Left            =   2160
         TabIndex        =   28
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Data Extension Val&ue:"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Data E&xtension Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.CommandButton Comm_Exit 
      Caption         =   "&Done"
      Height          =   372
      Left            =   4200
      TabIndex        =   34
      Top             =   600
      Width           =   2172
   End
   Begin VB.CommandButton Comm_View_Res 
      Caption         =   "Vi&ew qbXML Response"
      Height          =   372
      Left            =   4200
      TabIndex        =   33
      Top             =   1560
      Width           =   2172
   End
   Begin VB.CommandButton Comm_View_Req 
      Caption         =   "&View qbXML Request"
      Height          =   372
      Left            =   4200
      TabIndex        =   32
      Top             =   1080
      Width           =   2172
   End
   Begin VB.TextBox Text_CustomerName 
      Height          =   288
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   2532
   End
   Begin VB.TextBox Text_Phone 
      Height          =   288
      Left            =   1920
      TabIndex        =   8
      Top             =   3240
      Width           =   3975
   End
   Begin VB.CommandButton Comm_Submit 
      Caption         =   "&Add Customer"
      Default         =   -1  'True
      Height          =   372
      Left            =   4200
      TabIndex        =   31
      Top             =   120
      Width           =   2172
   End
   Begin VB.TextBox Text_LastName 
      Height          =   288
      Left            =   1920
      TabIndex        =   6
      Top             =   2760
      Width           =   3975
   End
   Begin VB.TextBox Text_FirstName 
      Height          =   288
      Left            =   1920
      TabIndex        =   4
      Top             =   2280
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      Caption         =   " Contact Information (optional) "
      Height          =   5175
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   6135
      Begin VB.Frame Frame2 
         Caption         =   "Billing Address"
         Height          =   3375
         Left            =   240
         TabIndex        =   9
         Top             =   1680
         Width           =   5655
         Begin VB.TextBox Text_Country 
            Height          =   285
            Left            =   1440
            TabIndex        =   25
            Top             =   2880
            Width           =   3975
         End
         Begin VB.TextBox Text_PostalCode 
            Height          =   285
            Left            =   1440
            TabIndex        =   23
            Top             =   2520
            Width           =   3975
         End
         Begin VB.TextBox Text_State 
            Height          =   285
            Left            =   1440
            TabIndex        =   21
            Top             =   2160
            Width           =   3975
         End
         Begin VB.TextBox Text_City 
            Height          =   285
            Left            =   1440
            TabIndex        =   19
            Top             =   1800
            Width           =   3975
         End
         Begin VB.TextBox Text_Addr4 
            Height          =   285
            Left            =   1440
            TabIndex        =   17
            Top             =   1440
            Width           =   3975
         End
         Begin VB.TextBox Text_Addr3 
            Height          =   285
            Left            =   1440
            TabIndex        =   15
            Top             =   1080
            Width           =   3975
         End
         Begin VB.TextBox Text_Addr2 
            Height          =   285
            Left            =   1440
            TabIndex        =   13
            Top             =   720
            Width           =   3975
         End
         Begin VB.TextBox Text_Addr1 
            Height          =   285
            Left            =   1440
            TabIndex        =   11
            Top             =   360
            Width           =   3975
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Countr&y:"
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   2880
            Width           =   1095
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "P&ostal Code:"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   2520
            Width           =   1095
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "&State:"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   2160
            Width           =   1095
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Ci&ty:"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   1800
            Width           =   1095
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Address Line &4:"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Address Line &3:"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Address Line &2:"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Address Line &1:"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "&Phone:"
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "&Last Name:"
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&First Name:"
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Label Label7 
      Caption         =   "&Customer to add:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   1200
      Width           =   1335
   End
End
Attribute VB_Name = "AddCust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------
' Form Module: AddCust
'
' Description:  This sample application demonstrates how to
'               construct a qbfc request, send it to QuickBooks,
'               and parse the qbfc response.  It prompts the
'               user for customer information and adds a new
'               customer to QuickBooks.
'
'               QuickBooks must be running with a data file open.
'               The current data file is used.
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/rdmgr/?ID=100
'
'----------------------------------------------------------

Option Explicit

' Request and response strings
Public requestXML  As String
Public responseXML As String

' Customer info
Public firstName   As String
Public lastName    As String
Public phoneNumber As String
Public custName    As String
Public Addr1       As String
Public Addr2       As String
Public Addr3       As String
Public Addr4       As String
Public City        As String
Public State       As String
Public PostalCode  As String
Public Country     As String
Public strDataExtName As String
Public strDataExtValue As String

' Submit button
Private Sub Comm_Submit_Click()

On Error GoTo ErrHandler

    ' Initialize
    requestXML = ""
    responseXML = ""
    OpenConnectionBeginSession
    ' Get input data
    If Not CollectFormData Then
        Exit Sub
    End If
    
    ' Send request to QuickBooks
    SendCustomerAddRequest
    EndSessionCloseConnection
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
    Exit Sub
    
End Sub


' View qbXML Request button
Private Sub Comm_View_Req_Click()

On Error GoTo ErrHandler
    
    If requestXML <> "" Then
        
        Dim reqFrm As DisplayXML
        Set reqFrm = New DisplayXML
        
        reqFrm.Text_Content = requestXML
        reqFrm.Caption = "Request XML"
        reqFrm.Show
        
    Else
        MsgBox "Request is empty.  Please add a customer first", vbInformation, "Request XML"
    End If
    Exit Sub
        

ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
    Exit Sub

End Sub

' View Response Button
'
Private Sub Comm_View_Res_Click()
 
 On Error GoTo ErrHandler
 
    ' Instantiate a RequestXML Obj
    '
        
    If responseXML <> "" Then
        
        Dim resFrm As New DisplayXML
        Dim tmpResponseXML As String
        
        ' replace Lf to CrLf, this is for display only
        tmpResponseXML = Replace(responseXML, vbLf, vbCrLf)
        resFrm.Text_Content = tmpResponseXML
        resFrm.Caption = "Response XML"
        resFrm.Show
        
    Else
        MsgBox "Response is empty.  Please add a customer first", vbInformation, "Response XML"
    End If
    
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
    Exit Sub

End Sub

' Exit button
'
Private Sub Comm_Exit_Click()
    Unload Me ' close the window
End Sub

' Form Data Collection
'
Private Function CollectFormData() As Boolean

On Error GoTo ErrHandler

    ' Get data from the form
    firstName = Text_FirstName.Text
    lastName = Text_LastName.Text
    phoneNumber = Text_Phone.Text
    custName = Text_CustomerName.Text
    Addr1 = Text_Addr1.Text
    Addr2 = Text_Addr2.Text
    Addr3 = Text_Addr3.Text
    Addr4 = Text_Addr4.Text
    City = Text_City.Text
    State = Text_State.Text
    PostalCode = Text_PostalCode.Text
    Country = Text_Country.Text
    strDataExtName = Text_DataExtName.Text
    strDataExtValue = Text_DataExtValue.Text
    
    ' Customer Name is required
    If custName = "" Then
        MsgBox "Customer Name is empty", vbOKOnly, "Error"
        CollectFormData = False
        GoTo ExitProc
    End If
    
    CollectFormData = True
    
ExitProc:
    Exit Function
    
ErrHandler:
    CollectFormData = False
    MsgBox Err.Description, vbExclamation, "Error"
    Exit Function
    
End Function





