VERSION 5.00
Begin VB.Form frmUS_CDN_CustomerAdd 
   Caption         =   "US / Canadian CustomerAdd"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   5550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   3720
      TabIndex        =   20
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton cmdClearForm 
      Caption         =   "Clear Form"
      Height          =   495
      Left            =   1920
      TabIndex        =   19
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton cmdAddCustomer 
      Caption         =   "Add Customer"
      Height          =   495
      Left            =   120
      TabIndex        =   18
      Top             =   4440
      Width           =   1695
   End
   Begin VB.TextBox txtPostalCode 
      Height          =   285
      Left            =   1320
      TabIndex        =   17
      Top             =   3840
      Width           =   2535
   End
   Begin VB.TextBox txtStateProvince 
      Height          =   285
      Left            =   1320
      TabIndex        =   16
      Top             =   3480
      Width           =   2535
   End
   Begin VB.TextBox txtCity 
      Height          =   285
      Left            =   1320
      TabIndex        =   15
      Top             =   3120
      Width           =   2535
   End
   Begin VB.TextBox txtAddr4 
      Height          =   285
      Left            =   1320
      TabIndex        =   14
      Top             =   2760
      Width           =   2535
   End
   Begin VB.TextBox txtAddr3 
      Height          =   285
      Left            =   1320
      TabIndex        =   13
      Top             =   2400
      Width           =   2535
   End
   Begin VB.TextBox txtAddr2 
      Height          =   285
      Left            =   1320
      TabIndex        =   12
      Top             =   2040
      Width           =   2535
   End
   Begin VB.TextBox txtAddr1 
      Height          =   285
      Left            =   1320
      TabIndex        =   11
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox txtCustomer 
      Height          =   285
      Left            =   1440
      TabIndex        =   10
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label Label8 
      Caption         =   "Postal Code"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Billing Address"
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
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblStateProvince 
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "City"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label lblAddr4 
      Caption         =   "Addr4"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Addr3"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Addr2"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Addr1"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Customer Name"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblNationality 
      Alignment       =   2  'Center
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
      TabIndex        =   0
      Top             =   240
      Width           =   5295
   End
End
Attribute VB_Name = "frmUS_CDN_CustomerAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------
' Form: frmUS_CDN_CustomerAdd
'
' Description: Displays the nationality of the QuickBooks version
'              connected to and allows the input of information for
'              adding a customer to the currently open company file
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'----------------------------------------------------------
Dim QBNationality As String
  
Private Sub cmdAddCustomer_Click()
'Check to make sure that the Addr fields are filled in properly
  If txtAddr1.Text = "" Then
    If txtAddr2.Text <> "" Then
      MsgBox "The address fields must be filled in order.  Addr2 must be blank if Addr1 isn't filled"
      Exit Sub
    Else
      If txtAddr3 <> "" Then
        MsgBox "The address fields must be filled in order.  Addr3 must be blank if Addr1 isn't filled"
        Exit Sub
      Else
        If txtAddr4.Text <> "" Then
          MsgBox "The address fields must be filled in order.  Addr4 must be blank if Addr1 isn't filled"
          Exit Sub
        End If
      End If
    End If
  Else
    If txtAddr2.Text = "" Then
      If txtAddr3 <> "" Then
        MsgBox "The address fields must be filled in order.  Addr3 must be blank if Addr2 isn't filled"
        Exit Sub
      Else
        If txtAddr4.Text <> "" Then
          MsgBox "The address fields must be filled in order.  Addr4 must be blank if Addr2 isn't filled"
          Exit Sub
        End If
      End If
    Else
      If txtAddr3 = "" And txtAddr4.Text <> "" Then
        MsgBox "The address fields must be filled in order.  Addr4 must be blank if Addr3 isn't filled"
        Exit Sub
      End If
    End If
  End If
  
  If AddCustomer(QBNationality, txtCustomer.Text, txtAddr1.Text, _
                 txtAddr2.Text, txtAddr3.Text, txtAddr4.Text, _
                 txtCity.Text, txtStateProvince.Text, _
                 txtPostalCode.Text) _
  Then
    ClearForm
  End If
End Sub

Private Sub cmdClearForm_Click()
  ClearForm
End Sub

Private Sub cmdExit_Click()
  EndQuickBooksSession
  End
End Sub

Private Sub Form_Load()
  StartQuickBooksSession
  QBNationality = GetQBNationality
  
  lblNationality.Caption = _
    "Running against " & QBNationality & " QuickBooks"

  If QBNationality = "US" Then
    lblStateProvince.Caption = "State"
  Else
    lblStateProvince.Caption = "Province"
  End If
End Sub

Private Sub ClearForm()
  txtCustomer.Text = ""
  txtAddr1.Text = ""
  txtAddr2.Text = ""
  txtAddr3.Text = ""
  txtAddr4.Text = ""
  txtCity.Text = ""
  txtStateProvince.Text = ""
  txtPostalCode.Text = ""
End Sub
