VERSION 5.00
Begin VB.Form frmModDataExtension 
   Caption         =   "Modify Data Extension Value for Customer"
   ClientHeight    =   7785
   ClientLeft      =   4380
   ClientTop       =   1815
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   5970
   Begin VB.CommandButton cmdModResponse 
      Caption         =   "Show Modify Response"
      Enabled         =   0   'False
      Height          =   615
      Left            =   2040
      TabIndex        =   10
      Top             =   7080
      Width           =   1815
   End
   Begin VB.CommandButton cmdModRequest 
      Caption         =   "Show Modify Request"
      Enabled         =   0   'False
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   7080
      Width           =   1815
   End
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
      Height          =   1095
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "frmModDataExtension.frx":0000
      Top             =   120
      Width           =   5655
   End
   Begin VB.ListBox lstCustomers 
      Height          =   2010
      ItemData        =   "frmModDataExtension.frx":00F5
      Left            =   120
      List            =   "frmModDataExtension.frx":00F7
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1560
      Width           =   5655
   End
   Begin VB.ListBox lstUsedDataExts 
      Height          =   1230
      ItemData        =   "frmModDataExtension.frx":00F9
      Left            =   120
      List            =   "frmModDataExtension.frx":00FB
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4080
      Width           =   5655
   End
   Begin VB.TextBox txtDataExtValue 
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      Top             =   5520
      Width           =   2655
   End
   Begin VB.CommandButton cmdModValue 
      Caption         =   "Modify Data Extension Value"
      Height          =   615
      Left            =   1800
      TabIndex        =   1
      Top             =   6120
      Width           =   2415
   End
   Begin VB.CommandButton cmdCloseWindow 
      Caption         =   "Close Window"
      Height          =   615
      Left            =   3960
      TabIndex        =   0
      Top             =   7080
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Customers"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Data Extension Values"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3840
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "New Data Extension Value"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   5520
      Width           =   1935
   End
End
Attribute VB_Name = "frmModDataExtension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------
' Form: frmModDataExtension
'
' Description: Allows the user to highlight a customer which brings
'              up a list of used data extensions for that customer.
'              The user may then highlight a data extension and modify
'              the value given for that customer.
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'----------------------------------------------------------
Private Sub cmdCloseWindow_Click()
  Unload frmModDataExtension
End Sub

Private Sub cmdModRequest_Click()
  frmqbXMLDisplay.txtqbXML = PrettyXMLString(modDataExtSample.strModRequest)
  Load frmqbXMLDisplay
  frmqbXMLDisplay.Show
End Sub

Private Sub cmdModResponse_Click()
  frmqbXMLDisplay.txtqbXML = PrettyXMLString(modDataExtSample.strModResponse)
  Load frmqbXMLDisplay
  frmqbXMLDisplay.Show
End Sub

Private Sub cmdModValue_Click()
  If lstUsedDataExts.ListIndex < 0 Then
    MsgBox "You must select a data extension to modify"
    Exit Sub
  End If
  
  If txtDataExtValue.Text = "" Then
    MsgBox "You must select a data extension to modify"
    Exit Sub
  End If
  
  If ModDataExt(lstCustomers.List(lstCustomers.ListIndex), _
                Left(lstUsedDataExts.List(lstUsedDataExts.ListIndex), _
                     InStr(lstUsedDataExts.List(lstUsedDataExts.ListIndex), "  ") - 1), _
                txtDataExtValue.Text) Then

    GetUsedCustomerDataExts _
      lstCustomers.List(lstCustomers.ListIndex), _
      lstUsedDataExts, _
      True
  End If
  
  cmdModRequest.Enabled = True
  cmdModResponse.Enabled = True
End Sub

Private Sub Form_Load()
  GetCustomers lstCustomers
End Sub

Private Sub lstCustomers_Click()
  lstUsedDataExts.Clear
  lstUsedDataExts.Refresh

  GetUsedCustomerDataExts _
    lstCustomers.List(lstCustomers.ListIndex), _
    lstUsedDataExts, _
    True

  If lstUsedDataExts.ListCount = 0 Then
    lstUsedDataExts.AddItem "This customer has no data extensions added to their record"
  End If
End Sub
