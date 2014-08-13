VERSION 5.00
Begin VB.Form frmAddDataExtension 
   Caption         =   "Add Data Extension To Customer"
   ClientHeight    =   7830
   ClientLeft      =   4140
   ClientTop       =   3030
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   5955
   Begin VB.CommandButton cmdShowResponse 
      Caption         =   "Show Add Response"
      Enabled         =   0   'False
      Height          =   615
      Left            =   2040
      TabIndex        =   12
      Top             =   7080
      Width           =   1695
   End
   Begin VB.CommandButton cmdShowRequest 
      Caption         =   "Show Add Request"
      Enabled         =   0   'False
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   7080
      Width           =   1695
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
      TabIndex        =   10
      Text            =   "frmAddDataExtension.frx":0000
      Top             =   120
      Width           =   5655
   End
   Begin VB.ListBox lstUsedDataExts 
      Height          =   255
      ItemData        =   "frmAddDataExtension.frx":00DD
      Left            =   4920
      List            =   "frmAddDataExtension.frx":00E4
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   5520
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox lstAvailableDataExts 
      Height          =   255
      ItemData        =   "frmAddDataExtension.frx":00F9
      Left            =   4680
      List            =   "frmAddDataExtension.frx":00FB
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   5520
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.CommandButton cmdCloseWindow 
      Caption         =   "Close Window"
      Height          =   615
      Left            =   3960
      TabIndex        =   7
      Top             =   7080
      Width           =   1815
   End
   Begin VB.CommandButton cmdAddValue 
      Caption         =   "Add Data Extension Value"
      Height          =   615
      Left            =   1800
      TabIndex        =   6
      Top             =   6120
      Width           =   2175
   End
   Begin VB.TextBox txtDataExtValue 
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Top             =   5520
      Width           =   2655
   End
   Begin VB.ListBox lstUnusedDataExtDefs 
      Height          =   1230
      ItemData        =   "frmAddDataExtension.frx":00FD
      Left            =   120
      List            =   "frmAddDataExtension.frx":00FF
      TabIndex        =   2
      Top             =   4080
      Width           =   5655
   End
   Begin VB.ListBox lstCustomers 
      Height          =   2010
      ItemData        =   "frmAddDataExtension.frx":0101
      Left            =   120
      List            =   "frmAddDataExtension.frx":0103
      TabIndex        =   0
      Top             =   1560
      Width           =   5655
   End
   Begin VB.Label Label3 
      Caption         =   "Data Extension Value"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Unused Data Extension Definitions"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Customers"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   855
   End
End
Attribute VB_Name = "frmAddDataExtension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------
' Form: frmAddDataExtension
'
' Description: Allows the user to select a select a customer causing
'              the form to list the unused data extensions for that
'              customer.  Then the user may highlight one of the
'              unused data extensions, provide a value for it and
'              add it to the customer record
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'----------------------------------------------------------
Private Sub cmdAddValue_Click()

  If lstCustomers.ListIndex < 0 Then
    MsgBox "You must select a customer first"
    Exit Sub
  End If
  
  If lstUnusedDataExtDefs.ListIndex < 0 Then
    MsgBox "You must select a Data Extension to add first"
    Exit Sub
  End If
  
  If txtDataExtValue.Text = "" Then
    MsgBox "You must supply a value for the Data Extension first"
    Exit Sub
  End If
  
  If AddDataExt(lstCustomers.List(lstCustomers.ListIndex), _
                Left(lstUnusedDataExtDefs.List(lstUnusedDataExtDefs.ListIndex), _
                     InStr(lstUnusedDataExtDefs.List(lstUnusedDataExtDefs.ListIndex), _
                           "  |") - 1), _
                txtDataExtValue.Text) Then
    lstUnusedDataExtDefs.RemoveItem (lstUnusedDataExtDefs.ListIndex)
    txtDataExtValue.Text = ""
    txtDataExtValue.Refresh
  End If
  
  cmdShowRequest.Enabled = True
  cmdShowResponse.Enabled = True
End Sub

Private Sub cmdCloseWindow_Click()
  Unload frmAddDataExtension
End Sub

Private Sub cmdShowRequest_Click()
  frmqbXMLDisplay.txtqbXML = PrettyXMLString(modDataExtSample.strAddRequest)
  Load frmqbXMLDisplay
  frmqbXMLDisplay.Show
End Sub

Private Sub cmdShowResponse_Click()
  frmqbXMLDisplay.txtqbXML = PrettyXMLString(modDataExtSample.strAddResponse)
  Load frmqbXMLDisplay
  frmqbXMLDisplay.Show
End Sub

Private Sub Form_Load()
  GetCustomers lstCustomers
End Sub

Private Sub lstCustomers_Click()
  lstUnusedDataExtDefs.Clear
  FillUnusedDataExts lstCustomers.List(lstCustomers.ListIndex), lstUnusedDataExtDefs
End Sub
