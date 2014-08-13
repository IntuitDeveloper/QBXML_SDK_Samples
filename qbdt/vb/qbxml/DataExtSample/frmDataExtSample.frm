VERSION 5.00
Begin VB.Form frmDataExtSample 
   Caption         =   "Data Extension Sample Main Window"
   ClientHeight    =   8520
   ClientLeft      =   4140
   ClientTop       =   2460
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   6270
   Begin VB.CommandButton cmdShowCustomResponse 
      Caption         =   "Show Custom Field Query Response"
      Height          =   735
      Left            =   4800
      TabIndex        =   13
      Top             =   7680
      Width           =   1335
   End
   Begin VB.CommandButton cmdShowCustomRequest 
      Caption         =   "Show Custom Field Query Request"
      Height          =   735
      Left            =   3240
      TabIndex        =   12
      Top             =   7680
      Width           =   1335
   End
   Begin VB.CommandButton cmdShowQueryResponse 
      Caption         =   "Show Data Extension Query Response"
      Height          =   735
      Left            =   1680
      TabIndex        =   11
      Top             =   7680
      Width           =   1335
   End
   Begin VB.CommandButton cmdShowQueryRequest 
      Caption         =   "Show Data Extension Query Request"
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   7680
      Width           =   1335
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
      Height          =   735
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "frmDataExtSample.frx":0000
      Top             =   480
      Width           =   5775
   End
   Begin VB.TextBox txtUtility 
      Height          =   195
      Left            =   6000
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "frmDataExtSample.frx":00C3
      Top             =   3960
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtDataExts 
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   1320
      Width           =   6015
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   735
      Left            =   4920
      TabIndex        =   6
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdModDataExt 
      Caption         =   "Modify Customer Data Extension"
      Height          =   735
      Left            =   3240
      TabIndex        =   5
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton cmdAddDataExtension 
      Caption         =   "Add Customer Data Extension"
      Height          =   735
      Left            =   1560
      TabIndex        =   4
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton cmdDefineDataExt 
      Caption         =   "Define Data Extension"
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   6600
      Width           =   1215
   End
   Begin VB.TextBox txtCustomFields 
      Enabled         =   0   'False
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   4320
      Width           =   6015
   End
   Begin VB.Label Label2 
      Caption         =   "Defined Custom Fields (read only, to show query for Custom Field Defs)"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   5775
   End
   Begin VB.Label Label1 
      Caption         =   "Data Extensions for OwnerID {E09C86CF-9D6E-4EF2-BCBE-4D66B6B0F754}"
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
      Width           =   5895
   End
End
Attribute VB_Name = "frmDataExtSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------
' Form: frmDataExtSample
'
' Description: This the main form and entry point for this sample
'              program.  It displays the currently defined data
'              extension definitions and custom fields for the open
'              company file.  From here the user may choose to
'              activate forms to add data extension definitions to the
'              company file, define values for data extension
'              for customers or modify data extension values for
'              customers
'
'              The form calls OpenSessionBeginSession to make sure
'              a company file is open.
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'----------------------------------------------------------
Private Sub cmdAddDataExtension_Click()
  If CustomersHaveDataExts Then
    Load frmAddDataExtension
    frmAddDataExtension.Show
  Else
    MsgBox "You must add a data extension for Customers before adding a data extension value to a specified customer"
  End If
End Sub

Private Sub cmdDefineDataExt_Click()
  Load frmAddDataExtDef
  frmAddDataExtDef.Show
End Sub

Private Sub cmdExit_Click()
  EndSessionCloseConnection
  End
End Sub

Private Sub cmdModDataExt_Click()
  If CustomersHaveDataExts Then
    Load frmModDataExtension
    frmModDataExtension.Show
  Else
    MsgBox "You must add a data extension for Customers before modifying a data extension value for a specific customer"
  End If
End Sub

Private Sub cmdShowCustomRequest_Click()
  frmqbXMLDisplay.txtqbXML = PrettyXMLString(modDataExtSample.strCustomRequest)
  Load frmqbXMLDisplay
  frmqbXMLDisplay.Show
End Sub

Private Sub cmdShowCustomResponse_Click()
  frmqbXMLDisplay.txtqbXML = PrettyXMLString(modDataExtSample.strCustomResponse)
  Load frmqbXMLDisplay
  frmqbXMLDisplay.Show
End Sub

Private Sub cmdShowQueryRequest_Click()
  frmqbXMLDisplay.txtqbXML = PrettyXMLString(modDataExtSample.strQueryRequest)
  Load frmqbXMLDisplay
  frmqbXMLDisplay.Show
End Sub

Private Sub cmdShowQueryResponse_Click()
  frmqbXMLDisplay.txtqbXML = PrettyXMLString(modDataExtSample.strQueryResponse)
  Load frmqbXMLDisplay
  frmqbXMLDisplay.Show
End Sub

Private Sub Form_Load()
  OpenConnectionBeginSession
  GetDataExts txtDataExts
  GetCustomFields txtCustomFields
End Sub
