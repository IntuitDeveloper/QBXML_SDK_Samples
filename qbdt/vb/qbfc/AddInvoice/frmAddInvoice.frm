VERSION 5.00
Begin VB.Form frmAddInvoice 
   Caption         =   "Add Invoice Sample"
   ClientHeight    =   4410
   ClientLeft      =   4230
   ClientTop       =   3450
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   6090
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   855
      Left            =   4080
      TabIndex        =   5
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton cmdAddInvoice 
      Caption         =   "Add Invoice"
      Height          =   855
      Left            =   1920
      TabIndex        =   4
      Top             =   3240
      Width           =   1695
   End
   Begin VB.OptionButton optQBFC 
      Caption         =   "Use QBFC"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   3720
      Width           =   1455
   End
   Begin VB.OptionButton optqbXMLRP 
      Caption         =   "Use qbXMLRP"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   3240
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "You may choose to have the program use either qbXML and the MSXML4 DOM parser to build the message or QBFC to build the message."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   5895
   End
   Begin VB.Label Label1 
      Caption         =   $"frmAddInvoice.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   5775
   End
End
Attribute VB_Name = "frmAddInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------
' Form: frmAddInvoice
'
' Description:  This sample demonstrates the use of qbXML 2.0 or
'               QBFC, by adding an invoice within.
'               It includes examples of the following:
'                   - Constructing an invoice add request including
'                     - Using a different shipping address
'                     - Adding multiple invoice lines including two
'                       item lines and one item group line
'
'               All non form code is found in modAddInvoice
'
' Created On: 09/10/2002
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'----------------------------------------------------------

Option Explicit

Private Sub cmdAddInvoice_Click()
  If optqbXMLRP.Value Then
    qbXML_AddInvoice
  Else
    QBFC_AddInvoice
  End If
End Sub

Private Sub cmdExit_Click()
  End
End Sub
