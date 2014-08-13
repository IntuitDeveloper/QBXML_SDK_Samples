VERSION 5.00
Begin VB.Form frmInvoiceModMain 
   Caption         =   "Invoice Modify Sample "
   ClientHeight    =   3630
   ClientLeft      =   5730
   ClientTop       =   4035
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   ScaleHeight     =   3630
   ScaleWidth      =   4710
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   3120
      TabIndex        =   4
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdQBFC 
      Caption         =   "Use QBFC"
      Height          =   975
      Left            =   2400
      TabIndex        =   3
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton cmdQBXMLRP 
      Caption         =   "Use qbXML Request Processor"
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   $"frmInvoiceModMain.frx":0000
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "This sample program demonstrates the use of the SDK version 2.1 invoice modify functionality."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmInvoiceModMain"
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

Private Sub cmdQBFC_Click()
  SetImplementation "QBFC"
  If QuickBooksVersionOK Then
    Load frmInvoiceQuery
    frmInvoiceQuery.Show
    frmInvoiceModMain.Hide
  End If
End Sub

Private Sub cmdQBXMLRP_Click()
  SetImplementation "QBXMLRP"
  If QuickBooksVersionOK Then
    Load frmInvoiceQuery
    frmInvoiceQuery.Show
    frmInvoiceModMain.Hide
  End If
End Sub

Private Sub cmdQuit_Click()
  End
End Sub

Private Sub Form_Load()
  Load frmPatience
  Load frmModifying
End Sub
