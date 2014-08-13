VERSION 5.00
Begin VB.Form frmShowRequest 
   Caption         =   "Form1"
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   ScaleHeight     =   8850
   ScaleWidth      =   8190
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   615
      Left            =   3240
      TabIndex        =   1
      Top             =   8160
      Width           =   1695
   End
   Begin VB.TextBox txtRequest 
      Height          =   7935
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   7935
   End
End
Attribute VB_Name = "frmShowRequest"
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

Private Sub Form_Load()
  txtRequest.Text = RequestInfo
End Sub


Private Sub cmdDone_Click()
  Unload frmShowRequest
End Sub


