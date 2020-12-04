VERSION 5.00
Begin VB.Form frmDepositAdd 
   Caption         =   "Deposit Add"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   7515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   855
      Left            =   4200
      TabIndex        =   2
      Top             =   5400
      Width           =   2295
   End
   Begin VB.CommandButton cmdDepositFunds 
      Caption         =   "Deposit Selected Funds"
      Height          =   855
      Left            =   960
      TabIndex        =   1
      Top             =   5400
      Width           =   2295
   End
   Begin VB.ListBox lstFundsForDeposit 
      Height          =   4935
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6975
   End
End
Attribute VB_Name = "frmDepositAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------
' Form Module: frmDepositAdd
'
' Description: this form is to display payments available for deposit
'              and allow the user to deposit them.
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'----------------------------------------------------------
Option Explicit

Dim booFundsSelected As Boolean

Private Sub cmdDepositFunds_Click()
  If Not booFundsSelected Then
    MsgBox "You must select funds to deposit before attempting to deposit them."
    Exit Sub
  End If
  
  DepositFunds lstFundsForDeposit.List(lstFundsForDeposit.ListIndex)
  GetFundsForDeposit lstFundsForDeposit
End Sub

Private Sub cmdExit_Click()
  EndSessionCloseConnection
  End
End Sub

Private Sub Form_Load()
  booFundsSelected = False
  Connect
  GetFundsForDeposit lstFundsForDeposit
End Sub

Private Sub lstFundsForDeposit_Click()
  booFundsSelected = True
End Sub
