VERSION 5.00
Begin VB.Form ShowItems 
   Caption         =   "Items List"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Done 
      Caption         =   "Done"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox ItemsList 
      Height          =   3015
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   720
      Width           =   6375
   End
   Begin VB.Label Label1 
      Caption         =   "Query result for the first 30 items in the current company file."
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   5895
   End
End
Attribute VB_Name = "ShowItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------
' ShowItems.frm
'
' This form is simply used to display items which are
' returned from QuickBooks.  All interaction with QB
' takes place in the Main Module.
'
' Last Updated: 08/2002
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'----------------------------------------------------------

Private Sub Done_Click()
    Unload Me
End Sub
