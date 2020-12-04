VERSION 5.00
Begin VB.Form Display 
   Caption         =   "Invoice list"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   8310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Comm_Exit 
      Caption         =   "&Done"
      Height          =   372
      Left            =   3420
      TabIndex        =   1
      Top             =   3525
      Width           =   1452
   End
   Begin VB.TextBox Text_Content 
      Height          =   3132
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   8055
   End
End
Attribute VB_Name = "Display"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------
' Form Module: Display
'
' Description:  Display invoice list
'
' Last Updated: 08/2002
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'----------------------------------------------------------

Option Explicit

Private Sub Comm_Exit_Click()
    Unload Me
End Sub
