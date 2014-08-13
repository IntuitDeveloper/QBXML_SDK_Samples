VERSION 5.00
Begin VB.Form DisplayXML 
   Caption         =   "Request XML"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   6105
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Comm_Exit 
      Caption         =   "&Done"
      Height          =   372
      Left            =   2160
      TabIndex        =   1
      Top             =   3480
      Width           =   1452
   End
   Begin VB.TextBox Text_Content 
      Height          =   3132
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   5772
   End
End
Attribute VB_Name = "DisplayXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------
' Form Module: DisplayFrm
'
' Description: this form is to display qbXML request/response.
'
' Created On: 09/15/2001
' Updated to SDK 2.0 On: 07/30/2002
'
' Unpublished work. INTUIT CONFIDENTIAL.
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/rdmgr/?ID=100
'
'----------------------------------------------------------

Option Explicit

Private Sub Comm_Exit_Click()
    Unload Me
End Sub
