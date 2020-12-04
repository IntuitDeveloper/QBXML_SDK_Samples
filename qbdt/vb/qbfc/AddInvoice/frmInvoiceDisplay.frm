VERSION 5.00
Begin VB.Form frmInvoiceDisplay 
   Caption         =   "Invoice Display"
   ClientHeight    =   5805
   ClientLeft      =   4380
   ClientTop       =   3300
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   5820
   Begin VB.TextBox txtInvoiceTotal 
      Height          =   285
      Left            =   1440
      TabIndex        =   15
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox txtInvoiceLines 
      Height          =   1815
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   3120
      Width           =   5535
   End
   Begin VB.TextBox txtSalesTax 
      Height          =   285
      Left            =   1440
      TabIndex        =   11
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox txtSubtotal 
      Height          =   285
      Left            =   1440
      TabIndex        =   10
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox txtDueDate 
      Height          =   285
      Left            =   1440
      TabIndex        =   9
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox txtInvoiceDate 
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox txtInvoiceNumber 
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton cmdCloseWindow 
      Caption         =   "Close Window"
      Height          =   615
      Left            =   2040
      TabIndex        =   6
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "Invoice Total"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Invoice Lines"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Invoice Successfully Added"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label Label5 
      Caption         =   "Sales Tax"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Subtotal"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Due Date"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Invoice Date"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Invoice Number"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmInvoiceDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------
' Form: frmInvoiceDisplay
'
' Description:  This form is used to display the invoice
'               information returned by the add invoice request
'               made from frmAddInvoice.
'
' Created On: 09/10/2002
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'----------------------------------------------------------

Option Explicit

Private Sub cmdCloseWindow_Click()
  Unload frmInvoiceDisplay
End Sub
