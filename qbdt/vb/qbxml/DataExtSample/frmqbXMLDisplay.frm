VERSION 5.00
Begin VB.Form frmqbXMLDisplay 
   Caption         =   "qbXML Display"
   ClientHeight    =   6870
   ClientLeft      =   4380
   ClientTop       =   2115
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   6630
   Begin VB.CommandButton cmdCloseWindow 
      Caption         =   "Close Window"
      Height          =   735
      Left            =   2280
      TabIndex        =   1
      Top             =   6120
      Width           =   1935
   End
   Begin VB.TextBox txtqbXML 
      Height          =   5895
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "frmqbXMLDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCloseWindow_Click()
  Unload frmqbXMLDisplay
End Sub
