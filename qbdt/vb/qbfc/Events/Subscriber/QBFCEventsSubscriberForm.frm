VERSION 5.00
Begin VB.Form QBFCEventsSubscriber 
   Caption         =   "QBFC Events Subscriber"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Subscribe 
      Caption         =   "Subscribe"
      Height          =   495
      Left            =   960
      TabIndex        =   8
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CommandButton Unsubscribe 
      Caption         =   "Unsubscribe"
      Height          =   495
      Left            =   2760
      TabIndex        =   7
      Top             =   6960
      Width           =   1335
   End
   Begin VB.TextBox ErrorMsg 
      Height          =   1455
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   5280
      Width           =   5175
   End
   Begin VB.TextBox ResponseXML 
      Height          =   1455
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   3240
      Width           =   5175
   End
   Begin VB.TextBox RequestXML 
      Height          =   1455
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   1200
      Width           =   5175
   End
   Begin VB.Label Errors 
      Caption         =   "Errors"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label Response 
      Caption         =   "Response"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Request 
      Caption         =   "Request"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "This sample shows how to subscribe and unsubscribe to Data Event:CustomerAdd and UI Event:Company File Close Events."
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "QBFCEventsSubscriber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Subscribe_Click()
    
    DoSubscribeEvents
        
End Sub

Private Sub Unsubscribe_Click()
    
    DoUnSubscribeEvents

End Sub





