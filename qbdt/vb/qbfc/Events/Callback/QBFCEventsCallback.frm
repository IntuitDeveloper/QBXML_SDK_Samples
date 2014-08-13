VERSION 5.00
Begin VB.Form QBFCEventsCallbackForm 
   Caption         =   "QBFCEventsCallback"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9210
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   9210
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Query 
      Caption         =   "Query"
      Height          =   615
      Left            =   1080
      TabIndex        =   8
      Top             =   4680
      Width           =   1575
   End
   Begin VB.TextBox QueryResult 
      Height          =   1455
      Left            =   3120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   4200
      Width           =   4935
   End
   Begin VB.TextBox ErrorMsg 
      Height          =   1095
      Left            =   1440
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Text            =   "QBFCEventsCallback.frx":0000
      Top             =   2520
      Width           =   5415
   End
   Begin VB.TextBox UIEvent 
      Height          =   1095
      Left            =   4680
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   1080
      Width           =   3975
   End
   Begin VB.TextBox DataEvent 
      Height          =   1095
      Left            =   600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Customer Query"
      Height          =   1815
      Left            =   600
      TabIndex        =   9
      Top             =   3960
      Width           =   8055
   End
   Begin VB.Label Label4 
      Caption         =   "Error Msg"
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Company File Close Event"
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Customer Add Event"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "This sample shows how to receive events.  This sample works with the QBFCEventsSubscriber sample."
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "QBFCEventsCallbackForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Query_Click()
    GetAllCustomers
End Sub
