VERSION 5.00
Begin VB.Form Macro 
   Caption         =   "UseMacro and DefMacro Sample Program"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6945
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   6945
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Press this button to send the InvoiceAdd and ReceivePayment in separate requests."
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   5400
      Width           =   6375
      Begin VB.CommandButton SendTwoRequests 
         Caption         =   "Send Two Requests"
         Height          =   495
         Index           =   0
         Left            =   2280
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.CommandButton SendOneRequest 
      Caption         =   "Send One Request"
      Height          =   495
      Index           =   1
      Left            =   2400
      TabIndex        =   5
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Press this button to send the InvoiceAdd and ReceivePayment in one request."
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   4320
      Width           =   6375
   End
   Begin VB.ListBox estimatesList 
      Height          =   2010
      Left            =   360
      TabIndex        =   1
      Top             =   1680
      Width           =   4815
   End
   Begin VB.CommandButton CloseWindow 
      Cancel          =   -1  'True
      Caption         =   "Close Window"
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   $"frmUseMacroDefMacro.frx":0000
      Height          =   975
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   6375
   End
   Begin VB.Label Label2 
      Caption         =   "An InvoiceAdd and ReceivePayment for that invoice will be sent to QuickBooks."
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   3840
      Width           =   6255
   End
   Begin VB.Label Label1 
      Caption         =   "Select an Estimate below.  The information from the estimate will be used in the InvoiceAdd."
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   8175
   End
End
Attribute VB_Name = "Macro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------
' Form: Macro
'
' Description: This sample program will show how to use the
'              "UseMacro" and "DefMacro".  The "Send One
'              Request" Button will build the InvoiceAdd
'              and ReceivePayment transactions in one
'              request.  The ReceivePayment uses the
'              UseMacro to refer to the InvoiceAdd
'              transaction.  The "Send Two Requests" button
'              will show that the DefMacro value is
'              persistent over separate requests.
'              Information from an Estimate transaction
'              is used to build the InvoiceAdd request.
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'----------------------------------------------------------

Private Sub CloseWindow_Click()
    EndSessionCloseConnection
    End
End Sub

Private Sub Form_Load()
    booSessionBegun = False
    booConnectionOpened = False
    OpenConnectionBeginSession
    GetEstimates estimatesList
End Sub

Private Sub SendOneRequest_Click(index As Integer)
    Dim i As Integer
    i = estimatesList.ListIndex
    If (i <> -1) Then
        SendOneQBRequest i
    Else
        MsgBox "An Estimate must be selected in the List Box."
    End If
End Sub

Private Sub SendTwoRequests_Click(index As Integer)
    Dim i As Integer
    i = estimatesList.ListIndex
    If (i <> -1) Then
        SendTwoQBRequests i
    Else
        MsgBox "An Estimate must be selected in the List Box."
    End If
End Sub
