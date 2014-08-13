VERSION 5.00
Begin VB.Form frmPurchaseOrderModifyMain 
   Caption         =   "Purchase Order Modify"
   ClientHeight    =   5205
   ClientLeft      =   5340
   ClientTop       =   3690
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   4695
   Begin VB.TextBox txtToDate 
      Height          =   285
      Left            =   2160
      TabIndex        =   9
      Text            =   "2038-01-18"
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox txtFromDate 
      Height          =   285
      Left            =   600
      TabIndex        =   7
      Text            =   "1970-01-01"
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdQBXMLRP 
      Caption         =   "Use qbXML Request Processor"
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CommandButton cmdQBFC 
      Caption         =   "Use QBFC"
      Height          =   975
      Left            =   2400
      TabIndex        =   1
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   3120
      TabIndex        =   0
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "To"
      Height          =   255
      Left            =   1800
      TabIndex        =   8
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "From"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Open Purchase Order Query Date(s)"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "This sample program demonstrates the use of the SDK version 2.1 purchase order modify functionality."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   $"frmPurchaseOrderModifyMain.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   4455
   End
End
Attribute VB_Name = "frmPurchaseOrderModifyMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------
' Copyright © 2003-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'----------------------------------------------------------

Option Explicit

Private Sub cmdQBFC_Click()
  If Not DatesOK Then
    Exit Sub
  End If
  
  SetDates CDate(txtFromDate.Text), CDate(txtToDate.Text)
  
  SetImplementation "QBFC"
  If QuickBooksVersionOK Then
    Load frmPurchaseOrderSelect
    frmPurchaseOrderSelect.Show
    frmPurchaseOrderModifyMain.Hide
  End If
End Sub

Private Sub cmdQBXMLRP_Click()
  If Not DatesOK Then
    Exit Sub
  End If
  
  SetDates CDate(txtFromDate.Text), CDate(txtToDate.Text)
  
  SetImplementation "QBXMLRP"
  If QuickBooksVersionOK Then
    Load frmPurchaseOrderSelect
    frmPurchaseOrderSelect.Show
    frmPurchaseOrderModifyMain.Hide
  End If
End Sub

Private Sub cmdQuit_Click()
  End
End Sub

Private Function DatesOK() As Boolean
  If Not IsDate(txtFromDate.Text) Then
    MsgBox "The From date you've supplied is illegal"
    DatesOK = False
    Exit Function
  End If

  If Not IsDate(txtToDate.Text) Then
    MsgBox "The To date you've supplied is illegal"
    DatesOK = False
    Exit Function
  End If
  
  If CDate(txtFromDate.Text) > CDate(txtToDate.Text) Then
    MsgBox "The To date you've supplied is earlier than the From date"
    DatesOK = False
    Exit Function
  End If
  
  If CDate(txtFromDate.Text) < CDate("1/1/1970") Or _
     CDate(txtFromDate.Text) > CDate("1/18/2038") Or _
     CDate(txtToDate.Text) > CDate("1/18/2038") Then
    MsgBox "The dates you supply must be between Jan 1, 1970 and Jan 18, 2038"
    DatesOK = False
    Exit Function
  End If
  
  DatesOK = True
End Function
