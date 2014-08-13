VERSION 5.00
Begin VB.Form frmPurchaseOrderSelect 
   Caption         =   "Select A Purchase Order for Modify"
   ClientHeight    =   4605
   ClientLeft      =   3195
   ClientTop       =   3510
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   9465
   Begin VB.CommandButton cmdPODetails 
      Caption         =   "View Purchase Order Details"
      Height          =   735
      Left            =   240
      TabIndex        =   77
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox txtSpace2 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   9
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   2640
      Width           =   365
   End
   Begin VB.TextBox txtSpace1 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   9
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox txtSpace2 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   8
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   2400
      Width           =   365
   End
   Begin VB.TextBox txtSpace1 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   8
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   2400
      Width           =   375
   End
   Begin VB.TextBox txtSpace2 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   7
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   2160
      Width           =   365
   End
   Begin VB.TextBox txtSpace1 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   7
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox txtSpace2 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   6
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   1920
      Width           =   365
   End
   Begin VB.TextBox txtSpace1 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   6
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox txtSpace2 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   5
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   1680
      Width           =   365
   End
   Begin VB.TextBox txtSpace1 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   5
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox txtSpace2 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   4
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   1440
      Width           =   365
   End
   Begin VB.TextBox txtSpace1 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   4
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox txtSpace2 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   3
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   1200
      Width           =   365
   End
   Begin VB.TextBox txtSpace1 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   3
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox txtSpace2 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   2
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   960
      Width           =   365
   End
   Begin VB.TextBox txtSpace1 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   2
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox txtSpace2 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   1
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   720
      Width           =   365
   End
   Begin VB.TextBox txtSpace1 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   1
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox txtSpace2 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   0
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   480
      Width           =   365
   End
   Begin VB.TextBox txtSpace1 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   0
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   735
      Left            =   7440
      TabIndex        =   56
      Top             =   3600
      Width           =   1575
   End
   Begin VB.VScrollBar vscPOScroll 
      Height          =   2415
      LargeChange     =   10
      Left            =   8760
      TabIndex        =   50
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox txtPOAmount 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   9
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   49
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox txtPODeliveryDate 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   9
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   48
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox txtPOVendor 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   9
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   47
      Top             =   2640
      Width           =   3375
   End
   Begin VB.TextBox txtPODate 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   9
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   46
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtPONumber 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   9
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   45
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox txtPOAmount 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   8
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox txtPODeliveryDate 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   8
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox txtPOVendor 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   8
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   42
      Top             =   2400
      Width           =   3375
   End
   Begin VB.TextBox txtPODate 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   8
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox txtPONumber 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   8
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   40
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox txtPOAmount 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   7
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox txtPODeliveryDate 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   7
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txtPOVendor 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   7
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   2160
      Width           =   3375
   End
   Begin VB.TextBox txtPODate 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   7
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txtPONumber 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   7
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox txtPOAmount 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   6
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox txtPODeliveryDate 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   6
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txtPOVendor 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   6
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   1920
      Width           =   3375
   End
   Begin VB.TextBox txtPODate 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   6
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox txtPONumber 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   6
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox txtPOAmount 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   5
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox txtPODeliveryDate 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   5
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtPOVendor 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   5
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   1680
      Width           =   3375
   End
   Begin VB.TextBox txtPODate 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   5
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox txtPONumber 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   5
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox txtPOAmount 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   4
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtPODeliveryDate 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   4
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtPOVendor 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   4
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   1440
      Width           =   3375
   End
   Begin VB.TextBox txtPODate 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   4
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox txtPONumber 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   4
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox txtPOAmount 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   3
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtPODeliveryDate 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   3
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtPOVendor 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   3
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1200
      Width           =   3375
   End
   Begin VB.TextBox txtPODate 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   3
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtPONumber 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   3
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox txtPOAmount 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   2
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtPODeliveryDate 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   2
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtPOVendor 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   2
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   960
      Width           =   3375
   End
   Begin VB.TextBox txtPODate 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   2
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txtPONumber 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   2
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox txtPOAmount 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   1
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox txtPODeliveryDate 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   1
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox txtPOVendor 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   1
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   720
      Width           =   3375
   End
   Begin VB.TextBox txtPODate 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   1
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox txtPONumber 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   1
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox txtPOAmount 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   0
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtPODeliveryDate 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   0
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtPOVendor 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   0
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   3375
   End
   Begin VB.TextBox txtPODate 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   0
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txtPONumber 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   0
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Amount"
      Height          =   255
      Left            =   7920
      TabIndex        =   55
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Delivery Date"
      Height          =   255
      Left            =   6120
      TabIndex        =   54
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Vendor"
      Height          =   255
      Left            =   2760
      TabIndex        =   53
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Date"
      Height          =   255
      Left            =   1440
      TabIndex        =   52
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "PO Number"
      Height          =   255
      Left            =   240
      TabIndex        =   51
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmPurchaseOrderSelect"
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

Dim strPOInfo(30) As String
Dim intPOCount As Integer
Dim intHighlightedPORow As Integer
Dim intActualHighlightedPO As Integer
Dim intTopDisplayedPO As Integer

Private Sub Form_Load()
  intHighlightedPORow = -1
  intActualHighlightedPO = -1
  
  RefreshPOList True
End Sub


Private Sub cmdExit_Click()
  EndSessionCloseConnection
  End
End Sub


Private Sub cmdPODetails_Click()
  If intActualHighlightedPO = -1 Then
    MsgBox "You must select a Purchase Order before you can view it's details"
    Exit Sub
  End If
  
  SetSelectedPOInfo strPOInfo(intActualHighlightedPO)
  Load frmPurchaseOrderDetail
  frmPurchaseOrderDetail.Show
End Sub


Private Sub FillPOInfoDisplay()

  intPOCount = 0
  
  If InStr(1, strPOInfo(1), "There were no open purchase orders returned") > 0 Then
    txtPOVendor(0).Text = "There were no open purchase orders returned"
    vscPOScroll.Enabled = False
    vscPOScroll.Visible = False
  End If
  
  intTopDisplayedPO = 1
  Dim strPOInfoSplit() As String
  Dim booDone As Boolean
  booDone = False
  Do
    intPOCount = intPOCount + 1
    If intPOCount < 11 Then
      strPOInfoSplit = Split(strPOInfo(intPOCount), "<spliter>")
      txtPONumber(intPOCount - 1).Text = strPOInfoSplit(0)
      txtPODate(intPOCount - 1).Text = strPOInfoSplit(1)
      txtPOVendor(intPOCount - 1).Text = strPOInfoSplit(2)
      txtPODeliveryDate(intPOCount - 1).Text = strPOInfoSplit(3)
      txtPOAmount(intPOCount - 1).Text = strPOInfoSplit(4)
    End If
    If intPOCount = UBound(strPOInfo) Then
      booDone = True
    ElseIf strPOInfo(intPOCount + 1) = Empty Then
      booDone = True
    End If
  Loop Until booDone
  
  vscPOScroll.Value = 1
  If intPOCount < 11 Then
    vscPOScroll.Enabled = False
    vscPOScroll.Visible = False
    Exit Sub
  End If
  
  vscPOScroll.Min = 1
  vscPOScroll.Max = intPOCount - 9
  vscPOScroll.Enabled = True
  vscPOScroll.Visible = True
  
End Sub


Private Sub ClearPOInfo()
  Dim i As Integer
  For i = 1 To 30
    strPOInfo(i) = Empty
  Next
  
  For i = 0 To 9
    txtPONumber(i).Text = Empty
    txtPODate(i).Text = Empty
    txtPOVendor(i).Text = Empty
    txtPODeliveryDate(i).Text = Empty
    txtPOAmount(i).Text = Empty
  Next i
End Sub

Private Sub HighlightPORow(intRowNum As Integer)
  If intHighlightedPORow >= 0 Then
    UnHighlightPORow intHighlightedPORow
  End If
  
  txtPONumber(intRowNum).BackColor = &HFFFF00
  txtSpace1(intRowNum).BackColor = &HFFFF00
  txtPODate(intRowNum).BackColor = &HFFFF00
  txtSpace2(intRowNum).BackColor = &HFFFF00
  txtPOVendor(intRowNum).BackColor = &HFFFF00
  txtPODeliveryDate(intRowNum).BackColor = &HFFFF00
  txtPOAmount(intRowNum).BackColor = &HFFFF00

  intHighlightedPORow = intRowNum
  intActualHighlightedPO = vscPOScroll.Value + intRowNum
End Sub


Private Sub UnHighlightPORow(intRowNum As Integer)
  If intHighlightedPORow >= 0 Then
    txtPONumber(intRowNum).BackColor = &H80000005
    txtSpace1(intRowNum).BackColor = &H80000005
    txtPODate(intRowNum).BackColor = &H80000005
    txtSpace2(intRowNum).BackColor = &H80000005
    txtPOVendor(intRowNum).BackColor = &H80000005
    txtPODeliveryDate(intRowNum).BackColor = &H80000005
    txtPOAmount(intRowNum).BackColor = &H80000005
  End If
  
  intHighlightedPORow = -1
End Sub

Private Sub txtPOAmount_Click(Index As Integer)
  If Index < intPOCount Then
    HighlightPORow (Index)
  End If
End Sub

Private Sub txtPODate_Click(Index As Integer)
  If Index < intPOCount Then
    HighlightPORow (Index)
  End If
End Sub

Private Sub txtPODeliveryDate_Click(Index As Integer)
  If Index < intPOCount Then
    HighlightPORow (Index)
  End If
End Sub

Private Sub txtPONumber_Click(Index As Integer)
  If Index < intPOCount Then
    HighlightPORow (Index)
  End If
End Sub

Private Sub txtPOVendor_Click(Index As Integer)
  If Index < intPOCount Then
    HighlightPORow (Index)
  End If
End Sub

Private Sub txtSpace1_Click(Index As Integer)
  If Index < intPOCount Then
    HighlightPORow (Index)
  End If
End Sub

Private Sub txtSpace2_Click(Index As Integer)
  If Index < intPOCount Then
    HighlightPORow (Index)
  End If
End Sub

Private Sub vscPOScroll_Change()
  UnHighlightPORow intHighlightedPORow
  
  Dim strPOInfoSplit() As String
  Dim i As Integer
  intTopDisplayedPO = vscPOScroll.Value
  For i = 0 To 9
    If strPOInfo(vscPOScroll.Value + i) <> Empty Then
      strPOInfoSplit = Split(strPOInfo(vscPOScroll.Value + i), "<spliter>")
      txtPONumber(i).Text = strPOInfoSplit(0)
      txtPODate(i).Text = strPOInfoSplit(1)
      txtPOVendor(i).Text = strPOInfoSplit(2)
      txtPODeliveryDate(i).Text = strPOInfoSplit(3)
      txtPOAmount(i).Text = strPOInfoSplit(4)
    End If
  Next
  
  If intActualHighlightedPO >= vscPOScroll.Value And _
     intActualHighlightedPO <= vscPOScroll.Value + 9 Then
    HighlightPORow intActualHighlightedPO - vscPOScroll.Value
  End If
  
  If frmPurchaseOrderSelect.Visible Then
    txtPONumber(0).SetFocus
  End If
End Sub

Public Sub RefreshPOList(booGiveWarning As Boolean)

  If intHighlightedPORow >= 0 Then
    UnHighlightPORow intHighlightedPORow
  End If
  
  ClearPOInfo
  LoadPOInfoArray strPOInfo, booGiveWarning
  FillPOInfoDisplay
  
  intHighlightedPORow = -1
  intActualHighlightedPO = -1
End Sub
