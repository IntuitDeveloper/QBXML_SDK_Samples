VERSION 5.00
Begin VB.Form frmDisplay 
   Caption         =   "SDKTestPlus3 qbXML Display"
   ClientHeight    =   9105
   ClientLeft      =   3285
   ClientTop       =   1320
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   ScaleHeight     =   9105
   ScaleWidth      =   8640
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   615
      Left            =   3345
      TabIndex        =   1
      Top             =   8385
      Width           =   1935
   End
   Begin VB.TextBox txtDisplay 
      Height          =   8175
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   105
      Width           =   8415
   End
End
Attribute VB_Name = "frmDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' (c) 2003 Intuit Inc. All Rights Reserved                          '
' Use is subject to IP Rights Notice and Restrictions available at: '
' http://developer.quickbooks.com/legal/IPRNotice_021201.html       '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub cmdClose_Click()
  Unload frmDisplay
End Sub

Private Sub Form_Load()
  Dim strInput As String
  Dim DisplayString As String
  
  FileNum = FreeFile
  txtDisplay.Text = Empty
  Open frmSDKTestPlus3.DisplayFile For Input As #FileNum
  Do Until EOF(FileNum)
    Line Input #FileNum, strInput
    txtDisplay.Text = txtDisplay.Text & strInput & vbCrLf
  Loop
  Close #FileNum
'  txtDisplay.Text = DisplayString
End Sub
