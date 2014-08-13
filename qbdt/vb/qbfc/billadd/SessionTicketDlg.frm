VERSION 5.00
Begin VB.Form SessionTicketDlg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Need Session Ticket"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox SessionTicketText 
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   2280
      Width           =   3855
   End
   Begin VB.CommandButton QBLogin 
      Caption         =   "Login to QuickBooks!"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   2415
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Enter Session Ticket:"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   $"SessionTicketDlg.frx":0000
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "SessionTicketDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' This is just a simple dialog that explains that the user must login to
' QuickBooks Online Edition and obtain a session ticket which they paste into
' this dialog so that we can tell the QBOESessionManager the session ticket that
' was obtained manually
'
' Note that this dialog and the CommPrefsDlg work hand-in-hand
'
Option Explicit

'
' Private variables
'
Private DidOK As Boolean  'Was OK pressed?
Private myLoginURL As String 'Passed in when dialog is invoked to tell us the QBOE URL to get a ticket
Private myAuthURL As String 'Passed in when dialog is invoked to tell us our connection ticket

'
' The "main" function of this dialog, this is the only way this dialog should be invoked
' (i.e. the program should not call SessionTicketDlg.Show itself, this function does it)
'
Public Sub GetSessionTicket(OESessionManager As QBOESessionManager, loginURL As String, authURL As String)
    '
    ' Save the URL for later
    '
    myLoginURL = loginURL
    myAuthURL = authURL
    '
    ' Show the dialog
    '
    SessionTicketDlg.Show vbModal
    '
    ' If the user clicked OK store the session ticket we got in the OESessionManager
    '
    If (DidOK) Then
        myAuthURL = myAuthURL & "&sessiontkt=" & SessionTicketDlg.SessionTicketText.Text
        OESessionManager.SessionTicket.setValue GetRealSessionTicket
    End If
End Sub

Private Sub CancelButton_Click()
    DidOK = False
    Me.Hide
End Sub

Private Sub OKButton_Click()
    DidOK = True
    Me.Hide
End Sub

'
' Open a browser to QuickBooks Online so the user can get the session ticket for us.
'
Private Sub QBLogin_Click()
    showURL myLoginURL
End Sub

Private Function GetRealSessionTicket() As String
    Dim http As New XMLHTTP40
    http.open "GET", myAuthURL, False
    http.send
    Dim resp As String
    resp = http.responseText
    resp = Mid(resp, 4)
    GetRealSessionTicket = resp
End Function

'
' Open an IE browser to a given URL to support Online Edition connection
' and session setup
'
Public Sub showURL(iFname As String)
    Dim IE1 As New InternetExplorer
    IE1.Visible = True
    IE1.Navigate iFname
End Sub


