VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSDKTestPlus3 
   Caption         =   "SDKTestPlus3"
   ClientHeight    =   7965
   ClientLeft      =   2910
   ClientTop       =   2190
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   ScaleHeight     =   7965
   ScaleWidth      =   9765
   Begin VB.TextBox lblStatus 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   41
      Top             =   6000
      Width           =   7935
   End
   Begin VB.Frame AuthFlagsFrame 
      Caption         =   "AuthFlags"
      Height          =   2055
      Left            =   240
      TabIndex        =   35
      Top             =   1800
      Width           =   1935
      Begin VB.CheckBox qbForceAuthDialog 
         Caption         =   "Force auth dialog"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CheckBox qbSimple 
         Caption         =   "QB Simple Start"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CheckBox qbPro 
         Caption         =   "QB Pro"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   960
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox qbPremier 
         Caption         =   "QB Premier"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   600
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox qbEnterprise 
         Caption         =   "QB Enterprise"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Value           =   1  'Checked
         Width           =   1575
      End
   End
   Begin VB.Frame SessionPrefsFrame 
      Caption         =   "Session  Prefs"
      Enabled         =   0   'False
      Height          =   2055
      Left            =   4440
      TabIndex        =   23
      Top             =   1800
      Width           =   3255
      Begin VB.Frame Frame1 
         Caption         =   "File Mode"
         Height          =   1335
         Left            =   1680
         TabIndex        =   30
         Top             =   600
         Width           =   1455
         Begin VB.OptionButton optDontCare 
            Caption         =   "Don't Care"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   960
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optMultiUser 
            Caption         =   "Multi-User"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optSingleUser 
            Caption         =   "Single User"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   600
            Width           =   1215
         End
      End
      Begin VB.Frame pdFrame 
         Caption         =   "Personal Data"
         Height          =   1335
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   1455
         Begin VB.OptionButton pdNotNeeded 
            Caption         =   "Not Needed"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   960
            Width           =   1215
         End
         Begin VB.OptionButton pdOptional 
            Caption         =   "Optional"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   600
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton pdRequired 
            Caption         =   "Required"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.CheckBox ReadOnly 
         Caption         =   "Read-only"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1680
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox Unattended 
         Caption         =   "Unattended"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame ConnFrame 
      Caption         =   "Connection Type"
      Height          =   2055
      Left            =   2520
      TabIndex        =   17
      Top             =   1800
      Width           =   1575
      Begin VB.OptionButton ConnLocalUI 
         Caption         =   "Local with UI"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1320
         Width           =   1335
      End
      Begin VB.OptionButton ConnQBOE 
         Caption         =   "QBOE"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1680
         Width           =   1335
      End
      Begin VB.OptionButton ConnRemote 
         Caption         =   "Remote QB"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton ConnLocal 
         Caption         =   "Local QB"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton ConnNone 
         Caption         =   "No Pref"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.CommandButton RequestBrowse 
      Caption         =   "Browse..."
      Height          =   375
      Left            =   7800
      TabIndex        =   16
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox RequestFile 
      Height          =   285
      Left            =   240
      TabIndex        =   14
      Top             =   1320
      Width           =   7455
   End
   Begin MSComDlg.CommonDialog OpenFileDialog 
      Left            =   9000
      Top             =   6720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CompanyBrowse 
      Caption         =   "Browse..."
      Height          =   375
      Left            =   7800
      TabIndex        =   13
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox CompanyFile 
      Height          =   285
      Left            =   240
      TabIndex        =   11
      Top             =   720
      Width           =   7455
   End
   Begin VB.CommandButton cmdProcessSubscription 
      Caption         =   "Send XML to Subscription Processor"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2760
      TabIndex        =   10
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton cmdSendXML 
      Caption         =   "Send XML to Request Processor"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6240
      TabIndex        =   8
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton cmdViewInput 
      Caption         =   "View Input"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7920
      TabIndex        =   7
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton cmdViewOutput 
      Caption         =   "View Output"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7920
      TabIndex        =   6
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   8280
      TabIndex        =   5
      Top             =   7350
      Width           =   1335
   End
   Begin VB.CommandButton cmdOpenConnection 
      Caption         =   "Open Connection"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton cmdBeginSession 
      Caption         =   "Begin Session"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton cmdEndSession 
      Caption         =   "End Session"
      Enabled         =   0   'False
      Height          =   495
      Left            =   8280
      TabIndex        =   2
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdCloseConnection 
      Caption         =   "Close Connection"
      Enabled         =   0   'False
      Height          =   495
      Left            =   8160
      TabIndex        =   1
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox txtApplicationName 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Text            =   "SDKTest Plus 3"
      Top             =   105
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Status:"
      Height          =   255
      Left            =   240
      TabIndex        =   34
      Top             =   5640
      Width           =   855
   End
   Begin VB.Line Line19 
      X1              =   8760
      X2              =   8880
      Y1              =   4920
      Y2              =   5040
   End
   Begin VB.Line Line18 
      X1              =   8880
      X2              =   9000
      Y1              =   5040
      Y2              =   4920
   End
   Begin VB.Line Line17 
      X1              =   8040
      X2              =   8160
      Y1              =   5400
      Y2              =   5280
   End
   Begin VB.Line Line16 
      X1              =   8160
      X2              =   8040
      Y1              =   5280
      Y2              =   5160
   End
   Begin VB.Line Line15 
      X1              =   8160
      X2              =   8280
      Y1              =   4560
      Y2              =   4440
   End
   Begin VB.Line Line14 
      X1              =   8280
      X2              =   8160
      Y1              =   4440
      Y2              =   4320
   End
   Begin VB.Line Line13 
      X1              =   6120
      X2              =   6240
      Y1              =   4560
      Y2              =   4440
   End
   Begin VB.Line Line12 
      X1              =   6240
      X2              =   6120
      Y1              =   4440
      Y2              =   4320
   End
   Begin VB.Line Line11 
      X1              =   2640
      X2              =   2760
      Y1              =   5400
      Y2              =   5280
   End
   Begin VB.Line Line10 
      X1              =   2760
      X2              =   2640
      Y1              =   5280
      Y2              =   5160
   End
   Begin VB.Line Line9 
      X1              =   2640
      X2              =   2760
      Y1              =   4560
      Y2              =   4440
   End
   Begin VB.Line Line8 
      X1              =   2760
      X2              =   2640
      Y1              =   4440
      Y2              =   4320
   End
   Begin VB.Line Line7 
      BorderWidth     =   3
      X1              =   8880
      X2              =   8880
      Y1              =   4680
      Y2              =   5040
   End
   Begin VB.Line Line6 
      BorderWidth     =   3
      X1              =   7800
      X2              =   8280
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line5 
      BorderWidth     =   3
      X1              =   4560
      X2              =   6240
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line4 
      BorderWidth     =   3
      X1              =   4560
      X2              =   8160
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   960
      X2              =   2760
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   960
      X2              =   960
      Y1              =   4680
      Y2              =   5280
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   1800
      X2              =   2760
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label Label3 
      Caption         =   "Request File:"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Company File: (leave blank to use current open file)"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Application Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmSDKTestPlus3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' (c) 2003 Intuit Inc. All Rights Reserved                          '
' Use is subject to IP Rights Notice and Restrictions available at: '
' http://developer.quickbooks.com/legal/IPRNotice_021201.html       '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strCompanyFilename As String
Dim strInXMLFile As String
Dim strOutXMLFile As String
Dim strError As String
Dim strAppName As String
Dim strOldDrive As String
Public AuthPrefsDirty As Boolean
Public ConnPrefsDirty As Boolean

Public DisplayFile As String

Private Sub FileError(FileType As String)
  MsgBox "You must select " & FileType & " file.", vbExclamation, "Beta qbXML Test"
End Sub
Private Sub Clear_Status_Area()
  lblStatus.Text = ""
  lblStatus.Refresh
End Sub

Private Sub EnableAuthFlagsSettings(EnableOrDisable As Boolean)

    AuthFlagsFrame.Enabled = EnableOrDisable
    qbEnterprise.Enabled = EnableOrDisable
    qbPremier.Enabled = EnableOrDisable
    qbPro.Enabled = EnableOrDisable
    qbSimple.Enabled = EnableOrDisable
    qbForceAuthDialog.Enabled = EnableOrDisable
End Sub

Public Sub showURL(iFname As String)
    Dim IE1 As New InternetExplorer
    IE1.Visible = True
    IE1.Navigate (iFname)
End Sub

Private Sub cmdBeginSession_Click()
  
  Dim strFileMode As String
    
  If optSingleUser.Value Then
    strFileMode = "Single"
  ElseIf optMultiUser.Value Then
    strFileMode = "Multi"
  Else
    strFileMode = "DoNotCare"
  End If
  
  strError = BeginSession(CompanyFile.Text, strFileMode, lblStatus)
  
  If strError = Empty Then
    cmdBeginSession.Enabled = False
    cmdCloseConnection.Enabled = False
    cmdEndSession.Enabled = True
    sessionPrefsControl (False)
    cmdSendXML.Enabled = (Not RequestFile.Text = "")
  End If
End Sub

Private Sub cmdCloseConnection_Click()
  CloseConnection lblStatus
  cmdBeginSession.Enabled = False
  cmdOpenConnection.Enabled = True
  cmdCloseConnection.Enabled = False
  cmdProcessSubscription.Enabled = False
  txtApplicationName.Enabled = True
End Sub

Private Sub cmdEndSession_Click()
    EndSession lblStatus
    cmdBeginSession.Enabled = True
    cmdCloseConnection.Enabled = True
    cmdEndSession.Enabled = False
    cmdSendXML.Enabled = False
    sessionPrefsControl (True)
End Sub

Private Sub cmdExit_Click()
  End
End Sub

Private Sub cmdOpenConnection_Click()
  Dim strFileMode As String
  
  If txtApplicationName.Text = Empty Then
    MsgBox "You must supply an application name before opening a connection"
    Exit Sub
  End If
  
  
  strAppName = txtApplicationName.Text
  strError = OpenConnection("", strAppName, lblStatus)
  
  If strError = Empty Then
    cmdBeginSession.Enabled = True
    cmdCloseConnection.Enabled = True
    cmdOpenConnection.Enabled = False
    txtApplicationName.Enabled = False
    cmdProcessSubscription.Enabled = (Not RequestFile.Text = "")
    sessionPrefsControl (True)
    lblStatus.Text = "OpenConnection call successful"
    lblStatus.Refresh
  End If
End Sub


Private Sub sessionPrefsControl(state As Boolean)
    SessionPrefsFrame.Enabled = state
    Unattended.Enabled = state
    ReadOnly.Enabled = state
    pdRequired.Enabled = state
    pdOptional.Enabled = state
    pdNotNeeded.Enabled = state
    optDontCare.Enabled = state
    optMultiUser.Enabled = state
    optSingleUser.Enabled = state
End Sub
Private Sub cmdProcessSubscription_Click()

  strOutXMLFile = CurDir & "\SubscriptionResponse.xml"
  
  strError = SendXMLFile("SubscriptionProcessor", RequestFile.Text, _
                         strOutXMLFile, lblStatus)
  If strError = Empty Then
    cmdViewOutput.Enabled = True
  End If
End Sub

Private Sub cmdSendXML_Click()

  strOutXMLFile = CurDir & "\QBResponse.xml"
  
  strError = SendXMLFile("RequestProcessor", RequestFile.Text, _
                         strOutXMLFile, lblStatus)
  If strError = Empty Then
    cmdViewOutput.Enabled = True
  End If
End Sub

Private Sub cmdViewInput_Click()
  showURL RequestFile.Text
End Sub

Private Sub cmdViewOutput_Click()
  showURL strOutXMLFile
End Sub


Private Sub CompanyBrowse_Click()
    OpenFileDialog.Filter = "QuickBooks Files|*.qbw"
    OpenFileDialog.InitDir = "c:\program files\Intuit"
    OpenFileDialog.ShowOpen
    CompanyFile.Text = OpenFileDialog.FileName
End Sub

Private Sub optUseCurrentFile_Click()
  lstCompanyFile.Refresh
End Sub

Private Sub ConnLocal_Click()
    ConnPrefsDirty = True
    EnableAuthFlagsSettings (True)
End Sub

Private Sub ConnLocalUI_Click()
    ConnPrefsDirty = True
    EnableAuthFlagsSettings (True)
End Sub

Private Sub ConnNone_Click()
    ConnPrefsDirty = True
    EnableAuthFlagsSettings (True)
End Sub

Private Sub ConnQBOE_Click()
    ConnPrefsDirty = True
    EnableAuthFlagsSettings (False)
End Sub

Private Sub ConnRemote_Click()
    ConnPrefsDirty = True
    EnableAuthFlagsSettings (False)
End Sub

Private Sub Form_Load()
    If (qbEnterprise.Value = 1 Or qbPremier.Value = 1 Or qbPro.Value = 1 Or qbSimple.Value = 1) Then
        AuthPrefsDirty = True
    Else
        AuthPrefsDirty = False
    End If
    ConnPrefsDirty = False
End Sub

Private Sub pdNotNeeded_Click()
    AuthPrefsDirty = True
End Sub

Private Sub pdOptional_Click()
    AuthPrefsDirty = True

End Sub

Private Sub pdRequired_Click()
    AuthPrefsDirty = True

End Sub


Private Sub ReadOnly_Click()
    AuthPrefsDirty = True
End Sub

Private Sub RequestBrowse_Click()
    OpenFileDialog.Filter = "XML Files|*.xml;*.qbxml"
    OpenFileDialog.InitDir = "../../../../xmlfiles"
    OpenFileDialog.ShowOpen
    RequestFile.Text = OpenFileDialog.FileName
End Sub

Private Sub RequestFile_Change()
    cmdSendXML.Enabled = (Not RequestFile.Text = "")
    cmdProcessSubscription.Enabled = (Not RequestFile.Text = "")
    cmdViewInput.Enabled = (Not RequestFile.Text = "")
End Sub

Private Sub Unattended_Click()
    AuthPrefsDirty = True
End Sub
