VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form UIForm 
   Caption         =   "SDKTest"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   6675
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox UseQBOE 
      Caption         =   "Send to QuickBooks Online Edition (QBOE)"
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   120
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Request"
      Height          =   495
      Left            =   1200
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox inputFile 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   4935
   End
   Begin VB.CommandButton browseButton 
      Caption         =   "Browse"
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog browseDialog 
      Left            =   5040
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox logBox 
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   1800
      Width           =   6375
   End
   Begin VB.CommandButton goBut 
      Caption         =   "Process Request"
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "qbXML Request File"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "UIForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'
' $Id: UIForm.frm,v 1.2 2002/08/05 13:00:33 giza Exp $
'
Option Explicit
'
'
'
Private Sub Command1_Click()
    ' Ensure that there's an input xml file.
    Dim InputFilename As String
    InputFilename = inputFile.Text
    Log.showURL (InputFilename)
End Sub

'
' This method is called by the VB runtime when the application is
' first started.
'
Private Sub Form_Load()
    Log.clear
    Dim args As Collection
    Set args = GetCommandLine
    If (args.Count > 0) Then
        inputFile.Text = args.Item(1)
    End If
End Sub

Private Sub Form_Terminate()
    Log.clearXML
End Sub

'
' This method is invoked when the user hits the Go button.
'
Private Sub goBut_Click()

    Log.clear

    ' Ensure that there's an input xml file.
    On Error GoTo problem
    Dim InputFilename As String
    InputFilename = inputFile.Text
    Dim fnum As Integer
    Dim isOpen As Boolean
    fnum = FreeFile()
    Open InputFilename For Input As #fnum
    isOpen = True
        
    ' Load the file into a string
    Dim xmldoc As String
    xmldoc = Input(LOF(fnum), fnum)
    Close #fnum

    ' Send the request to QuickBooks and display the response.
    Dim response As String
    response = qbooks.post(xmldoc)
    If (response = "") Then
        Exit Sub
    End If
    Log.showXMLStream (response)
    Exit Sub
problem:
    If isOpen Then Close #fnum
    If Err Then Err.Raise Err.Number, , Err.Description
End Sub
'
' Vanilla Windows UI code for file browsing.
'
Private Sub browseButton_Click()
  On Error GoTo ErrorHandler

  ' Set CancelError to True
  browseDialog.CancelError = True
  
  ' Set filters
  browseDialog.Filter = "All Files (*.*)|*.*|qbXML Request Files" & _
        "(*.xml)|*.xml"
  
  ' Specify default filter
  browseDialog.FilterIndex = 2
  
  ' Set flags
  browseDialog.Flags = cdlOFNHideReadOnly
  
  browseDialog.InitDir = App.Path & "\..\..\..\..\xmlfiles"

    
  ' Display the Open dialog box
  browseDialog.ShowOpen
  
  ' Set up the selected file in textBox
  inputFile.Text = browseDialog.FileName
  Exit Sub
  
ErrorHandler:
  Exit Sub
End Sub

Private Function GetCommandLine() As Collection
    'Declare variables.
    Dim C, CmdLine, CmdLnLen, InArg, I, NumArgs
    Dim InQuote As Boolean
    Dim ArgArray As New Collection
    Dim Arg As String
    Arg = ""
    'Get command line arguments and make sure we have a space at end.
    CmdLine = Command() & " "
    CmdLnLen = Len(CmdLine)
    'Go thru command line one character
    'at a time.
    For I = 1 To CmdLnLen
        C = Mid(CmdLine, I, 1)
        '  Test for a quote
        If (C = Chr(34)) Then
            If (InQuote) Then
                InQuote = False
            Else
                InQuote = True
            End If
        Else
            'Test for space or tab.
            If (InQuote Or (C <> " " And C <> vbTab)) Then
                'Neither space nor tab or in a quoted string
                'Test if already in argument.
                If Not InArg Then
                    'New argument begins.
                    'Test for too many arguments.
                    InArg = True
                End If
                'Concatenate character to current argument.
                Arg = Arg & C
            Else
                'Found a space or tab.
                'Set InArg flag to False.
                If (InArg) Then ArgArray.Add Arg
                InArg = False
                Arg = ""
            End If
        End If
    Next I
    Set GetCommandLine = ArgArray
End Function


