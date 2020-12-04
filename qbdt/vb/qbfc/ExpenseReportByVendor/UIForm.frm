VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form UIForm 
   Caption         =   "QBSDK 2.0 --- Expense By Vendor Summary Report"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12420
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   12420
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton showResButton 
      Caption         =   "Show Response qbXML"
      Height          =   375
      Left            =   6300
      TabIndex        =   6
      Top             =   1260
      Width           =   2175
   End
   Begin VB.CommandButton showReqButton 
      Caption         =   "Show Request qbXML"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   1260
      Width           =   2175
   End
   Begin VB.CommandButton exitButton 
      Caption         =   "&Exit Program"
      Height          =   375
      Left            =   8880
      TabIndex        =   4
      Top             =   1260
      Width           =   2175
   End
   Begin VB.CommandButton goButton 
      Caption         =   "&Generate Report"
      Height          =   375
      Left            =   1140
      TabIndex        =   2
      Top             =   1260
      Width           =   2175
   End
   Begin VB.TextBox htmlFile 
      Height          =   375
      Left            =   4380
      TabIndex        =   1
      Top             =   720
      Width           =   5235
   End
   Begin SHDocVwCtl.WebBrowser embedBrowser 
      Height          =   6135
      Left            =   300
      TabIndex        =   0
      Top             =   1740
      Width           =   11835
      ExtentX         =   20876
      ExtentY         =   10821
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin MSComDlg.CommonDialog browseDialog 
      Left            =   360
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Select Output Html File:"
      Height          =   255
      Left            =   1860
      TabIndex        =   3
      Top             =   780
      Width           =   2475
   End
End
Attribute VB_Name = "UIForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' UIForm.frm
' Created July, 2002
'
' This file contains the code for the User Interface form of
' the project.  The UI allows the user to first select a
' QuickBooks company file and a location to store the output
' .html file.  When the GenerateReport button is pressed,
' it kicks off all of the main action.
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'
'



' Variables needed for displaying qbXML request and response
' in DisplayForm forms.
Private requestDisplay As New DisplayForm
Private responseDisplay As New DisplayForm
Private displayInit As Boolean



'
' Set default values when the form is loaded.
'
Private Sub Form_Load()

    ' Load a blank page on the embedded browser to avoid
    ' file not found pages
    embedBrowser.Navigate2 App.Path & "\blank.html"
    
    ' Default location for html file
    htmlFile.Text = App.Path & "\output.html"

End Sub



'
' Browse for a company file using the MS Common Dialog
' Control
'
Private Sub browseButton_Click()
  On Error GoTo ErrorHandler

  ' Set CancelError to True
  browseDialog.CancelError = True
  
  ' Set flags
  browseDialog.Flags = cdlOFNHideReadOnly
  
  ' Set filters
  browseDialog.Filter = "All Files (*.*)|*.*|QB Company Files" & _
    "(*.qbw)|*.qbw"
  
  ' Specify default filter
  browseDialog.FilterIndex = 2
  
  ' Display the Open dialog box
  browseDialog.ShowOpen
  
  ' Set up the selected file in textBox
  qbFile.Text = browseDialog.fileName
  Exit Sub
  
ErrorHandler:
  ' If an error occurs, do nothing
  Exit Sub
End Sub



'
' Generate and display the report when the go button is
' clicked if a company file has already been selected.
'
Private Sub goButton_Click()
    On Error GoTo ErrorHandler

    ' Make sure we have an output file path
    Dim url As String
    
    If ("" = htmlFile.Text) Then
        MsgBox "Please fill in a path for the html output file."
        Exit Sub
    Else
        url = htmlFile.Text
    End If

    UIForm.MousePointer = vbHourglass

    Dim xmlResponse As IMsgSetResponse
    Dim xmlRequest As IMsgSetRequest

    ' create the request and send the request to QB
    ' receive the response back from QB
    Dim bError As Boolean
    bError = sendReqToQB("", xmlRequest, xmlResponse)

    ' Exit sub if error from QuickBooks was already
    ' reported in sendReqToQB
    If (bError) Then
        GoTo ErrorHandler2
    End If

    ' Set the two DisplayForm forms so that the user will be
    ' able to view the request and response by clicking on the
    ' Show Request and Show Response buttons.
    requestDisplay.setXML (xmlRequest.ToXMLString())
    responseDisplay.setXML (xmlResponse.ToXMLString())
    displayInit = True

    ' process the response and create the html page to display
    Dim bSuccess As Boolean
    bSuccess = parseResponse(xmlResponse, url)

    ' Exit sub if the xml was not parsed correctly
    If (Not bSuccess) Then
        GoTo ErrorHandler2
    End If
    
    ' Finally, display the result html in the embedded browser
    embedBrowser.Navigate2 url
    
    UIForm.MousePointer = vbDefault
    Exit Sub

ErrorHandler:
    UIForm.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical
    Exit Sub

' Used if another function has already displayed
' and error message
ErrorHandler2:
    UIForm.MousePointer = vbDefault
    Exit Sub
End Sub



'
' Display XML Request if one has been generated already.
'
Private Sub showReqButton_Click()
    On Error GoTo showReqErrorHandler
    
    If displayInit Then
        requestDisplay.showXML "qbXML Request:"
    Else
        MsgBox "Please click Generate Report before trying to " _
            & "view the qbXML request."
    End If

    Exit Sub

showReqErrorHandler:
    MsgBox Err.Description
    Exit Sub
End Sub



'
' Display XML Response if one has been generated already.
'
Private Sub showResButton_Click()
    On Error GoTo showResErrorHandler
    
    If displayInit Then
        responseDisplay.showXML "qbXML Response:"
    Else
        MsgBox "Please click Generate Report before trying to " _
            & "view the qbXML response."
    End If
    
    Exit Sub
    
showResErrorHandler:
    MsgBox Err.Description
    Exit Sub
End Sub



'
' Quit the program.
'
Private Sub exitButton_Click()
    Unload Me
End Sub
