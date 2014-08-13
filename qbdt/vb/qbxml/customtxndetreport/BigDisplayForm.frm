VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form BigDisplayForm 
   Caption         =   "Your Report"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12090
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   12090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton closeButton 
      Caption         =   "Close Window"
      Height          =   435
      Left            =   4860
      TabIndex        =   1
      Top             =   7020
      Width           =   2415
   End
   Begin SHDocVwCtl.WebBrowser browser 
      Height          =   6615
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   11715
      ExtentX         =   20664
      ExtentY         =   11668
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
      Location        =   "http:///"
   End
End
Attribute VB_Name = "BigDisplayForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' BigDisplayForm.frm
' Created July, 2002
'
' This form contains an embedded browser which is
' used to display the report we generate to the user.
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'



'
' Display the html page at url in the embedded browser.
'
Public Sub displayResults(url As String)
    On Error GoTo ErrorHandler
    
    Me.Show
    browser.Navigate2 url
    Exit Sub
    
ErrorHandler:
    MsgBox "Error displaying results"
    Unload Me
    Exit Sub
End Sub



'
' Close the form when the user clicks the close button.
'
Private Sub closeButton_Click()
    Unload Me
End Sub

