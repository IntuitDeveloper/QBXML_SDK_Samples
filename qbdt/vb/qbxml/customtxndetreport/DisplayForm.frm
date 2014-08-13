VERSION 5.00
Begin VB.Form DisplayForm 
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   7395
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton closeButton 
      Caption         =   "Close"
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   5640
      Width           =   2355
   End
   Begin VB.TextBox displayText 
      Height          =   5355
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   60
      Width           =   7275
   End
End
Attribute VB_Name = "DisplayForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' DisplayForm.frm
' Created July, 2002
'
' This is a simple form used to display QBXML requests
' and responses to the user.
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'


Private XMLString As String


'
' Set the xml string to be displayed later.
'
Public Sub setXML(xmlToShow As String)
    XMLString = xmlToShow
End Sub



'
' Display xml to the user.
'
Public Sub showXML(title As String)
    Me.Caption = title
    displayText.Text = XMLString
    Me.Show
End Sub



'
' Close the form.
'
Private Sub closeButton_Click()
    Unload Me
End Sub
