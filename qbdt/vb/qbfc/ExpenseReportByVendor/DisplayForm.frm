VERSION 5.00
Begin VB.Form DisplayForm 
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   5865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton closeButton 
      Caption         =   "Close"
      Height          =   255
      Left            =   1740
      TabIndex        =   1
      Top             =   4080
      Width           =   2355
   End
   Begin VB.TextBox displayText 
      Height          =   3915
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   60
      Width           =   5715
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
' This form can be used to display the request and response
' qbxml to the user for illustrational purposes.
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'
'


Private xmlString As String


'
' Set the xml string to be displayed later.
'
Public Sub setXML(xmlToShow As String)
    xmlString = xmlToShow
End Sub



'
' Display xml to the user.
'
Public Sub showXML(title As String)
    Me.Caption = title
    displayText.Text = xmlString
    Me.Show
End Sub



'
' Close the form.
'
Private Sub closeButton_Click()
    Unload Me
End Sub

