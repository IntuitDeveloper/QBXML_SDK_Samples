VERSION 5.00
Begin VB.Form CustomerList 
   Caption         =   "qbXML Sample: Select Customer"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7155
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   7155
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo_CustList 
      Height          =   315
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   825
      Width           =   2532
   End
   Begin VB.CommandButton Comm_Exit 
      Caption         =   "&Exit"
      Height          =   372
      Left            =   5160
      TabIndex        =   3
      Top             =   960
      Width           =   1812
   End
   Begin VB.CommandButton Comm_View_res 
      Caption         =   "View qbXML Re&sponse"
      Height          =   372
      Left            =   5160
      TabIndex        =   2
      Top             =   2280
      Width           =   1812
   End
   Begin VB.CommandButton Comm_View_Req 
      Caption         =   "View qbXML  Re&quest"
      Height          =   372
      Left            =   5160
      TabIndex        =   1
      Top             =   1800
      Width           =   1812
   End
   Begin VB.CommandButton Comm_Submit 
      Caption         =   "&Modify"
      Height          =   372
      Left            =   5160
      TabIndex        =   0
      Top             =   480
      Width           =   1812
   End
   Begin VB.Frame Frame1 
      Caption         =   "Please select the customer whose data you want to modify: "
      Height          =   1932
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   4692
      Begin VB.Label Label7 
         Caption         =   "Customer Full Name List:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   405
         Width           =   1815
      End
   End
   Begin VB.Image Image_QBBANNER 
      DragMode        =   1  'Automatic
      Height          =   345
      Left            =   240
      Top             =   3000
      Width           =   5040
   End
End
Attribute VB_Name = "CustomerList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------
' Form Module: CustomerList
'
' Description:  This form allows the user to select a customer to
'               be modified.
'
' Created On: 11/08/2001
' Updated to SDK 2.0: 08/05/2002
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'----------------------------------------------------------

Option Explicit

'Selected Customer Obj
Dim selectedCust As Customer


' Submit button
Private Sub Comm_Submit_Click()
    
On Error GoTo ErrHandler
 
    ' Collect Form Data
    If CollectFormData Then
    
        ' Load CustomerMod Form
        Dim CustModForm As New CustomerMod
        
        ' Pass the selected Customer to CustomerMod Form
        CustModForm.Customer selectedCust
        
        ' Set up the Form Caption with this customer's Name
        CustModForm.Caption = "qbXML Sample: Customer " & selectedCust.Name & " Information"
        
        ' Show the CustomerMod form
        CustModForm.Show vbModal, Me
        
        ' Close this form to go back to the main form
        Unload Me
        
    End If
    
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
    Exit Sub
    
End Sub

' Display request XML
Private Sub Comm_View_Req_Click()

On Error GoTo ErrHandler
        
    ' Instantiate a RequestXML Obj
    Dim reqFrm As Display
    Set reqFrm = New Display
    
    reqFrm.Text_Content = Main.GetRequestXML
    reqFrm.Caption = "Request XML"
    reqFrm.Show vbModal, Me
    
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
    Exit Sub
    
End Sub

' Display response XML
Private Sub Comm_View_Res_Click()
    
On Error GoTo ErrHandler
    
    Dim resFrm As New Display
    
    If Main.GetResponseXML <> "" Then
        Dim tmpRes As String
        tmpRes = Replace(Main.GetResponseXML, vbLf, vbCrLf)
        resFrm.Text_Content = tmpRes
    Else
        resFrm.Text_Content = ""
    End If
    resFrm.Caption = "Response XML"
    resFrm.Show vbModal, Me
    
    Exit Sub

ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
    Exit Sub
 
End Sub

Private Sub Comm_Exit_Click()
    Unload Me
End Sub


' Form Data Collection
Private Function CollectFormData() As Boolean

On Error GoTo ErrHandler:

    ' Customer Name
    Dim custFullName As String
    
    custFullName = Combo_CustList.Text
    If custFullName = "" Then
        MsgBox "Please select a customer"
        CollectFormData = False
        Exit Function
    End If
    
    ' Set the customer obj
    Set selectedCust = Main.customerCollection.Item(custFullName)
    
    CollectFormData = True
    Exit Function
    
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
    CollectFormData = False
    Exit Function
    
End Function


Private Sub Form_Load()
    
On Error GoTo ErrHandler

    ' Load the image
    Dim appPath
    appPath = App.Path
    Image_QBBANNER.Picture = LoadPicture(appPath & "/qbbanner.bmp")

    ' Initialize the customer combo list
    Dim cust As Customer
    For Each cust In Main.customerCollection
        Combo_CustList.AddItem cust.FullName
    Next
    
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
    Exit Sub

End Sub


Public Property Get SelectedCustomer() As Customer
        Set SelectedCustomer = selectedCust
End Property
        

