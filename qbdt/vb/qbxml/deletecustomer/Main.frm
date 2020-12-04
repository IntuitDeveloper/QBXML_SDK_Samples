VERSION 5.00
Begin VB.Form Mainfrm 
   Caption         =   "qbXML Sample: Deleting a Customer"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   9885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Comm_View_Req 
      Caption         =   "&View qbXML Request"
      Height          =   372
      Left            =   7680
      TabIndex        =   12
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton Comm_View_Res 
      Caption         =   "Vi&ew qbXML Response"
      Height          =   372
      Left            =   7680
      TabIndex        =   11
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CommandButton Comm_Exit 
      Caption         =   "&Done"
      Height          =   372
      Left            =   7680
      TabIndex        =   10
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton Comm_Submit_Add 
      Caption         =   "&Add Customer"
      Default         =   -1  'True
      Height          =   372
      Left            =   4680
      TabIndex        =   4
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox Text_CustomerName 
      Height          =   288
      Left            =   4680
      TabIndex        =   3
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton Comm_Submit_Remove 
      Caption         =   "Delete Customer"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   5520
      Width           =   2055
   End
   Begin VB.TextBox Text_ListID 
      Height          =   288
      Left            =   600
      TabIndex        =   1
      Top             =   5640
      Width           =   2535
   End
   Begin VB.CommandButton Comm_Submit_Find 
      Caption         =   "Find An ""In Use"" Customer"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Line Line5 
      X1              =   840
      X2              =   9600
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line4 
      X1              =   7320
      X2              =   7320
      Y1              =   960
      Y2              =   6600
   End
   Begin VB.Image Image_QBBANNER 
      Height          =   435
      Left            =   720
      Picture         =   "Main.frx":0000
      Top             =   6720
      Width           =   6450
   End
   Begin VB.Label Label7 
      Caption         =   "&Customer (Required for Add):"
      Height          =   255
      Left            =   4740
      TabIndex        =   9
      Top             =   1200
      Width           =   2595
   End
   Begin VB.Image Image_QBCUST 
      Height          =   495
      Left            =   240
      Picture         =   "Main.frx":3532
      Top             =   360
      Width           =   450
   End
   Begin VB.Label Label4 
      Caption         =   "List&ID (Required for Delete):"
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   5400
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "This sample program illustrates how to delete a customer using the QuickBooks SDK."
      Height          =   435
      Left            =   720
      TabIndex        =   7
      Top             =   360
      Width           =   7275
   End
   Begin VB.Label Label2 
      Caption         =   $"Main.frx":3D94
      Height          =   1215
      Left            =   660
      TabIndex        =   6
      Top             =   1320
      Width           =   3555
   End
   Begin VB.Label Label3 
      Caption         =   $"Main.frx":3E87
      Height          =   1695
      Left            =   600
      TabIndex        =   5
      Top             =   3240
      Width           =   3615
   End
   Begin VB.Line Line1 
      X1              =   600
      X2              =   6960
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line2 
      X1              =   480
      X2              =   6840
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line3 
      X1              =   660
      X2              =   6840
      Y1              =   6240
      Y2              =   6240
   End
End
Attribute VB_Name = "Mainfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Main.frm
'
' This is the main form of the Delete Customer sample application which
' uses QuickBooks SDK 2.0.  This form contains most of the data for the User
' Interface. The code for communicating with QuickBooks is handled in a
' separate module qbModule.bas. Likewise, the XML generating and parsing code
' is handled by the separate module XMLBuilder.
'
' This sample application illustrates how to use the DeleteCustomer
' functionality of the SDK.  It shows, first, how to delete
' a customer who is inactive.  Next it allows the user to attempt to delete
' a customer who is currently active and view the errors reported by QuickBooks
' in this case.
'
'
' Created: February, 2002
' Updated to SDK 2.0 August, 2002
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'----------------------------------------------------------

Option Explicit

' Declare variables for Customer info
Dim custName    As String
Dim ListID      As String

' Declare variables for Response info
Dim resCustName     As String
Dim resCustFullName As String
Dim resListID       As String


' Submit button for adding a customer
Private Sub Comm_Submit_Add_Click()

On Error GoTo ErrHandler

    ' Initialize variables
    requestXML = ""
    responseXML = ""

    ' Get input data
    If Not CollectFormData("Add") Then
        Exit Sub
    End If
    
    If Not blnIsOpenConnection Then
        'Call OpenConnection to establish a QuickBooks connection and session
        If Not OpenConnection Then
         Exit Sub
        End If
    End If
    
    ' Build the request XML to add a customer
    requestXML = CreateCustomerAddRq
    
        
    ' Send request to QuickBooks
    DoRequest
    
    
    ' Parse response
    If Not ParseResponseXML("CustomerAddRs") Then
        CloseConnection
        Exit Sub
    End If
                
    
    Text_ListID.Text = resListID
                
    ' Display the result
    MsgBox "Customer " & custName & " has been successfully created." & vbCr & vbCr & _
                "  ListID = " & resListID & vbCr & _
                "  FullName = " & resCustFullName & vbCr & vbCr & "The ListID test box has been automatically populated with this value.", vbOKOnly, "Success"
             
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
    CloseConnection
    Exit Sub
    
End Sub


' Query QuickBooks for open Invoices, pull out a customer ListID from
' one of them, and fill in the ListID box with this value.
Private Sub Comm_Submit_Find_Click()

On Error GoTo ErrHandler

    ' Initialize
    requestXML = ""
    responseXML = ""
    
    
    If Not blnIsOpenConnection Then
        'Call OpenConnection
        If Not OpenConnection Then
         Exit Sub
        End If
    End If
    
    ' Build the request XML to query QuickBooks for open Invoices
    requestXML = CreateInvoiceQueryRq
    
    ' Send request to QuickBooks
    DoRequest
    
    ' Parse response
    If Not ParseResponseXML("InvoiceQueryRs") Then
        Exit Sub
    End If
                
    Mainfrm.Text_ListID.Text = resListID
    Mainfrm.Text_CustomerName.Text = ""
                
    ' Display the result
    If resListID <> "" Then
        MsgBox "There is an invoice open for customer with ListID " & resListID _
            & ". The ListID test box has been automatically populated with this value.", vbOKOnly, "Success"
    Else
        MsgBox "No open invoices were found in QuickBooks data file.", vbOKOnly, "Success"
    End If
             
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
    CloseConnection
    Exit Sub

End Sub


' Submit button for deleting a customer
Private Sub Comm_Submit_Remove_Click()
    On Error GoTo ErrHandler

    ' Get input data
    If Not CollectFormData("Delete") Then
        Exit Sub
    End If
    
    If Not blnIsOpenConnection Then
        'Call OpenConnection
        If Not OpenConnection Then
         Exit Sub
        End If
    End If
   
    ' Build the request XML to delete a customer
    requestXML = CreateCustomerDelRq
       
    ' Send request to QuickBooks
    DoRequest
    
    
    ' Parse response
    If Not ParseResponseXML("ListDelRs") Then
        Exit Sub
    End If
                
    ' Display the result
    MsgBox "Customer " & custName & " has been successfully deleted.", , "Success"
             
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
    CloseConnection
    Exit Sub

End Sub


' View qbXML Request button
Private Sub Comm_View_Req_Click()

On Error GoTo ErrHandler
    
    If requestXML <> "" Then
        
        Dim reqFrm As DisplayXML
        Set reqFrm = New DisplayXML
        
        reqFrm.Text_Content = requestXML
        reqFrm.Caption = "Request XML"
        reqFrm.Show
        
    Else
        MsgBox "Request is empty.  Please add a customer first", vbInformation, "Request XML"
    End If
    Exit Sub
        

ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
    Exit Sub

End Sub

' View Response Button
Private Sub Comm_View_Res_Click()
 
 On Error GoTo ErrHandler
        
    If responseXML <> "" Then
        
        Dim resFrm As New DisplayXML
        Dim tmpResponseXML As String
        
        ' replace Lf to CrLf, this is for display only
        tmpResponseXML = Replace(responseXML, vbLf, vbCrLf)
        resFrm.Text_Content = tmpResponseXML
        resFrm.Caption = "Response XML"
        resFrm.Show
        
    Else
        MsgBox "Response is empty.  Please add a customer first", vbInformation, "Response XML"
    End If
    
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
    Exit Sub

End Sub

' Exit button
Private Sub Comm_Exit_Click()
    Unload Me ' close the window
End Sub


' Create CustomerAdd request
Private Function CreateCustomerAddRq() As String

    Dim doc As New DOMDocument40
    Dim QBXML As IXMLDOMNode
    Dim MsgsRq As IXMLDOMElement
    Dim CustomerAdd As IXMLDOMElement
    Dim Name As IXMLDOMElement
    
           
    Set QBXML = doc.appendChild(doc.createElement("QBXML"))
    Set MsgsRq = QBXML.appendChild(doc.createElement("QBXMLMsgsRq"))
    MsgsRq.setAttribute "onError", "continueOnError"
    Set CustomerAdd = MsgsRq.appendChild(doc.createElement("CustomerAddRq"))
    Set CustomerAdd = CustomerAdd.appendChild(doc.createElement("CustomerAdd"))
    Set Name = CustomerAdd.appendChild(doc.createElement("Name"))
    Name.appendChild doc.createTextNode(custName)
    CreateCustomerAddRq = doc.xml
    
End Function

' Create an invoice query request
Private Function CreateInvoiceQueryRq() As String

    Dim doc As New DOMDocument40
    Dim QBXML As IXMLDOMNode
    Dim MsgsRq As IXMLDOMElement
    Dim InvoiceQuery As IXMLDOMElement
    Dim PaidStatus As IXMLDOMElement
    Dim ListID As IXMLDOMElement
  
           
    Set QBXML = doc.appendChild(doc.createElement("QBXML"))
    Set MsgsRq = QBXML.appendChild(doc.createElement("QBXMLMsgsRq"))
    MsgsRq.setAttribute "onError", "continueOnError"
    Set InvoiceQuery = MsgsRq.appendChild(doc.createElement("InvoiceQueryRq"))
    Set PaidStatus = InvoiceQuery.appendChild(doc.createElement("PaidStatus"))
    PaidStatus.appendChild doc.createTextNode("NotPaidOnly")
    CreateInvoiceQueryRq = doc.xml
End Function


' Create Customer Delete request
Private Function CreateCustomerDelRq() As String

    Dim doc As New DOMDocument40
    Dim QBXML As IXMLDOMNode
    Dim MsgsRq As IXMLDOMElement
    Dim ListDel As IXMLDOMElement
    Dim dataElement As IXMLDOMElement
    
    Set QBXML = doc.appendChild(doc.createElement("QBXML"))
    Set MsgsRq = QBXML.appendChild(doc.createElement("QBXMLMsgsRq"))
    MsgsRq.setAttribute "onError", "continueOnError"
    Set ListDel = MsgsRq.appendChild(doc.createElement("ListDelRq"))
    Set dataElement = ListDel.appendChild(doc.createElement("ListDelType"))
    dataElement.appendChild doc.createTextNode("Customer")
    Set dataElement = ListDel.appendChild(doc.createElement("ListID"))
    dataElement.appendChild doc.createTextNode(ListID)
    CreateCustomerDelRq = doc.xml
    
End Function



' Parse response XML from QuickBooks using MSXML parser
Public Function ParseResponseXML(elementName As String) As Boolean

On Error GoTo ErrHandler
    
    resListID = ""
    
    Dim retStatusCode       As String
    Dim retStatusMessage    As String
    Dim retStatusSeverity   As String
    
    ' Create xmlDoc Obj
    
    ' DOM Document Object
    Dim xmlDoc        As New msxml2.DOMDocument40
    
    ' DOM Node list object for looping through
    Dim objNodeList   As IXMLDOMNodeList
    
    ' Node objects
    Dim objChild            As IXMLDOMNode
    Dim custChildNode       As IXMLDOMNode
    Dim invoiceChildNode    As IXMLDOMNode
    
    ' Attributes Name Mapping
    Dim attrNamedNodeMap As IXMLDOMNamedNodeMap

    Dim i As Integer
    Dim ret As Boolean
    Dim errorMsg As String
    
    errorMsg = ""
    
    ' Load xml doc
    ret = xmlDoc.loadXML(responseXML)
    If Not ret Then
        errorMsg = "loadXML failed, reason: " & xmlDoc.parseError.reason
        GoTo ErrHandler
    End If
    
    ' Get nodes list
    Set objNodeList = xmlDoc.getElementsByTagName(elementName)
     
    ' Loop through each node
    ' Since we have only one request, we should only have one
    ' response.  The loop is actually unnecessary, but it
    ' is a good programming practice
    For i = 0 To (objNodeList.length - 1)
    
        ' Get the CustomerRetRs
        Set attrNamedNodeMap = objNodeList.Item(i).Attributes
        
        ' Get the status Code, info and Severity
        '
        retStatusCode = attrNamedNodeMap.getNamedItem("statusCode").nodeValue
        retStatusSeverity = attrNamedNodeMap.getNamedItem("statusSeverity").nodeValue
        retStatusMessage = attrNamedNodeMap.getNamedItem("statusMessage").nodeValue
                
        ' Check status code to see if there is error or warning
        If retStatusCode <> "0" Then
            ' Checking for Warning is a good practice, although unlikely to happen
            ' on an add request.
            If retStatusSeverity = "Warning" Then
                ' Show the warning, then continue normal processing
                MsgBox retStatusMessage, vbExclamation, "Warning from QuickBooks"
            ElseIf retStatusSeverity = "Error" Then
                MsgBox retStatusMessage, vbExclamation, "Error from QuickBooks"
                ' We only have one response thus we will exit.  If we have multiple
                ' responses, then we may want to continue with the loop.
                ParseResponseXML = False
                Exit Function
            End If
        End If
        
        ' Look at the child nodes
        For Each objChild In objNodeList.Item(i).childNodes
            
            ' Get the CustomerRet block if we were adding a customer
            If objChild.nodeName = "CustomerRet" Then
                
                ' Get the elements in this block
                For Each custChildNode In objChild.childNodes
                    If custChildNode.nodeName = "ListID" Then
                        resListID = custChildNode.Text
                    ElseIf custChildNode.nodeName = "Name" Then
                        resCustName = custChildNode.Text
                    ElseIf custChildNode.nodeName = "FullName" Then
                        resCustFullName = custChildNode.Text
                    End If
                Next
                
            End If ' End of customerRet
            
            
            ' Get the "InvoiceQueryRet" if we were looking for
            ' an in-use customer by querying for invoices -- if
            ' we find one of these, we'll have all the information
            ' we need and can then break from the function.
            If objChild.nodeName = "InvoiceRet" Then
            
                For Each invoiceChildNode In objChild.childNodes
                    If invoiceChildNode.nodeName = "CustomerRef" Then
                        
                        ' Get the elements in this block
                        For Each custChildNode In invoiceChildNode.childNodes
                            If custChildNode.nodeName = "ListID" Then
                                resListID = custChildNode.Text
                            ElseIf custChildNode.nodeName = "Name" Then
                                resCustName = custChildNode.Text
                            ElseIf custChildNode.nodeName = "FullName" Then
                                resCustFullName = custChildNode.Text
                            End If
                        Next
                        
                        ' If we get here, we have all the information we need
                        GoTo BreakPoint
                    
                    End If
                Next
            
            End If ' end of InvoiceQueryRet
        
        
            ' Get the Customer Name if we were deleting a customer
            If elementName = "ListDelRs" And objChild.nodeName = "FullName" Then
                custName = objChild.Text
            End If ' end of ListDelType
        
        
        Next
    Next
    
BreakPoint:
    ParseResponseXML = True
    Exit Function
    
ErrHandler:
    If errorMsg <> "" Then
        MsgBox errorMsg, vbExclamation, "Error"
    Else
        MsgBox Err.Description, vbExclamation, "Error"
    End If
    ParseResponseXML = False
    Exit Function

End Function

' Form Data collection and verification
'
Private Function CollectFormData(operation As String) As Boolean

On Error GoTo ErrHandler

    ' Get data from the form
    custName = Text_CustomerName.Text
    ListID = Text_ListID.Text
    
    ' Customer Name is required for an Add
    If (custName = "") And ("Add" = operation) Then
        MsgBox "Customer Name is empty", vbOKOnly, "Error"
        CollectFormData = False
        GoTo ExitProc
    End If
    
    If (ListID = "") And ("Delete" = operation) Then
        MsgBox "ListID is empty", vbOKOnly, "Error"
        CollectFormData = False
        GoTo ExitProc
    End If
        
    CollectFormData = True
    
ExitProc:
    Exit Function
    
ErrHandler:
    CollectFormData = False
    MsgBox Err.Description, vbExclamation, "Error"
    Exit Function
    
End Function



Private Sub Form_Unload(Cancel As Integer)

    If blnIsOpenConnection Then
        'Call CloseConnection
        CloseConnection
    End If
    
End Sub

