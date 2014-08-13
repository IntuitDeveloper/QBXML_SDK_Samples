VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Main 
   Caption         =   "qbXML Sample: Modify Customer"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   ScaleHeight     =   3435
   ScaleWidth      =   7185
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Comm_Browse 
      Caption         =   "&Browse..."
      Height          =   372
      Left            =   5640
      TabIndex        =   4
      Top             =   1080
      Width           =   1092
   End
   Begin MSComDlg.CommonDialog CommDlg_Browse 
      Left            =   6000
      Top             =   1560
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.TextBox CompanyFileNameInput 
      Height          =   288
      Left            =   480
      TabIndex        =   3
      Text            =   "c:\program files\intuit\quickbooks pro\sample_product-based business.qbw"
      Top             =   1080
      Width           =   4932
   End
   Begin VB.CommandButton Comm_Exit 
      Caption         =   "&Exit"
      Height          =   372
      Left            =   4080
      TabIndex        =   2
      Top             =   1800
      Width           =   1332
   End
   Begin VB.CommandButton Comm_Submit 
      Caption         =   "&Next"
      Height          =   372
      Left            =   1320
      TabIndex        =   1
      Top             =   1800
      Width           =   1332
   End
   Begin VB.Image Image_QBBANNER 
      Height          =   465
      Left            =   360
      Top             =   2640
      Width           =   4920
   End
   Begin VB.Label Label2 
      Caption         =   "Please enter the QB file name in full path:"
      Height          =   252
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   3972
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------
' Form Module: Main
'
' Description:  This sample program demonstrates the following:
'               -   Use of DOMXMLBuilder to build a request
'               -   Parsing the response
'               -   Creation of qbXML query and modify requests
'               -   Sending a request to QuickBooks
'
' Updated to SDK 2.0: 08/05/2002
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'----------------------------------------------------------

'----------------------------------------------------------
' Data Declaration
'
'----------------------------------------------------------
Option Explicit


Const cAppID = "1234"
Const cAppName = "VB Modify Customer Sample App"
    
Dim requestXML  As String
Dim responseXML As String

' Collection object, to store the list of customers
Public customerCollection As Collection
Public companyFile As String
' Make rpWrapper public so it can be used by CustomerMod
' We want to have a single session to QuickBooks
Public rpWrapper As qbXMLRPWrapper



Private Sub Comm_Browse_Click()
    
On Error GoTo ErrHandler
    
    ' Set CancelError is True
    CommDlg_Browse.CancelError = True
  
    ' Set flags
    CommDlg_Browse.Flags = cdlOFNHideReadOnly
    ' Set filters
    CommDlg_Browse.Filter = "All Files (*.*)|*.*|QuickBooks data file" & _
    "(*.qbw)|*.qbw"
    ' Specify default filter
    CommDlg_Browse.FilterIndex = 2
    ' Display the Open dialog box
    CommDlg_Browse.ShowOpen
    ' Display name of selected file
    CompanyFileNameInput.Text = CommDlg_Browse.FileName
    
    Exit Sub
  
ErrHandler:
    'User pressed the Cancel button
    Exit Sub
  
End Sub

' Submit buttom
'
Private Sub Comm_Submit_Click()

    
 On Error GoTo ErrorHandler
 
    ' Get QB data file name
    If Not CollectFormData Then
        Exit Sub
    End If
    
    ' Build the request XML
    BuildXML

    ' Send request to QuickBooks
    Dim success As Boolean
    Dim errNumber As Long
    Dim errMsg As String
    
    Set rpWrapper = New qbXMLRPWrapper
        
    ' Start session with QuickBooks
    success = rpWrapper.Start(cAppID, cAppName, companyFile)
    If Not success Then
        rpWrapper.GetErrorInfo errNumber, errMsg
        Err.Raise Number:=errNumber, Description:=errMsg
    End If
    
    
    'Send request
    success = rpWrapper.DoRequest(requestXML, responseXML)
    If Not success Then
        rpWrapper.GetErrorInfo errNumber, errMsg
        Err.Raise Number:=errNumber, Description:=errMsg
    End If
        
    Set customerCollection = Nothing    ' Reset previous collection
    Set customerCollection = New Collection
    
    If Not ParseResponseXML Then
        Exit Sub
    End If
    
    ' Load CustomerList Form
    Dim CustListForm As New CustomerList
    CustListForm.Show vbModal, Me
    
    ' Close session with QuickBooks
    success = rpWrapper.Finish()
    If Not success Then
        rpWrapper.GetErrorInfo errNumber, errMsg
        Err.Raise Number:=errNumber, Description:=errMsg
    End If
    Set rpWrapper = Nothing
    
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Error"
    Exit Sub
    
End Sub

' exit button
'
Private Sub Comm_Exit_Click()
    ' close windows
    Unload Me
End Sub


' Form Data Collcetion
'
Private Function CollectFormData() As Boolean

    companyFile = CompanyFileNameInput.Text
    
    ' Verify file existence
    If Dir(companyFile) = "" Then
        MsgBox "The company file doesn't exist!" & vbCr _
                    & companyFile, vbExclamation, "Error"
        CollectFormData = False
        Exit Function
    End If

    CollectFormData = True
    Exit Function
    
End Function



Private Sub Form_Load()

On Error GoTo ErrHandler

    Dim appPath
    appPath = App.Path
    Image_QBBANNER.Picture = LoadPicture(appPath & "/qbbanner.bmp")
    
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
    Exit Sub
    
End Sub

'----------------------------------------------------------
'Build XML request
'----------------------------------------------------------
Private Sub BuildXML()

    Dim doc As New DOMDocument
    Dim QBXML As IXMLDOMNode
    Dim MsgsRq As IXMLDOMElement
    Dim CustomerQuery As IXMLDOMElement
    
    Set QBXML = doc.appendChild(doc.createElement("QBXML"))
    Set MsgsRq = QBXML.appendChild(doc.createElement("QBXMLMsgsRq"))
    MsgsRq.setAttribute "onError", "continueOnError"
    Set CustomerQuery = MsgsRq.appendChild(doc.createElement("CustomerQueryRq"))
    
    requestXML = doc.xml
    
End Sub
'
' ParseResponseXML: To Use MSXML to parse response XML
'
Private Function ParseResponseXML() As Boolean

On Error GoTo ErrHandler
    
    ' DOM Document object
    Dim xmlDoc        As New MSXML2.DOMDocument40
    
    ' Node List
    Dim objNodeList   As IXMLDOMNodeList
    Dim objChild      As IXMLDOMNode
    Dim custChildNode As IXMLDOMNode
    
    xmlDoc.async = False
    
    ' Attributes Name Mapping
    Dim attrNamedNodeMap As IXMLDOMNamedNodeMap

    Dim i As Integer
    
    ' QBXML status values
    
    Dim retStatusCode     As String
    Dim retStatusMessage  As String
    Dim retStatusSeverity As String
    
    Dim ret As Boolean
    Dim errorMsg As String
    
On Error GoTo ErrHandler

    ' Load xml document
    ret = xmlDoc.loadXML(responseXML)
    
    If Not ret Then
        errorMsg = "loadXML failed, reason: " & xmlDoc.parseError.reason
        GoTo ErrHandler
    End If
    
    ' Get CustomerQueryRs Node list
    Set objNodeList = xmlDoc.getElementsByTagName("CustomerQueryRs")
     
    ' The following loop is actually unnecessary for this case.
    ' We have only one request, so we should only have one response.
    For i = 0 To (objNodeList.length - 1)
    
        ' Get the CustomerRetRs
        Set attrNamedNodeMap = objNodeList.Item(i).Attributes
        
        retStatusCode = attrNamedNodeMap.getNamedItem("statusCode").nodeValue
        retStatusSeverity = attrNamedNodeMap.getNamedItem("statusSeverity").nodeValue
        retStatusMessage = attrNamedNodeMap.getNamedItem("statusMessage").nodeValue
                
        ' Check status code to see if there is an error
        If retStatusCode <> "0" Then
            ' For a query response, we want to have special check for status code = 1
            If retStatusCode = "1" Then     ' Status code for no record found
                MsgBox "No customer is found", vbInformation, "Message from QuickBooks"
                ParseResponseXML = False
                Exit Function
            Else
                ' Error or warning
                ' retStatusSeverity can be used to distinguish error from warning
                MsgBox retStatusMessage, vbExclamation, "Error from QuickBooks"
                ' We have only one response thus we will exit.  If we have multiple
                ' responses, then we need to continue looping.
                ParseResponseXML = False
                Exit Function
            End If
        End If
        
        ' Walk through the child nodes to get the detail Customer info
        For Each objChild In objNodeList.Item(i).childNodes
            
            ' Get the CustomerRet block
            If objChild.nodeName = "CustomerRet" Then
                
                Dim cust As New Customer
                
                ' Get the elements in this block
                With cust
                For Each custChildNode In objChild.childNodes
                    If custChildNode.nodeName = "ListID" Then
                        .ListID = custChildNode.Text
                    ElseIf custChildNode.nodeName = "FullName" Then
                        .FullName = custChildNode.Text
                    ElseIf custChildNode.nodeName = "TimeCreated" Then
                        .TimeCreated = custChildNode.Text
                    ElseIf custChildNode.nodeName = "TimeModified" Then
                        .TimeModified = custChildNode.Text
                    ElseIf custChildNode.nodeName = "Balance" Then
                        .Balance = custChildNode.Text
                    ElseIf custChildNode.nodeName = "Phone" Then
                        .Phone = custChildNode.Text
                    ElseIf custChildNode.nodeName = "FirstName" Then
                        .FirstName = custChildNode.Text
                    ElseIf custChildNode.nodeName = "LastName" Then
                        .LastName = custChildNode.Text
                    ElseIf custChildNode.nodeName = "Name" Then
                        .Name = custChildNode.Text
                    ElseIf custChildNode.nodeName = "CompanyName" Then
                        .CompanyName = custChildNode.Text
                    ElseIf custChildNode.nodeName = "EditSequence" Then
                        .EditSequence = custChildNode.Text
                    ElseIf custChildNode.nodeName = "Email" Then
                        .Email = custChildNode.Text
                    ElseIf custChildNode.nodeName = "TermsRef_ListID" Then
                        .TermsRef_ListID = custChildNode.Text
                    ElseIf custChildNode.nodeName = "TermsRef_FullName" Then
                        .TermsRef_FullName = custChildNode.Text
                    ElseIf custChildNode.nodeName = "Sublevel" Then
                        .Sublevel = custChildNode.Text
                    End If
                                  
                Next ' end of one CustomerRet
                End With
                
                ' Add customer object to the collection with FullName as the key
                customerCollection.Add cust, cust.FullName
                
                ' Reset the customer object
                Set cust = Nothing
                
            End If
            
        Next
    Next
        
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


Public Function GetRequestXML() As String

    GetRequestXML = requestXML

End Function
Public Function GetResponseXML() As String

    GetResponseXML = responseXML

End Function
