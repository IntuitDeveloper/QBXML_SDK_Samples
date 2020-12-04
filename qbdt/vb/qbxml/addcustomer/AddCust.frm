VERSION 5.00
Begin VB.Form AddCust 
   Caption         =   "qbXML Sample: Create a New Customer"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   7545
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Comm_Exit 
      Caption         =   "&Done"
      Height          =   372
      Left            =   5040
      TabIndex        =   8
      Top             =   1200
      Width           =   2172
   End
   Begin VB.CommandButton Comm_View_Res 
      Caption         =   "Vi&ew qbXML Response"
      Height          =   372
      Left            =   5040
      TabIndex        =   7
      Top             =   2760
      Width           =   2172
   End
   Begin VB.CommandButton Comm_View_Req 
      Caption         =   "&View qbXML Request"
      Height          =   372
      Left            =   5040
      TabIndex        =   6
      Top             =   2280
      Width           =   2172
   End
   Begin VB.TextBox Text_CustomerName 
      Height          =   288
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   2532
   End
   Begin VB.TextBox Text_Phone 
      Height          =   288
      Left            =   1680
      TabIndex        =   4
      Top             =   3000
      Width           =   1332
   End
   Begin VB.CommandButton Comm_Submit 
      Caption         =   "&Add Customer"
      Default         =   -1  'True
      Height          =   372
      Left            =   5040
      TabIndex        =   5
      Top             =   720
      Width           =   2172
   End
   Begin VB.TextBox Text_LastName 
      Height          =   288
      Left            =   1680
      TabIndex        =   3
      Top             =   2520
      Width           =   1932
   End
   Begin VB.TextBox Text_FirstName 
      Height          =   288
      Left            =   1680
      TabIndex        =   2
      Top             =   2040
      Width           =   1932
   End
   Begin VB.Frame Frame1 
      Caption         =   " Contact Information (optional) "
      Height          =   2172
      Left            =   480
      TabIndex        =   9
      Top             =   1680
      Width           =   3972
      Begin VB.Label Label3 
         Caption         =   "&Phone:"
         Height          =   252
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   852
      End
      Begin VB.Label Label2 
         Caption         =   "&Last Name:"
         Height          =   252
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   972
      End
      Begin VB.Label Label1 
         Caption         =   "&First Name:"
         Height          =   252
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   972
      End
   End
   Begin VB.Image Image_QBBANNER 
      Height          =   435
      Left            =   600
      Picture         =   "AddCust.frx":0000
      Top             =   4200
      Width           =   6450
   End
   Begin VB.Image Image_QBCUST 
      Height          =   495
      Left            =   3360
      Picture         =   "AddCust.frx":3532
      Top             =   960
      Width           =   450
   End
   Begin VB.Label Label7 
      Caption         =   "&Customer to add:"
      Height          =   252
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   1332
   End
End
Attribute VB_Name = "AddCust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------
' Form Module: AddCust
'
' Description:  This sample application demonstrates how to
'               construct a qbXML request, send it to QuickBooks,
'               and parse the qbXML response.  It prompts the
'               user for customer information and adds a new
'               customer to QuickBooks.
'
'               QuickBooks must be running with a data file open.
'               The current data file is used.
'
' Created On: 11/05/2001
' Updated to SDK 2.0 On: 07/30/2002
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'----------------------------------------------------------

Option Explicit

' AppID and AppName sent to QuickBooks
Const cAppID = "123"
Const cAppName = "IDN VB AddCustomer Sample"

' We only have one request, thus we simply use "1".  For multiple
' requests, it is recommended you use unique requestID for each
' request
Const cRequestID = "1"

' Request and response strings
Dim requestXML  As String
Dim responseXML As String

' Customer info
Dim firstName   As String
Dim lastName    As String
Dim phoneNumber As String
Dim custName    As String

' Response info
Dim resCustName         As String
Dim resCustFullName     As String
Dim resListID           As String


' Submit button
Private Sub Comm_Submit_Click()

On Error GoTo ErrHandler

    ' Initialize
    requestXML = ""
    responseXML = ""

    ' Get input data
    If Not CollectFormData Then
        Exit Sub
    End If

    '
    ' Set up to communicate with QuickBooks
    '
    ' Create qbXMLRP2 COM object
    '
    Dim qbXMLRP As New QBXMLRP2Lib.RequestProcessor2
    
    Dim ticket  As String
    
    '
    ' Open connection to QuickBooks
    '
    qbXMLRP.OpenConnection cAppID, cAppName
    '
    ' Begin Session
    ' Pass empty string for the data file name to use the currently
    ' open data file.
    '
    ticket = qbXMLRP.BeginSession("", QBXMLRP2Lib.qbFileOpenDoNotCare)
           
    '
    ' Build the request XML
    '
    ' XMLBuilder
    Dim builder As New DOMDocument40
    
    Dim QBXML As IXMLDOMNode
    Set QBXML = builder.createElement("QBXML")
    builder.appendChild QBXML
    Dim msgsRq As IXMLDOMElement
    Set msgsRq = QBXML.appendChild(builder.createElement("QBXMLMsgsRq"))
    msgsRq.setAttribute "onError", "continueOnError"
    Dim CustomerAddRq As IXMLDOMElement
    Dim CustomerAdd As IXMLDOMElement
    Set CustomerAddRq = msgsRq.appendChild(builder.createElement("CustomerAddRq"))
    Set CustomerAdd = CustomerAddRq.appendChild(builder.createElement("CustomerAdd"))
    Dim dataElement As IXMLDOMElement
    Set dataElement = CustomerAdd.appendChild(builder.createElement("Name"))
    dataElement.appendChild builder.createTextNode(custName)
    If firstName <> "" Then
        Set dataElement = CustomerAdd.appendChild(builder.createElement("FirstName"))
        dataElement.appendChild builder.createTextNode(firstName)
    End If
    If lastName <> "" Then
        Set dataElement = CustomerAdd.appendChild(builder.createElement("LastName"))
        dataElement.appendChild builder.createTextNode(lastName)
    End If
    If phoneNumber <> "" Then
        Set dataElement = CustomerAdd.appendChild(builder.createElement("Phone"))
        dataElement.appendChild builder.createTextNode(phoneNumber)
    End If

    Dim supportedVersion As String
    supportedVersion = qbXMLLatestVersion(qbXMLRP, ticket)
    requestXML = qbXMLAddProlog(supportedVersion, builder.xml)
    
    '
    ' Send request to QuickBooks
    '
    responseXML = qbXMLRP.ProcessRequest(ticket, requestXML)
    
    
    '
    ' Done communicating with QuickBooks, so close up the connection
    '
    ' End the session
    '
    qbXMLRP.EndSession ticket
    '
    ' Close the connection
    '
    qbXMLRP.CloseConnection
    
    '
    ' Now we'll parse the response from QuickBooks
    '
    If Not ParseResponseXML Then
        Exit Sub
    End If
                
    '
    ' Display the result
    '
    MsgBox "Customer " & custName & " has been sucessfully created" & vbCr & _
                "ListID = " & resListID & vbCr & _
                "FullName = " & resCustFullName, vbOKOnly, "Success"
             
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
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
'
Private Sub Comm_View_Res_Click()
 
 On Error GoTo ErrHandler
 
    ' Instantiate a RequestXML Obj
    '
        
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
'
Private Sub Comm_Exit_Click()
    Unload Me ' close the window
End Sub



' Parse response XML from QuickBooks using MSXML parser
Private Function ParseResponseXML() As Boolean

On Error GoTo ErrHandler
    
    Dim retStatusCode       As String
    Dim retStatusMessage    As String
    Dim retStatusSeverity   As String
    
    ' Create xmlDoc Obj
    
    ' DOM Document Object
    Dim xmlDoc        As New DOMDocument40
    
    ' DOM Node list object for looping through
    Dim objNodeList   As IXMLDOMNodeList
    
    ' Node objects
    Dim objChild      As IXMLDOMNode
    Dim custChildNode As IXMLDOMNode
    
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
    
    ' Get CustomerAddRs nodes list
    Set objNodeList = xmlDoc.getElementsByTagName("CustomerAddRs")
     
    ' Loop through each CustomerAddRs node
    ' Since we have only one request, we should only have one
    ' CustomerAddRs.  The loop is actually unnecessary, but it
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
        
        ' Walk through the child nodes of CustomerAddRs node, to get
        ' the detail info
        ' This loop is not necessary actually, since we only have one
        ' CustomerRet.
        For Each objChild In objNodeList.Item(i).childNodes
            
            ' Get the CustomerRet block
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
            
        Next ' End of customerAddret
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

' Form Data Collection
'
Private Function CollectFormData() As Boolean

On Error GoTo ErrHandler

    ' Get data from the form
    firstName = Text_FirstName.Text
    lastName = Text_LastName.Text
    phoneNumber = Text_Phone.Text
    custName = Text_CustomerName.Text
    
    ' Customer Name is required
    If custName = "" Then
        MsgBox "Customer Name is empty", vbOKOnly, "Error"
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

Function qbXMLLatestVersion(rp As RequestProcessor2, ticket As String) As String
    Dim strXMLVersions() As String
    'Should be able to use this, but there appears to be a bug that may cause 2.0 to be returned
    'when it should not.
    'strXMLVersions = rp.QBXMLVersionsForSession(ticket)
    
    'Instead we use HostQuery
    'Create a DOM document object for creating our request.
    Dim xml As New DOMDocument

    'Create the QBXML aggregate
    Dim rootElement As IXMLDOMNode
    Set rootElement = xml.createElement("QBXML")
    xml.appendChild rootElement
  
    'Add the QBXMLMsgsRq aggregate to the QBXML aggregate
    Dim QBXMLMsgsRqNode As IXMLDOMNode
    Set QBXMLMsgsRqNode = xml.createElement("QBXMLMsgsRq")
    rootElement.appendChild QBXMLMsgsRqNode

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'If we were writing a real application this is where we would add
    'a newMessageSetID so we could perform error recovery.  Any time a
    'request contains an add, delete, modify or void request developers
    'should use the error recovery mechanisms.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Set the QBXMLMsgsRq onError attribute to continueOnError
    Dim onErrorAttr As IXMLDOMAttribute
    Set onErrorAttr = xml.createAttribute("onError")
    onErrorAttr.Text = "stopOnError"
    QBXMLMsgsRqNode.Attributes.setNamedItem onErrorAttr
  
    'Add the InvoiceAddRq aggregate to QBXMLMsgsRq aggregate
    Dim HostQuery As IXMLDOMNode
    Set HostQuery = xml.createElement("HostQueryRq")
    QBXMLMsgsRqNode.appendChild HostQuery
    
    Dim strXMLRequest As String
    strXMLRequest = _
        "<?xml version=""1.0"" ?>" & _
        "<!DOCTYPE QBXML PUBLIC '-//INTUIT//DTD QBXML QBD 1.0//EN' 'http://developer.intuit.com'>" _
        & rootElement.xml

    Dim strXMLResponse As String
    strXMLResponse = rp.ProcessRequest(ticket, strXMLRequest)
    Dim QueryResponse As New DOMDocument

    'Parse the response XML
    QueryResponse.async = False
    QueryResponse.loadXML (strXMLResponse)

    Dim supportedVersions As IXMLDOMNodeList
    Set supportedVersions = QueryResponse.getElementsByTagName("SupportedQBXMLVersion")
    
    Dim VersNode As IXMLDOMNode
    
    Dim i As Long
    Dim vers As Double
    Dim LastVers As Double
    LastVers = 0
    For i = 0 To supportedVersions.length - 1
        Set VersNode = supportedVersions.Item(i)
        vers = VersNode.firstChild.Text
        If (vers > LastVers) Then
            LastVers = vers
            qbXMLLatestVersion = VersNode.firstChild.Text
        End If
    Next i
End Function

Function qbXMLAddProlog(supportedVersion As String, xml As String) As String
    Dim qbXMLVersionSpec As String
    If (Val(supportedVersion) >= 2) Then
        qbXMLVersionSpec = "<?qbxml version=""" & supportedVersion & """?>"
    ElseIf (supportedVersion = "1.1") Then
        qbXMLVersionSpec = "<!DOCTYPE QBXML PUBLIC '-//INTUIT//DTD QBXML QBD " _
                           & supportedVersion & "//EN' 'http://developer.intuit.com'>"
    Else
        MsgBox "You are apparently running QuickBooks 2002 Release 1, we strongly recommend that you use QuickBooks' online update feature to obtain the latest fixes and enhancements", vbExclamation
        qbXMLVersionSpec = "<!DOCTYPE QBXML PUBLIC '-//INTUIT//DTD QBXML QBD " _
                           & supportedVersion & "//EN' 'http://developer.intuit.com'>"
    End If
    qbXMLAddProlog = "<?xml version=""1.0""?>" & vbCrLf & qbXMLVersionSpec & xml
End Function

