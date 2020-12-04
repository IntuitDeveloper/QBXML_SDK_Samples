VERSION 5.00
Begin VB.Form CustomerMod 
   Caption         =   "qbXML Sample: Modify Customer"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text_Phone 
      Height          =   288
      Left            =   1920
      TabIndex        =   4
      Top             =   3480
      Width           =   1212
   End
   Begin VB.TextBox Text_LastName 
      Height          =   288
      Left            =   1920
      TabIndex        =   3
      Top             =   3120
      Width           =   2412
   End
   Begin VB.TextBox Text_FirstName 
      Height          =   288
      Left            =   1920
      TabIndex        =   2
      Top             =   2760
      Width           =   2412
   End
   Begin VB.CommandButton Comm_Exit 
      Caption         =   "&Done"
      Height          =   372
      Left            =   5640
      TabIndex        =   8
      Top             =   2760
      Width           =   1932
   End
   Begin VB.CommandButton Comm_View_Res 
      Caption         =   "View qbXML  Re&sponse"
      Height          =   372
      Left            =   5640
      TabIndex        =   7
      Top             =   3960
      Width           =   1932
   End
   Begin VB.CommandButton Comm_View_Req 
      Caption         =   "View qbXML Re&quest"
      Height          =   372
      Left            =   5640
      TabIndex        =   6
      Top             =   3360
      Width           =   1932
   End
   Begin VB.CommandButton Comm_Submit 
      Caption         =   "&Modify Customer"
      Height          =   372
      Left            =   5640
      TabIndex        =   5
      Top             =   2160
      Width           =   1932
   End
   Begin VB.TextBox Text_CompanyName 
      Height          =   288
      Left            =   1920
      TabIndex        =   1
      Top             =   3840
      Width           =   2412
   End
   Begin VB.TextBox Text_Name 
      Height          =   288
      Left            =   1920
      TabIndex        =   0
      Top             =   2400
      Width           =   2412
   End
   Begin VB.Frame Frame1 
      Caption         =   " &Customer Info: "
      Height          =   1452
      Left            =   360
      TabIndex        =   9
      Top             =   240
      Width           =   4812
      Begin VB.Label Label10 
         Caption         =   "&Full Name:"
         Height          =   252
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   852
      End
      Begin VB.Label Label2 
         Caption         =   "M&odified Timed:"
         Height          =   252
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   1212
      End
      Begin VB.Label L_ModifiedTime 
         Caption         =   "Mtime"
         Height          =   252
         Left            =   1440
         TabIndex        =   15
         Top             =   1080
         Width           =   2172
      End
      Begin VB.Label L_FullName 
         Caption         =   "Fullname"
         Height          =   252
         Left            =   1440
         TabIndex        =   14
         Top             =   600
         Width           =   2532
      End
      Begin VB.Label L_ListID 
         Caption         =   "ListID"
         Height          =   252
         Left            =   1440
         TabIndex        =   13
         Top             =   360
         Width           =   2412
      End
      Begin VB.Label L_CreatedTime 
         Caption         =   "Ctime"
         Height          =   252
         Left            =   1440
         TabIndex        =   12
         Top             =   840
         Width           =   2172
      End
      Begin VB.Label Label1 
         Caption         =   "C&reated Time:"
         Height          =   252
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   1092
      End
      Begin VB.Label Label3 
         Caption         =   "&List ID:"
         Height          =   252
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   492
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " &You can modify the following fields: "
      Height          =   2292
      Left            =   360
      TabIndex        =   18
      Top             =   2040
      Width           =   4812
      Begin VB.Label Label9 
         Caption         =   "&Phone:"
         Height          =   252
         Left            =   120
         TabIndex        =   23
         Top             =   1440
         Width           =   732
      End
      Begin VB.Label Label7 
         Caption         =   "&Last Name:"
         Height          =   252
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   972
      End
      Begin VB.Label Label8 
         Caption         =   "&Company Name:"
         Height          =   252
         Left            =   120
         TabIndex        =   21
         Top             =   1800
         Width           =   1332
      End
      Begin VB.Label Label6 
         Caption         =   "&First Name:"
         Height          =   252
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   852
      End
      Begin VB.Label Label5 
         Caption         =   "&Name:"
         Height          =   252
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   612
      End
   End
   Begin VB.Image Image_QBBANNER 
      DragMode        =   1  'Automatic
      Height          =   465
      Left            =   360
      Top             =   4800
      Width           =   5640
   End
End
Attribute VB_Name = "CustomerMod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------
' Form Module: CustomerMod
'
' Description:  This form allows the user to modify customer
'               information.
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

Dim requestXML      As String
Dim responseXML     As String

Dim currEditSeq     As String   ' Current EditSequesce
Dim currName        As String   ' Current Customer Name
Dim currFullName    As String   ' Current FullName
Dim currListID      As String   ' Current List ID
Dim currTimeModified As String  ' Current Modified Time


' Existing Customer object
Dim existCustObj As Customer

' User entered data
Dim modifiedCustObj As Customer



Private Sub Comm_Submit_Click()
    
On Error GoTo ErrHandler
    
    requestXML = ""
    responseXML = ""
 
    ' Get data
    If Not CollectFormData Then
        Exit Sub
    End If
    
    ' Build XML
    BuildXML
    
    ' Send request to QuickBooks
    DoRequest
                    
    If Not ParseResponseXML Then
        Exit Sub
    End If
    
    ' Repaint the form
    Me.Caption = "qbXML Sample: Customer " & currName & " Information"
    
    L_FullName.Caption = currFullName
    L_ListID.Caption = currListID
    L_ModifiedTime.Caption = TimeFormat(currTimeModified)
    
    MsgBox "Customer  " & currFullName & _
           " has been successfully modified in QuickBooks."
    
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
    Exit Sub
End Sub

Private Sub Comm_View_Req_Click()

On Error GoTo ErrHandler

    Dim reqFrm As New Display
        
    reqFrm.Text_Content = requestXML
    reqFrm.Caption = "Request XML"
    reqFrm.Show vbModal, Me
    
    Exit Sub

ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
    Exit Sub
    
End Sub

Private Sub Comm_View_Res_Click()

On Error GoTo ErrHandler
        
    Dim resFrm As New Display
    
    If responseXML <> "" Then
        Dim tmpRes As String
        tmpRes = Replace(responseXML, vbLf, vbCrLf)
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


' Form initialization
Private Sub Form_Load()

On Error GoTo ErrHandler

    ' allocate objs
    Set modifiedCustObj = New Customer
    
    With existCustObj
    
        ' Set the default values
        
        ' Label fields
        L_FullName.Caption = .FullName
        L_ListID.Caption = .ListID
        L_CreatedTime.Caption = TimeFormat(.TimeCreated)
        L_ModifiedTime.Caption = TimeFormat(.TimeModified)
    
        ' Text fields
        Text_Phone = .Phone
        Text_LastName = .LastName
        Text_FirstName = .FirstName
        Text_CompanyName = .CompanyName
        Text_Name = .Name
        
        ' Set the modifiedCustObj
        modifiedCustObj.ListID = .ListID
        modifiedCustObj.FullName = .FullName
        modifiedCustObj.Name = .Name
        modifiedCustObj.CompanyName = .CompanyName
        modifiedCustObj.EditSequence = .EditSequence
        
        currEditSeq = .EditSequence
        
    End With
        
    ' Load the banner
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
    Dim CustomerMod As IXMLDOMElement
    
    '
    ' Set up the basic request frame
    '
    Set QBXML = doc.appendChild(doc.createElement("QBXML"))
    Set MsgsRq = QBXML.appendChild(doc.createElement("QBXMLMsgsRq"))
    MsgsRq.setAttribute "onError", "continueOnError"
    '
    ' Note that after we create the CustomerModRq we don't need to reference it other
    ' than to add in the CustomerMod element, so we re-use the CustomerMod var
    '
    Set CustomerMod = MsgsRq.appendChild(doc.createElement("CustomerModRq"))
    Set CustomerMod = CustomerMod.appendChild(doc.createElement("CustomerMod"))
    
    '
    ' Now we just need to stuff in the data we are changing.
    ' Note again that we just re-use the dataelement since everything is getting
    ' appended into the structure in the right place.  Order IS important in
    ' qbXML.
    '
    With modifiedCustObj
        Dim DataElement As IXMLDOMElement
        Set DataElement = CustomerMod.appendChild(doc.createElement("ListID"))
        DataElement.appendChild doc.createTextNode(.ListID)
        Set DataElement = CustomerMod.appendChild(doc.createElement("EditSequence"))
        DataElement.appendChild doc.createTextNode(currEditSeq)
        Set DataElement = CustomerMod.appendChild(doc.createElement("Name"))
        DataElement.appendChild doc.createTextNode(.Name)
        Set DataElement = CustomerMod.appendChild(doc.createElement("CompanyName"))
        DataElement.appendChild doc.createTextNode(.CompanyName)
        Set DataElement = CustomerMod.appendChild(doc.createElement("FirstName"))
        DataElement.appendChild doc.createTextNode(.FirstName)
        Set DataElement = CustomerMod.appendChild(doc.createElement("LastName"))
        DataElement.appendChild doc.createTextNode(.LastName)
        Set DataElement = CustomerMod.appendChild(doc.createElement("Phone"))
        DataElement.appendChild doc.createTextNode(.Phone)
    End With

    requestXML = doc.xml

End Sub

' Send request to QuickBooks
' We use the existing session that was initiated in Main
Private Sub DoRequest()

    Dim success As Boolean
    Dim errNumber As Long
    Dim errMsg As String
    
    success = Main.rpWrapper.DoRequest(requestXML, responseXML)
    If Not success Then
        Main.rpWrapper.GetErrorInfo errNumber, errMsg
        Err.Raise Number:=errNumber, Description:=errMsg
    End If
        
End Sub

' Parse the response XML
Private Function ParseResponseXML() As Boolean

On Error GoTo ErrHandler
    
    Dim xmlDoc As New MSXML2.DOMDocument40
    Dim objNodeList As IXMLDOMNodeList
    Dim objChild    As IXMLDOMNode
    Dim custChildNode As IXMLDOMNode
    
    xmlDoc.async = False
    
    ' Attributes Name Mapping
    Dim attrNamedNodeMap As IXMLDOMNamedNodeMap
    
    ' QBXML values
    Dim retStatusCode As String
    Dim retStatusMessage As String
    Dim retStatusSeverity As String
    
    Dim i As Integer
    Dim ret As Boolean
    Dim errorMsg As String
    
    errorMsg = ""

    ' Load xml Doc
    ret = xmlDoc.loadXML(responseXML)
    
    If Not ret Then
        errorMsg = "loadXML failed, reason: " & xmlDoc.parseError.reason
        GoTo ErrHandler
    End If
    
    Set objNodeList = xmlDoc.getElementsByTagName("CustomerModRs")
         
    ' The following loop is actually unnecessary for this case.
    ' We have only one request, so we should only have one response.
    For i = 0 To (objNodeList.length - 1)
    
        Set attrNamedNodeMap = objNodeList.Item(i).Attributes
        
        retStatusCode = attrNamedNodeMap.getNamedItem("statusCode").nodeValue
        retStatusSeverity = attrNamedNodeMap.getNamedItem("statusSeverity").nodeValue
        retStatusMessage = attrNamedNodeMap.getNamedItem("statusMessage").nodeValue
                
        ' Check status code to see if there is error or warning
        If retStatusCode <> "0" Then
            If retStatusSeverity = "Warning" Then
                ' Show the warning, then continue normal processing
                MsgBox retStatusMessage, vbExclamation, "Warning from QuickBooks"
            ElseIf retStatusSeverity = "Error" Then
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
                
                For Each custChildNode In objChild.childNodes
                    If custChildNode.nodeName = "EditSequence" Then
                        currEditSeq = custChildNode.Text
                    ElseIf custChildNode.nodeName = "FullName" Then
                        currFullName = custChildNode.Text
                    ElseIf custChildNode.nodeName = "TimeModified" Then
                        currTimeModified = custChildNode.Text
                    ElseIf custChildNode.nodeName = "ListID" Then
                        currListID = custChildNode.Text
                    ElseIf custChildNode.nodeName = "Name" Then
                        currName = custChildNode.Text
                    End If
                                  
                Next
                        
            End If
            
        Next ' end of CustomerRet loop
    Next ' end of CustomerModRs loop
    
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


Private Function CollectFormData() As Boolean
    
On Error GoTo ErrHandler

    ' Get data from the form
    With modifiedCustObj
        .Name = Text_Name.Text
        .Phone = Text_Phone.Text
        .LastName = Text_LastName.Text
        .FirstName = Text_FirstName.Text
        .CompanyName = Text_CompanyName.Text
    End With
    
    CollectFormData = True
    Exit Function
    
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
    CollectFormData = False
    Exit Function

End Function

Public Sub Customer(cust As Customer)
    Set existCustObj = cust
End Sub


' Convert qbXML time format to a more a readable format for display purpose
Private Function TimeFormat(inTime As String) As String
    Dim outTime As String
    If inTime = "" Or Len(inTime) < 20 Then
        ' Unknow format
        outTime = ""
    Else
        Dim ymdStr As String
        Dim hmsStr As String
        Dim tmp As String
        
        tmp = Left(inTime, 19)
        ymdStr = Left(tmp, 10)
        hmsStr = Right(tmp, 8)
        
        outTime = ymdStr & " " & hmsStr
    End If
    TimeFormat = outTime
    
End Function

' Clean up
Private Sub Form_Terminate()
    Set modifiedCustObj = Nothing
End Sub

