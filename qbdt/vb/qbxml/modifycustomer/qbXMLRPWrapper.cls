VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "qbXMLRPWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------------------
' Class Module: qbXMLRPWrapper
'
' Description:  Encapsulates the calls to the Request Processor which
'               is the interface to communicate with QuickBooks.
'               The proper calling sequence is
'                   Start
'                   DoRequest 1
'                   DoRequest 2
'                   ...
'                   DoRequest n
'                   Finish
'               Each method call returns a status that must be checked
'               to determine whether the call was successful or not.
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

Private m_ErrNumber As Long
Private m_ErrMsg As String

Private m_RP As QBXMLRP2Lib.RequestProcessor2
Private m_Ticket As String

' Initiate connection to QuickBooks (qbXMLRP)
' Return true if successful, false otherwise
Public Function Start(appID As String, appName As String, _
                    companyFile As String) As Boolean

On Error GoTo ErrHandler

    Set m_RP = New QBXMLRP2Lib.RequestProcessor2
    
    ' Open connection to qbXMLRp COM
    m_RP.OpenConnection appID, appName
    
    ' Begin Session
    m_Ticket = m_RP.BeginSession(companyFile, QBXMLRP2Lib.qbFileOpenDoNotCare)
    
    Start = True
    Exit Function

ErrHandler:
    Start = False
    m_ErrNumber = Err.Number
    m_ErrMsg = Err.Description
    Exit Function

End Function

' Send qbXML request.  The response is returned in responseXML.
' Return true if successful, false otherwise
Public Function DoRequest(requestXML As String, responseXML As String) As Boolean

On Error GoTo ErrHandler

    '
    ' Determine the maximum SDK version supported and use it
    '
    Dim supportedVersion As String
    supportedVersion = qbXMLLatestVersion(m_RP, m_Ticket)
    requestXML = qbXMLAddProlog(supportedVersion, requestXML)
    
    ' Get the responseXML.  Will throw error if ticket is invalid
    responseXML = m_RP.ProcessRequest(m_Ticket, requestXML)
                
    DoRequest = True
    Exit Function

ErrHandler:
    DoRequest = False
    m_ErrNumber = Err.Number
    m_ErrMsg = Err.Description
    Exit Function

End Function


' Close connection to QuickBooks
' Return true if successful, false otherwise
Public Function Finish() As Boolean

On Error GoTo ErrHandler

    If m_Ticket <> "" Then
        ' End the session
        m_RP.EndSession m_Ticket
    
        ' Close the connection
        m_RP.CloseConnection
    End If
    
    m_Ticket = ""
    Set m_RP = Nothing

    Finish = True
    Exit Function

ErrHandler:
    Finish = False
    m_ErrNumber = Err.Number
    m_ErrMsg = Err.Description
    Exit Function

End Function

' Return detail error information
Public Sub GetErrorInfo(errNumber As Long, errMsg As String)
    errNumber = m_ErrNumber
    errMsg = m_ErrMsg
End Sub

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


