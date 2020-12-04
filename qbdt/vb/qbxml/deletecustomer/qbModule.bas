Attribute VB_Name = "qbModule"
' qbModule.bas
'
' This module handles the code for communicating with QuickBooks. The code is
' primarily for opening a QuickBooks connection, establishing a session and
' calling the ProcessRequest method with the given qbXML data. The module
' also provides code for ending a QuickBooks session and closing the connection.
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

'Connection info
' qbXMLRP COM object
Public qbXMLRP As New QBXMLRP2Lib.RequestProcessor2
Public blnIsOpenConnection As Boolean

' Ticket
Public ticket As String

' AppID and AppName sent to QuickBooks
Private Const cAppID = "SDK 2.0 Delete Customer Desktop Sample"
Private Const cAppName = "IDN SDK 2.0 Delete Customer Desktop Sample"

' We only have one qbXML request, so let's simply use "1".  For multiple
' requests, it is recommended that unique requestID for each request be used
Public Const cRequestID = "1"

' Request and response strings
Public requestXML  As String
Public responseXML As String


Public Function OpenConnection() As Boolean
    
On Error GoTo ErrHandler
   
    If blnIsOpenConnection Then
        OpenConnection = True
        Exit Function
    End If
   
   
    ' If any of the call to qbXMLRP fails, error will be thrown.
    ' This error must be catched by the caller, in this case
    ' Comm_Submit_Click
    
    ' Open connection to qbXMLRP COM
    qbXMLRP.OpenConnection cAppID, cAppName
        
    ' Begin Session
    ' Pass empty string for the data file name to use the currently
    ' open data file.
    
On Error GoTo ErrHandlerBeginSession
    
    ticket = qbXMLRP.BeginSession("", QBXMLRP2Lib.qbFileOpenSingleUser)
    blnIsOpenConnection = True
    OpenConnection = True
    
    Exit Function
        
ErrHandler:
    blnIsOpenConnection = False
    OpenConnection = False
    MsgBox Err.Description, vbExclamation, "Error"
    Exit Function

ErrHandlerBeginSession:
    blnIsOpenConnection = False
    OpenConnection = False
    qbXMLRP.CloseConnection
    MsgBox Err.Description, vbExclamation, "Error"
    Exit Function


End Function


Public Sub CloseConnection()
' Ends session and closes connection
    
If Not blnIsOpenConnection Then
        Exit Sub
End If

On Error GoTo ErrHandler
    
    ' End the session
    qbXMLRP.EndSession ticket
    
    ' Close the connection
    qbXMLRP.CloseConnection

    blnIsOpenConnection = False

    Exit Sub

ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
    Exit Sub

End Sub

' Send request to QuickBooks
Public Sub DoRequest()
        
    '
    ' Set up an appropriate qbXML prolog depending on the version of QB
    '
    Dim supportedVersion As String
    supportedVersion = qbXMLLatestVersion(qbXMLRP, ticket)
    requestXML = qbXMLAddProlog(supportedVersion, requestXML)
    
    '
    ' Send request
    '
    responseXML = qbXMLRP.ProcessRequest(ticket, requestXML)
                    
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



