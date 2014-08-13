Attribute VB_Name = "modSDKTestPlus3"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' (c) 2003 Intuit Inc. All Rights Reserved                          '
' Use is subject to IP Rights Notice and Restrictions available at: '
' http://developer.quickbooks.com/legal/IPRNotice_021201.html       '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim CompanyFilename As String
Dim InXMLFile As String
Dim OutXMLFile As String
Dim InXMLString As String
Dim OutXMLString As String
Public DisplayFile As String

'Module variables
  Dim booSetDone   As Boolean
  Dim booConnected As Boolean
  Dim booSessionBegun As Boolean
  Dim booIsRemote As Boolean
  
  Dim strTicket As String

'Module objects
  Dim qbXMLCOM As QBXMLRP2Lib.RequestProcessor2

Public Function OpenConnection( _
                   strAppID As String, _
                   strAppName As String, _
                   lblStatusOutput As TextBox) As String

    Dim strStatus As String
  
    If booSetDone = False Then
        Set qbXMLCOM = New QBXMLRP2Lib.RequestProcessor2
        booSetDone = True
    End If
  
    lblStatusOutput.Text = Empty
    lblStatusOutput.ForeColor = &HC00C&
    lblStatusOutput.Text = "Starting to Open Connection"
    lblStatusOutput.Refresh
  
    On Error GoTo ErrorProcessing
    

    If frmSDKTestPlus3.ConnPrefsDirty Then
        booIsRemote = False
        Dim connPref As QBXMLRP2Lib.QBXMLRPConnectionType
        connPref = unknown
        If (frmSDKTestPlus3.ConnLocal.Value) Then
            connPref = localQBD
        ElseIf (frmSDKTestPlus3.ConnLocalUI.Value) Then
            connPref = localQBDLaunchUI
        ElseIf (frmSDKTestPlus3.ConnRemote.Value) Then
            connPref = remoteQBD
            booIsRemote = True
        ElseIf (frmSDKTestPlus3.ConnQBOE.Value) Then
            connPref = remoteQBOE
            booIsRemote = True
        End If
        qbXMLCOM.OpenConnection2 strAppID, strAppName, connPref
    Else
        qbXMLCOM.OpenConnection strAppID, strAppName
    End If
  
    OpenConnection = Empty
    booConnected = True
    Exit Function
  
ErrorProcessing:
    lblStatusOutput.Text = Empty
    lblStatusOutput.ForeColor = &HF0&
    lblStatusOutput.Text = Err.Number & "  " & Err.Description
    lblStatusOutput.Refresh
    OpenConnection = Err.Description
End Function

Public Function BeginSession( _
                   strCompanyFilename As String, _
                   strFileMode As String, _
                   lblStatusOutput As TextBox) As String

    On Error GoTo ErrSetPrefs
    Dim authFlags As Long
    
    ' remoteQBD and remoteQBOE don't allow AuthFlags setting
    If (frmSDKTestPlus3.AuthPrefsDirty And Not booIsRemote) Then
        lblStatusOutput.Text = Empty
        lblStatusOutput.ForeColor = &HC00C&
        lblStatusOutput.Text = "Starting to Set Auth Preferences"
        lblStatusOutput.Refresh
        Dim prefs As QBXMLRP2Lib.AuthPreferences
        Set prefs = qbXMLCOM.AuthPreferences
        If (frmSDKTestPlus3.Unattended.Value) Then
            prefs.PutUnattendedModePref umpRequired
        Else
            prefs.PutUnattendedModePref umpOptional
        End If
        
        authFlags = 0
        If (frmSDKTestPlus3.qbEnterprise.Value = 1) Then
            authFlags = authFlags Or &H8&
        End If
        
        If (frmSDKTestPlus3.qbPremier.Value = 1) Then
            authFlags = authFlags Or &H4&
        End If
        
        If (frmSDKTestPlus3.qbPro.Value = 1) Then
            authFlags = authFlags Or &H2&
        End If
        
        If (frmSDKTestPlus3.qbSimple.Value = 1) Then
            authFlags = authFlags Or &H1&
        End If
        
        If (frmSDKTestPlus3.qbForceAuthDialog.Value = 1) Then
            authFlags = authFlags Or &H80000000
        End If
        
        prefs.PutAuthFlags (authFlags)
        
        prefs.PutIsReadOnly frmSDKTestPlus3.ReadOnly.Value
        
        If (frmSDKTestPlus3.pdRequired.Value) Then
            prefs.PutPersonalDataPref pdpRequired
        ElseIf (frmSDKTestPlus3.pdNotNeeded.Value) Then
            prefs.PutPersonalDataPref pdpNotNeeded
        Else
            prefs.PutPersonalDataPref pdpOptional
        End If
    End If
    
    On Error GoTo ErrBeginSession
    lblStatusOutput.Text = Empty
    lblStatusOutput.ForeColor = &HC00C&
    lblStatusOutput.Text = "Starting to Begin Session"
    lblStatusOutput.Refresh
  
    Dim openMode As QBXMLRP2Lib.QBFileMode
    If strFileMode = "Single" Then
        openMode = qbFileOpenSingleUser
    ElseIf strFileMode = "Multi" Then
        openMode = qbFileOpenMultiUser
    Else
        openMode = qbFileOpenDoNotCare
    End If
    
    strTicket = qbXMLCOM.BeginSession(strCompanyFilename, openMode)
  
    On Error GoTo ErrGetFileName
  
    If strCompanyFilename = Empty Then
        strCompanyFilename = qbXMLCOM.GetCurrentCompanyFileName(strTicket)
    End If

    lblStatusOutput.Text = Empty
    lblStatusOutput.ForeColor = &HC00C&
    lblStatusOutput.Text = "Session begun with " & strCompanyFilename
    lblStatusOutput.Refresh
    booSessionBegun = True
    BeginSession = Empty
    Exit Function
  
ErrSetPrefs:
    lblStatusOutput.Text = Empty
    lblStatusOutput.ForeColor = &HF0&
    lblStatusOutput.Text = Err.Number & " " & Err.Description
    lblStatusOutput.Refresh
    BeginSession = Err.Description
    Exit Function
    
ErrBeginSession:
    lblStatusOutput.Text = Empty
    lblStatusOutput.ForeColor = &HF0&
    lblStatusOutput.Text = Err.Number & "  " & Err.Description
    lblStatusOutput.Refresh
    BeginSession = Err.Description
    Exit Function

ErrGetFileName:
    lblStatusOutput.Text = Empty
    lblStatusOutput.ForeColor = &HF0&
    lblStatusOutput.Text = Err.Number & "  " & Err.Description
    lblStatusOutput.Refresh
    BeginSession = Err.Description
    Exit Function

End Function

Public Sub EndSession(lblStatusOutput As TextBox)

  On Error GoTo ErrorDisplay
  
  lblStatusOutput.Text = Empty
  lblStatusOutput.ForeColor = &HC00C&
  lblStatusOutput.Text = "Starting to End Session"
  lblStatusOutput.Refresh
  
  If booSessionBegun Then
    qbXMLCOM.EndSession (strTicket)
    booSessionBegun = False
  End If
  
  lblStatusOutput.Text = Empty
  lblStatusOutput.ForeColor = &HC00C&
  lblStatusOutput.Text = "Session Ended"
  lblStatusOutput.Refresh
  Exit Sub
  
ErrorDisplay:
  MsgBox Err.Description
End Sub

Public Sub CloseConnection(lblStatusOutput As TextBox)

  On Error GoTo ErrorDisplay
  
  lblStatusOutput.Text = Empty
  lblStatusOutput.ForeColor = &HC00C&
  lblStatusOutput.Text = "Starting to Close Connection"
  lblStatusOutput.Refresh
  
  If booConnected Then
    qbXMLCOM.CloseConnection
    booConnected = False
  End If
  
  lblStatusOutput.Text = Empty
  lblStatusOutput.ForeColor = &HC00C&
  lblStatusOutput.Text = "Not Connected"
  lblStatusOutput.Refresh
  
  Exit Sub
  
ErrorDisplay:
  MsgBox Err.Description
End Sub


'----------------------------------------------------------
'
' Routine: PrettyPrintXMLToFile
'
' Description
'
' Takes an XML message set string and a file name as input.
' Creates a new copy of the file and pretty prints the XML
' message set into the file.
'
'----------------------------------------------------------
Public Sub PrettyPrintXMLToFile(XMLString As String, _
                                 XMLFile As String)
                                       
  Dim SplitXMLString() As String
  Dim IndentString As String
  Dim XMLStringLength As Long
  Dim SplitIndex
  
  IndentString = Empty
  
  FileNum = FreeFile
  Open XMLFile For Output As FileNum
  
'Remove the linefeeds from the XML output string
  XMLString = Replace(XMLString, vbLf, vbNullString)
  
  SplitXMLString = Split(XMLString, "<")
  
'We're expecting the first character of the XML output to be "<"
'which result in an empty first array element, so skip it.
  SplitIndex = LBound(SplitXMLString) + 1
  
  Do
    If Left(SplitXMLString(SplitIndex), 1) = "/" Then
      IndentString = Left(IndentString, Len(IndentString) - 3)
      Print #FileNum, IndentString & "<" & _
                      SplitXMLString(SplitIndex)
      SplitIndex = SplitIndex + 1
    ElseIf Left(SplitXMLString(SplitIndex + 1), 1) = "/" Then
      If InStr(1, Left(SplitXMLString(SplitIndex), _
               InStr(1, SplitXMLString(SplitIndex), ">")), _
                " ") > 0 Then
        Print #FileNum, IndentString & "<" & _
                        SplitXMLString(SplitIndex)
        SplitIndex = SplitIndex + 1
      Else
        Print #FileNum, IndentString & "<" & _
                        SplitXMLString(SplitIndex) & "<" & _
                        SplitXMLString(SplitIndex + 1)
        SplitIndex = SplitIndex + 2
      End If
    Else
      Print #FileNum, IndentString & "<" & _
                      SplitXMLString(SplitIndex)
      If SplitXMLString(SplitIndex) <> "?xml version=""1.0"" ?>" And _
         SplitXMLString(SplitIndex) <> "!DOCTYPE QBXML PUBLIC '-//INTUIT//DTD QBXML QBD 1.0//EN' >" And _
         Right(SplitXMLString(SplitIndex), 2) <> "/>" Then
        IndentString = IndentString & "   "
      End If
      SplitIndex = SplitIndex + 1
    End If
  Loop Until SplitIndex >= UBound(SplitXMLString)
  
  If Left(SplitXMLString(UBound(SplitXMLString)), 1) = "/" Then
    IndentString = Left(IndentString, Len(IndentString) - 3)
  End If
  
  Print #FileNum, IndentString & "<" & _
                  SplitXMLString(UBound(SplitXMLString))
  
  Close FileNum
End Sub

Private Sub XMLFileToXMLString(XMLFileName As String, _
                               XMLString As String)
  'Clear out the XML String
  XMLString = Empty
  
  'Open the file, read each line and concatonate each to
  'the XML String
  FileNum = FreeFile
  Open XMLFileName For Input As #FileNum
  Do Until EOF(FileNum)
    Line Input #FileNum, strInput
    XMLString = XMLString & strInput
  Loop
  Close #FileNum
End Sub

Public Function SendXMLFile(strProcessor As String, _
                            strInXMLFile As String, _
                            strOutXMLFile As String, _
                            lblStatus As TextBox) As String

    Dim strInXML As String
    Dim strOutXML As String

    XMLFileToXMLString strInXMLFile, strInXML
  
    On Error GoTo ErrorDisplay
  
    lblStatus.ForeColor = &HE0F0&
    lblStatus.Text = "Your request is being processed"
    lblStatus.Refresh
  
    'Process the XML
    If strProcessor = "RequestProcessor" Then
        strOutXML = qbXMLCOM.ProcessRequest(strTicket, strInXML)
    Else
        strOutXML = qbXMLCOM.ProcessSubscription(strInXML)
    End If

    PrettyPrintXMLToFile strOutXML, strOutXMLFile
  
    lblStatus.Text = Empty
    lblStatus.ForeColor = &HC00C&
    lblStatus.Text = "Processing complete"
    lblStatus.Refresh
    SendXMLFile = Empty
    Exit Function
  
ErrorDisplay:
'   MsgBox Err.Description
    lblStatus.Text = Empty
    lblStatus.ForeColor = &HF0&
    lblStatus.Text = Err.Number & "  " & Err.Description
    lblStatus.Refresh
    SendXMLFile = Err.Description
End Function
