Option Strict Off
Option Explicit On
Module modSDKTestPlus3
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
	Dim booSetDone As Boolean
	Dim booConnected As Boolean
	Dim booSessionBegun As Boolean
	Dim booIsRemote As Boolean
	
	Dim strTicket As String

    'Module objects
    Dim qbXMLCOM As QBXMLRP2Lib.RequestProcessor2

    Public Function OpenConnection(ByRef strAppID As String, ByRef strAppName As String, ByRef lblStatusOutput As System.Windows.Forms.TextBox) As String
		
		Dim strStatus As String
		
		If booSetDone = False Then
			qbXMLCOM = New QBXMLRP2Lib.RequestProcessor2
			booSetDone = True
		End If
		
		lblStatusOutput.Text = CStr(Nothing)
		lblStatusOutput.ForeColor = System.Drawing.ColorTranslator.FromOle(&HC00C)
		lblStatusOutput.Text = "Starting to Open Connection"
		lblStatusOutput.Refresh()
		
		On Error GoTo ErrorProcessing
		
		
		Dim connPref As QBXMLRP2Lib.QBXMLRPConnectionType
		If frmSDKTestPlus3.ConnPrefsDirty Then
			booIsRemote = False
			connPref = QBXMLRP2Lib.QBXMLRPConnectionType.unknown
			If (frmSDKTestPlus3.ConnLocal.Checked) Then
				connPref = QBXMLRP2Lib.QBXMLRPConnectionType.localQBD
			ElseIf (frmSDKTestPlus3.ConnLocalUI.Checked) Then 
				connPref = QBXMLRP2Lib.QBXMLRPConnectionType.localQBDLaunchUI
			ElseIf (frmSDKTestPlus3.ConnRemote.Checked) Then 
				connPref = QBXMLRP2Lib.QBXMLRPConnectionType.remoteQBD
				booIsRemote = True
			ElseIf (frmSDKTestPlus3.ConnQBOE.Checked) Then 
				connPref = QBXMLRP2Lib.QBXMLRPConnectionType.remoteQBOE
				booIsRemote = True
			End If
			qbXMLCOM.OpenConnection2(strAppID, strAppName, connPref)
		Else
			qbXMLCOM.OpenConnection(strAppID, strAppName)
		End If
		
		OpenConnection = CStr(Nothing)
		booConnected = True
		Exit Function
		
ErrorProcessing: 
		lblStatusOutput.Text = CStr(Nothing)
		lblStatusOutput.ForeColor = System.Drawing.ColorTranslator.FromOle(&HF0)
		lblStatusOutput.Text = Err.Number & "  " & Err.Description
		lblStatusOutput.Refresh()
		OpenConnection = Err.Description
	End Function
	
	Public Function BeginSession(ByRef strCompanyFilename As String, ByRef strFileMode As String, ByRef lblStatusOutput As System.Windows.Forms.TextBox) As String
		
		On Error GoTo ErrSetPrefs
        Dim authFlags As Integer

        ' remoteQBD and remoteQBOE don't allow AuthFlags setting
        Dim prefs As QBXMLRP2Lib.AuthPreferences
		If (frmSDKTestPlus3.AuthPrefsDirty And Not booIsRemote) Then
			lblStatusOutput.Text = CStr(Nothing)
			lblStatusOutput.ForeColor = System.Drawing.ColorTranslator.FromOle(&HC00C)
			lblStatusOutput.Text = "Starting to Set Auth Preferences"
			lblStatusOutput.Refresh()
			prefs = qbXMLCOM.AuthPreferences
			If (frmSDKTestPlus3.Unattended.CheckState) Then
				prefs.PutUnattendedModePref(QBXMLRP2Lib.QBXMLRPUnattendedModePrefType.umpRequired)
			Else
				prefs.PutUnattendedModePref(QBXMLRP2Lib.QBXMLRPUnattendedModePrefType.umpOptional)
			End If
			
			authFlags = 0
			If (frmSDKTestPlus3.qbEnterprise.CheckState = 1) Then
				authFlags = authFlags Or &H8
			End If
			
			If (frmSDKTestPlus3.qbPremier.CheckState = 1) Then
				authFlags = authFlags Or &H4
			End If
			
			If (frmSDKTestPlus3.qbPro.CheckState = 1) Then
				authFlags = authFlags Or &H2
			End If
			
			If (frmSDKTestPlus3.qbSimple.CheckState = 1) Then
				authFlags = authFlags Or &H1
			End If
			
			If (frmSDKTestPlus3.qbForceAuthDialog.CheckState = 1) Then
				authFlags = authFlags Or &H80000000
			End If
			
			prefs.PutAuthFlags((authFlags))
			
			prefs.PutIsReadOnly(frmSDKTestPlus3.ReadOnly_Renamed.CheckState)
			
			If (frmSDKTestPlus3.pdRequired.Checked) Then
				prefs.PutPersonalDataPref(QBXMLRP2Lib.QBXMLRPPersonalDataPrefType.pdpRequired)
			ElseIf (frmSDKTestPlus3.pdNotNeeded.Checked) Then 
				prefs.PutPersonalDataPref(QBXMLRP2Lib.QBXMLRPPersonalDataPrefType.pdpNotNeeded)
			Else
				prefs.PutPersonalDataPref(QBXMLRP2Lib.QBXMLRPPersonalDataPrefType.pdpOptional)
			End If
		End If
		
		On Error GoTo ErrBeginSession
		lblStatusOutput.Text = CStr(Nothing)
		lblStatusOutput.ForeColor = System.Drawing.ColorTranslator.FromOle(&HC00C)
		lblStatusOutput.Text = "Starting to Begin Session"
		lblStatusOutput.Refresh()
		
		Dim openMode As QBXMLRP2Lib.QBFileMode
		If strFileMode = "Single" Then
			openMode = QBXMLRP2Lib.QBFileMode.qbFileOpenSingleUser
		ElseIf strFileMode = "Multi" Then 
			openMode = QBXMLRP2Lib.QBFileMode.qbFileOpenMultiUser
		Else
			openMode = QBXMLRP2Lib.QBFileMode.qbFileOpenDoNotCare
		End If
		
		strTicket = qbXMLCOM.BeginSession(strCompanyFilename, openMode)
		
		On Error GoTo ErrGetFileName


        If String.IsNullOrEmpty(strCompanyFilename) Then
            strCompanyFilename = qbXMLCOM.GetCurrentCompanyFileName(strTicket)
        End If

        lblStatusOutput.Text = CStr(Nothing)
		lblStatusOutput.ForeColor = System.Drawing.ColorTranslator.FromOle(&HC00C)
		lblStatusOutput.Text = "Session begun with " & strCompanyFilename
		lblStatusOutput.Refresh()
		booSessionBegun = True
		BeginSession = CStr(Nothing)
		Exit Function
		
ErrSetPrefs: 
		lblStatusOutput.Text = CStr(Nothing)
		lblStatusOutput.ForeColor = System.Drawing.ColorTranslator.FromOle(&HF0)
		lblStatusOutput.Text = Err.Number & " " & Err.Description
		lblStatusOutput.Refresh()
		BeginSession = Err.Description
		Exit Function
		
ErrBeginSession: 
		lblStatusOutput.Text = CStr(Nothing)
		lblStatusOutput.ForeColor = System.Drawing.ColorTranslator.FromOle(&HF0)
		lblStatusOutput.Text = Err.Number & "  " & Err.Description
		lblStatusOutput.Refresh()
		BeginSession = Err.Description
		Exit Function
		
ErrGetFileName: 
		lblStatusOutput.Text = CStr(Nothing)
		lblStatusOutput.ForeColor = System.Drawing.ColorTranslator.FromOle(&HF0)
		lblStatusOutput.Text = Err.Number & "  " & Err.Description
		lblStatusOutput.Refresh()
		BeginSession = Err.Description
		Exit Function
		
	End Function
	
	Public Sub EndSession(ByRef lblStatusOutput As System.Windows.Forms.TextBox)
		
		On Error GoTo ErrorDisplay
		
		lblStatusOutput.Text = CStr(Nothing)
		lblStatusOutput.ForeColor = System.Drawing.ColorTranslator.FromOle(&HC00C)
		lblStatusOutput.Text = "Starting to End Session"
		lblStatusOutput.Refresh()
		
		If booSessionBegun Then
			qbXMLCOM.EndSession((strTicket))
			booSessionBegun = False
		End If
		
		lblStatusOutput.Text = CStr(Nothing)
		lblStatusOutput.ForeColor = System.Drawing.ColorTranslator.FromOle(&HC00C)
		lblStatusOutput.Text = "Session Ended"
		lblStatusOutput.Refresh()
		Exit Sub
		
ErrorDisplay: 
		MsgBox(Err.Description)
	End Sub
	
	Public Sub CloseConnection(ByRef lblStatusOutput As System.Windows.Forms.TextBox)
		
		On Error GoTo ErrorDisplay
		
		lblStatusOutput.Text = CStr(Nothing)
		lblStatusOutput.ForeColor = System.Drawing.ColorTranslator.FromOle(&HC00C)
		lblStatusOutput.Text = "Starting to Close Connection"
		lblStatusOutput.Refresh()
		
		If booConnected Then
			qbXMLCOM.CloseConnection()
			booConnected = False
		End If
		
		lblStatusOutput.Text = CStr(Nothing)
		lblStatusOutput.ForeColor = System.Drawing.ColorTranslator.FromOle(&HC00C)
		lblStatusOutput.Text = "Not Connected"
		lblStatusOutput.Refresh()
		
		Exit Sub
		
ErrorDisplay: 
		MsgBox(Err.Description)
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
	Public Sub PrettyPrintXMLToFile(ByRef XMLString As String, ByRef XMLFile As String)
		Dim FileNum As Object
		
		Dim SplitXMLString() As String
		Dim IndentString As String
		Dim XMLStringLength As Integer
		Dim SplitIndex As Object
		
		IndentString = CStr(Nothing)

        FileNum = FreeFile()
        FileOpen(FileNum, XMLFile, OpenMode.Output)

        'Remove the linefeeds from the XML output string
        XMLString = Replace(XMLString, vbLf, vbNullString)
		
		SplitXMLString = Split(XMLString, "<")

        'We're expecting the first character of the XML output to be "<"
        'which result in an empty first array element, so skip it.
        SplitIndex = LBound(SplitXMLString) + 1

        Do
            If Left(SplitXMLString(SplitIndex), 1) = "/" Then
                IndentString = Left(IndentString, Len(IndentString) - 3)
                PrintLine(FileNum, IndentString & "<" & SplitXMLString(SplitIndex))
                SplitIndex = SplitIndex + 1
            ElseIf Left(SplitXMLString(SplitIndex + 1), 1) = "/" Then
                If InStr(1, Left(SplitXMLString(SplitIndex), InStr(1, SplitXMLString(SplitIndex), ">")), " ") > 0 Then
                    PrintLine(FileNum, IndentString & "<" & SplitXMLString(SplitIndex))
                    SplitIndex = SplitIndex + 1
                Else
                    PrintLine(FileNum, IndentString & "<" & SplitXMLString(SplitIndex) & "<" & SplitXMLString(SplitIndex + 1))
                    SplitIndex = SplitIndex + 2
                End If
            Else
                PrintLine(FileNum, IndentString & "<" & SplitXMLString(SplitIndex))
                If SplitXMLString(SplitIndex) <> "?xml version=""1.0"" ?>" And SplitXMLString(SplitIndex) <> "!DOCTYPE QBXML PUBLIC '-//INTUIT//DTD QBXML QBD 1.0//EN' >" And Right(SplitXMLString(SplitIndex), 2) <> "/>" Then
                    IndentString = IndentString & "   "
                End If
                SplitIndex = SplitIndex + 1
            End If
        Loop Until SplitIndex >= UBound(SplitXMLString)

        If Left(SplitXMLString(UBound(SplitXMLString)), 1) = "/" Then
			IndentString = Left(IndentString, Len(IndentString) - 3)
		End If

        PrintLine(FileNum, IndentString & "<" & SplitXMLString(UBound(SplitXMLString)))

        FileClose(FileNum)
    End Sub
	
	Private Sub XMLFileToXMLString(ByRef XMLFileName As String, ByRef XMLString As String)
		Dim strInput As Object
		Dim FileNum As Object
		'Clear out the XML String
		XMLString = CStr(Nothing)

        'Open the file, read each line and concatonate each to
        'the XML String
        FileNum = FreeFile()
        FileOpen(FileNum, XMLFileName, OpenMode.Input)
        Do Until EOF(FileNum)
            strInput = LineInput(FileNum)
            XMLString = XMLString & strInput
        Loop
        FileClose(FileNum)
    End Sub
	
	Public Function SendXMLFile(ByRef strProcessor As String, ByRef strInXMLFile As String, ByRef strOutXMLFile As String, ByRef lblStatus As System.Windows.Forms.TextBox) As String
		
		Dim strInXML As String
		Dim strOutXML As String
		
		XMLFileToXMLString(strInXMLFile, strInXML)
		
		On Error GoTo ErrorDisplay
		
		lblStatus.ForeColor = System.Drawing.ColorTranslator.FromOle(&HE0F0)
		lblStatus.Text = "Your request is being processed"
		lblStatus.Refresh()
		
		'Process the XML
		If strProcessor = "RequestProcessor" Then
			strOutXML = qbXMLCOM.ProcessRequest(strTicket, strInXML)
		Else
			strOutXML = qbXMLCOM.ProcessSubscription(strInXML)
		End If
		
		PrettyPrintXMLToFile(strOutXML, strOutXMLFile)
		
		lblStatus.Text = CStr(Nothing)
		lblStatus.ForeColor = System.Drawing.ColorTranslator.FromOle(&HC00C)
		lblStatus.Text = "Processing complete"
		lblStatus.Refresh()
		SendXMLFile = CStr(Nothing)
		Exit Function
		
ErrorDisplay: 
		'   MsgBox Err.Description
		lblStatus.Text = CStr(Nothing)
		lblStatus.ForeColor = System.Drawing.ColorTranslator.FromOle(&HF0)
		lblStatus.Text = Err.Number & "  " & Err.Description
		lblStatus.Refresh()
		SendXMLFile = Err.Description
	End Function
End Module