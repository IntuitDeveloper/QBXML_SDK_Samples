Option Strict Off
Option Explicit On
Module QBooks
    ' qbooks.bas
    '
    ' This module is part of the Invoice sample program
    ' for the QuickBooks SDK Version CA2.0.
    ' Created September, 2002
    '
    ' Copyright © 2021-2022 Intuit Inc. All rights reserved.
    ' Use is subject to the terms specified at:
    '      http://developer.intuit.com/legal/devsite_tos.html
    '
    '-------------------------------------------------------------


    'Connection info
    ' qbXMLRP COM object
    Public qbXMLRP As New QBXMLRPLib.RequestProcessor
    Public blnIsOpenConnection As Boolean
	Public blnSessionBegun As Boolean
	' Ticket
	Public ticket As String
	
	' Request and response strings
	Public requestXML As String
	Public responseXML As String
	
	
	Public Function OpenConnection() As Boolean
		'Open the connection
		On Error GoTo ErrHandler
		
		If blnIsOpenConnection Then
			OpenConnection = True
			Exit Function
		End If
		
		
		blnSessionBegun = False
		blnIsOpenConnection = False
		
		' Open connection to qbXMLRP COM
		qbXMLRP.OpenConnection("QBRAddEmp", "IDN QuickBooks Sample Add Invoice")
		blnIsOpenConnection = True
        ' Begin Session
        ' Pass empty string for the data file name to use the currently
        ' open data file.


        ticket = qbXMLRP.BeginSession("", QBXMLRPLib.QBFileMode.qbFileOpenSingleUser)
        blnSessionBegun = True
		
		OpenConnection = True 'return that the connection was successful
		
		'Verifying which version of qbXML QuickBooks is supporting If you want to do a US/Canadian APP,
		'This is where you would find the version supported by QuickBooks.  You would modify the requests accordingly
		'to the version of QuickBooks
		Dim VersionSupportedArray() As String
        VersionSupportedArray = qbXMLRP.QBXMLVersionsForSession(ticket) ' This return an array of string containing all the version of qbXML
        'supported by QuickBooks

        'Checking that QuickBooks support the Canadian SDK (version CA2.0)
        Dim strCanadianVersion As String
		Dim blnCanadianVersionFound As Boolean
		
		Dim str_Renamed As Object


        '    Dim nArrayUpperBound As Integer


        strCanadianVersion = "2.0"
        blnCanadianVersionFound = True

        For	Each str_Renamed In VersionSupportedArray
            
            'MessageBox.Show(str_Renamed)
            If strCanadianVersion = str_Renamed Then
                blnCanadianVersionFound = True
            End If

        Next str_Renamed
		
		If blnCanadianVersionFound = False Then 'If version CA 2.0 not found...
            MsgBox("This QuickBooks does not support the version 2.0 of qbXML",  , "qbXML version not supported""")
            If blnSessionBegun = True Then
				qbXMLRP.EndSession(ticket)
			End If
			' Close the connection
			If blnIsOpenConnection = True Then
				blnIsOpenConnection = False
				qbXMLRP.CloseConnection()
			End If
			OpenConnection = False
			
		End If
		
		
		Exit Function
		
ErrHandler: 
		blnIsOpenConnection = False
		OpenConnection = False
		' End the session
		If blnSessionBegun = True Then
			qbXMLRP.EndSession(ticket)
		End If
		' Close the connection
		If blnIsOpenConnection = True Then
			qbXMLRP.CloseConnection()
		End If
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
		Exit Function
		
	End Function
	
	
	Public Sub CloseConnection()
		' Ends session and closes connection
		
		If Not blnIsOpenConnection Then
			Exit Sub
		End If
		
		On Error GoTo ErrHandler
		
		If blnSessionBegun = True Then
			qbXMLRP.EndSession(ticket)
		End If
		' Close the connection
		If blnIsOpenConnection = True Then
			qbXMLRP.CloseConnection()
		End If
		
		Exit Sub
		
ErrHandler: 
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
		Exit Sub
		
	End Sub
	
	
	' This subroutine is available for error checking.  It is sometimes
	' useful to print the XML which QuickBooks returns to a file so that
	' any problems can be uncovered easily.  Although this subroutine is
	' not currently in use in the ReceievePayment sample code, it is
	' encouraged that you add it in if you would like to see the precise
	' XML that is being sent to or received from QuickBooks.
	'
	Sub PrintXMLToFile(ByRef xmlString As String, ByRef XMLFile As String)
		
		Dim SplitXMLString() As String
		Dim IndentString As String
		Dim xmlStringLength As Integer
		Dim SplitIndex As Object
		
		IndentString = ""
		
		Dim FileNum As Object
		
		FileNum = FreeFile
		
		FileOpen(FileNum, XMLFile, OpenMode.Output)
		
		' Remove the linefeeds from the XML output string
		xmlString = Replace(xmlString, vbLf, vbNullString)
		
		SplitXMLString = Split(xmlString, "<")
		
		' We're expecting the first character of the XML output to be "<"
		' which result in an empty first array element, so skip it.
		
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
				IndentString = IndentString & "   "
				
				SplitIndex = SplitIndex + 1
			End If
			
		Loop Until SplitIndex >= UBound(SplitXMLString)
		
		If Left(SplitXMLString(UBound(SplitXMLString)), 1) = "/" Then
			IndentString = Left(IndentString, Len(IndentString) - 3)
		End If
		
		
		PrintLine(FileNum, IndentString & "<" & SplitXMLString(UBound(SplitXMLString)))
		
		
		FileClose(FileNum)
	End Sub
End Module