Option Strict Off
Option Explicit On
Module qbooks
    '
    ' Copyright © 2021-2022 Intuit Inc. All rights reserved.
    ' Use is subject to the terms specified at:
    '      http://developer.intuit.com/legal/devsite_tos.html
    '
    ' $Id: qbooks.bas,v 1.1 2002/08/05 12:55:40 giza Exp $
    '
    '
    ' Deliver the qbXML request to QuickBooks and return the
    ' response.
    '
    Public Function post(ByRef xmlStream As String) As String
		
		On Error GoTo errHandler
		
		' Open connection to qbXMLRP COM
		Dim qbXMLRP As New QBXMLRP2Lib.RequestProcessor2
		Dim connType As QBXMLRP2Lib.QBXMLRPConnectionType
		connType = QBXMLRP2Lib.QBXMLRPConnectionType.localQBD
		If (UIForm.UseQBOE).CheckState Then
			connType = QBXMLRP2Lib.QBXMLRPConnectionType.remoteQBOE
		End If
		qbXMLRP.OpenConnection2("", "SDKTest", connType)
		
		
		' Begin Session
		' Pass empty string for the data file name to use the currently
		' open data file.
		Dim ticket As String
		ticket = qbXMLRP.BeginSession("", QBXMLRP2Lib.QBFileMode.qbFileOpenDoNotCare)
		
		' Send request to QuickBooks
		post = qbXMLRP.ProcessRequest(ticket, xmlStream)
		
		' End the session
		qbXMLRP.EndSession(ticket)
		
		' Close the connection
		qbXMLRP.CloseConnection()
		Exit Function
		
errHandler: 
		post = ""
		Log.out((Err.Description))
	End Function
End Module