Option Strict Off
Option Explicit On
Friend Class Query
	Inherits System.Windows.Forms.Form
	'----------------------------------------------------------
	' Form: InvoiceQuery
	'
	' Description:  This sample demonstrates the use of QBFC,
	'               by querying for invoices within a given date range.
	'               It includes examples of the following:
	'                   - Constructing a complex query
	'                   - Looping through the list of invoices in a response
	'                   - Getting detailed invoice information (invoice lines)
	'                   - Reading data from an OR object
	'                   - Checking fields that are not guaranteed in the response
	'                     and obtaining data from them
	'
	' Created On: 01/17/2002
	' Updated to QBFC 2.0: 08/2002
	' Updated to QBFC 5.0: 09/2005
	' Updated to QBFC 12.0 and fixed setting max QBXML version: 09/2012
	'
	' Copyright © 2021-2022 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	'----------------------------------------------------------
	
	
	Dim fromDate As String
	Dim toDate As String
	Dim invoiceCollection As Collection
	
	' These constants hold the application name and ID plus the
	' major and minor versions for QBFC 2.0.
	Const cAppID As String = ""
	Const cAppName As String = "IDN Desktop VB QBFC InvoiceQuery"
	
	
	
	
	'
	' When the user clicks Submit, validate the data that has been entered
	' and send the request to QuickBooks for processing.
	'
	Private Sub Submit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Submit.Click
		
		On Error GoTo ErrHandler
		
		If GetDates Then 'Check the validity of the input data
			If QueryInvoices Then 'If the dates are valid, use them in query
				ShowInvoices() 'If there are any invoices, show them
			End If
		End If
		Exit Sub
		
ErrHandler: 
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
		Exit Sub
		
	End Sub
	
	
	
	'
	' Make sure input data is valid, in particular the dates.
	'
	Private Function GetDates() As Object
		
		On Error GoTo ErrHandler
		
		fromDate = VB6.Format(FromTxnDate.Text, "mm/dd/yyyy")
		toDate = VB6.Format(ToTxnDate.Text, "mm/dd/yyyy")
		
		If fromDate <> "" And Not IsDate(fromDate) Then
			MsgBox("From Txn Date format is wrong!")
			
			GetDates = False
			Exit Function
		End If
		
		If toDate <> "" And Not IsDate(toDate) Then
			MsgBox("To Txn Date format is wrong!")
			
			GetDates = False
			Exit Function
		End If
		
		
		GetDates = True
		Exit Function
		
ErrHandler: 
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
		
		GetDates = False
		Exit Function
		
	End Function
	
	
	
	'
	' Query QuickBooks for invoices between the specified dates
	' using QBFC.
	'
	Private Function QueryInvoices() As Object
		
		On Error GoTo ErrHandler
		
		' Step 1: Start session with QuickBooks
		Dim SessionManager As New QBFC15Lib.QBSessionManager
		Dim connType As QBFC15Lib.ENConnectionType
		connType = QBFC15Lib.ENConnectionType.ctLocalQBDLaunchUI
		If (UseQBOE.CheckState = 1) Then
			connType = QBFC15Lib.ENConnectionType.ctRemoteQBOE
		End If
		SessionManager.OpenConnection2(cAppID, cAppName, connType)
		SessionManager.BeginSession("", QBFC15Lib.ENOpenMode.omDontCare)
		
		' Step 2: Create Message Set request
		Dim requestMsgSet As QBFC15Lib.IMsgSetRequest
		requestMsgSet = GetLatestMsgSetRequest(SessionManager)
		requestMsgSet.Attributes.OnError = QBFC15Lib.ENRqOnError.roeContinue
		
		' Step 3: Create the query object needed to perform InvoiceQueryRq
		Dim invQuery As QBFC15Lib.IInvoiceQuery
		invQuery = requestMsgSet.AppendInvoiceQueryRq
		
		Dim invFilter As QBFC15Lib.IInvoiceFilter
		invFilter = invQuery.ORInvoiceQuery.InvoiceFilter
		
		' If there's date info, create the filter and put it in the query
		If fromDate <> "" Or toDate <> "" Then
			If fromDate <> "" Then
				invFilter.ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter.TxnDateFilter.FromTxnDate.SetValue((CDate(fromDate)))
			End If
			If toDate <> "" Then
				invFilter.ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter.TxnDateFilter.ToTxnDate.SetValue((CDate(toDate)))
			End If
		End If
		invQuery.IncludeLineItems.SetValue((Me.IncludeLineItems.CheckState))
		
		' Step 4: Do the request
		Dim responseMsgSet As QBFC15Lib.IMsgSetResponse
		responseMsgSet = SessionManager.DoRequests(requestMsgSet)
		
		' Terminate the session and connection,
		' since we are done with the session manager
		SessionManager.EndSession()
		SessionManager.CloseConnection()
		
		' Uncomment the following to see the request and response XML for debugging
		' MsgBox requestMsgSet.ToXMLString, vbOKOnly, "RequestXML"
		' MsgBox responseMsgSet.ToXMLString, vbOKOnly, "ResponseXML"
		
		' Step 5: Interpret the response
		Dim rsList As QBFC15Lib.IResponseList
		rsList = responseMsgSet.ResponseList
		
		Dim response As QBFC15Lib.IResponse
		' Retrieve the one response corresponding to our single request
		response = rsList.GetAt(0)
		
		Dim msg As Object
		Dim invoiceList As QBFC15Lib.IInvoiceRetList
		Dim ndx As Object
		Dim invoiceRet As QBFC15Lib.IInvoiceRet
		If (response.StatusCode >= 1000) Then
			If (response.StatusCode = 1) Then ' No record found
				MsgBox("No invoice is found", MsgBoxStyle.Information, "Message from QuickBooks")
			Else
				
				msg = "Error occured.  Status Code = " & CStr(response.StatusCode) & ", Status Message = " & response.StatusMessage & ", Status Severity = " & response.StatusSeverity
				MsgBox(msg, MsgBoxStyle.Exclamation, "Message from QuickBooks")
			End If
			
			QueryInvoices = False
			Exit Function
		Else
			' We have one or more invoices in the invoice list, which is the response.Detail
			invoiceList = response.Detail
			invoiceCollection = New Collection
			For ndx = 0 To (invoiceList.Count - 1)
				
				invoiceRet = invoiceList.GetAt(ndx)
				' Add to the collection
				invoiceCollection.Add(invoiceRet, invoiceRet.TxnID.GetValue)
			Next 
		End If
		
		
		QueryInvoices = True
		Exit Function
		
ErrHandler: 
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
		
		QueryInvoices = False
		Exit Function
	End Function
	
	
	
	'
	' Display invoices using the Display.frm form.
	'
	Private Sub ShowInvoices()
		
		Dim msg As String
		Dim invRet As QBFC15Lib.IInvoiceRet
		Dim i As Short
		Const cMaxShown As Short = 30 ' Show only a maximum of 30 invoices
		
		If invoiceCollection.Count() > cMaxShown Then
			msg = "Showing " & CStr(cMaxShown) & " out of " & CStr(invoiceCollection.Count()) & " invoices" & vbCrLf
		End If
		
		i = 1
		For	Each invRet In invoiceCollection
			msg = msg & vbCrLf & GetInvoiceRetDetail(invRet)
			If i = cMaxShown Then
				Exit For
			End If
			i = i + 1
		Next invRet
		
		Dim frmDisplay As New Display
		frmDisplay.Text_Content.Text = msg
		VB6.ShowForm(frmDisplay, VB6.FormShowConstants.Modal, Me)
		
		'Clear the collection
		
		invoiceCollection = Nothing
		Exit Sub
		
	End Sub
	
	
	
	'
	' Retrieve details of an invoice from a particular IInvoiceRet object.
	'
	Private Function GetInvoiceRetDetail(ByRef invRet As QBFC15Lib.IInvoiceRet) As String
		
		Dim msg As Object
		
		'Retrieve guaranteed fields
		
		msg = " TxnNumber = " & invRet.TxnNumber.GetValue & ", Customer = " & invRet.CustomerRef.FullName.GetValue
		
		'Retrive non-guaranteed fields
		If (Not (invRet.RefNumber Is Nothing)) Then
			
			msg = msg & ", RefNumber = " & invRet.RefNumber.GetValue
		End If
		
		If (Not (invRet.Memo Is Nothing)) Then
			
			msg = msg & ", Memo = " & invRet.RefNumber.GetValue
		End If
		
		'Retrieve invoice line list
		'Each line can be either InvoiceLineRet OR InvoiceLineGroupRet
		Dim orInvoiceLineRetList As QBFC15Lib.IORInvoiceLineRetList
		orInvoiceLineRetList = invRet.orInvoiceLineRetList
		Dim linendx, linendxMax As Object
		Dim orInvoiceLineRet As QBFC15Lib.IORInvoiceLineRet
		Dim orRate As QBFC15Lib.IORRate
		If (Not (orInvoiceLineRetList Is Nothing)) Then
			
			
			linendxMax = orInvoiceLineRetList.Count - 1
			
			For linendx = 0 To linendxMax
				
				orInvoiceLineRet = orInvoiceLineRetList.GetAt(linendx)
				
				
				
				msg = msg & vbCrLf & vbTab & " Line: " & CStr(linendx)
				'Check what to retrieve from the orInvoiceLineRet object
				'based on the "ortype" property
				If (orInvoiceLineRet.ortype = QBFC15Lib.ENORInvoiceLineRet.orilrInvoiceLineRet) Then
					
					If (Not (orInvoiceLineRet.InvoiceLineRet.Desc Is Nothing)) Then
						
						msg = msg & ", Desc: " & orInvoiceLineRet.InvoiceLineRet.Desc.GetValue
					End If
					
					If (Not (orInvoiceLineRet.InvoiceLineRet.Amount Is Nothing)) Then
						
						msg = msg & ", Amount: " & orInvoiceLineRet.InvoiceLineRet.Amount.GetValue
					End If
					
					If (Not (orInvoiceLineRet.InvoiceLineRet.ItemRef Is Nothing)) Then
						
						msg = msg & ", Quantity: " & orInvoiceLineRet.InvoiceLineRet.ItemRef.FullName.GetValue
					End If
					
					orRate = orInvoiceLineRet.InvoiceLineRet.orRate
					If (Not (orRate Is Nothing)) Then
						If orRate.ortype = QBFC15Lib.ENORRate.orrRate Then
							
							msg = msg & ", Rate: " & CStr(orRate.Rate.GetValue)
						Else
							
							msg = msg & ", RatePercent: " & CStr(orRate.RatePercent.GetValue)
						End If
					End If
					
				ElseIf (orInvoiceLineRet.ortype = QBFC15Lib.ENORInvoiceLineRet.orilrInvoiceLineGroupRet) Then 
					
					msg = msg & ", Group Name: " & orInvoiceLineRet.InvoiceLineGroupRet.ItemGroupRef.FullName.GetValue
					
					msg = msg & ", Total Amount: " & CStr(orInvoiceLineRet.InvoiceLineGroupRet.TotalAmount.GetValue)
					
					If (Not (orInvoiceLineRet.InvoiceLineGroupRet.Desc Is Nothing)) Then
						
						msg = msg & ", Desc: " & orInvoiceLineRet.InvoiceLineGroupRet.Desc.GetValue
					End If
					
				End If
			Next 
		End If
		
		
		GetInvoiceRetDetail = msg
	End Function
	
	
	
	'
	' Exit program.
	'
	Private Sub Exit_Renamed_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Exit_Renamed.Click
		Me.Close()
	End Sub
	
	
	
	Public Function GetLatestMsgSetRequest(ByRef SessionManager As QBFC15Lib.QBSessionManager) As QBFC15Lib.IMsgSetRequest
		Dim supportedVersion As String
		supportedVersion = CStr(Val(QBFCLatestVersion(SessionManager)))
		If (supportedVersion >= CStr(6#)) Then
			GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 6, 0)
		ElseIf (supportedVersion >= CStr(5#)) Then 
			GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 5, 0)
		ElseIf (supportedVersion >= CStr(4#)) Then 
			GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 4, 0)
		ElseIf (supportedVersion >= CStr(3#)) Then 
			GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 3, 0)
		ElseIf (supportedVersion >= CStr(2#)) Then 
			GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 2, 0)
		ElseIf (supportedVersion = CStr(1.1)) Then 
			GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 1, 1)
		Else
			MsgBox("You are apparently running QuickBooks 2002 Release 1, we strongly recommend that you use QuickBooks' online update feature to obtain the latest fixes and enhancements", MsgBoxStyle.Exclamation)
			GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 1, 0)
		End If
	End Function
	
	Function QBFCLatestVersion(ByRef SessionManager As QBFC15Lib.QBSessionManager) As String
		Dim strXMLVersions() As String
		strXMLVersions = VB6.CopyArray(SessionManager.QBXMLVersionsForSession)
		
		
		Dim i As Integer
		Dim vers As Double
		Dim LastVers As Double
		LastVers = 0
		For i = LBound(strXMLVersions) To UBound(strXMLVersions)
			vers = Val(strXMLVersions(i))
			If (vers > LastVers) Then
				LastVers = vers
				QBFCLatestVersion = strXMLVersions(i)
			End If
		Next i
	End Function
End Class