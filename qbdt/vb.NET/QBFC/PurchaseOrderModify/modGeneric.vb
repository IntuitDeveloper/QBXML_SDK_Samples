Option Strict Off
Option Explicit On
Module modGeneric
	'----------------------------------------------------------
	' Copyright © 2021-2022 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	'----------------------------------------------------------
	
	
	Dim Implementation As String
	
	Dim strqbXMLLevel As String
	Dim booSupportsModify As Boolean
	
	Dim dateFrom As Date
	Dim dateTo As Date
	
	Dim strSelectedPOInfo As String
	
	
	Public Sub SetImplementation(ByRef Value As String)
		Implementation = Value
	End Sub
	
	
	Public Function GetImplementation() As String
		GetImplementation = Implementation
	End Function
	
	
	Public Sub SetDates(ByRef dateInFrom As Date, ByRef dateInTo As Date)
		
		dateFrom = dateInFrom
		dateTo = dateInTo
	End Sub
	
	
	Public Sub SetSelectedPOInfo(ByRef strPOInfo As String)
		strSelectedPOInfo = strPOInfo
	End Sub
	
	
	Public Function SupportsModify() As Boolean
		SupportsModify = booSupportsModify
	End Function
	
	Public Function QuickBooksVersionOK() As Boolean
		'This used to show error checking, but it did so badly, so we just
		'made a dummy out of this function.  If you run this against a QB
		'version less than QB2003 R3 you'll get a runtime error.
		If Implementation = "QBXMLRP" Then
			QBXMLRP_OpenConnectionBeginSession()
		Else
			QBFC_OpenConnectionBeginSession()
		End If
		
		strqbXMLLevel = GetMaxVersionSupported
		
		booSupportsModify = True
		
		QuickBooksVersionOK = True
	End Function
	
	Public Sub EndSessionCloseConnection()
		If Implementation = "QBXMLRP" Then
			QBXMLRP_EndSessionCloseConnection()
		ElseIf Implementation = "QBFC" Then 
			QBFC_EndSessionCloseConnection()
			'  Else
			'    QBFCCA_EndSessionCloseConnection
		End If
	End Sub
	
	Function GetMaxVersionSupported() As String
		If Implementation = "QBXMLRP" Then
			GetMaxVersionSupported = "3.0"
		ElseIf Implementation = "QBFC" Then 
			GetMaxVersionSupported = "3.0"
		End If
	End Function
	
	Public Sub LoadPOInfoArray(ByRef strPOInfo() As String, ByRef booGiveWarning As Boolean)
		If Implementation = "QBXMLRP" Then
			QBXMLRP_LoadPOInfoArray(strPOInfo, dateFrom, dateTo, booGiveWarning)
		ElseIf Implementation = "QBFC" Then 
			QBFC_LoadPOInfoArray(strPOInfo, dateFrom, dateTo, booGiveWarning)
			'  Else
			'    QBFCCA_LoadPOInfoArray strPOInfo, dateFrom, dateTo, booGiveWarning
		End If
	End Sub
	
	Public Sub GetPOInfo(ByRef strTxnID As String, ByRef strEditSequence As String, ByRef strRefNumber As String, ByRef strTxnDate As String, ByRef strVendor As String, ByRef strPOLines() As String)
		
		If Implementation = "QBXMLRP" Then
			QBXMLRP_GetPOInfo(strTxnID, strEditSequence, strRefNumber, strTxnDate, strVendor, strPOLines, strSelectedPOInfo)
		ElseIf Implementation = "QBFC" Then 
			QBFC_GetPOInfo(strTxnID, strEditSequence, strRefNumber, strTxnDate, strVendor, strPOLines, strSelectedPOInfo)
			'  Else
			'    QBFCCA_GetPOInfo strTxnID, strEditSequence, strRefNumber, strTxnDate, _
			''                     strVendor, strPOLines, strSelectedPOInfo
		End If
	End Sub
	
	
	Public Sub ClosePO(ByRef strTxnID As String, ByRef strEditSequence As String)
		If Implementation = "QBXMLRP" Then
			QBXMLRP_ClosePO(strTxnID, strEditSequence)
		ElseIf Implementation = "QBFC" Then 
			QBFC_ClosePO(strTxnID, strEditSequence)
			'  Else
			'    QBFCCA_ClosePO strTxnID, strEditSequence
		End If
	End Sub
	
	
	Public Sub ChangePOLine(ByRef strAction As String, ByRef strTxnID As String, ByRef strEditSequence As String, ByRef strPOLines() As String, ByRef intSelectedPOLine As Short)
		
		Dim intGroupLine As Short
		intGroupLine = -1
		If strAction <> "ReceiveAll" Then
			If LineType(strPOLines(intSelectedPOLine)) = "GroupSubItem" Then
				intGroupLine = intSelectedPOLine - 1
				Do While LineType(strPOLines(intGroupLine)) <> "GroupItem"
					intGroupLine = intGroupLine - 1
				Loop 
			End If
		End If
		
		If Implementation = "QBXMLRP" Then
			QBXMLRP_ChangePOLine(strAction, strTxnID, strEditSequence, strPOLines, intSelectedPOLine, intGroupLine)
		ElseIf Implementation = "QBFC" Then 
			QBFC_ChangePOLine(strAction, strTxnID, strEditSequence, strPOLines, intSelectedPOLine, intGroupLine)
			'  Else
			'    QBFCCA_ChangePOLine strAction, strTxnID, strEditSequence, _
			''                        strPOLines, intSelectedPOLine, intGroupLine
		End If
	End Sub
	
	
	Public Sub SetQuantitiesAndBillForRemainingItems(ByRef strTxnID As String, ByRef strEditSequence As String, ByRef strVendor As String, ByRef strRefNumber As String, ByRef strTxnDate As String, ByRef strPOLines() As String, ByRef intPOLine As Short)
		
		If Implementation = "QBXMLRP" Then
			QBXMLRP_SetQuantitiesAndBillForRemainingItems(strTxnID, strEditSequence, strVendor, strRefNumber, strTxnDate, strPOLines, intPOLine)
		ElseIf Implementation = "QBFC" Then 
			QBFC_SetQuantitiesAndBillForRemainingItems(strTxnID, strEditSequence, strVendor, strRefNumber, strTxnDate, strPOLines, intPOLine)
			'  Else
			'    QBFCCA_SetQuantitiesAndBillForRemainingItems _
			''      strTxnID, strEditSequence, strVendor, strRefNumber, _
			''      strTxnDate, strPOLines, intPOLine
		End If
	End Sub
	
	
	Public Function LineType(ByRef strPOLine As String) As String
		
		Dim Splits() As String
		
		Splits = Split(strPOLine, "<spliter>")
		
		LineType = Splits(9)
	End Function
End Module