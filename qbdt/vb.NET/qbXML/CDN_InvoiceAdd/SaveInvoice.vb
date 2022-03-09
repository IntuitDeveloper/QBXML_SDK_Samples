Option Strict Off
Option Explicit On
Module SaveInvoice
	' SaveInvoice.bas
	' This module is part of the Invoice sample program
	' for the QuickBooks SDK Version CA2.0.
	' Created September, 2002
	'
	' Copyright © 2021-2022 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	'-------------------------------------------------------------
	
	
	Public Sub FuncTionSaveInvoice(ByRef requestXML As String)
		'This function will save the invoice by sending a request to QuickBooks
		On Error GoTo ErrHandler
		'Call OpenConnection to establish a QuickBooks connection and session
		If Not blnIsOpenConnection Then
			If Not OpenConnection Then
				Exit Sub
			End If
		End If
		
		'PrintXMLToFile requestXML, "C:\request.xml" '*** remove Comment character to produce a xml request file
		
		' Send request to QuickBooks
		responseXML = qbXMLRP.ProcessRequest(ticket, requestXML)
		
		'PrintXMLToFile responseXML, "C:\response.xml" '*** remove Comment character to produce a xml response file
		
		' The function processInvoiceAddResponse takes the
		' qbXML which QuickBooks has returned and parses it in order to find out
		' whether the call was successful or whether there was an error.
		If Not processInvoiceAddResponse(responseXML) Then
			Exit Sub
		End If
		Exit Sub
ErrHandler: 
		' In case something goes wrong along the way..
		MsgBox("Error adding invoice to QuickBooks: " & Err.Description)
		CloseConnection()
	End Sub
	
	' Process the qbXML response and print out a success or error
	' message about the added Invoice.  The parsing in this function
	' looks only at whether the response returned with an error or
	' with a good status.
	'
	Public Function processInvoiceAddResponse(ByRef strResponseXML As String) As Boolean

        Dim doc As New MSXML2.DOMDocument60

        ' For status codes of the InvoiceAdd call
        Dim statusCode As String
		
		' For status message of InvoiceAdd
		Dim statusMsg As String
		
		' These are needed for the status code attributes later
		Dim rprAttr As MSXML2.IXMLDOMNamedNodeMap
		Dim statCodeNode As MSXML2.IXMLDOMNode
		Dim statMsgNode As MSXML2.IXMLDOMNode
		
		' Variable recAddInvNodeList will hold all elements named
		' InvoiceAddRs, and recAddInvRsNode will hold the
		' particular element
		Dim recAddInvNodeList As MSXML2.IXMLDOMNodeList
		Dim recAddInvRsNode As MSXML2.IXMLDOMNode
		
		
		doc.async = False
		doc.loadXML(strResponseXML)
		
		'
		' PARSE THE ADD Invoice RESPONSE:
		'
		
		' recAddInvRsNode will grab the InvoiceAddRs element we need
		' from the array of 1 element
		recAddInvNodeList = doc.getElementsByTagName("InvoiceAddRs")
		recAddInvRsNode = recAddInvNodeList.Item(0)
		
		' InvoiceAddRs has four attributes:
		' requestID, statusCode, statusSeverity, statusMessage
		rprAttr = recAddInvRsNode.Attributes
		
		' We only want the statusCode attribute and the statusMessage
		If rprAttr.length > 0 Then
			statCodeNode = rprAttr.getNamedItem("statusCode")
            statusCode = statCodeNode.nodeValue
            statMsgNode = rprAttr.getNamedItem("statusMessage")
            statusMsg = statMsgNode.nodeValue
        End If
		
		
		' A statusCode of 0 indicates that there is no error.
		' Return a success message if this is the case.  Otherwise
		' return an error message with the status message from QB.
		If 0 = CDbl(statusCode) Then
			processInvoiceAddResponse = True
			MsgBox("The new invoice has been successfully posted to QuickBooks")
		Else
			processInvoiceAddResponse = False
			MsgBox("Error from QuickBooks: " & statusMsg)
		End If
		
	End Function
End Module