Option Strict Off
Option Explicit On
Module TaxCodeInfo
	' TaxCodeInfo.bas
	' This module is part of the Invoice sample program
	' for the QuickBooks SDK Version CA2.0.
	' Created September, 2002
	'
	' Copyright © 2021-2022 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	'-------------------------------------------------------------
	
	
	'global strings containing information associated with the Tax code
	Public colTaxCodeList As Collection
	Sub GetTaxCodeList()
		'The GetTaxCodeList function generates a request to QuickBooks and call the ParseTaxCode procedure to parse the
		'TaxCode's list
		On Error GoTo ErrHandler
		'Call OpenConnection to establish a QuickBooks connection and session
		If Not blnIsOpenConnection Then
			If Not OpenConnection Then
				Exit Sub
			End If
		End If


        ' We generate the request qbXML in order to get the TaxCode list.
        requestXML = "<?xml version=""1.0"" ?>" & "<?qbxml version=""CA2.0""?>" & "<QBXML><QBXMLMsgsRq onError=""continueOnError"">" & "<TaxCodeQueryRq requestID=""1""/>" & "</QBXMLMsgsRq></QBXML>"


        '    PrintXMLToFile requestXML, "C:\request.xml" '*** remove Comment character to produce a xml request file

        ' Send request to QuickBooks
        responseXML = qbXMLRP.ProcessRequest(ticket, requestXML)
		
		'    PrintXMLToFile responseXML, "C:\response.xml" '*** remove Comment character to produce a xml response file
		
		'The ParseTaxCode procedure parses the response coming from QuickBooks and populates the colTaxCodeList collection
		ParseTaxCode((responseXML))
		Exit Sub
ErrHandler: 
		' In case something goes wrong along the way..
		MsgBox("Error sending request to QuickBooks: " & Err.Description)
		CloseConnection()
	End Sub
	
	Sub ParseTaxCode(ByRef sXML As Object)
        'This function parse the response to the TaxCodeQuery request

        Dim objXmlDoc As New MSXML2.DOMDocument60
        Dim objRootElement As MSXML2.IXMLDOMElement ' root element of the XML document (QBXML Block)
		Dim objMsgsRsNode As MSXML2.IXMLDOMNode ' QBXMLMsgsRs grouping (opening tag of message)
		Dim objMessageNode As MSXML2.IXMLDOMNode ' Message node
		Dim objTaxCodeNode As MSXML2.IXMLDOMNode ' TaxCode node
		Dim objTaxCodeItem As MSXML2.IXMLDOMNode ' TaxCode node
		Dim objElementItem As MSXML2.IXMLDOMNode ' Element node
		
		Dim strTaxCodeArray(2) As String
		'strTaxCodeArray(0) is TaxCode Name
		'strTaxCodeArray (1) is GST
		'strTaxCodeArray(2) is PST
		' Load XML Script
		objXmlDoc.async = False
        If objXmlDoc.loadXML(sXML) = True And Len(objXmlDoc.xml) > 0 Then 'verify that there is some values in the documents

            objRootElement = objXmlDoc.documentElement

            For Each objMsgsRsNode In objRootElement.childNodes

                If objMsgsRsNode.nodeName = "QBXMLMsgsRs" Then

                    For Each objMessageNode In objMsgsRsNode.childNodes

                        ' If we find the TaxCode list, let's parse it.
                        If objMessageNode.nodeName = "TaxCodeQueryRs" Then

                            For Each objTaxCodeNode In objMessageNode.childNodes
                                strTaxCodeArray(0) = ""
                                strTaxCodeArray(1) = ""
                                strTaxCodeArray(2) = ""

                                For Each objTaxCodeItem In objTaxCodeNode.childNodes
                                    'Transfer the TaxCode information into the TaxCode collection

                                    If objTaxCodeItem.nodeName = "Name" Then
                                        strTaxCodeArray(0) = objTaxCodeItem.nodeTypedValue
                                    ElseIf objTaxCodeItem.nodeName = "Tax1Rate" Then
                                        strTaxCodeArray(1) = objTaxCodeItem.nodeTypedValue
                                    ElseIf objTaxCodeItem.nodeName = "Tax2Rate" Then
                                        strTaxCodeArray(2) = objTaxCodeItem.nodeTypedValue
                                    End If

                                Next objTaxCodeItem
                                colTaxCodeList.Add(strTaxCodeArray, strTaxCodeArray(0))
                            Next objTaxCodeNode

                        End If
                    Next objMessageNode
                End If
            Next objMsgsRsNode
        End If
    End Sub
End Module