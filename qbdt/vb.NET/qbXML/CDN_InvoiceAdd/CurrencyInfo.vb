Option Strict Off
Option Explicit On
Module CurrencyInfo
	' CurrencyInfo.bas
	' This module is part of the Invoice sample program
	' for the QuickBooks SDK Version CA2.0.
	' Created September, 2002
	'
	' Copyright © 2021-2022 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	'-------------------------------------------------------------
	
	Function GetCurExRate(ByRef strCurrencyID As String) As String
		'This function create a request to QuickBooks to return information on a single currency
		On Error GoTo ErrHandler
		
		
		'Call OpenConnection to establish a QuickBooks connection and session
		If Not blnIsOpenConnection Then
			If Not OpenConnection Then
				Exit Function
			End If
		End If


        ' This request ask QuickBooks to return information on a certain currency
        requestXML = "<?xml version=""1.0"" ?>" & "<?qbxml version=""2.0""?>" & "<QBXML><QBXMLMsgsRq onError=""continueOnError"">" & "<CurrencyQueryRq requestID=""1"">" & "<ListID>" & strCurrencyID & "</ListID>" & "</CurrencyQueryRq>" & "</QBXMLMsgsRq></QBXML>"

        '    PrintXMLToFile requestXML, "C:\request.xml" '*** remove Comment character to produce a xml request file

        ' Send request to QuickBooks
        responseXML = qbXMLRP.ProcessRequest(ticket, requestXML)
		'    PrintXMLToFile responseXML, "C:\response.xml"  '*** remove Comment character to produce a xml response file
		
		'Parse the response
		GetCurExRate = ParseCurrency(responseXML)
		
		Exit Function
		
ErrHandler: 
		' In case something goes wrong along the way..
		MsgBox("Error getting the currency rate from QuickBooks: " & Err.Description)
		CloseConnection()
	End Function
	
	Function ParseCurrency(ByRef sXML As Object) As String
        'This function parse the response from QuickBooks in order to find the Exchange Rate
        Dim objXmlDoc As New MSXML2.DOMDocument60
        Dim objRootElement As MSXML2.IXMLDOMElement ' root element of the XML document (QBXML Block)
		Dim objMsgsRsNode As MSXML2.IXMLDOMNode ' QBXMLMsgsRs grouping (opening tag of message)
		Dim objMessageNode As MSXML2.IXMLDOMNode ' Message node
		Dim objCurrencyNode As MSXML2.IXMLDOMNode ' Currency node
		Dim objCurrencyItem As MSXML2.IXMLDOMNode ' CurrencyItem node
		
		
		' Load XML Script
		objXmlDoc.async = False
		
		If objXmlDoc.loadXML(sXML) = True And Len(objXmlDoc.xml) > 0 Then 'verify that there is some values in the documents
			
			objRootElement = objXmlDoc.documentElement
			
			For	Each objMsgsRsNode In objRootElement.childNodes
				
				If objMsgsRsNode.nodeName = "QBXMLMsgsRs" Then
					
					For	Each objMessageNode In objMsgsRsNode.childNodes
						
						' If we find the Currency, let's parse it.
						If objMessageNode.nodeName = "CurrencyQueryRs" Then
							
							
							For	Each objCurrencyNode In objMessageNode.childNodes 'For each currency in the list(in this example, there is only one currency returned)
								
								For	Each objCurrencyItem In objCurrencyNode.childNodes
									If objCurrencyItem.nodeName = "ExchangeRate" Then
										
										ParseCurrency = objCurrencyItem.nodeTypedValue 'Add the exchange rate to the collection
										Exit For
									End If
								Next objCurrencyItem
							Next objCurrencyNode
							
						End If
					Next objMessageNode
				End If
			Next objMsgsRsNode
		End If
		
	End Function
End Module