Option Strict Off
Option Explicit On
Module CustomerInfo
    ' CustomerInfo.bas
    ' This module is part of the Invoice sample program
    ' for the QuickBooks SDK Version 2.0.
    ' Created September, 2002
    '
    ' Copyright © 2021-2022 Intuit Inc. All rights reserved.
    ' Use is subject to the terms specified at:
    '      http://developer.intuit.com/legal/devsite_tos.html
    '
    '-------------------------------------------------------------


    'This public collection contain the Currency ListIDs and names associated to the Customer
    Public colCustomerCurrencyList As Collection
	Sub GetCustomerList()
		'The GetCustomerList function generates a request to QuickBooks and call the ParseCustomer function to parse the
		'Customer's list
		On Error GoTo ErrHandler
		'Call OpenConnection to establish a QuickBooks connection and session
		If Not blnIsOpenConnection Then
			If Not OpenConnection Then
				Exit Sub
			End If
		End If


        ' We generate the request qbXML in order to get the customer list.
        requestXML = "<?xml version=""1.0"" ?>" & "<?qbxml version=""2.0""?>" & "<QBXML><QBXMLMsgsRq onError=""continueOnError"">" & "<CustomerQueryRq requestID=""1""/>" & "</QBXMLMsgsRq></QBXML>"


        '    PrintXMLToFile requestXML, "C:\request.xml" '*** remove Comment character to produce a xml request file

        ' Send request to QuickBooks
        responseXML = qbXMLRP.ProcessRequest(ticket, requestXML)
		
		'PrintXMLToFile responseXML, "C:\response.xml" '*** remove Comment character to produce a xml response file
		
		'The ParseCustomer procedure parses the response coming from QuickBooks and returns the customers list inside the colCustomerCurrencyList collection
		ParseCustomer((responseXML))
		Exit Sub
ErrHandler: 
		' In case something goes wrong along the way..
		MsgBox("Error sending request to QuickBooks: " & Err.Description)
		CloseConnection()
	End Sub
	
	Sub ParseCustomer(ByRef sXML As Object)
        'This function parse the response to the CustomerQuery request
        'Some info is store in a global variable and some are returned by the function
        Dim objXmlDoc As New MSXML2.DOMDocument60
        Dim objRootElement As MSXML2.IXMLDOMElement ' root element of the XML document (QBXML Block)
		Dim objMsgsRsNode As MSXML2.IXMLDOMNode ' QBXMLMsgsRs grouping (opening tag of message)
		Dim objMessageNode As MSXML2.IXMLDOMNode ' Message node
		Dim objCustomerNode As MSXML2.IXMLDOMNode ' Customer node
		Dim objCustomerItem As MSXML2.IXMLDOMNode ' Customer node
		Dim objElementItem As MSXML2.IXMLDOMNode ' Element node
		
		Dim strFullName As String
		Dim strCustomerCurrencyArray(1) As String
		'strCustomerCurrencyArray(0) is the customer's name
		'strCustomerCurrencyArray(1) is the currency ref
		
		' Load XML Script
		objXmlDoc.async = False
		
		If objXmlDoc.loadXML(sXML) = True And Len(objXmlDoc.xml) > 0 Then 'verify that there is some values in the documents
			
			objRootElement = objXmlDoc.documentElement
			
			For	Each objMsgsRsNode In objRootElement.childNodes
				
				If objMsgsRsNode.nodeName = "QBXMLMsgsRs" Then
					
					For	Each objMessageNode In objMsgsRsNode.childNodes
						
						' If we find the customer list, let's parse it.
						If objMessageNode.nodeName = "CustomerQueryRs" Then
							
							For	Each objCustomerNode In objMessageNode.childNodes ' Go through each customer nodes returned by the query
								strCustomerCurrencyArray(0) = ""
								strCustomerCurrencyArray(1) = ""
								
								For	Each objCustomerItem In objCustomerNode.childNodes
                                    'Transfer the customer information into the collection

                                    If objCustomerItem.nodeName = "FullName" Then
                                        
                                        strFullName = objCustomerItem.nodeTypedValue
                                        strCustomerCurrencyArray(0) = strFullName
                                    ElseIf objCustomerItem.nodeName = "CurrencyRef" Then
                                        ' Parses the CurrencyRef node
                                        For	Each objElementItem In objCustomerItem.childNodes
											If objElementItem.nodeName = "ListID" Then
												
												strCustomerCurrencyArray(1) = objElementItem.nodeTypedValue
												Exit For
											End If
										Next objElementItem
										
									End If
								Next objCustomerItem
								colCustomerCurrencyList.Add(strCustomerCurrencyArray, strFullName)
							Next objCustomerNode
							
						End If
					Next objMessageNode
				End If
			Next objMsgsRsNode
		End If
	End Sub
End Module