Attribute VB_Name = "CurrencyInfo"
' CurrencyInfo.bas
' This module is part of the Invoice sample program
' for the QuickBooks SDK Version CA2.0.
' Created September, 2002
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'-------------------------------------------------------------

Function GetCurExRate(strCurrencyID As String) As String
'This function create a request to QuickBooks to return information on a single currency
On Error GoTo ErrHandler
    
    
    'Call OpenConnection to establish a QuickBooks connection and session
    If Not blnIsOpenConnection Then
        If Not OpenConnection Then
         Exit Function
        End If
    End If
    
     
    ' This request ask QuickBooks to return information on a certain currency
    requestXML = "<?xml version=""1.0"" ?>" & _
    "<?qbxml version=""CA2.0""?>" & _
    "<QBXML><QBXMLMsgsRq onError=""continueOnError"">" & _
    "<CurrencyQueryRq requestID=""1"">" & _
    "<ListID>" & strCurrencyID & "</ListID>" & _
    "</CurrencyQueryRq>" & _
    "</QBXMLMsgsRq></QBXML>"
    
'    PrintXMLToFile requestXML, "C:\request.xml" '*** remove Comment character to produce a xml request file
        
    ' Send request to QuickBooks
    responseXML = qbXMLRP.ProcessRequest(ticket, requestXML)
'    PrintXMLToFile responseXML, "C:\response.xml"  '*** remove Comment character to produce a xml response file
    
    'Parse the response
    GetCurExRate = ParseCurrency(responseXML)
    
    Exit Function

ErrHandler:
    ' In case something goes wrong along the way..
    MsgBox "Error getting the currency rate from QuickBooks: " & Err.Description
    CloseConnection
End Function

Function ParseCurrency(sXML As Variant) As String
'This function parse the response from QuickBooks in order to find the Exchange Rate
    Dim objXmlDoc As New msxml2.DOMDocument40
    Dim objRootElement As IXMLDOMElement        ' root element of the XML document (QBXML Block)
    Dim objMsgsRsNode As IXMLDOMNode              ' QBXMLMsgsRs grouping (opening tag of message)
    Dim objMessageNode As IXMLDOMNode           ' Message node
    Dim objCurrencyNode As IXMLDOMNode            ' Currency node
    Dim objCurrencyItem As IXMLDOMNode            ' CurrencyItem node
        
        
    ' Load XML Script
    objXmlDoc.async = False
    If objXmlDoc.loadXML(sXML) = True And Len(objXmlDoc.xml) > 0 Then 'verify that there is some values in the documents
                    
        Set objRootElement = objXmlDoc.documentElement
        
        For Each objMsgsRsNode In objRootElement.childNodes
        
            If objMsgsRsNode.nodeName = "QBXMLMsgsRs" Then
            
                For Each objMessageNode In objMsgsRsNode.childNodes
                    
                    ' If we find the Currency, let's parse it.
                    If objMessageNode.nodeName = "CurrencyQueryRs" Then
                    
                        
                        For Each objCurrencyNode In objMessageNode.childNodes 'For each currency in the list(in this example, there is only one currency returned)
                            
                            For Each objCurrencyItem In objCurrencyNode.childNodes
                                If objCurrencyItem.nodeName = "ExchangeRate" Then
                                    ParseCurrency = objCurrencyItem.nodeTypedValue 'Add the exchange rate to the collection
                                    Exit For
                                End If
                            Next
                        Next
                        
                    End If
                Next
            End If
        Next
    End If

End Function
