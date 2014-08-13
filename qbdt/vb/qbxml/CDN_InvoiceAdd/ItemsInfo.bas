Attribute VB_Name = "ItemsInfo"
' ItemsInfo.bas
' This module is part of the Invoice sample program
' for the QuickBooks SDK Version CA2.0.
'
' Created September, 2002
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'-------------------------------------------------------------

'These global variable contain the information associated with the item
'This public collection contain the item iformation Customer
Public colItemList As Collection

Sub GetItemList()
'The GetItemList function generates a request to QuickBooks and call the ParseItem function to parse the
'Item's list

On Error GoTo ErrHandler
    
    
    'Call OpenConnection to establish a QuickBooks connection and session
    If Not blnIsOpenConnection Then
        If Not OpenConnection Then
         Exit Sub
        End If
    End If
    
    'Create the query to find the Item list
    requestXML = "<?xml version=""1.0"" ?>" & _
    "<?qbxml version=""CA2.0""?>" & _
    "<QBXML><QBXMLMsgsRq onError=""continueOnError"">" & _
    "<ItemQueryRq requestID=""1""/>" & _
    "</QBXMLMsgsRq></QBXML>"
    
    'PrintToFile requestXML
'    PrintXMLToFile requestXML, "C:\request.xml" '*** remove comment character to produce a xml request file

    
    
        
    ' Send request to QuickBooks
    responseXML = qbXMLRP.ProcessRequest(ticket, requestXML)
'    PrintXMLToFile responseXML, "C:\response.xml" '*** remove Comment character to produce a xml response file
    
    'Call the Parse item Procedure in order to populate the colItemList collection
    ParseItem (responseXML)
    
    Exit Sub

ErrHandler:
    ' In case something goes wrong along the way..
    MsgBox "Error sending request to QuickBooks: " & Err.Description
    CloseConnection
End Sub
Sub ParseItem(sXML As Variant)
'This function parse the response to the ItemQuery request
'The parsed information will be inserted in global variables
'The only Items that are collected in this function are the Item Services, Item non Inventory and Item Other charges,
'The other Items are ignored
    Dim objXmlDoc As New msxml2.DOMDocument40
    Dim objRootElement As IXMLDOMElement        ' root element of the XML document (QBXML Block)
    Dim objMsgsRsNode As IXMLDOMNode              ' QBXMLMsgsRs grouping (opening tag of message)
    Dim objMessageNode As IXMLDOMNode           ' Message node
    Dim objItemNode As IXMLDOMNode            ' Item node
    Dim objItemGroupItem As IXMLDOMNode            ' Item Group node
    Dim objElementItem As IXMLDOMNode           ' Element node
    Dim objPurchaseItem As IXMLDOMNode           ' Purchase node
   
    Dim blnDescFound As Boolean
    Dim strCustomerCurrencyArray(4) As String
    'strCustomerCurrencyArray(0) is the name of the item
    'strCustomerCurrencyArray(1) is the taxcode associated to this currency
    'strCustomerCurrencyArray(2) is Description associated to this item
    'strCustomerCurrencyArray(3) is the price associated to this item
    'strCustomerCurrencyArray(4) tell us if it's a sale or a purchase desc and price
    
    ' Load XML Script
    objXmlDoc.async = False
    If objXmlDoc.loadXML(sXML) = True And Len(objXmlDoc.xml) > 0 Then 'verify that there is some values in the documents
                    
        Set objRootElement = objXmlDoc.documentElement
        
        For Each objMsgsRsNode In objRootElement.childNodes
        
            If objMsgsRsNode.nodeName = "QBXMLMsgsRs" Then
            
                For Each objMessageNode In objMsgsRsNode.childNodes
                    
                    ' If we find the Item list, let's parse it.
                    If objMessageNode.nodeName = "ItemQueryRs" Then
                    
                        
                        For Each objItemNode In objMessageNode.childNodes
                            strCustomerCurrencyArray(0) = ""
                            strCustomerCurrencyArray(1) = ""
                            strCustomerCurrencyArray(2) = ""
                            strCustomerCurrencyArray(3) = ""
                            strCustomerCurrencyArray(4) = ""
                            ' This if statement filter the items that we want to retrieve information about
                            If objItemNode.nodeName = "ItemServiceRet" Or objItemNode.nodeName = "ItemNonInventoryRet" Or objItemNode.nodeName = "ItemOtherChargeRet" Then
                                blnDescFound = False
                                For Each objItemGroupItem In objItemNode.childNodes
                                    If objItemGroupItem.nodeName = "Name" Then
                                        strCustomerCurrencyArray(0) = objItemGroupItem.nodeTypedValue
                                    ElseIf objItemGroupItem.nodeName = "TaxCodeRef" Then
                                        ' Parses the TaxCodeRef node
                                        For Each objElementItem In objItemGroupItem.childNodes
                                            If objElementItem.nodeName = "FullName" Then
                                                strCustomerCurrencyArray(1) = objElementItem.nodeTypedValue
                                            End If
                                        Next
                                    ElseIf objItemGroupItem.nodeName = "SalesAndPurchase" Then
                                        ' Parses the SalesAndPurchase node
                                        For Each objElementItem In objItemGroupItem.childNodes
                                            If objElementItem.nodeName = "SalesDesc" Then
                                                strCustomerCurrencyArray(2) = objElementItem.nodeTypedValue
                                                blnDescFound = True
                                            ElseIf objElementItem.nodeName = "SalesPrice" Then
                                                strCustomerCurrencyArray(3) = objElementItem.nodeTypedValue
                                                strCustomerCurrencyArray(4) = "SalesAndPurchase"
                                                Exit For
                                            End If
                                        Next
                                        
                                    ElseIf objItemGroupItem.nodeName = "SalesOrPurchase" Then
                                        ' Parses the SalesOrPurchase node
                                        For Each objElementItem In objItemGroupItem.childNodes
                                            If objElementItem.nodeName = "Desc" Then
                                                strCustomerCurrencyArray(2) = objElementItem.nodeTypedValue
                                                blnDescFound = True
                                            ElseIf objElementItem.nodeName = "Price" Then
                                                strCustomerCurrencyArray(3) = objElementItem.nodeTypedValue
                                                strCustomerCurrencyArray(4) = "SalesOrPurchase"
                                                Exit For
                                            End If
                                        Next
                                    End If
                                Next
                                colItemList.Add strCustomerCurrencyArray, strCustomerCurrencyArray(0)
                            End If
                        Next
                        
                    End If
                Next
                
            End If
                
        Next
    
    End If
End Sub




