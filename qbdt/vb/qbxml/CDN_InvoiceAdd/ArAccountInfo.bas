Attribute VB_Name = "ArAccountInfo"
' ArAccountInfo.bas
' This module is part of the Invoice sample program
' for the QuickBooks SDK Version CA2.0.
' Created September, 2002
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'-------------------------------------------------------------


'Public String variable that contains the currency associated to the accounts
Public colArAccount As Collection
Function GetArAccountList()
'The GetArAccountList function generates a request to QuickBooks and call the ParseArAccount procedure to parse the
'ArAccount's list
'On Error GoTo ErrHandler
    'Call OpenConnection to establish a QuickBooks connection and session
    If Not blnIsOpenConnection Then
        If Not OpenConnection Then
         Exit Function
        End If
    End If
    
     
    ' We generate the request qbXML in order to get the ArAccount list and filtering to have just the account receivables.
    requestXML = "<?xml version=""1.0"" ?>" & _
    "<?qbxml version=""CA2.0""?>" & _
    "<QBXML><QBXMLMsgsRq onError=""continueOnError"">" & _
    "<AccountQueryRq requestID=""1"">" & _
    "<AccountType>AccountsReceivable</AccountType>" & _
    "</AccountQueryRq>" & _
    "</QBXMLMsgsRq></QBXML>"
    
  
'    PrintXMLToFile requestXML, "C:\request.xml" '*** remove comment character to produce a xml request file
        
    ' Send request to QuickBooks
    responseXML = qbXMLRP.ProcessRequest(ticket, requestXML)
    
    'PrintXMLToFile responseXML, "C:\response.xml" '*** remove Comment character to produce a xml response file
    
    'The ParseArAccount function parses the response coming from QuickBooks and populates the colArAccount collection
    ParseArAccount (responseXML)
    Exit Function
ErrHandler:
    ' In case something goes wrong along the way..
    MsgBox "Error sending request to QuickBooks: " & Err.Description
    CloseConnection
End Function

Sub ParseArAccount(sXML As Variant)
'This function parse the response to the ArAccountQuery request
'The parsed information is save in the public colArAccount collection

    Dim objXmlDoc As New msxml2.DOMDocument40
    Dim objRootElement As IXMLDOMElement        ' root element of the XML document (QBXML Block)
    Dim objMsgsRsNode As IXMLDOMNode              ' QBXMLMsgsRs grouping (opening tag of message)
    Dim objMessageNode As IXMLDOMNode           ' Message node
    Dim objArAccountNode As IXMLDOMNode            ' ArAccount node
    Dim objArAccountItem As IXMLDOMNode            ' ArAccount node
    Dim objElementItem As IXMLDOMNode           ' Element node
   
    Dim strFullName As String
    Dim strArAccountArray(1) As String
    'strArAccountArray(0) is the name of the ArAccount
    'strArAccountArray(1) is the currencyref
    ' Load XML Script
    objXmlDoc.async = False
    If objXmlDoc.loadXML(sXML) = True And Len(objXmlDoc.xml) > 0 Then 'verify that there is some values in the documents
                    
        Set objRootElement = objXmlDoc.documentElement
        
        For Each objMsgsRsNode In objRootElement.childNodes
        
            If objMsgsRsNode.nodeName = "QBXMLMsgsRs" Then
            
                For Each objMessageNode In objMsgsRsNode.childNodes
                    
                    ' If we find the Account list, let's parse it.
                    If objMessageNode.nodeName = "AccountQueryRs" Then
                    
                        For Each objArAccountNode In objMessageNode.childNodes
                            strArAccountArray(0) = ""
                            strArAccountArray(1) = ""
                        
                            For Each objArAccountItem In objArAccountNode.childNodes
                                'Transfer the ArAccount information into the collection
                                If objArAccountItem.nodeName = "FullName" Then
                                    strFullName = objArAccountItem.nodeTypedValue
                                    strArAccountArray(0) = strFullName
                                ElseIf objArAccountItem.nodeName = "CurrencyRef" Then
                                    ' Parses the CurrencyRef node
                                    For Each objElementItem In objArAccountItem.childNodes
                                        If objElementItem.nodeName = "ListID" Then
                                        'Insert the ListID into the colArAccount collection.  One property of the a collection is to accept
                                        'a second parameter that is alphanumeric and will used as an identifier of the object.  The name of the ArAccount is used as
                                        ' the identifier

                                           strArAccountArray(1) = objElementItem.nodeTypedValue
                                           Exit For 'We have the list ID, go to the next element
                                        End If
                                    Next
                                End If
                            Next
                            colArAccount.Add strArAccountArray, strFullName
                        Next
                        
                    End If
                Next
            End If
        Next
    End If
End Sub





