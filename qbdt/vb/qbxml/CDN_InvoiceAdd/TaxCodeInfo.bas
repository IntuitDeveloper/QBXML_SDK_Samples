Attribute VB_Name = "TaxCodeInfo"
' TaxCodeInfo.bas
' This module is part of the Invoice sample program
' for the QuickBooks SDK Version CA2.0.
' Created September, 2002
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
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
    requestXML = "<?xml version=""1.0"" ?>" & _
    "<?qbxml version=""CA2.0""?>" & _
    "<QBXML><QBXMLMsgsRq onError=""continueOnError"">" & _
    "<TaxCodeQueryRq requestID=""1""/>" & _
    "</QBXMLMsgsRq></QBXML>"
    
    
'    PrintXMLToFile requestXML, "C:\request.xml" '*** remove Comment character to produce a xml request file
    
    ' Send request to QuickBooks
    responseXML = qbXMLRP.ProcessRequest(ticket, requestXML)
    
'    PrintXMLToFile responseXML, "C:\response.xml" '*** remove Comment character to produce a xml response file
    
    'The ParseTaxCode procedure parses the response coming from QuickBooks and populates the colTaxCodeList collection
    ParseTaxCode (responseXML)
    Exit Sub
ErrHandler:
    ' In case something goes wrong along the way..
    MsgBox "Error sending request to QuickBooks: " & Err.Description
    CloseConnection
End Sub

Sub ParseTaxCode(sXML As Variant)
'This function parse the response to the TaxCodeQuery request

    Dim objXmlDoc As New msxml2.DOMDocument40
    Dim objRootElement As IXMLDOMElement        ' root element of the XML document (QBXML Block)
    Dim objMsgsRsNode As IXMLDOMNode              ' QBXMLMsgsRs grouping (opening tag of message)
    Dim objMessageNode As IXMLDOMNode           ' Message node
    Dim objTaxCodeNode As IXMLDOMNode            ' TaxCode node
    Dim objTaxCodeItem As IXMLDOMNode            ' TaxCode node
    Dim objElementItem As IXMLDOMNode           ' Element node
   
    Dim strTaxCodeArray(2) As String
    'strTaxCodeArray(0) is TaxCode Name
    'strTaxCodeArray (1) is GST
    'strTaxCodeArray(2) is PST
    ' Load XML Script
    objXmlDoc.async = False
    If objXmlDoc.loadXML(sXML) = True And Len(objXmlDoc.xml) > 0 Then 'verify that there is some values in the documents
                    
        Set objRootElement = objXmlDoc.documentElement
        
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
                                
                            Next
                            colTaxCodeList.Add strTaxCodeArray, strTaxCodeArray(0)
                        Next

                    End If
                Next
            End If
        Next
    End If
End Sub



