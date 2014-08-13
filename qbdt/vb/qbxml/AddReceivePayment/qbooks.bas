Attribute VB_Name = "qbooks"
' qbooks.bas
'
' This module is part of the Receive Payments sample program
' for the QuickBooks SDK Version 2.0.  All of the communication
' with QuickBooks is handled in this file.
'
' In addition, this file contains examples of how to
' generate and parse qbXML using MSXML v4.0.  MSXML 4.0 must
' be installed on your computer in order to run the sample
' application.  It can be downloaded from microsoft.com.
'
' Created February, 2002
' Updated to SDK 2.0 July, 2002
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'-------------------------------------------------------------

Option Explicit

'Connection info
' qbXMLRP COM object
Public qbXMLRP As New RequestProcessor2
Public blnIsOpenConnection As Boolean

' Ticket
Public ticket As String

' Request and response strings
Public requestXML  As String
Public responseXML As String


Public Function OpenConnection() As Boolean
    
On Error GoTo ErrHandler
   
    If blnIsOpenConnection Then
        OpenConnection = True
        Exit Function
    End If
   
   
    ' If any of the call to qbXMLRP fails, error will be thrown.
    ' This error must be catched by the caller, in this case
    ' Comm_Submit_Click
    
    ' Open connection to qbXMLRP COM
    qbXMLRP.OpenConnection2 "", "IDN QuickBooks Sample Receive Payment", MainForm.connType
        
    ' Begin Session
    ' Pass empty string for the data file name to use the currently
    ' open data file.
    
On Error GoTo ErrHandlerBeginSession
    
    ticket = qbXMLRP.BeginSession("", qbFileOpenSingleUser)
    blnIsOpenConnection = True
    OpenConnection = True
    
    ' make sure ReceivePaymentAdd is supported
    Dim supportedVersion As String
    supportedVersion = qbXMLLatestVersion(qbXMLRP, ticket)
    If (supportedVersion < "1.1") Then
        MsgBox "The ReceivePaymentAdd request requires support for qbXML 1.1 which was released with R2P of QuickBooks 2002, please perform an update of QuickBooks", vbExclamation
        End
    End If
    
    Exit Function
        
ErrHandler:
    blnIsOpenConnection = False
    OpenConnection = False
    MsgBox Err.Description, vbExclamation, "Error"
    Exit Function

ErrHandlerBeginSession:
    blnIsOpenConnection = False
    OpenConnection = False
    qbXMLRP.CloseConnection
    MsgBox Err.Description, vbExclamation, "Error"
    Exit Function

End Function


Public Sub CloseConnection()
' Ends session and closes connection
    
If Not blnIsOpenConnection Then
        Exit Sub
End If

On Error GoTo ErrHandler
    
    ' End the session
    qbXMLRP.EndSession ticket
    
    ' Close the connection
    qbXMLRP.CloseConnection

    blnIsOpenConnection = False

    Exit Sub

ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error"
    Exit Sub

End Sub


' Once the user has selected a customer, we need to collect
' all of the open invoices for this customer and add them
' to the invoices list box using AddInvoiceToList().  We
' also need to add all credit memos, if there are any,
' to the credit memo list.
'
Public Sub displayLists(customer As String)
    On Error GoTo ErrHandler


    ' This call to genListsRequest will generate a block of qbXML
    ' querying for open invoices and credit memos in the name of
    ' the selected customer.
    requestXML = genListsRequest(customer)
    requestXML = qbXMLAddProlog(qbXMLLatestVersion(qbXMLRP, ticket), requestXML)
    
    'Call OpenConnection to establish a QuickBooks connection and session
    If Not blnIsOpenConnection Then
        If Not OpenConnection Then
         Exit Sub
        End If
    End If
    
    ' Send request to QuickBooks
    responseXML = qbXMLRP.ProcessRequest(ticket, requestXML)

    ' The response will contain elements for each open invoice or
    ' credit memo which the customer has.  This call to
    ' processListsResponse will store the important information
    ' from these elements and add each choice to the drop-down lists.

    processListsResponse responseXML, customer
    Exit Sub

ErrHandler:
    ' In case QuickBooks is not running, or something else goes wrong
    MsgBox "Error retrieving lists from QuickBooks: " & Err.Description
    CloseConnection
End Sub

Public Sub fillCustomerList()
    requestXML = genCustomerListRq
    requestXML = qbXMLAddProlog(qbXMLLatestVersion(qbXMLRP, ticket), requestXML)

    'Call OpenConnection to establish a QuickBooks connection and session
    If Not blnIsOpenConnection Then
        If Not OpenConnection Then
         Exit Sub
        End If
    End If
    
    ' Send request to QuickBooks
    responseXML = qbXMLRP.ProcessRequest(ticket, requestXML)

    processCustomerList (responseXML)
    
End Sub

Private Function genCustomerListRq() As String
    Dim doc As New DOMDocument40
    Dim top As IXMLDOMElement
    Dim query As IXMLDOMElement
    Dim request As IXMLDOMElement
    
    Set top = doc.createElement("QBXML")
    doc.appendChild top
    Set request = doc.createElement("QBXMLMsgsRq")
    Dim onErrorAttr As IXMLDOMAttribute
    Set onErrorAttr = doc.createAttribute("onError")
    onErrorAttr.Text = "continueOnError"
    request.Attributes.setNamedItem onErrorAttr
    top.appendChild request
    Set query = doc.createElement("CustomerQueryRq")
    Dim reqid As IXMLDOMAttribute
    Set reqid = doc.createAttribute("requestID")
    reqid.Text = "1"
    query.Attributes.setNamedItem reqid
    
    request.appendChild query
    
    genCustomerListRq = doc.xml
    
End Function

Private Sub processCustomerList(responseXML As String)
    Dim doc As New DOMDocument40
    doc.async = False
    doc.loadXML (responseXML)
    Dim customers As IXMLDOMNodeList
    Set customers = doc.getElementsByTagName("CustomerRet")
    Dim curCust As IXMLDOMNode
    Set curCust = customers.nextNode
    While Not curCust Is Nothing
        Dim FullName As IXMLDOMNode
        Set FullName = curCust.selectSingleNode("FullName")
        MainForm.AddCustomerToList (FullName.Text)
        Set curCust = customers.nextNode
    Wend
End Sub
' Once all of the proper error checking has been completed on the
' form and we have all of the information necessary to make a
' payment in QuickBooks, this subroutine will generate the qbXML
' request using ReceivePayment, then will submit this request to
' QuickBooks and parse the response for errors.
'
Public Sub SendPaymentToQB(customer As String, payDate As String, _
    refNum As String, payMethod As String, payAmt As Currency, _
    memoAmt As Currency, discountAmt As Currency, _
    invoiceTxnId As String, credTxnID As String)
    On Error GoTo ErrHandler
    
    
    'Call OpenConnection to establish a QuickBooks connection and session
    If Not blnIsOpenConnection Then
        If Not OpenConnection Then
         Exit Sub
        End If
    End If
    
     
    ' This call to genRecPaymentRequest generates the request qbXML
    ' telling QuickBooks to receive the payment.
    requestXML = genRecPaymentRequest(customer, payDate, refNum, _
        payMethod, payAmt, memoAmt, discountAmt, invoiceTxnId, _
        credTxnID)
    
    requestXML = qbXMLAddProlog(qbXMLLatestVersion(qbXMLRP, ticket), requestXML)
        
    ' Send request to QuickBooks
    responseXML = qbXMLRP.ProcessRequest(ticket, requestXML)
    
    
    ' Parse the response. The function processRecPaymentResponse takes the
    ' qbXML which QuickBooks has returned and parses it in order to find out
    ' whether the call was successful or whether there was an error.
    If Not processRecPaymentResponse(responseXML) Then
        CloseConnection
        Exit Sub
    End If
    

    Exit Sub

ErrHandler:
    ' In case something goes wrong along the way..
    MsgBox "Error sending payment to QuickBooks: " & Err.Description
    CloseConnection
End Sub


' Takes a customer and returns the qbXML request needed to
' receive a list of open invoices and credit memos for the
' customer from QuickBooks.  Notice the use of MSXML 4.0
' in creating the xml document.
'
Private Function genListsRequest(customer As String) _
    As String
    
    ' Requires MSXML v4.0
    Dim doc As New DOMDocument40
    
    Dim rootElement As IXMLDOMNode
    Dim msgsRqNode As IXMLDOMNode
    Dim invQueryNode As IXMLDOMNode
    Dim credQueryNode As IXMLDOMNode
     
    Dim paidStatNode As IXMLDOMNode
    Dim entityNode1 As IXMLDOMNode
    Dim fullNameNode1 As IXMLDOMNode
    Dim entityNode2 As IXMLDOMNode
    Dim fullNameNode2 As IXMLDOMNode
    Dim includeTxnNode As IXMLDOMNode
    
    Dim requestIDAttr1 As IXMLDOMAttribute
    Dim requestIDAttr2 As IXMLDOMAttribute
    Dim onErrorAttr As IXMLDOMAttribute
    
    Dim xmlRequest As String
    
    Set rootElement = doc.createElement("QBXML")
    doc.appendChild rootElement
    
    ' The following creates a block of XML which looks basically
    ' like this:
    
    ' <?xml version="1.0" ?>
    ' <!DOCTYPE QBXML PUBLIC '-//INTUIT//DTD QBXML QBD 2.0//EN'
    '                 'NEW_URL_WILL_GO_HERE'>
    ' <QBXML>
    '   <QBXMLMsgsRq onError="continueOnError">
    '       <InvoiceQueryRq requestID="1">
    '           <EntityFilter>
    '              <FullNameWithChildren>Customer Name Here</FullNameWithChildren>
    '           </EntityFilter>
    '           <PaidStatus>NotPaidOnly</PaidStatus>
    '       </InvoiceQueryRq>
    '       <CreditMemoQueryRq requestID="2">
    '           <EntityFilter>
    '              <FullNameWithChildren>Customer Name Here</FullNameWithChildren>
    '           </EntityFilter>
    '           <IncludeLinkedTxns>true</IncludeLinkedTxns>
    '       </CreditMemoQueryRq>
    '   </QBXMLMsgsRq>
    ' </QBXML>
    
    Set msgsRqNode = doc.createElement("QBXMLMsgsRq")
    rootElement.appendChild msgsRqNode
    
    Set onErrorAttr = doc.createAttribute("onError")
    onErrorAttr.Text = "continueOnError"
    msgsRqNode.Attributes.setNamedItem onErrorAttr
    
    ' Create the Invoice Query
    Set invQueryNode = doc.createElement("InvoiceQueryRq")
    msgsRqNode.appendChild invQueryNode
    
    Set requestIDAttr1 = doc.createAttribute("requestID")
    requestIDAttr1.Text = "1"
    invQueryNode.Attributes.setNamedItem requestIDAttr1
    
    ' Filter based on the current customer
    Set entityNode1 = doc.createElement("EntityFilter")
    invQueryNode.appendChild entityNode1
    
    ' Need FullNameWithChildren as opposed to just FullName
    ' This way, if the customer is "Cook, Brian", for example,
    ' the filter will pick up invoices for "Cook, Brian:Kitchen"
    ' and so on
    Set fullNameNode1 = doc.createElement("FullNameWithChildren")
    fullNameNode1.Text = customer
    entityNode1.appendChild fullNameNode1
    
    ' Filter out invoices that have already been paid
    Set paidStatNode = doc.createElement("PaidStatus")
    paidStatNode.Text = "NotPaidOnly"
    invQueryNode.appendChild paidStatNode
    
    ' Now create the Credit Memo Query
    Set credQueryNode = doc.createElement("CreditMemoQueryRq")
    msgsRqNode.appendChild credQueryNode
    
    Set requestIDAttr2 = doc.createAttribute("requestID")
    requestIDAttr2.Text = "2"
    credQueryNode.Attributes.setNamedItem requestIDAttr2
    
    ' Filter so we only receive credit memos for the
    ' chosen customer
    Set entityNode2 = doc.createElement("EntityFilter")
    credQueryNode.appendChild entityNode2
    
    ' Again we need FullNameWithChildren instead of just FullName
    Set fullNameNode2 = doc.createElement("FullNameWithChildren")
    fullNameNode2.Text = customer
    entityNode2.appendChild fullNameNode2
    
    ' We need a history of linked transactions on the credit memo
    ' in order to determine if it has been used up
    Set includeTxnNode = doc.createElement("IncludeLinkedTxns")
    includeTxnNode.Text = "true"
    credQueryNode.appendChild includeTxnNode
        
        
    ' Now xmlRequest contains the full request for filtered
    ' lists of invoices and credit memos
    genListsRequest = doc.xml

End Function


' Parses the response from QuickBooks, adding all of the
' new invoices to the invoice list and all of the
' credit memos to the credit memo list.  The bulk of
' this subroutine is actually taken up by variable
' declarations of MSXML DOM objects needed to parse
' the qbXML we have been handed back from QuickBooks.
'
Private Sub processListsResponse(outXML As String, _
    customer As String)

    ' The MSXML DOM document
    Dim doc As New DOMDocument40

    ' For status codes of the invoice query and credit memo query
    Dim statusCode1 As String
    Dim statusCode2 As String
    
    ' For corresponding status messages
    Dim statusMsg1 As String
    Dim statusMsg2 As String

    Dim numInvs As Integer      ' Will hold the number of open invoices
    Dim numCredMemos As Integer ' Will hold the number of credit memos
    Dim i As Integer            ' Will be a for-loop index
    Dim index As Integer        ' And another

    ' These are all needed to parse the InvoiceQueryRs
    Dim invNodeList As IXMLDOMNodeList
    Dim invQRsNode As IXMLDOMNode
    Dim iqrAttr As IXMLDOMNamedNodeMap
    Dim statCodeNode1 As IXMLDOMNode
    Dim statMsgNode1 As IXMLDOMNode
    Dim invoiceNodeList As IXMLDOMNodeList
            
    ' These will hold info about invoices extracted from the qbXML
    Dim txnId As String
    Dim appliedAmt As Currency
    Dim balanceRemaining As Currency
    Dim suggDiscount As Currency
    Dim invRefNum As String
            
    ' More MSXML objects needed to parse the invoice information
    Dim invoiceNode As IXMLDOMNode
    Dim custElement As IXMLDOMElement
    Dim txnIdElement As IXMLDOMElement
    Dim invRefElement As IXMLDOMElement
    Dim aaElement As IXMLDOMElement
    Dim brElement As IXMLDOMElement
    Dim disElement As IXMLDOMElement
    
    ' These are all needed to parse the CreditMemoRs
    Dim credQRsNode As IXMLDOMNode
    Dim cqrAttr As IXMLDOMNamedNodeMap
    Dim statCodeNode2 As IXMLDOMNode
    Dim statMsgNode2 As IXMLDOMNode
    
    ' These will hold information extracted from the
    ' qbXML about the particular credit memos
    Dim credTxnID As String
    Dim amount As Currency
    Dim amountLeft As Currency
    
    ' More MSXML DOM objects needed for extracting credit
    ' memo information from the qbXML
    Dim credNodeList As IXMLDOMNodeList
    Dim credMemoNodeList As IXMLDOMNodeList
    Dim credMemoNode As IXMLDOMNode
    Dim custElement2 As IXMLDOMElement
    Dim credTxnIdElement As IXMLDOMElement
    Dim amountElement As IXMLDOMElement
    Dim linkedTxnNodeList As IXMLDOMNodeList
    Dim nextTxnNode As IXMLDOMElement
    Dim amtTxnNode As IXMLDOMElement
    
    ' Needed later to test if various elements exist before
    ' trying to access them
    Dim testNodeList As IXMLDOMNodeList

    '
    ' Begin by creating the MSXML object representing the entire
    ' qbXML document, and storing lists of all of the
    ' InvoiceQueryRs and CreditMemoQueryRs nodes:
    '
            
    doc.async = False
    doc.loadXML (outXML)
    
    ' Get lists of InvoiceQueryRs nodes and CreditMemoQueryRs nodes
    Set invNodeList = doc.getElementsByTagName("InvoiceQueryRs")
    Set credNodeList = doc.getElementsByTagName("CreditMemoQueryRs")
            
    '
    ' Next, parse the InvoiceQuery response:
    '
    
    ' InvQRsNode will grab the InvoiceQueryRs element we need
    ' from the array of 1 element
    Set invQRsNode = invNodeList.Item(0)
    
    ' InvoiceQueryRs has four attributes:
    ' requestID, statusCode, statusSeverity, statusMessage
    Set iqrAttr = invQRsNode.Attributes
    
    ' We only want the statusCode attribute and the statusMessage
    If iqrAttr.length > 0 Then
        Set statCodeNode1 = iqrAttr.getNamedItem("statusCode")
        statusCode1 = statCodeNode1.nodeValue
        Set statMsgNode1 = iqrAttr.getNamedItem("statusMessage")
        statusMsg1 = statMsgNode1.nodeValue
    End If
        
    ' A statusCode of 0 indicates that there is no error.  We only want to
    ' continue on if there is no error in InvoiceQueryRs
    
    If statusCode1 < 1000 Then
        
        Set invoiceNodeList = doc.getElementsByTagName("InvoiceRet")
        numInvs = invoiceNodeList.length
        
        ' Allocate memory for storing the invoice information
        MainForm.SetInvoiceArrayLengths numInvs
        
        ' Iterate through invoices, picking out TxnIDs, reference numbers,
        ' balances, and other information we need to store
        For i = 0 To numInvs - 1
            Set invoiceNode = invoiceNodeList.Item(i)
            Set custElement = invoiceNode.selectSingleNode("CustomerRef//FullName")
                
            ' There should always be a TxnID element
            Set txnIdElement = invoiceNode.selectSingleNode("TxnID")
            txnId = txnIdElement.Text
                
            ' Get the invoice reference number if it exists
            Set testNodeList = invoiceNode.selectNodes("RefNumber")
            If 1 = testNodeList.length Then
                Set invRefElement = invoiceNode.selectSingleNode("RefNumber")
                invRefNum = invRefElement.Text
            Else
                invRefNum = 0
            End If
                
            ' Get the AppliedAmount element if it exists
            Set testNodeList = invoiceNode.selectNodes("AppliedAmount")
            If 1 = testNodeList.length Then
                Set aaElement = invoiceNode.selectSingleNode("AppliedAmount")
                appliedAmt = aaElement.Text
            Else
                appliedAmt = 0
            End If
                
            ' Get the Balance Remaining Element if it exists
            ' (This should always exist..)
            Set testNodeList = invoiceNode.selectNodes("BalanceRemaining")
            If 1 = testNodeList.length Then
                Set brElement = invoiceNode.selectSingleNode("BalanceRemaining")
                balanceRemaining = brElement.Text
            Else
                balanceRemaining = 0
            End If
                
            ' Get the SuggestedDiscountAmount element if it exists
            Set testNodeList = invoiceNode.selectNodes("SuggestedDiscountAmount")
            If 1 = testNodeList.length Then
                Set disElement = invoiceNode.selectSingleNode("SuggestedDiscountAmount")
                suggDiscount = disElement.Text
            Else
                suggDiscount = 0
            End If
            
            ' Add the invoice to the list of invoices the user can choose
            ' from on the form.  This also adds it to the global variables
            ' in MainForm.frm so we can retrieve it later.
            MainForm.AddInvoiceToList txnId, appliedAmt, balanceRemaining, _
                suggDiscount, invRefNum, i
        Next
        
    
        '
        ' Now parse the credit memo query response:
        '
    
        ' CredQRsNode will grab the CreditMemoQueryRs element we need
        ' from the array of 1 element
        Set credQRsNode = credNodeList.Item(0)
    
        ' CreditMemoQueryRs has four attributes:
        ' requestID, statusCode, statusSeverity, statusMessage
        Set cqrAttr = credQRsNode.Attributes
    
        ' We only want the statusCode attribute and the statusMessage
        If cqrAttr.length > 0 Then
            Set statCodeNode2 = cqrAttr.getNamedItem("statusCode")
            statusCode2 = statCodeNode2.nodeValue
            Set statMsgNode2 = cqrAttr.getNamedItem("statusMessage")
            statusMsg2 = statMsgNode2.nodeValue
        End If
        
        ' A statusCode of 0 indicates that there is no error.  We only want to
        ' continue on if there is no error in CreditMemoQueryRs
        
        If 0 = statusCode2 Then
            
            Set credMemoNodeList = doc.getElementsByTagName("CreditMemoRet")
            numCredMemos = credMemoNodeList.length
            MainForm.SetCreditMemoArrayLengths numCredMemos
            
            ' Iterate through items, picking out their names and other
            ' info
            For i = 0 To numCredMemos - 1
                Set credMemoNode = credMemoNodeList.Item(i)
                Set custElement2 = credMemoNode.selectSingleNode("CustomerRef//FullName")
                
                ' There should always be a TxnID element
                Set credTxnIdElement = credMemoNode.selectSingleNode("TxnID")
                credTxnID = credTxnIdElement.Text
                
                ' Get the total amount if this element exists
                Set testNodeList = credMemoNode.selectNodes("TotalAmount")
                If 1 = testNodeList.length Then
                    Set amountElement = credMemoNode.selectSingleNode("TotalAmount")
                    amount = amountElement.Text
                Else
                    amount = 0
                End If
                    
                ' If there are linked transactions attached, we need to add
                ' up how much they were for to subtract this from the total
                ' amount that the credit memo was originally worth
                Set linkedTxnNodeList = credMemoNode.selectNodes("LinkedTxn")
                amountLeft = amount
                
                ' Loop through each linked transaction and subtract the
                ' amount from the total amount to get the amount left
                If (linkedTxnNodeList.length > 0) Then
                    For index = 0 To (linkedTxnNodeList.length - 1)
                        Set nextTxnNode = linkedTxnNodeList.Item(index)
                        Set amtTxnNode = nextTxnNode.selectSingleNode("Amount")
                        ' Amount element is actually negative, so we add
                        amountLeft = amountLeft + amtTxnNode.Text
                    Next
                End If
                
                ' Add the credit memo to the list that the user can choose
                ' from on the form.  This also adds it to the global variables
                ' in MainForm.frm so we can retrieve it later.
                MainForm.AddCreditMemoToList credTxnID, amount, amountLeft, i
                
            Next
            
            ' In case the credit memo boxes were previously turned off
            MainForm.TurnOnCreditMemos
            
            MsgBox "The invoice and credit memo lists below have been populated. " _
                & "Fill out the form below and hit Apply Payment to make the " _
                & "payment for " & customer & "."

        Else
                
            ' If the second status code is 1, then QuickBooks has told
            ' us there are not any credit memos for the current customer.
            ' This is ok!
            If 1 = statusCode2 Then
                
                If (0 = numCredMemos) Then
                    MainForm.TurnOffCreditMemos
                End If
                
                MsgBox "The invoice list below has been populated. " & _
                    "Note: There are not currently credit memos for " _
                    & customer & ".  However, you may still receive a" _
                    & " payment from this customer.  Fill out the form " _
                    & "below and hit Apply Payment."
            
            Else
            
                ' Displays erros if statusCode is not 0 or 1
                MsgBox "Error from QuickBooks: " & statusMsg2
                
            End If
        
        End If
    
    Else
            ' Back to the first else from the invoice query...
            ' If the statusCode is 1, this means that no open invoices
            ' were found for the customer.  We should display a message
            ' about this, but not throw up an error
        If (1 = statusCode1) Then
        
            MsgBox "There are not currently open invoices for " & customer _
                & ".  Please select another customer."
        
        Else
        
            ' Displays errors if the statusCode is not equal to 0 or 1.
            MsgBox "Error from QuickBooks: " & statusMsg1
    
        End If
    
    End If


End Sub



' Generate the request qbXML which will send the payment to QuickBooks
' using the new ReceivePaymentAdd functionality.
'
Public Function genRecPaymentRequest(customer As String, _
    payDate As String, refNum As String, payMethod As String, payAmt As Currency, _
    memoAmt As Currency, discountAmt As Currency, _
    invoiceTxnId As String, credTxnID As String)

    Dim doc As New DOMDocument40
    
    Dim rootElement As IXMLDOMNode
    Dim msgsRqNode As IXMLDOMNode
    Dim recPayRqNode As IXMLDOMNode
    Dim recPayNode As IXMLDOMNode

    ' MSXML elements to hold attributes
    Dim requestIDAttr As IXMLDOMAttribute
    Dim onErrorAttr As IXMLDOMAttribute
    
    ' Elements which will hold information about the customer,
    ' the TxnID, and everything else entered on the form
    Dim custRefNode As IXMLDOMNode
    Dim custFullNameNode As IXMLDOMNode
    Dim txnDateNode As IXMLDOMNode
    Dim refNumNode As IXMLDOMNode
    Dim totAmtNode As IXMLDOMNode
    Dim payMethNode As IXMLDOMNode
    Dim methFullName As IXMLDOMNode
    Dim appliedToTxnNode As IXMLDOMNode
    Dim txnIdNode As IXMLDOMNode
    Dim payAmtNode As IXMLDOMNode

    ' Elements for the credit memo and discount amount
    Dim setCreditNode As IXMLDOMNode
    Dim creditTxnIdNode As IXMLDOMNode
    Dim appAmtNode As IXMLDOMNode
    Dim disAmtNode As IXMLDOMNode
    
    Dim xmlRequest As String
        
    
    ' <QBXML> ... </QBXML>
    Set rootElement = doc.createElement("QBXML")
    doc.appendChild rootElement
    
    ' <QBXMLMsgsRq onError=continueOnError> ... </QBXMLMsgsRq>
    Set msgsRqNode = doc.createElement("QBXMLMsgsRq")
    rootElement.appendChild msgsRqNode
    
    Set onErrorAttr = doc.createAttribute("onError")
    onErrorAttr.Text = "continueOnError"
    msgsRqNode.Attributes.setNamedItem onErrorAttr
       
    ' Create the ReceievePaymentAddRq element, and so on
    Set recPayRqNode = doc.createElement("ReceivePaymentAddRq")
    msgsRqNode.appendChild recPayRqNode
    
    Set requestIDAttr = doc.createAttribute("requestID")
    requestIDAttr.Text = "1"
    recPayRqNode.Attributes.setNamedItem requestIDAttr
    
    Set recPayNode = doc.createElement("ReceivePaymentAdd")
    recPayRqNode.appendChild recPayNode

    ' Now we have to add in all of the information we received
    ' from the user input, one node at a time -- Make sure these
    ' remain in the right order!!

    ' Elements for the Customer:
    Set custRefNode = doc.createElement("CustomerRef")
    recPayNode.appendChild custRefNode

    Set custFullNameNode = doc.createElement("FullName")
    custFullNameNode.Text = customer
    custRefNode.appendChild custFullNameNode
    
    ' Element for the Date:
    Set txnDateNode = doc.createElement("TxnDate")
    txnDateNode.Text = payDate
    recPayNode.appendChild txnDateNode
    
    ' Element for the reference number
    Set refNumNode = doc.createElement("RefNumber")
    refNumNode.Text = refNum
    recPayNode.appendChild refNumNode
    
    ' Element for total amount
    Set totAmtNode = doc.createElement("TotalAmount")
    totAmtNode.Text = currToString(payAmt)
    recPayNode.appendChild totAmtNode
        
    Set payMethNode = doc.createElement("PaymentMethodRef")
    recPayNode.appendChild payMethNode
    
    Set methFullName = doc.createElement("FullName")
    methFullName.Text = payMethod
    payMethNode.appendChild methFullName

    ' The AppliedToTxnAdd node contains information about the
    ' credit memos and discounts
    Set appliedToTxnNode = doc.createElement("AppliedToTxnAdd")
    recPayNode.appendChild appliedToTxnNode

    ' This TxnId identifies which invoice is being paid off
    Set txnIdNode = doc.createElement("TxnID")
    txnIdNode.Text = invoiceTxnId
    appliedToTxnNode.appendChild txnIdNode
    
    Set payAmtNode = doc.createElement("PaymentAmount")
    payAmtNode.Text = currToString(payAmt)
    appliedToTxnNode.appendChild payAmtNode
    
    ' If the user is applying a credit memo...
    If (memoAmt > 0) Then
    
        Set setCreditNode = doc.createElement("SetCredit")
        appliedToTxnNode.appendChild setCreditNode
    
        ' The TxnId identifies which credit memo is being used
        Set creditTxnIdNode = doc.createElement("CreditTxnID")
        creditTxnIdNode.Text = credTxnID
        setCreditNode.appendChild creditTxnIdNode
    
        ' This will set the amount of the credit memo
        Set appAmtNode = doc.createElement("AppliedAmount")
        appAmtNode.Text = currToString(memoAmt)
        setCreditNode.appendChild appAmtNode
    
    End If
    
    ' If the user is applying a discount
    If (discountAmt > 0) Then
    
        ' This will set the amount of the discount
        Set disAmtNode = doc.createElement("DiscountAmount")
        disAmtNode.Text = currToString(discountAmt)
        appliedToTxnNode.appendChild disAmtNode
    
    End If
                  
    ' Now xmlRequest contains the full request for filtered
    ' lists of invoices and credit memos
    genRecPaymentRequest = doc.xml

End Function


' Process the qbXML response and print out a success or error
' message about the received payment.  The parsing in this function
' looks only at whether the response returned with an error or
' with a good status.
'
Public Function processRecPaymentResponse(strResponseXML As String) As Boolean

    Dim doc As New DOMDocument40

    ' For status codes of the ReceivePaymentAdd call
    Dim statusCode As String
    
    ' For status message of ReceivePaymentAdd
    Dim statusMsg As String

    ' These are needed for the status code attributes later
    Dim rprAttr As IXMLDOMNamedNodeMap
    Dim statCodeNode As IXMLDOMNode
    Dim statMsgNode As IXMLDOMNode

    ' Variable recPayNodeList will hold all elements named
    ' ReceivePaymentRs, and recPayRsNode will hold the
    ' particular element
    Dim recPayNodeList As IXMLDOMNodeList
    Dim recPayRsNode As IXMLDOMNode
    
    
    doc.async = False
    doc.loadXML (strResponseXML)
        
    '
    ' PARSE THE RECEIVE PAYMENT RESPONSE:
    '
    
    ' RecPayRsNode will grab the ReceivePaymentRs element we need
    ' from the array of 1 element
    Set recPayNodeList = doc.getElementsByTagName("ReceivePaymentAddRs")
    Set recPayRsNode = recPayNodeList.Item(0)
    
    ' ReceievePaymentRs has four attributes:
    ' requestID, statusCode, statusSeverity, statusMessage
    Set rprAttr = recPayRsNode.Attributes
    
    ' We only want the statusCode attribute and the statusMessage
    If rprAttr.length > 0 Then
        Set statCodeNode = rprAttr.getNamedItem("statusCode")
        statusCode = statCodeNode.nodeValue
        Set statMsgNode = rprAttr.getNamedItem("statusMessage")
        statusMsg = statMsgNode.nodeValue
    End If
        
        
    ' A statusCode of 0 indicates that there is no error.
    ' Return a success message if this is the case.  Otherwise
    ' return an error message with the status message from QB.
    If 0 = statusCode Then
        processRecPaymentResponse = True
        MsgBox "The Receive Payment entry has been successfully posted to QuickBooks"
    Else
        processRecPaymentResponse = False
        MsgBox "Error from QuickBooks: " & _
            statusMsg
    End If

End Function


' Convert a standard date to a date that QuickBooks will accept
'
Public Function toQBDate(theDate As String) As String
    
    On Error GoTo ErrHandler
    toQBDate = Format(theDate, "yyyy-mm-dd")
    
    If toQBDate <> "" And Not IsDate(toQBDate) Then
        toQBDate = "error"
        Exit Function
    End If
    
    Exit Function

ErrHandler:
    toQBDate = "error"
End Function




' Converts a currency value to a string which QuickBooks can
' handle.  For example, if the amount is $6.50, this will convert
' it to 6.50 so it will not show up as 6.5.
'
Public Function currToString(curr As Currency) As String

    Dim newString As String
    Dim currStrings() As String
    
    newString = curr & "."
    currStrings = Split(newString, ".")

    If ("" = currStrings(1)) Then
        ' If the amount of money is an even dollar amount
        newString = currStrings(0) & ".00"
    Else
        newString = currStrings(0) & "." & currStrings(1)
    End If
    
    currToString = newString
    
End Function


' This subroutine is available for error checking.  It is sometimes
' useful to print the XML which QuickBooks returns to a file so that
' any problems can be uncovered easily.  Although this subroutine is
' not currently in use in the ReceievePayment sample code, it is
' encouraged that you add it in if you would like to see the precise
' XML that is being sent to or receieved from QuickBooks.
'
Private Sub PrettyPrintXMLToFile(xmlString As String, XMLFile As String)
                                       
  Dim SplitXMLString() As String
  Dim IndentString As String
  Dim xmlStringLength As Long
  Dim SplitIndex
  
  IndentString = ""
  
  Dim FileNum
  FileNum = FreeFile
  Open XMLFile For Output As FileNum
  
  ' Remove the linefeeds from the XML output string
  xmlString = Replace(xmlString, vbLf, vbNullString)
  
  SplitXMLString = Split(xmlString, "<")
  
  ' We're expecting the first character of the XML output to be "<"
  ' which result in an empty first array element, so skip it.
  SplitIndex = LBound(SplitXMLString) + 1
  
  Do
    If Left(SplitXMLString(SplitIndex), 1) = "/" Then
      IndentString = Left(IndentString, Len(IndentString) - 3)
      Print #FileNum, IndentString & "<" & _
                      SplitXMLString(SplitIndex)
      SplitIndex = SplitIndex + 1
    ElseIf Left(SplitXMLString(SplitIndex + 1), 1) = "/" Then
      If InStr(1, _
               Left(SplitXMLString(SplitIndex), _
                    InStr(1, SplitXMLString(SplitIndex), ">")), _
                " ") > 0 Then
        Print #FileNum, IndentString & "<" & _
                        SplitXMLString(SplitIndex)
        SplitIndex = SplitIndex + 1
      Else
        Print #FileNum, IndentString & "<" & _
                        SplitXMLString(SplitIndex) & "<" & _
                        SplitXMLString(SplitIndex + 1)
        SplitIndex = SplitIndex + 2
      End If
    Else
      Print #FileNum, IndentString & "<" & _
                      SplitXMLString(SplitIndex)
      IndentString = IndentString & "   "
      SplitIndex = SplitIndex + 1
    End If
  Loop Until SplitIndex >= UBound(SplitXMLString)
  
  If Left(SplitXMLString(UBound(SplitXMLString)), 1) = "/" Then
    IndentString = Left(IndentString, Len(IndentString) - 3)
  End If
  
  Print #FileNum, IndentString & "<" & _
                  SplitXMLString(UBound(SplitXMLString))
  
  Close FileNum
End Sub

Function qbXMLAddProlog(supportedVersion As String, xml As String) As String
    Dim qbXMLVersionSpec As String
    If (Val(supportedVersion) >= 2) Then
        qbXMLVersionSpec = "<?qbxml version=""" & supportedVersion & """?>"
    ElseIf (supportedVersion = "1.1") Then
        qbXMLVersionSpec = "<!DOCTYPE QBXML PUBLIC '-//INTUIT//DTD QBXML QBD " _
                           & supportedVersion & "//EN' 'http://developer.intuit.com'>"
    Else
        MsgBox "You are apparently running QuickBooks 2002 Release 1, we strongly recommend that you use QuickBooks' online update feature to obtain the latest fixes and enhancements", vbExclamation
        qbXMLVersionSpec = "<!DOCTYPE QBXML PUBLIC '-//INTUIT//DTD QBXML QBD " _
                           & supportedVersion & "//EN' 'http://developer.intuit.com'>"
    End If
    qbXMLAddProlog = "<?xml version=""1.0""?>" & vbCrLf & qbXMLVersionSpec & xml
End Function

Function qbXMLLatestVersion(rp As RequestProcessor2, ticket As String) As String
    Dim strXMLVersions() As String
    strXMLVersions = rp.QBXMLVersionsForSession(ticket)
    
    
    Dim i As Long
    Dim vers As String
    Dim LastVers As String
    LastVers = "0"
    For i = LBound(strXMLVersions) To UBound(strXMLVersions)
        vers = strXMLVersions(i)
        If (vers > LastVers) Then
            LastVers = vers
            qbXMLLatestVersion = vers
        End If
    Next i
End Function

