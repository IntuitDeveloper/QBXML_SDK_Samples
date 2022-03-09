Option Strict Off
Option Explicit On
Imports MSXML2

Module modQBXMLRP
    '----------------------------------------------------------
    ' Copyright © 2021-2022 Intuit Inc. All rights reserved.
    ' Use is subject to the terms specified at:
    '      http://developer.intuit.com/rdmgr/?ID=100
    '
    '----------------------------------------------------------


    
    Dim objRequestProcessor As QBXMLRP2Lib.RequestProcessor2

    Dim strTicket As String
    Dim booSessionBegun As Boolean

    Dim strqbXMLVersionLine As String


    Dim objSavedDOMDocument As New DOMDocument60
    Dim objSavedInvoiceRet As IXMLDOMNode

    Dim strSavedRequest As String

    Public Sub QBXMLRP_OpenConnectionBeginSession()

        On Error GoTo Errs

        objRequestProcessor = New QBXMLRP2Lib.RequestProcessor2

        objRequestProcessor.OpenConnection("", "IDN Invoice Modify Sample Application")
        strTicket = objRequestProcessor.BeginSession("", QBXMLRP2Lib.QBFileMode.qbFileOpenDoNotCare)
        booSessionBegun = True
        Exit Sub

Errs:
        If Err.Number = &H80040416 Then
            MsgBox("You must have QuickBooks running with a company" & vbCrLf & "file open to use this program.")
            objRequestProcessor.CloseConnection()
            End
        ElseIf Err.Number = &H80040422 Then
            MsgBox("This QuickBooks company file is open in single user mode and" & vbCrLf & "another application is already accessing it.  Please exit the" & vbCrLf & "other application and run this application again.")
            objRequestProcessor.CloseConnection()
            End
        ElseIf Err.Number = &H1AD Then
            MsgBox("It appears that the qbXML Request Processor has not" & vbCrLf & "been installed, indicating QuickBooks 2002 or later" & vbCrLf & "may not have been installed.  Please run this sample" & vbCrLf & "after installing QuickBooks 2003 and running the Upgrade.")
            End
        ElseIf Err.Number = &H1AE Then
            MsgBox("It appears that you have QuickBooks 2002 R1P installed." & vbCrLf & "This program requires QuickBooks 2003 to work.")
            End
        Else
            MsgBox(Err.Number & vbCrLf & Hex(Err.Number) & vbCrLf & Err.Description)
            End
        End If
    End Sub


    Public Sub QBXMLRP_EndSessionCloseConnection()
        If booSessionBegun Then
            objRequestProcessor.EndSession(strTicket)
            objRequestProcessor.CloseConnection()
        End If
    End Sub


    Function QBXMLRP_MaxVersionSupported() As String

        'This section contained obsolete code that has been stubbed out

        QBXMLRP_MaxVersionSupported = "15.0"
    End Function


    Public Sub QBXMLRP_FillComboBox(ByRef cmbComboBox As System.Windows.Forms.ComboBox, ByRef strQueryType As String, ByRef strNameElement As String, ByRef strFilter As String, ByRef booMarkGroupItems As Boolean)

        'Clear the combo box
        cmbComboBox.Items.Clear()

        Dim strSplits() As String
        strSplits = Split(strQueryType, ",")

        Dim strNameElementSplits() As String
        strNameElementSplits = Split(strNameElement, ",")

        Dim strXMLRequest As String
        Dim strXMLResponse As String

        Dim i As Short
        Dim objDOMResponseDoc As New DOMDocument60
        Dim objQueryRs As IXMLDOMNodeList
        Dim rsStatusAttr As IXMLDOMNamedNodeMap
        Dim strQueryStatus As String
        Dim QueryRsNode As IXMLDOMNode
        Dim NodeList As IXMLDOMNodeList
        Dim numItems As Integer
        Dim itemNode As IXMLDOMNode
        Dim j As Short
        For i = 0 To UBound(strSplits)
            strXMLRequest = "<?xml version=""1.0""?>" & "<?qbxml version=""15.0""?>" & "<QBXML><QBXMLMsgsRq onError=""stopOnError""><" & strSplits(i) & "QueryRq>" & strFilter & "</" & strSplits(i) & "QueryRq>" & "</QBXMLMsgsRq></QBXML>"

            strXMLResponse = objRequestProcessor.ProcessRequest(strTicket, strXMLRequest)

            'Set up a DOM document object to load the response into

            'Parse the response XML
            objDOMResponseDoc.async = False
            objDOMResponseDoc.loadXML(strXMLResponse)

            'Get the status for our query request
            objQueryRs = objDOMResponseDoc.getElementsByTagName(strSplits(i) & "QueryRs")

            rsStatusAttr = objQueryRs.item(0).attributes
            
            strQueryStatus = rsStatusAttr.getNamedItem("statusCode").nodeValue

            'If the status is bad, report it to the user
            If strQueryStatus <> "0" Then
                If strQueryStatus = "1" Then
                    Exit Sub
                Else
                    MsgBox(rsStatusAttr.getNamedItem("statusMessage").nodeValue)
                    Exit Sub
                End If
            End If

            'Parse the query response and add the query values to the combo box
            QueryRsNode = objQueryRs.item(0)

            NodeList = objDOMResponseDoc.getElementsByTagName(strSplits(i) & "Ret")

            numItems = NodeList.length

            If UBound(strNameElementSplits) > 0 Then
                strNameElement = strNameElementSplits(i)
            End If

            'Declare the XML objects outside of the loop
            For j = 0 To numItems - 1
                itemNode = NodeList.item(j)
                If strSplits(i) = "ItemGroup" And booMarkGroupItems Then
                    cmbComboBox.Items.Add(itemNode.selectSingleNode(strNameElement).text & " - Group Item")
                Else
                    cmbComboBox.Items.Add(itemNode.selectSingleNode(strNameElement).text)
                End If
            Next j
        Next i
    End Sub


    Public Sub QBXMLRP_FillInvoiceList(ByRef lstInvoices As System.Windows.Forms.ListBox, ByRef strRefNumber As String, ByRef strFromDateTime As String, ByRef strToDateTime As String, ByRef strDateQueryType As String, ByRef strDateMacro As String, ByRef strCustomerJob As String, ByRef booCustomerWithChildren As Boolean, ByRef strAccount As String, ByRef booAccountWithChildren As Boolean, ByRef strFromRefNumberRange As String, ByRef strToRefNumberRange As String, ByRef strRefNumberPiece As String, ByRef strRefNumberCriteria As String, ByRef strPaidStatus As String)

        Dim strXMLRequest As String

        Dim objDOMDocument As New MSXML2.DOMDocument60
        Dim objRootNode As MSXML2.IXMLDOMNode
        Dim objRequestNode As MSXML2.IXMLDOMNode
        Dim objElement As MSXML2.IXMLDOMElement
        Dim objAttribute As MSXML2.IXMLDOMAttribute
        Dim objInvoiceQueryNode As MSXML2.IXMLDOMNode
        Dim objDateTimeFilter As MSXML2.IXMLDOMNode
        Dim objDateFilter As MSXML2.IXMLDOMNode
        Dim objEntityFilter As MSXML2.IXMLDOMNode
        Dim objAccountFilter As MSXML2.IXMLDOMNode
        Dim objRefNumberRangeFilter As MSXML2.IXMLDOMNode
        Dim objRefNumberFilter As MSXML2.IXMLDOMNode
        If strRefNumber <> "" Then
            'We only need to query for the ref number, don't bother building
            'the XML with MSXML
            strXMLRequest = "<?xml version=""1.0""?>" & "<?qbxml version=""15.0""?>" & "<QBXML><QBXMLMsgsRq onError=""stopOnError""><InvoiceQueryRq>" & "<RefNumber>" & strRefNumber & "</RefNumber>" & "<IncludeLineItems>true</IncludeLineItems></InvoiceQueryRq>" & "</QBXMLMsgsRq></QBXML>"
        Else 'strRefNumber <> ""


            
            CreateStandardRequestNode(False, "continueOnError", objDOMDocument, objRootNode, objRequestNode, objAttribute)

            
            AddMSXMLNode("InvoiceQueryRq", objDOMDocument, objRequestNode, objInvoiceQueryNode)

            'Limit our response to 30 invoices
            
            AddMSXMLElement("MaxReturned", "30", objDOMDocument, objInvoiceQueryNode, objElement)

            
            If Not String.IsNullOrEmpty(strFromDateTime) Or IsNothing(strToDateTime) Then 'now not added
                
                AddMSXMLNode(strDateQueryType, objDOMDocument, objInvoiceQueryNode, objDateTimeFilter)
                If strDateQueryType = "ModifiedDateRangeFilter" Then
                    
                    If Not IsNothing(strFromDateTime) Then
                        
                        AddMSXMLElement("FromModifiedDate", strFromDateTime, objDOMDocument, objDateTimeFilter, objElement)
                    End If
                    
                    If Not String.IsNullOrEmpty(strToDateTime) Then
                        
                        AddMSXMLElement("ToModifiedDate", strToDateTime, objDOMDocument, objDateTimeFilter, objElement)
                    End If
                Else 'strDateQueryType = "ModifiedDateRangeFilter"
                    
                    If Not String.IsNullOrEmpty(strFromDateTime) Then
                        
                        AddMSXMLElement("FromTxnDate", strFromDateTime, objDOMDocument, objDateTimeFilter, objElement)
                    End If
                    
                    If Not String.IsNullOrEmpty(strToDateTime) Then
                        
                        AddMSXMLElement("ToTxnDate", strToDateTime, objDOMDocument, objDateTimeFilter, objElement)
                    End If
                End If 'strDateQueryType = "ModifiedDateRangeFilter"
            End If 'strFromDateTime <> Empty Or strToDateTime <> Empty

            
            If Not String.IsNullOrEmpty(strDateMacro) Then
                
                AddMSXMLNode("TxnDateRangeFilter", objDOMDocument, objInvoiceQueryNode, objDateFilter)
                
                AddMSXMLElement("DateMacro", strDateMacro, objDOMDocument, objDateFilter, objElement)
            End If

            
            If Not String.IsNullOrEmpty(strCustomerJob) Then
                
                AddMSXMLNode("EntityFilter", objDOMDocument, objInvoiceQueryNode, objEntityFilter)

                If booCustomerWithChildren Then
                    
                    AddMSXMLElement("FullNameWithChildren", strCustomerJob, objDOMDocument, objEntityFilter, objElement)
                Else
                    
                    AddMSXMLElement("FullName", strCustomerJob, objDOMDocument, objEntityFilter, objElement)
                End If
            End If

            
            If Not String.IsNullOrEmpty(strAccount) Then
                
                AddMSXMLNode("AccountFilter", objDOMDocument, objInvoiceQueryNode, objAccountFilter)

                If booAccountWithChildren Then
                    
                    AddMSXMLElement("FullNameWithChildren", strAccount, objDOMDocument, objAccountFilter, objElement)
                Else
                    
                    AddMSXMLElement("FullName", strAccount, objDOMDocument, objAccountFilter, objElement)
                End If
            End If

            
            If Not String.IsNullOrEmpty(strFromRefNumberRange) Or Not String.IsNullOrEmpty(strToRefNumberRange) Then
                
                AddMSXMLNode("RefNumberRangeFilter", objDOMDocument, objInvoiceQueryNode, objRefNumberRangeFilter)
                
                If Not String.IsNullOrEmpty(strFromRefNumberRange) Then
                    
                    AddMSXMLElement("FromRefNumber", strFromRefNumberRange, objDOMDocument, objRefNumberRangeFilter, objElement)
                End If
                
                If Not String.IsNullOrEmpty(strToRefNumberRange) Then
                    
                    AddMSXMLElement("ToRefNumber", strToRefNumberRange, objDOMDocument, objRefNumberRangeFilter, objElement)
                End If
            End If 'strFromRefNumberRange <> Empty Or strToRefNumberRange <> Empty

            
            If Not String.IsNullOrEmpty(strRefNumberPiece) Then
                
                AddMSXMLNode("RefNumberFilter", objDOMDocument, objInvoiceQueryNode, objRefNumberFilter)
                
                AddMSXMLElement("MatchCriterion", strRefNumberCriteria, objDOMDocument, objRefNumberFilter, objElement)
                
                AddMSXMLElement("RefNumber", strRefNumberPiece, objDOMDocument, objRefNumberFilter, objElement)
            End If

            
            If Not String.IsNullOrEmpty(strPaidStatus) Then
                
                AddMSXMLElement("PaidStatus", strPaidStatus, objDOMDocument, objInvoiceQueryNode, objElement)
            End If

            
            AddMSXMLElement("IncludeLineItems", "true", objDOMDocument, objInvoiceQueryNode, objElement)

            strXMLRequest = "<?xml version=""1.0""?>" & "<?qbxml version=""15.0""?>" & objRootNode.xml
        End If 'strRefNumber <> ""
        PrettyPrintXMLToFile(strXMLRequest, "C:\debugrq.xml")
        strSavedRequest = PrettyPrintXMLToString(strXMLRequest)

        Dim strXMLResponse As String
        strXMLResponse = objRequestProcessor.ProcessRequest(strTicket, strXMLRequest)
        PrettyPrintXMLToFile(strXMLResponse, "C:\debugrs.xml")

        objDOMDocument.async = False
        objDOMDocument.loadXML(strXMLResponse)

        Dim objInvoiceQueryNodeList As IXMLDOMNodeList
        objInvoiceQueryNodeList = objDOMDocument.getElementsByTagName("InvoiceQueryRs")

        objInvoiceQueryNode = objInvoiceQueryNodeList.item(0)

        Dim objInvoiceQueryAttributes As IXMLDOMNamedNodeMap
        objInvoiceQueryAttributes = objInvoiceQueryNode.attributes

        If objInvoiceQueryAttributes.getNamedItem("statusCode").nodeValue <> "0" Then
            MsgBox("Error getting Invoices" & vbCrLf & vbCrLf & "Error = " & objInvoiceQueryAttributes.getNamedItem("statusCode").nodeValue & vbCrLf & vbCrLf & "Message = " & objInvoiceQueryAttributes.getNamedItem("statusMessage").nodeValue)

            lstInvoices.Items.Add("No invoices match the query filter used")
            Exit Sub
        End If

        Dim objInvoiceRetNodeList As IXMLDOMNodeList
        objInvoiceRetNodeList = objDOMDocument.getElementsByTagName("InvoiceRet")

        Dim objInvoiceRet As IXMLDOMNode
        Dim objNodeList As IXMLDOMNodeList
        Dim intItems As Short
        Dim strReturnedRefNumber As String
        Dim strItems As String

        Dim i As Short
        For i = 0 To objInvoiceRetNodeList.length - 1
            objInvoiceRet = objInvoiceRetNodeList.item(i)

            If Not objInvoiceRet.selectSingleNode("RefNumber") Is Nothing Then
                strReturnedRefNumber = "Invoice " & objInvoiceRet.selectSingleNode("RefNumber").text
            Else
                strReturnedRefNumber = "Un-numbered "
            End If

            objNodeList = objInvoiceRet.selectNodes("InvoiceLineRet")
            If objNodeList Is Nothing Then
                intItems = 0
            Else
                intItems = objNodeList.length
            End If
            objNodeList = objInvoiceRet.selectNodes("InvoiceLineGroupRet")
            If Not objNodeList Is Nothing Then
                intItems = intItems + objNodeList.length
            End If

            strItems = Str(intItems)
            If Len(strItems) = 1 Then strItems = "  " & strItems
            If Len(strItems) = 2 Then strItems = " " & strItems
            lstInvoices.Items.Add(strReturnedRefNumber & "     " & objInvoiceRet.selectSingleNode("TxnDate").text & "     " & strItems & " items     " & objInvoiceRet.selectSingleNode("CustomerRef").selectSingleNode("FullName").text & "     Balance " & objInvoiceRet.selectSingleNode("BalanceRemaining").text & "     " & objInvoiceRet.selectSingleNode("TxnID").text)
        Next i
    End Sub


    Public Sub QBXMLRP_GetInvoice(ByRef TxnID As String)

        Dim strXMLRequest As String

        'We only need to query for the TxnID, don't bother building
        'the XML with MSXML
        strXMLRequest = "<?xml version=""1.0""?>" & "<?qbxml version=""15.0""?>" & "<QBXML><QBXMLMsgsRq onError=""stopOnError""><InvoiceQueryRq>" & "<TxnID>" & TxnID & "</TxnID><IncludeLineItems>1</IncludeLineItems>" & "</InvoiceQueryRq></QBXMLMsgsRq></QBXML>"

        Dim strXMLResponse As String
        strXMLResponse = objRequestProcessor.ProcessRequest(strTicket, strXMLRequest)

        objSavedDOMDocument.async = False
        objSavedDOMDocument.loadXML(strXMLResponse)

        Dim objInvoiceQueryNodeList As IXMLDOMNodeList
        objInvoiceQueryNodeList = objSavedDOMDocument.getElementsByTagName("InvoiceQueryRs")

        Dim objInvoiceQueryNode As IXMLDOMNode
        objInvoiceQueryNode = objInvoiceQueryNodeList.item(0)

        Dim objInvoiceQueryAttributes As IXMLDOMNamedNodeMap
        objInvoiceQueryAttributes = objInvoiceQueryNode.attributes

        If objInvoiceQueryAttributes.getNamedItem("statusCode").nodeValue <> "0" Then
            MsgBox("Error getting Invoices" & vbCrLf & vbCrLf & "Error = " & objInvoiceQueryAttributes.getNamedItem("statusCode").nodeValue & vbCrLf & vbCrLf & "Message = " & objInvoiceQueryAttributes.getNamedItem("statusMessage").nodeValue)
            Exit Sub
        End If

        Dim objInvoiceRetNodeList As IXMLDOMNodeList
        objInvoiceRetNodeList = objSavedDOMDocument.getElementsByTagName("InvoiceRet")

        objSavedInvoiceRet = objInvoiceRetNodeList.item(0)
    End Sub


    Public Sub QBXMLRP_LoadInvoiceModifyForm()
        frmInvoiceModify.txtTxnID.Text = objSavedInvoiceRet.selectSingleNode("TxnID").text

        frmInvoiceModify.txtEditSequence.Text = objSavedInvoiceRet.selectSingleNode("EditSequence").text

        If Not objSavedInvoiceRet.selectSingleNode("RefNumber") Is Nothing Then
            frmInvoiceModify.txtRefNumber.Text = objSavedInvoiceRet.selectSingleNode("RefNumber").text
        End If

        frmInvoiceModify.txtTxnDate.Text = objSavedInvoiceRet.selectSingleNode("TxnDate").text

        If Not objSavedInvoiceRet.selectSingleNode("IsPending") Is Nothing Then
            If objSavedInvoiceRet.selectSingleNode("IsPending").text = "true" Then
                frmInvoiceModify.chkPending.CheckState = System.Windows.Forms.CheckState.Checked 'Checked
            End If
        End If

        If Not objSavedInvoiceRet.selectSingleNode("IsToBePrinted") Is Nothing Then
            If objSavedInvoiceRet.selectSingleNode("IsToBePrinted").text = "true" Then
                frmInvoiceModify.chkToBePrinted.CheckState = System.Windows.Forms.CheckState.Checked 'Checked
            End If
        End If

        frmInvoiceModify.cmbCustomer.Text = objSavedInvoiceRet.selectSingleNode("CustomerRef").selectSingleNode("FullName").text

        If Not objSavedInvoiceRet.selectSingleNode("ClassRef") Is Nothing Then
            frmInvoiceModify.cmbClass.Text = objSavedInvoiceRet.selectSingleNode("ClassRef").selectSingleNode("FullName").text
        End If

        If Not objSavedInvoiceRet.selectSingleNode("BillAddress") Is Nothing Then
            If Not objSavedInvoiceRet.selectSingleNode("BillAddress").selectSingleNode("Addr1") Is Nothing Then
                frmInvoiceModify.txtBillAddr1.Text = objSavedInvoiceRet.selectSingleNode("BillAddress").selectSingleNode("Addr1").text
            End If

            If Not objSavedInvoiceRet.selectSingleNode("BillAddress").selectSingleNode("Addr2") Is Nothing Then
                frmInvoiceModify.txtBillAddr2.Text = objSavedInvoiceRet.selectSingleNode("BillAddress").selectSingleNode("Addr2").text
            End If

            If Not objSavedInvoiceRet.selectSingleNode("BillAddress").selectSingleNode("Addr3") Is Nothing Then
                frmInvoiceModify.txtBillAddr3.Text = objSavedInvoiceRet.selectSingleNode("BillAddress").selectSingleNode("Addr3").text
            End If

            If Not objSavedInvoiceRet.selectSingleNode("BillAddress").selectSingleNode("Addr4") Is Nothing Then
                frmInvoiceModify.txtBillAddr4.Text = objSavedInvoiceRet.selectSingleNode("BillAddress").selectSingleNode("Addr4").text
            End If

            If Not objSavedInvoiceRet.selectSingleNode("BillAddress").selectSingleNode("City") Is Nothing Then
                frmInvoiceModify.txtBillCity.Text = objSavedInvoiceRet.selectSingleNode("BillAddress").selectSingleNode("City").text
            End If

            If Not objSavedInvoiceRet.selectSingleNode("BillAddress").selectSingleNode("State") Is Nothing Then
                frmInvoiceModify.txtBillState.Text = objSavedInvoiceRet.selectSingleNode("BillAddress").selectSingleNode("State").text
            End If

            If Not objSavedInvoiceRet.selectSingleNode("BillAddress").selectSingleNode("PostalCode") Is Nothing Then
                frmInvoiceModify.txtBillPostalCode.Text = objSavedInvoiceRet.selectSingleNode("BillAddress").selectSingleNode("PostalCode").text
            End If

            If Not objSavedInvoiceRet.selectSingleNode("BillAddress").selectSingleNode("Country") Is Nothing Then
                frmInvoiceModify.txtBillCountry.Text = objSavedInvoiceRet.selectSingleNode("BillAddress").selectSingleNode("Country").text
            End If
        End If

        If Not objSavedInvoiceRet.selectSingleNode("ShipAddress") Is Nothing Then
            If Not objSavedInvoiceRet.selectSingleNode("ShipAddress").selectSingleNode("Addr1") Is Nothing Then
                frmInvoiceModify.txtShipAddr1.Text = objSavedInvoiceRet.selectSingleNode("ShipAddress").selectSingleNode("Addr1").text
            End If

            If Not objSavedInvoiceRet.selectSingleNode("ShipAddress").selectSingleNode("Addr2") Is Nothing Then
                frmInvoiceModify.txtShipAddr2.Text = objSavedInvoiceRet.selectSingleNode("ShipAddress").selectSingleNode("Addr2").text
            End If

            If Not objSavedInvoiceRet.selectSingleNode("ShipAddress").selectSingleNode("Addr3") Is Nothing Then
                frmInvoiceModify.txtShipAddr3.Text = objSavedInvoiceRet.selectSingleNode("ShipAddress").selectSingleNode("Addr3").text
            End If

            If Not objSavedInvoiceRet.selectSingleNode("ShipAddress").selectSingleNode("Addr4") Is Nothing Then
                frmInvoiceModify.txtShipAddr4.Text = objSavedInvoiceRet.selectSingleNode("ShipAddress").selectSingleNode("Addr4").text
            End If

            If Not objSavedInvoiceRet.selectSingleNode("ShipAddress").selectSingleNode("City") Is Nothing Then
                frmInvoiceModify.txtShipCity.Text = objSavedInvoiceRet.selectSingleNode("ShipAddress").selectSingleNode("City").text
            End If

            If Not objSavedInvoiceRet.selectSingleNode("ShipAddress").selectSingleNode("State") Is Nothing Then
                frmInvoiceModify.txtShipState.Text = objSavedInvoiceRet.selectSingleNode("ShipAddress").selectSingleNode("State").text
            End If

            If Not objSavedInvoiceRet.selectSingleNode("ShipAddress").selectSingleNode("PostalCode") Is Nothing Then
                frmInvoiceModify.txtShipPostalCode.Text = objSavedInvoiceRet.selectSingleNode("ShipAddress").selectSingleNode("PostalCode").text
            End If

            If Not objSavedInvoiceRet.selectSingleNode("ShipAddress").selectSingleNode("Country") Is Nothing Then
                frmInvoiceModify.txtShipCountry.Text = objSavedInvoiceRet.selectSingleNode("ShipAddress").selectSingleNode("Country").text
            End If
        End If

        If Not objSavedInvoiceRet.selectSingleNode("ARAccountRef") Is Nothing Then
            frmInvoiceModify.cmbARAccount.Text = objSavedInvoiceRet.selectSingleNode("ARAccountRef").selectSingleNode("FullName").text
        End If

        If Not objSavedInvoiceRet.selectSingleNode("TermsRef") Is Nothing Then
            frmInvoiceModify.cmbTerms.Text = objSavedInvoiceRet.selectSingleNode("TermsRef").selectSingleNode("FullName").text
        End If

        If Not objSavedInvoiceRet.selectSingleNode("PONumber") Is Nothing Then
            frmInvoiceModify.txtPONumber.Text = objSavedInvoiceRet.selectSingleNode("PONumber").text
        End If

        If Not objSavedInvoiceRet.selectSingleNode("DueDate") Is Nothing Then
            frmInvoiceModify.txtDueDate.Text = objSavedInvoiceRet.selectSingleNode("DueDate").text
        End If

        If Not objSavedInvoiceRet.selectSingleNode("ShipDate") Is Nothing Then
            frmInvoiceModify.txtShipDate.Text = objSavedInvoiceRet.selectSingleNode("ShipDate").text
        End If

        If Not objSavedInvoiceRet.selectSingleNode("FOB") Is Nothing Then
            frmInvoiceModify.txtFOB.Text = objSavedInvoiceRet.selectSingleNode("FOB").text
        End If

        If Not objSavedInvoiceRet.selectSingleNode("SalesRepRef") Is Nothing Then
            frmInvoiceModify.cmbSalesRep.Text = objSavedInvoiceRet.selectSingleNode("SalesRepRef").selectSingleNode("FullName").text
        End If

        If Not objSavedInvoiceRet.selectSingleNode("ShipMethodRef") Is Nothing Then
            frmInvoiceModify.cmbShipMethod.Text = objSavedInvoiceRet.selectSingleNode("ShipMethodRef").selectSingleNode("FullName").text
        End If

        If Not objSavedInvoiceRet.selectSingleNode("ItemSalesTaxRef") Is Nothing Then
            frmInvoiceModify.cmbItemSalesTax.Text = objSavedInvoiceRet.selectSingleNode("ItemSalesTaxRef").selectSingleNode("FullName").text
        End If

        If Not objSavedInvoiceRet.selectSingleNode("CustomerSalesTaxCodeRef") Is Nothing Then
            frmInvoiceModify.cmbCustTaxCode.Text = objSavedInvoiceRet.selectSingleNode("CustomerSalesTaxCodeRef").selectSingleNode("FullName").text
        End If

        If Not objSavedInvoiceRet.selectSingleNode("CustomerMsgRef") Is Nothing Then
            frmInvoiceModify.cmbCustomerMsg.Text = objSavedInvoiceRet.selectSingleNode("CustomerMsgRef").selectSingleNode("FullName").text
        End If

        If Not objSavedInvoiceRet.selectSingleNode("Memo") Is Nothing Then
            frmInvoiceModify.txtMemo.Text = objSavedInvoiceRet.selectSingleNode("Memo").text
        End If
    End Sub


    Public Sub QBXMLRP_LoadInvoiceLineArray(ByRef strLineArray() As String)

        Dim objNodeList As IXMLDOMNodeList
        objNodeList = objSavedInvoiceRet.selectNodes("InvoiceLineRet")

        Dim objNode As IXMLDOMNode
        Dim strFoo As String
        If objNodeList.length > 0 Then
            objNode = objNodeList.item(0)
            strFoo = objNode.previousSibling.nodeName
            If strFoo = "InvoiceLineGroupRet" Then
                objNodeList = objSavedInvoiceRet.selectNodes("InvoiceLineGroupRet")
            End If
        Else
            objNodeList = objSavedInvoiceRet.selectNodes("InvoiceLineGroupRet")
        End If

        objNode = objNodeList.item(0)
        Dim objGroupLines As IXMLDOMNodeList
        Dim objLine As IXMLDOMNode
        Dim i As Short
        Dim j As Short
        i = 1
        Do While objNode.nodeName = "InvoiceLineRet" Or objNode.nodeName = "InvoiceLineGroupRet"
            If objNode.nodeName = "InvoiceLineRet" Then
                strLineArray(i) = LineInfo(objNode) & "Item<spliter>Original"
            Else
                strLineArray(i) = LineInfo(objNode) & "Group<spliter>Original"

                objGroupLines = objNode.selectNodes("InvoiceLineRet")
                If objGroupLines.length > 0 Then
                    For j = 0 To objGroupLines.length - 1
                        i = i + 1
                        objLine = objGroupLines.item(j)
                        strLineArray(i) = LineInfo(objLine) & "SubItem<spliter>Original"
                    Next j
                End If 'objGroupLines.length > 0
            End If 'objNode.nodeName = "InvoiceLineRet"

            If objNode.nextSibling Is Nothing Then
                objNode = objSavedInvoiceRet.selectSingleNode("TxnID")
            Else
                objNode = objNode.nextSibling
                i = i + 1
            End If
        Loop
    End Sub


    Private Function LineInfo(ByRef objNode As IXMLDOMNode) As String

        Dim strLineInfo As String
        Dim strRateOrPercent As String

        strLineInfo = objNode.selectSingleNode("TxnLineID").text & "<spliter>"

        If Not objNode.selectSingleNode("Quantity") Is Nothing Then
            strLineInfo = strLineInfo & objNode.selectSingleNode("Quantity").text
        End If
        strLineInfo = strLineInfo & "<spliter>"

        If Not objNode.selectSingleNode("ItemRef") Is Nothing Then
            strLineInfo = strLineInfo & objNode.selectSingleNode("ItemRef").selectSingleNode("FullName").text
        ElseIf Not objNode.selectSingleNode("ItemGroupRef") Is Nothing Then
            strLineInfo = strLineInfo & objNode.selectSingleNode("ItemGroupRef").selectSingleNode("FullName").text
        End If
        strLineInfo = strLineInfo & "<spliter>"

        If Not objNode.selectSingleNode("Desc") Is Nothing Then
            strLineInfo = strLineInfo & objNode.selectSingleNode("Desc").text
        End If
        strLineInfo = strLineInfo & "<spliter>"

        If Not objNode.selectSingleNode("Rate") Is Nothing Then
            strLineInfo = strLineInfo & objNode.selectSingleNode("Rate").text
            strRateOrPercent = "Rate"
        ElseIf Not objNode.selectSingleNode("RatePercent") Is Nothing Then
            strLineInfo = strLineInfo & objNode.selectSingleNode("RatePercent").text
            strRateOrPercent = "RatePercent"
        Else
            strRateOrPercent = "Neither"
        End If
        strLineInfo = strLineInfo & "<spliter>"

        If Not objNode.selectSingleNode("Amount") Is Nothing Then
            strLineInfo = strLineInfo & objNode.selectSingleNode("Amount").text
        End If
        strLineInfo = strLineInfo & "<spliter>"

        If Not objNode.selectSingleNode("ClassRef") Is Nothing Then
            strLineInfo = strLineInfo & objNode.selectSingleNode("ClassRef").selectSingleNode("FullName").text
        End If
        strLineInfo = strLineInfo & "<spliter>"

        If Not objNode.selectSingleNode("ServiceDate") Is Nothing Then
            strLineInfo = strLineInfo & objNode.selectSingleNode("ServiceDate").text
        End If
        strLineInfo = strLineInfo & "<spliter>"

        If Not objNode.selectSingleNode("SalesTaxCodeRef") Is Nothing Then
            strLineInfo = strLineInfo & objNode.selectSingleNode("SalesTaxCodeRef").selectSingleNode("FullName").text
        End If
        strLineInfo = strLineInfo & "<spliter>" & strRateOrPercent & "<spliter><spliter>"


        LineInfo = strLineInfo
    End Function


    Public Sub QBXMLRP_ModifyInvoice(ByRef strInvoiceChangeString As String)

        Dim strTemp As String
        strTemp = strInvoiceChangeString

        Dim objDOMDoc As New MSXML2.DOMDocument60 'DOMDocument60
        Dim objRootNode As IXMLDOMNode
        Dim objRequestNode As IXMLDOMNode
        Dim objElement As IXMLDOMElement
        Dim objAttribute As IXMLDOMAttribute

        
        CreateStandardRequestNode(True, "continueOnError", objDOMDoc, objRootNode, objRequestNode, objAttribute)

        Dim objInvoiceModRqNode As IXMLDOMNode
        
        AddMSXMLNode("InvoiceModRq", objDOMDoc, objRequestNode, objInvoiceModRqNode)

        Dim objInvoiceModNode As IXMLDOMNode
        
        AddMSXMLNode("InvoiceMod", objDOMDoc, objInvoiceModRqNode, objInvoiceModNode)

        'We know the invoice change string starts out as <TxnID>value</TxnID>
        'then <EditSequence>value</EditSequence> so use that knowledge to pull
        'out the values and put them in the request

        strTemp = Right(strTemp, Len(strTemp) - 7)
        
        AddMSXMLElement("TxnID", Left(strTemp, InStr(1, strTemp, "</") - 1), objDOMDoc, objInvoiceModNode, objElement)

        strTemp = Right(strTemp, Len(strTemp) - (InStr(1, strTemp, "<EditS") + 13))
        
        AddMSXMLElement("EditSequence", Left(strTemp, InStr(1, strTemp, "</") - 1), objDOMDoc, objInvoiceModNode, objElement)

        strTemp = Right(strTemp, Len(strTemp) - (InStr(1, strTemp, "</Edit") + 14))

        'The rest of our invoice change string is either going to be elements to
        'add to the InvoiceModRq aggregate or aggregates to add to the aggregate.
        'We'll pull each one out of the invoice change string and treat them the
        'same letting the recursive procedure we call figure it out.

        Dim strStartTag As String
        Dim strEndTag As String
        Dim strSegment As String
        Dim intTagLength As Short

        Dim objNode As IXMLDOMNode

        Do
            GetTags(strTemp, strStartTag, strEndTag, intTagLength)
            If intTagLength = 0 Then
                MsgBox("Error processing invoice change string, exiting")
                End
            End If

            strSegment = Left(strTemp, InStr(1, strTemp, strEndTag) + intTagLength)
            strTemp = Right(strTemp, Len(strTemp) - Len(strSegment))

            AddElementOrAggregate(strSegment, strStartTag, intTagLength, objDOMDoc, objInvoiceModNode, objNode, objElement)

            
        Loop Until strTemp Is String.Empty

        Dim strXMLRequest As String
        strXMLRequest = "<?xml version=""1.0""?>" & "<?qbxml version=""15.0""?>" & objRootNode.xml
        PrettyPrintXMLToFile(strXMLRequest, "C:\InvoiceModRq.xml")
        strSavedRequest = PrettyPrintXMLToString(strXMLRequest)

        Dim strXMLResponse As String
        strXMLResponse = objRequestProcessor.ProcessRequest(strTicket, strXMLRequest)
        PrettyPrintXMLToFile(strXMLResponse, "C:\InvoiceModRs.xml")

        objDOMDoc.async = False
        objDOMDoc.loadXML(strXMLResponse)

        Dim objInvoiceModNodeList As IXMLDOMNodeList
        objInvoiceModNodeList = objDOMDoc.getElementsByTagName("InvoiceModRs")

        objInvoiceModNode = objInvoiceModNodeList.item(0)

        Dim objInvoiceModAttributes As IXMLDOMNamedNodeMap
        objInvoiceModAttributes = objInvoiceModNode.attributes

        If objInvoiceModAttributes.getNamedItem("statusCode").nodeValue <> "0" Then
            MsgBox("Error modifying invoice" & vbCrLf & vbCrLf & "Error = " & objInvoiceModAttributes.getNamedItem("statusCode").nodeValue & vbCrLf & vbCrLf & "Message = " & objInvoiceModAttributes.getNamedItem("statusMessage").nodeValue)

            Exit Sub
        Else
            MsgBox("Successfully modified invoice")
        End If

    End Sub


    Private Sub AddElementOrAggregate(ByRef strElOrAggString As String, ByRef strStartTag As String, ByRef intTagLength As Short, ByRef objDOMDoc As DOMDocument60, ByRef objParentNode As IXMLDOMNode, ByRef objNode As IXMLDOMNode, ByRef objElement As IXMLDOMElement)

        Dim strTemp As String
        strTemp = strElOrAggString
        Dim strTagName As String
        strTagName = Left(strStartTag, Len(strStartTag) - 1)
        strTagName = Right(strTagName, Len(strTagName) - 1)

        Dim strValue As String
        strValue = Left(strTemp, Len(strTemp) - (intTagLength + 1))
        strValue = Right(strValue, Len(strValue) - intTagLength)

        Dim strInnerStartTag As String
        Dim strInnerEndTag As String
        Dim intInnerTagLength As Short

        GetTags(strValue, strInnerStartTag, strInnerEndTag, intInnerTagLength)

        
        Dim objChildNode As IXMLDOMNode
        Dim strSegment As String
        If String.IsNullOrEmpty(strInnerStartTag) Then
            
            AddMSXMLElement(strTagName, strValue, objDOMDoc, objParentNode, objElement)
        Else
            
            AddMSXMLNode(strTagName, objDOMDoc, objParentNode, objChildNode)

            Do
                strSegment = Left(strValue, InStr(1, strValue, strInnerEndTag) + intInnerTagLength)
                strValue = Right(strValue, Len(strValue) - Len(strSegment))

                AddElementOrAggregate(strSegment, strInnerStartTag, intInnerTagLength, objDOMDoc, objChildNode, objNode, objElement)

                

                If strValue IsNot String.Empty Then
                    GetTags(strValue, strInnerStartTag, strInnerEndTag, intInnerTagLength)
                End If
                
                '
            Loop Until strValue Is String.Empty
        End If
    End Sub


    Public Sub QBXMLRP_GetItemInfo(ByRef strItemFullName As String, ByRef strDesc As String, ByRef strRate As String, ByRef strRateOrPercent As String, ByRef strSalesTaxCode As String)

        Dim objDOMDoc As New DOMDocument60
        Dim objRootNode As IXMLDOMNode
        Dim objRequestNode As IXMLDOMNode
        Dim objElement As IXMLDOMElement
        Dim objAttribute As IXMLDOMAttribute

        
        CreateStandardRequestNode(False, "continueOnError", objDOMDoc, objRootNode, objRequestNode, objAttribute)

        Dim objQueryNode As IXMLDOMNode
        
        AddMSXMLNode("ItemQueryRq", objDOMDoc, objRequestNode, objQueryNode)

        
        AddMSXMLElement("FullName", strItemFullName, objDOMDoc, objQueryNode, objElement)

        Dim strXMLRequest As String
        strXMLRequest = "<?xml version=""1.0""?>" & "<?qbxml version=""15.0""?>" & objRootNode.xml
        PrettyPrintXMLToFile(strXMLRequest, "C:\debugrq.xml")

        Dim strXMLResponse As String
        strXMLResponse = objRequestProcessor.ProcessRequest(strTicket, strXMLRequest)
        PrettyPrintXMLToFile(strXMLResponse, "C:\debugrs.xml")

        objDOMDoc.async = False
        objDOMDoc.loadXML(strXMLResponse)

        Dim objItemQueryNodeList As IXMLDOMNodeList
        objItemQueryNodeList = objDOMDoc.getElementsByTagName("ItemQueryRs")

        Dim objItemQueryNode As IXMLDOMNode
        objItemQueryNode = objItemQueryNodeList.item(0)

        Dim objItemQueryAttributes As IXMLDOMNamedNodeMap
        objItemQueryAttributes = objItemQueryNode.attributes

        If objItemQueryAttributes.getNamedItem("statusCode").nodeValue <> "0" Then
            MsgBox("Error getting Item information" & vbCrLf & vbCrLf & "Error = " & objItemQueryAttributes.getNamedItem("statusCode").nodeValue & vbCrLf & vbCrLf & "Message = " & objItemQueryAttributes.getNamedItem("statusMessage").nodeValue)

            Exit Sub
        End If

        Dim objItemRetNode As IXMLDOMNode
        objItemRetNode = objItemQueryNode.firstChild

        Dim strItemType As String
        strItemType = objItemRetNode.nodeName

        If strItemType = "ItemInventoryRet" Or strItemType = "ItemInventoryAssemblyRet" Then

            If Not objItemRetNode.selectSingleNode("SalesDesc") Is Nothing Then
                strDesc = objItemRetNode.selectSingleNode("SalesDesc").text
            End If

            If Not objItemRetNode.selectSingleNode("SalesPrice") Is Nothing Then
                strRate = objItemRetNode.selectSingleNode("SalesPrice").text
            End If

            If Not objItemRetNode.selectSingleNode("SalesTaxCodeRef") Is Nothing Then
                strSalesTaxCode = objItemRetNode.selectSingleNode("SalesTaxCodeRef").selectSingleNode("FullName").text
            End If

        ElseIf strItemType = "ItemSubtotalRet" Or strItemType = "ItemPaymentRet" Or strItemType = "ItemSalesTaxRet" Or strItemType = "ItemGroupRet" Then

            If Not objItemRetNode.selectSingleNode("ItemDesc") Is Nothing Then
                strDesc = objItemRetNode.selectSingleNode("ItemDesc").text
            End If

        ElseIf strItemType = "ItemDiscountRet" Then

            If Not objItemRetNode.selectSingleNode("ItemDesc") Is Nothing Then
                strDesc = objItemRetNode.selectSingleNode("ItemDesc").text
            End If

            If Not objItemRetNode.selectSingleNode("SalesTaxCodeRef") Is Nothing Then
                strSalesTaxCode = objItemRetNode.selectSingleNode("SalesTaxCodeRef").selectSingleNode("FullName").text
            End If

            If Not objItemRetNode.selectSingleNode("DiscountRate") Is Nothing Then
                strRate = objItemRetNode.selectSingleNode("DiscountRate").text
            End If

            If Not objItemRetNode.selectSingleNode("DiscountRatePercent") Is Nothing Then
                strRate = objItemRetNode.selectSingleNode("DiscountRatePercent").text
                strRateOrPercent = "RatePercent"
            End If

        Else

            If Not objItemRetNode.selectSingleNode("SalesTaxCodeRef") Is Nothing Then
                strSalesTaxCode = objItemRetNode.selectSingleNode("SalesTaxCodeRef").selectSingleNode("FullName").text
            End If

            If Not objItemRetNode.selectSingleNode("SalesOrPurchase") Is Nothing Then
                If Not objItemRetNode.selectSingleNode("SalesOrPurchase").selectSingleNode("Desc") Is Nothing Then
                    strDesc = objItemRetNode.selectSingleNode("SalesOrPurchase").selectSingleNode("Desc").text
                End If

                If Not objItemRetNode.selectSingleNode("SalesOrPurchase").selectSingleNode("Price") Is Nothing Then
                    strRate = objItemRetNode.selectSingleNode("SalesOrPurchase").selectSingleNode("Price").text
                End If

                If Not objItemRetNode.selectSingleNode("SalesOrPurchase").selectSingleNode("PricePercent") Is Nothing Then
                    strRate = objItemRetNode.selectSingleNode("SalesOrPurchase").selectSingleNode("PricePercent").text
                    strRateOrPercent = "RatePercent"
                End If
            Else

                If Not objItemRetNode.selectSingleNode("SalesAndPurchase").selectSingleNode("SalesDesc") Is Nothing Then
                    strDesc = objItemRetNode.selectSingleNode("SalesAndPurchase").selectSingleNode("SalesDesc").text
                End If

                If Not objItemRetNode.selectSingleNode("SalesAndPurchase").selectSingleNode("SalesPrice") Is Nothing Then
                    strRate = objItemRetNode.selectSingleNode("SalesAndPurchase").selectSingleNode("SalesPrice").text
                    strRateOrPercent = "Rate"
                End If
            End If
        End If
    End Sub


    Public Sub QBXMLRP_LoadRequest(ByRef strRequestText As String)
        strRequestText = strSavedRequest
    End Sub
End Module