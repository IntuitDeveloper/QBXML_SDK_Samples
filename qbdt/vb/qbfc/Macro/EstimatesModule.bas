Attribute VB_Name = "EstimatesModule"
'----------------------------------------------------------
' Module: EstimatesModule
'
' Description: This module contains the code which provides support
'               functions for the UseMacroDefMacro project.  These
'               functions provide additional example code for queries
'               and processing results and filling in fields for a
'               request.
'
' Routines:
'
'           GetEstimates
'             This procedure will query all Estimates.  The estimates
'             will be queried in bunches using the MaxReturned field.
'             The TxnDateRangeFilter will be used to specify where
'             to start off on each new query of a set of estimates.
'
'           FillEstimatesListBox
'             This procedure will go thru the response from each estimate
'             query done.  A collection is kept of the TxnID/RefNumber for
'             each estimate.  We will use this collection to determine if any
'             of the estimates returned in this response were part of a
'             previous response from the query.  The RefNumber for each
'             estimate is displayed in the Estimate List Box.
'
'           EstimateInList
'             This procedure will search the collection to see if the
'             estimate's TxnID is already in the collection.  The TxnID
'             is used as a key to the collection.
'
'           ProcessResponse
'             This procedure will look at the status.code of the returned
'             response to see if the request(s) were successful.
'
'           GetSpecificEstimate
'             This procedure will query QB for a specific estimate
'             by specifying the TxnID of the estimate
'
'           FillInFieldsFromEstimate
'             This procedure will set many fields of the InvoiceAdd
'             with the fields returned from the estimateQuery.
'
'           EstimateLineHasData
'             This function checks to see whether the estimate line has any data
'             QuickBooks can have estimates with empty lines
'
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'----------------------------------------------------------

    Dim estimateCollection As Collection
    Const MAX_RETURNED = 20

    
    Public Sub GetEstimates(ByRef estimateList As ListBox)
        On Error GoTo Errs

        ' get a IMsgSetRequest to use for the EstimateQuery
        Dim estimateQueryRequest As IMsgSetRequest
        Dim sessionManager As QBSessionManager
        Set sessionManager = GetSessionManager
        Set estimateQueryRequest = sessionManager.CreateMsgSetRequest("US", 2, 0)
        estimateQueryRequest.ClearRequests
        
        ' set the OnError attribute to continueOnError
        estimateQueryRequest.Attributes.OnError = ENRqOnError.roeContinue

        ' Add the EstimatesQuery request
        Dim estimatesQuery As IEstimateQuery
        Set estimatesQuery = estimateQueryRequest.AppendEstimateQueryRq()
        If (estimatesQuery Is Nothing) Then
            Exit Sub
        End If

        Dim bDone As Boolean
        Dim beginTxnDate As Date
        Dim maxReturnedMultiplier As Integer
        bDone = False
        beginTxnDate = #1/1/1901#
        maxReturnedMultiplier = 1

        'Clear the list box and collections
        If (estimateCollection Is Nothing) Then
            Set estimateCollection = New Collection
        End If
        estimateList.Clear
        Do While (estimateCollection.count > 0)
            estimateCollection.Remove (1)
        Loop

        Do While (Not bDone)

            ' start looking for estimate next in the list
            ' we may have some overlap
            estimatesQuery.ORTxnQuery.TxnFilter.ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter.TxnDateFilter.FromTxnDate.SetValue beginTxnDate

            ' we are going to set the number of entries returned to limit the size of the return structure
            estimatesQuery.ORTxnQuery.TxnFilter.MaxReturned.SetValue (MAX_RETURNED * maxReturnedMultiplier)

            ' send the request to QB
            Dim msgSetResponse As IMsgSetResponse
            Set msgSetResponse = sessionManager.DoRequests(estimateQueryRequest)

            FillEstimatesListBox msgSetResponse, estimateList, bDone, beginTxnDate, maxReturnedMultiplier

        Loop

        Exit Sub

Errs:
        MsgBox "Error in GetEstimates" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description

    End Sub

    Public Sub FillEstimatesListBox(ByRef msgSetResponse As IMsgSetResponse, _
                    ByRef estimatesList As ListBox, _
                    ByRef bDone As Boolean, _
                    ByRef lastTxnDate As Date, _
                    ByRef maxReturnedMultiplier As Integer)
        On Error GoTo Errs

        ' check to make sure we have objects to access first
        ' and that there are responses in the list
        If (msgSetResponse Is Nothing) Then
            bDone = True
            Exit Sub
        End If
        If (msgSetResponse.responseList Is Nothing) Then
            bDone = True
            Exit Sub
        End If
        If (msgSetResponse.responseList.count <= 0) Then
            bDone = True
            Exit Sub
        End If

        ' Start parsing the response list
        Dim responseList As IResponseList
        Set responseList = msgSetResponse.responseList

        ' go thru each response and process the response.
        ' this example will only have one response in the list
        ' so we will look at index=0
        Dim response As IResponse
        Set response = responseList.GetAt(0)
        If (Not response Is Nothing) Then
            If response.StatusCode <> "0" Then
                'If the status is bad, report it to the user
                MsgBox "FillEstimatesListBox unexpected Error - " & response.StatusMessage
                bDone = True
                Exit Sub
            End If
        End If

        ' first make sure we have a response object to handle
        If (response Is Nothing) Then
            bDone = True
            Exit Sub
        End If
        If (response.Type Is Nothing) Then
            bDone = True
            Exit Sub
        End If
        If (response.Detail Is Nothing) Then
            bDone = True
            Exit Sub
        End If
        If (response.Detail.Type Is Nothing) Then
            bDone = True
            Exit Sub
        End If

        ' make sure we are processing the EstimateQueryRs and
        ' the EstimateRetList responses in this response list
        Dim estimateRetList As IEstimateRetList
        Dim responseType As ENResponseType
        Dim responseDetailType As ENObjectType
        responseType = response.Type.GetValue()
        responseDetailType = response.Detail.Type.GetValue()
        If (responseType = ENResponseType.rtEstimateQueryRs) And _
            (responseDetailType = ENObjectType.otEstimateRetList) Then
            ' save the response detail in the appropriate object type
            ' since we have first verified the type of the response object
            Set estimateRetList = response.Detail
        Else
            ' bail, we do not have the responses we were expecting
            bDone = True
            Exit Sub
        End If

        'Parse the query response and add the Estimate to the Estimate list box
        Dim count As Integer
        Dim index As Integer
        Dim estimateRet As IEstimateRet
        Dim estimateInfo As InfoClass
        count = estimateRetList.count

        ' we are done with the estimateQueries if we have not received the MaxReturned
        If (count < MAX_RETURNED) Then
            bDone = True
        Else
            ' if the last element returned is already in our list
            ' do not bother to go back thru the returns.
            ' we must go back and increase the number of elements
            ' returned because there are too many txns with the same date
            Dim lastIndex As Long
            lastIndex = count - 1
            Set estimateRet = estimateRetList.GetAt(lastIndex)
            If (estimateRet Is Nothing) Then
                bDone = True
                Exit Sub
            End If
            If (estimateRet.TxnID Is Nothing) Then
                bDone = True
                Exit Sub
            End If
            Dim bInList As Boolean
            EstimateInList estimateRet.TxnID.GetValue(), bInList
            If (bInList) Then
                maxReturnedMultiplier = maxReturnedMultiplier + 1
                Exit Sub
            End If
        End If

        For index = 0 To count - 1
            Set estimateRet = estimateRetList.GetAt(index)
            If (estimateRet Is Nothing) Then
                bDone = True
                Exit Sub
            End If
            If (estimateRet.RefNumber Is Nothing) Then
                bDone = True
                Exit Sub
            End If
            If (estimateRet.TxnID Is Nothing) Then
                bDone = True
                Exit Sub
            End If
            ' lets check to make sure we do not have the estimate
            ' in our list already
            ' skip this estimate if this is a repeat from the last query
            EstimateInList estimateRet.TxnID.GetValue(), bInList
            If (Not bInList) Then
                ' we are saving the RefNumber and TxnID pairs for each Estimate
                ' this is good practice since the RefNumber can change but the
                ' TxnID will not change for an Estimate.
                Set estimateInfo = New InfoClass
                ' save all the field information for this customer
                estimateInfo.SetRefNumber estimateRet.RefNumber.GetValue()
                estimateInfo.SetTxnID estimateRet.TxnID.GetValue()
                ' add the estimateInfo to our list
                ' add the estimate RefNumber to the list box
                estimatesList.AddItem estimateInfo.GetRefNumber
                estimateCollection.Add estimateInfo, estimateInfo.GetTxnID
            End If
            lastTxnDate = estimateRet.TxnDate.GetValue
        Next

        Exit Sub
Errs:
        MsgBox "Error in FillEstimatesListBox" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description

        bDone = True
    End Sub

    Private Function EstimateInList(ByRef estimateTxnID As String, ByRef bInList As Boolean)
        On Error GoTo Errs

        Dim estimateInfo As InfoClass
        
        Set estimateInfo = estimateCollection.Item(estimateTxnID)
        
        ' if the above line is successful, then the estimate is in the list
        bInList = True

        Exit Function
Errs:
        bInList = False

    End Function

    Public Function ProcessResponse(ByRef msgSetResponse As IMsgSetResponse) As Boolean
        On Error GoTo Errs
        
        ' check to make sure we have objects to access first
        ' and that there are responses in the list
        If (msgSetResponse Is Nothing) Then
            ProcessResponse = False
            Exit Function
        End If
        If (msgSetResponse.responseList Is Nothing) Then
            ProcessResponse = False
            Exit Function
        End If
        If (msgSetResponse.responseList.count <= 0) Then
            ProcessResponse = False
            Exit Function
        End If

        ' Start parsing the response list
        Dim responseList As IResponseList
        Set responseList = msgSetResponse.responseList

        ' go thru each response and process the response.
        Dim response As IResponse
        Dim index As Integer
        Dim count As Integer
        count = responseList.count
        For index = 0 To count - 1
            Set response = responseList.GetAt(index)
            If (response Is Nothing) Then
                ProcessResponse = False
                Exit Function
            End If
            If response.StatusCode <> "0" Then
                'If the status is bad, report it to the user
                MsgBox ("ProcessResponse unexpected Error - " & response.StatusMessage)
                ProcessResponse = False
                Exit Function
            End If
        Next
        
        ProcessResponse = True
        
        Exit Function
Errs:
        MsgBox "Error in ProcessResponse" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description

        ProcessResponse = False
    End Function

    Public Function GetSpecificEstimate(ByRef selectedIndex As Integer) As IEstimateRet
        On Error GoTo Errs
        
        ' get the selected estimate
        ' the collection is 1 based, the listbox is 0 based
        Dim estimateInfo As InfoClass
        Dim estimateTxnID As String
        Set estimateInfo = estimateCollection.Item(selectedIndex + 1)
        If (estimateInfo Is Nothing) Then
            Exit Function
        Else
            estimateTxnID = estimateInfo.GetTxnID
        End If
                
        ' get a IMsgSetRequest to use for the EstimateQuery
        Dim estimateQueryRequest As IMsgSetRequest
        Dim sessionManager As QBSessionManager
        Set sessionManager = GetSessionManager
        Set estimateQueryRequest = sessionManager.CreateMsgSetRequest("US", 2, 0)
        
        ' Add the EstimateQuery request
        Dim estimateQuery As IEstimateQuery
        Set estimateQuery = estimateQueryRequest.AppendEstimateQueryRq()
        
        ' specify the estimate we want
        estimateQuery.ORTxnQuery.TxnIDList.Add (estimateTxnID)
        
        ' we want all the details of this estimate so include the lines
        estimateQuery.IncludeLineItems.SetValue True
        
        Dim msgSetResponse As IMsgSetResponse
        Set msgSetResponse = sessionManager.DoRequests(estimateQueryRequest)
        
        ' check to make sure we have objects to access first
        ' and that there are responses in the list
        If (msgSetResponse Is Nothing) Then
            Exit Function
        End If
        If (msgSetResponse.responseList Is Nothing) Then
            Exit Function
        End If
        If (msgSetResponse.responseList.count <= 0) Then
            Exit Function
        End If

        ' Start parsing the response list
        Dim responseList As IResponseList
        Set responseList = msgSetResponse.responseList

        ' go thru the one response and process the response.
        Dim response As IResponse
        Set response = responseList.GetAt(0)
        If (response Is Nothing) Then
            Exit Function
        End If
        If response.StatusCode <> "0" Then
            'If the status is bad, report it to the user
            MsgBox ("GetSpecificEstimate unexpected Error - " & response.StatusMessage)
            Exit Function
        End If
        
        ' first make sure we have a response object to handle
        If (response Is Nothing) Then
            Exit Function
        End If
        If (response.Type Is Nothing) Then
            Exit Function
        End If
        If (response.Detail Is Nothing) Then
            Exit Function
        End If
        If (response.Detail.Type Is Nothing) Then
            Exit Function
        End If

        ' make sure we are processing the EstimateQueryRs and
        ' the EstimateRetList responses in this response list
        Dim estimateRetList As IEstimateRetList
        Dim responseType As ENResponseType
        Dim responseDetailType As ENObjectType
        responseType = response.Type.GetValue()
        responseDetailType = response.Detail.Type.GetValue()
        If (responseType = ENResponseType.rtEstimateQueryRs) And _
            (responseDetailType = ENObjectType.otEstimateRetList) Then
            ' save the response detail in the appropriate object type
            ' since we have first verified the type of the response object
            Set estimateRetList = response.Detail
        Else
            ' bail, we do not have the responses we were expecting
            Exit Function
        End If

        'get the query response return the estimateRet variable
        If (Not estimateRetList Is Nothing) Then
            If (estimateRetList.count = 1) Then
                Set GetSpecificEstimate = estimateRetList.GetAt(0)
            End If
        End If
            
        Exit Function
Errs:
        MsgBox "Error in GetSpecificEstimate" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description
        
    End Function

    Public Function FillInFieldsFromEstimate(ByRef invoiceAdd As IInvoiceAdd, _
                    ByRef selectedIndex As Integer, ByRef customerID As String, _
                    ByRef paymentAmount As String)
        On Error GoTo Errs
        
        ' Get the EstimateRet for the selected Estimate
        Dim estimateRet As IEstimateRet
        Set estimateRet = GetSpecificEstimate(selectedIndex)
        
        If (estimateRet Is Nothing) Or (invoiceAdd Is Nothing) Then
            Exit Function
        End If
        
        '<CustomerRef>
        If (Not estimateRet.CustomerRef Is Nothing) Then
            If (Not estimateRet.CustomerRef.ListID Is Nothing) Then
                invoiceAdd.CustomerRef.ListID.SetValue (estimateRet.CustomerRef.ListID.GetValue())
                customerID = estimateRet.CustomerRef.ListID.GetValue
            End If
        End If
        '<ClassRef>                                          <!-- opt -->
        If (Not estimateRet.ClassRef Is Nothing) Then
            If (Not estimateRet.ClassRef.ListID Is Nothing) Then
                invoiceAdd.ClassRef.ListID.SetValue (estimateRet.ClassRef.ListID.GetValue())
            End If
        End If
        '<BillAddress>                                       <!-- opt -->
        If (Not estimateRet.BillAddress Is Nothing) Then
            '<Addr1>STRTYPE</Addr1>                            <!-- opt, QBD max = 41, QBO max = 500 -->
            If (Not estimateRet.BillAddress.Addr1 Is Nothing) Then
                invoiceAdd.BillAddress.Addr1.SetValue (estimateRet.BillAddress.Addr1.GetValue())
            End If
          '<Addr2>STRTYPE</Addr2>                            <!-- opt, QBD max = 41, QBO max = 500 -->
            If (Not estimateRet.BillAddress.Addr2 Is Nothing) Then
                invoiceAdd.BillAddress.Addr2.SetValue (estimateRet.BillAddress.Addr2.GetValue())
            End If
          '<Addr3>STRTYPE</Addr3>                            <!-- opt, QBD max = 41, QBO max = 500 -->
            If (Not estimateRet.BillAddress.Addr3 Is Nothing) Then
                invoiceAdd.BillAddress.Addr3.SetValue (estimateRet.BillAddress.Addr3.GetValue())
            End If
          '<Addr4>STRTYPE</Addr4>                            <!-- opt, QBD max = 41, QBO max = 500 -->
            If (Not estimateRet.BillAddress.Addr4 Is Nothing) Then
                invoiceAdd.BillAddress.Addr4.SetValue (estimateRet.BillAddress.Addr4.GetValue())
            End If
          '<City>STRTYPE</City>                              <!-- opt, QBD max = 31, QBO max = 255 -->
            If (Not estimateRet.BillAddress.City Is Nothing) Then
                invoiceAdd.BillAddress.City.SetValue (estimateRet.BillAddress.City.GetValue())
            End If
          '<State>STRTYPE</State>                            <!-- opt, QBD max = 21, QBO max = 255 -->
            If (Not estimateRet.BillAddress.State Is Nothing) Then
                invoiceAdd.BillAddress.State.SetValue (estimateRet.BillAddress.State.GetValue())
            End If
          '<PostalCode>STRTYPE</PostalCode>                  <!-- opt, QBD max = 13, QBO max = 30 -->
            If (Not estimateRet.BillAddress.PostalCode Is Nothing) Then
                invoiceAdd.BillAddress.PostalCode.SetValue (estimateRet.BillAddress.PostalCode.GetValue())
            End If
          '<Country>STRTYPE</Country>                        <!-- opt, QBD max = 31, QBO max = 255 -->
            If (Not estimateRet.BillAddress.Country Is Nothing) Then
                invoiceAdd.BillAddress.Country.SetValue (estimateRet.BillAddress.Country.GetValue())
            End If
        '</BillAddress>
        End If
        '<PONumber>STRTYPE</PONumber>                        <!-- opt, QBD max = 25 -->
        If (Not estimateRet.PONumber Is Nothing) Then
            invoiceAdd.PONumber.SetValue (estimateRet.PONumber.GetValue())
        End If
        '<TermsRef>                                          <!-- opt -->
        If (Not estimateRet.TermsRef Is Nothing) Then
            If (Not estimateRet.TermsRef.ListID Is Nothing) Then
                invoiceAdd.TermsRef.ListID.SetValue (estimateRet.TermsRef.ListID.GetValue())
            End If
        End If
        '<DueDate>DATETYPE</DueDate>                         <!-- opt -->
        If (Not estimateRet.DueDate Is Nothing) Then
            invoiceAdd.DueDate.SetValue (estimateRet.DueDate.GetValue())
        End If
        '<SalesRepRef>                                       <!-- opt -->
        If (Not estimateRet.SalesRepRef Is Nothing) Then
            If (Not estimateRet.SalesRepRef.ListID Is Nothing) Then
                invoiceAdd.SalesRepRef.ListID.SetValue (estimateRet.SalesRepRef.ListID.GetValue())
            End If
        End If
        '<FOB>STRTYPE</FOB>                                  <!-- opt, QBD max = 13 -->
        If (Not estimateRet.FOB Is Nothing) Then
            invoiceAdd.FOB.SetValue (estimateRet.FOB.GetValue())
        End If
        '<ItemSalesTaxRef>                                   <!-- opt -->
        If (Not estimateRet.ItemSalesTaxRef Is Nothing) Then
            If (Not estimateRet.ItemSalesTaxRef.ListID Is Nothing) Then
                invoiceAdd.ItemSalesTaxRef.ListID.SetValue (estimateRet.ItemSalesTaxRef.ListID.GetValue())
            End If
        End If
        '<Memo>STRTYPE</Memo>                                <!-- opt, QBD max = 4095, QBO max = 4000 -->
        If (Not estimateRet.Memo Is Nothing) Then
            invoiceAdd.Memo.SetValue (estimateRet.Memo.GetValue())
        End If
        '<CustomerMsgRef>                                    <!-- opt -->
        If (Not estimateRet.CustomerMsgRef Is Nothing) Then
            If (Not estimateRet.CustomerMsgRef.ListID Is Nothing) Then
                invoiceAdd.CustomerMsgRef.ListID.SetValue (estimateRet.CustomerMsgRef.ListID.GetValue())
            End If
        End If
        '<CustomerSalesTaxCodeRef>                           <!-- opt -->
        If (Not estimateRet.CustomerSalesTaxCodeRef Is Nothing) Then
            If (Not estimateRet.CustomerSalesTaxCodeRef.ListID Is Nothing) Then
                invoiceAdd.CustomerSalesTaxCodeRef.ListID.SetValue (estimateRet.CustomerSalesTaxCodeRef.ListID.GetValue())
            End If
        End If
        If (Not estimateRet.OREstimateLineRetList Is Nothing) Then
            Dim index As Long
            Dim count As Long
            Dim orEstimateLineRet As IOREstimateLineRet
            Dim estimateLineRet As IEstimateLineRet
            Dim estimateLineGroupRet As IEstimateLineGroupRet
            Dim orInvoiceLineAdd As IORInvoiceLineAdd
            Dim invoiceLineAdd As IInvoiceLineAdd
            Dim invoiceLineGroupAdd As IInvoiceLineGroupAdd
            count = estimateRet.OREstimateLineRetList.count
            For index = 0 To count - 1
                Set orEstimateLineRet = estimateRet.OREstimateLineRetList.GetAt(index)
                If (Not orEstimateLineRet Is Nothing) Then
                    ' determine if we have an EstimateLineRet type
                    If (orEstimateLineRet.ortype = orelrEstimateLineRet) Then
                        Set estimateLineRet = orEstimateLineRet.estimateLineRet
                        If (Not estimateLineRet Is Nothing) Then
                            
                            'verify that the line contains any data
                            'before adding an invoice line - QuickBooks may have empty lines
                            If (EstimateLineHasData(estimateLineRet)) Then
                         
                                Set orInvoiceLineAdd = invoiceAdd.ORInvoiceLineAddList.Append
                                Set invoiceLineAdd = orInvoiceLineAdd.invoiceLineAdd
                            
                                '<ItemRef>                                         <!-- opt -->
                                If (Not estimateLineRet.ItemRef Is Nothing) Then
                                    If (Not estimateLineRet.ItemRef.ListID Is Nothing) Then
                                        invoiceLineAdd.ItemRef.ListID.SetValue estimateLineRet.ItemRef.ListID.GetValue()
                                    End If
                                End If
                                '<Desc>STRTYPE</Desc>                              <!-- opt, QBD max = 4095, QBO max = 4000 -->
                                If (Not estimateLineRet.Desc Is Nothing) Then
                                    invoiceLineAdd.Desc.SetValue estimateLineRet.Desc.GetValue()
                                End If
                                '<Quantity>QUANTYPE</Quantity>                     <!-- opt -->
                                If (Not estimateLineRet.Quantity Is Nothing) Then
                                    invoiceLineAdd.Quantity.SetValue estimateLineRet.Quantity.GetValue()
                                End If
                                If (Not estimateLineRet.ORRate Is Nothing) Then
                                    If (estimateLineRet.ORRate.ortype = orrRate) Then
                                        '<Rate>PRICETYPE</Rate>
                                        If (Not estimateLineRet.ORRate.Rate Is Nothing) Then
                                            invoiceLineAdd.ORRatePriceLevel.Rate.SetValue estimateLineRet.ORRate.Rate.GetValue()
                                        End If
                                    ElseIf (estimateLineRet.ORRate.ortype = orrRatePercent) Then
                                        '<RatePercent>PERCENTTYPE</RatePercent>
                                        If (Not estimateLineRet.ORRate.RatePercent Is Nothing) Then
                                            invoiceLineAdd.ORRatePriceLevel.RatePercent.SetValue estimateLineRet.ORRate.RatePercent.GetValue()
                                        End If
                                    End If
                                End If
                            
                                '<ClassRef>                                        <!-- opt -->
                                If (Not estimateLineRet.ClassRef Is Nothing) Then
                                    If (Not estimateLineRet.ClassRef.ListID Is Nothing) Then
                                        invoiceLineAdd.ClassRef.ListID.SetValue estimateLineRet.ClassRef.ListID.GetValue()
                                    End If
                                End If
                                '<Amount>AMTTYPE</Amount>                          <!-- opt -->
                                If (Not estimateLineRet.Amount Is Nothing) Then
                                    invoiceLineAdd.Amount.SetValue estimateLineRet.Amount.GetValue()
                                End If
                                '<SalesTaxCodeRef>                                 <!-- opt -->
                                If (Not estimateLineRet.SalesTaxCodeRef Is Nothing) Then
                                    If (Not estimateLineRet.SalesTaxCodeRef.ListID Is Nothing) Then
                                        invoiceLineAdd.SalesTaxCodeRef.ListID.SetValue estimateLineRet.SalesTaxCodeRef.ListID.GetValue()
                                    End If
                                End If
                            End If
                        End If
                    ' determine if we have an EstimateGroupLineRet type
                    ElseIf (orEstimateLineRet.ortype = orelrEstimateLineGroupRet) Then
                        Set orInvoiceLineAdd = invoiceAdd.ORInvoiceLineAddList.Append
                        Set invoiceLineGroupAdd = orInvoiceLineAdd.invoiceLineGroupAdd
                        Set estimateLineGroupRet = orEstimateLineRet.estimateLineGroupRet
                        If (Not estimateLineGroupRet Is Nothing) Then
                            '<ItemGroupRef>
                            If (Not estimateLineGroupRet.ItemGroupRef Is Nothing) Then
                                If (Not estimateLineGroupRet.ItemGroupRef.ListID Is Nothing) Then
                                    invoiceLineGroupAdd.ItemGroupRef.ListID.SetValue estimateLineGroupRet.ItemGroupRef.ListID.GetValue()
                                End If
                            End If
                            '<Quantity>QUANTYPE</Quantity>                     <!-- opt -->
                            If (Not estimateLineGroupRet.Quantity Is Nothing) Then
                                invoiceLineGroupAdd.Quantity.SetValue estimateLineGroupRet.Quantity.GetValue()
                            End If
                            '<ServiceDate>DATETYPE</ServiceDate>               <!-- opt,  not supported by QBD -->
                        End If
                    End If
                End If
            Next
        End If
        
        ' get the total amount to use on the ReceivePaymentAdd request
        If (Not estimateRet.TotalAmount Is Nothing) Then
            paymentAmount = estimateRet.TotalAmount.GetAsString()
        End If
        
        Exit Function
Errs:
        MsgBox "Error in FillInFieldsFromEstimate" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description

    End Function


Public Function EstimateLineHasData(estimateLineRet As IEstimateLineRet) As Boolean

    If (Not estimateLineRet.Amount Is Nothing) Then
       EstimateLineHasData = True
       Exit Function
    End If
    
    If (Not estimateLineRet.ClassRef Is Nothing) Then
       EstimateLineHasData = True
       Exit Function
    End If
    
    If (Not estimateLineRet.Desc Is Nothing) Then
       EstimateLineHasData = True
       Exit Function
    End If
    
    If (Not estimateLineRet.ItemRef Is Nothing) Then
       EstimateLineHasData = True
       Exit Function
    End If
    
    If (Not estimateLineRet.ORMarkupRate Is Nothing) Then
       EstimateLineHasData = True
       Exit Function
    End If
    
    If (Not estimateLineRet.ORRate Is Nothing) Then
       EstimateLineHasData = True
       Exit Function
    End If
    
    If (Not estimateLineRet.Quantity Is Nothing) Then
       EstimateLineHasData = True
       Exit Function
    End If
    
    If (Not estimateLineRet.SalesTaxCodeRef Is Nothing) Then
       EstimateLineHasData = True
       Exit Function
    End If
       
    EstimateLineHasData = False
End Function

