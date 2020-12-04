'----------------------------------------------------------
' Module: SyncCustomerListModule
'
' Description: This module contains the code which creates QBFC
'              messages, exchanges them with QuickBooks, interprets
'              the responses and loads information into form objects.
'
' Routines: OpenConnectionBeginSession
'             Opens a connection and begins a sesson with the
'             currently open company file.  If a company isn't open,
'             the routine will display a message and then exit the
'             program.
'
'           EndSessionCloseConnection
'             Calls EndSession and CloseConnection if the boolean
'             booSessionBegun is true.
'
'           GetCustomers
'             Send a CustomerQuery to QuickBooks.  Call FillCustomerListBox 
'             to fill in the list box.
'
'           FillCustomerListBox
'             Parses a CustomerQueryRet response and 
'             fills the Customer List Box with the FullName/ListID 
'             object for each Item.  The Customer's FullName is displayed in
'             list box.
'
'           SyncCustomerList
'             Calls the CompanyFileRestored, CheckDeletedCustomers and 
'             CheckModifiedCustomer functions to keep out Customer List
'             up to date
'
'           CompanyFileRestored
'             Calls the CompanyActivityQuery to determine if the company 
'             file has been restored since our last sync.  If the restore
'             has happened, we will requery all customers and fill in the
'             list box again by calling the GetCustomers function.
'
'           CheckDeletedCustomers
'             Calls the ListDeletedQuery to determine if any customers have
'             been deleted since the last time we sync'd our customer list
'             with QB.  If some customers have been deleted, we will find 
'             the customers in our list box and delete them.
'
'           CheckModifiedCustomer
'             Calls the CustomerQuery with a FromModifiedDate filter to 
'             determine which customer records have been modified since our 
'             last sync with QB.  If some customers have been deleted, we 
'             will find the customer in our list and update our list data
'             with the data from QB.
'
'           FoundCustomerInListBox
'             Loops thru the objects in the list box to see if we have a match.
'           
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'----------------------------------------------------------
Imports Interop.QBFC13

Module SyncCustomerListModule

    Dim booSessionBegun As Boolean
    Dim MAX_RETURNED As Integer = 20

    'Module objects
    Dim qbSessionManager As QBSessionManager
    Dim msgSetRequest As IMsgSetRequest
    Dim timeLastCustomerSync As System.DateTime

    Public Sub OpenConnectionBeginSession()

        booSessionBegun = False

        On Error GoTo Errs

        qbSessionManager = New QBSessionManager()

        qbSessionManager.OpenConnection("", "IDN Sync Customer List Sample - QBFC")

        qbSessionManager.BeginSession("", ENOpenMode.omDontCare)
        booSessionBegun = True

        'Check to make sure the QuickBooks we're working with supports version 2 of the SDK
        Dim strXMLVersions() As String
        strXMLVersions = qbSessionManager.QBXMLVersionsForSession

        Dim booSupports2dot0 As Boolean
        booSupports2dot0 = False
        Dim i As Long
        For i = LBound(strXMLVersions) To UBound(strXMLVersions)
            If (strXMLVersions(i) = "2.0") Then
                booSupports2dot0 = True
                msgSetRequest = qbSessionManager.CreateMsgSetRequest("US", 2, 0)
                Exit For
            End If
        Next

        If Not booSupports2dot0 Then
            MsgBox("This program only runs against QuickBooks installations which support the 2.0 qbXML spec.  Your version of QuickBooks does not support qbXML 2.0")
            End
        End If
        Exit Sub

Errs:
        If Err.Number = &H80040416 Then
            MsgBox("You must have QuickBooks running with the company" & vbCrLf & "file open to use this program.")
            qbSessionManager.CloseConnection()
            End
        ElseIf Err.Number = &H80040422 Then
            MsgBox("This QuickBooks company file is open in single user mode and" & vbCrLf & "another application is already accessing it.  Please exit the" & vbCrLf & "other application and run this application again.")
            qbSessionManager.CloseConnection()
            End
        Else
            MsgBox("HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, MsgBoxStyle.Critical, "Error in OpenConnectionBeginSession")

            If booSessionBegun Then
                qbSessionManager.EndSession()
            End If

            qbSessionManager.CloseConnection()
            End
        End If
    End Sub

    Public Sub EndSessionCloseConnection()
        On Error GoTo Errs

        If booSessionBegun Then
            qbSessionManager.EndSession()
            qbSessionManager.CloseConnection()
        End If

        Exit Sub
Errs:
        MsgBox("HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    MsgBoxStyle.Critical, _
                    "Error in EndSessionCloseConnection")
    End Sub

    Public Sub GetCustomers(ByRef lstCustomers As System.Windows.Forms.ListBox, _
                        ByRef bError As Boolean)
        On Error GoTo Errs

        ' make sure we do not have any old requests still defined
        msgSetRequest.ClearRequests()

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

        ' Add the CustomerQuery request
        Dim customerQuery As ICustomerQuery
        customerQuery = msgSetRequest.AppendCustomerQueryRq()

        ' in this example we are going to keep a list of only "active" status Customers
        ' the default is receiving only the active status customers, but we will 
        ' set the value anyway for illustration.
        customerQuery.ORCustomerListQuery.CustomerListFilter.ActiveStatus.SetValue(ENActiveStatus.asActiveOnly)

        ' we are going to set the number of entries returned to limit the size of the return structure
        customerQuery.ORCustomerListQuery.CustomerListFilter.MaxReturned.SetValue(MAX_RETURNED)

        Dim bDone As Boolean = False
        Dim firstFullName As String = "!"

        'Clear the list box
        lstCustomers.Items.Clear()

        Do While (Not bDone)

            ' start looking for customer next in the list
            ' we will have one overlap
            customerQuery.ORCustomerListQuery.CustomerListFilter.ORNameFilter.NameRangeFilter.FromName.SetValue(firstFullName)

            ' send the request to QB
            Dim msgSetResponse As IMsgSetResponse
            msgSetResponse = qbSessionManager.DoRequests(msgSetRequest)

            'MsgBox(msgSetRequest.ToXMLString())

            FillCustomerListBox(msgSetResponse, lstCustomers, bDone, firstFullName, bError)

        Loop

        Exit Sub

Errs:
        MsgBox("HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    MsgBoxStyle.Critical, _
                    "Error in GetCustomers")
        bError = True
    End Sub

    Public Sub FillCustomerListBox(ByRef msgSetResponse As IMsgSetResponse, _
                    ByRef lstCustomers As System.Windows.Forms.ListBox, _
                    ByRef bDone As Boolean, _
                    ByRef lastFullName As String, _
                    ByRef bError As Boolean)
        On Error GoTo Errs

        ' check to make sure we have objects to access first
        ' and that there are responses in the list
        If (msgSetResponse Is Nothing) Or _
            (msgSetResponse.ResponseList Is Nothing) Or _
            (msgSetResponse.ResponseList.Count <= 0) Then
            bDone = True
            bError = True
            Exit Sub
        End If

        ' Start parsing the response list
        Dim responseList As IResponseList
        responseList = msgSetResponse.ResponseList

        ' go thru each response and process the response.
        ' this example will only have one response in the list
        ' so we will look at index=0
        Dim response As IResponse
        response = responseList.GetAt(0)
        If (Not response Is Nothing) Then
            If response.StatusCode <> "0" Then
                'If the status is bad, report it to the user
                MsgBox("FillCustomerListBox unexpexcted Error - " & response.StatusMessage)
                bDone = True
                bError = True
                Exit Sub
            End If
        End If

        ' first make sure we have a response object to handle
        If (response Is Nothing) Or _
            (response.Type Is Nothing) Or _
            (response.Detail Is Nothing) Or _
            (response.Detail.Type Is Nothing) Then
            bDone = True
            bError = True
            Exit Sub
        End If

        ' make sure we are processing the CustomerQueryRs and 
        ' the CustomerRetList responses in this response list
        Dim customerRetList As ICustomerRetList
        Dim responseType As ENResponseType
        Dim responseDetailType As ENObjectType
        responseType = response.Type.GetValue()
        responseDetailType = response.Detail.Type.GetValue()
        If (responseType = ENResponseType.rtCustomerQueryRs) And _
            (responseDetailType = ENObjectType.otCustomerRetList) Then
            ' save the response detail in the appropriate object type
            ' since we have first verified the type of the response object
            customerRetList = response.Detail
        Else
            ' bail, we do not have the responses we were expecting
            bDone = True
            bError = True
            Exit Sub
        End If

        'Parse the query response and add the Customers to the Customer list box
        Dim count As Short
        Dim index As Short
        Dim customerRet As ICustomerRet
        Dim customerInfo As CustomerClass
        count = customerRetList.Count

        ' we are done with the customerQueries if we have not received the MaxReturned
        If (count < MAX_RETURNED) Then
            bDone = True
        End If

        For index = 0 To count - 1
            ' skip this customer if this is a repeat from the last query
            customerRet = customerRetList.GetAt(index)
            If (customerRet Is Nothing) Or _
                    (customerRet.FullName Is Nothing) Or _
                    (customerRet.ListID Is Nothing) Then
                bDone = True
                bError = True
                Exit Sub
            End If
            ' only the first customerRet should be repeating and then
            ' lets just check to make sure we do not have the customer
            ' just in case another app changed a customer right between our
            ' queries.
            If (index <> 0) Or (Not FoundCustomerInListBox(lstCustomers, customerRet.ListID.GetValue())) Then
                ' we are saving the FullName and ListID pairs for each Customer
                ' this is good practice since the FullName can change but the 
                ' ListID will not change for an Customer.
                customerInfo = New CustomerClass()
                ' save all the field information for this customer
                customerInfo.customerRet = customerRet
                lstCustomers.Items.Add(customerInfo)
            End If
            lastFullName = customerRet.FullName.GetValue()
        Next

        Exit Sub
Errs:
        MsgBox("HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    MsgBoxStyle.Critical, _
                    "Error in FillCustomerListBox")
        bDone = True
        bError = True
    End Sub

    Public Sub SyncCustomerList(ByRef lstCustomers As System.Windows.Forms.ListBox, _
                ByRef bError As Boolean)
        On Error GoTo Errs

        Dim today, tmpDate As Date
        today = New Date().Now()
        tmpDate = timeLastCustomerSync

        ' if last resync is more than 3 months ago, we need to grab the whole list 
        ' again.  The ListDeletedQuery will only save the last 3 months worth of deletes
        If (tmpDate.AddMonths(3) <= today) Then
            GetCustomers(lstCustomers, bError)
            Exit Sub
        End If

        ' see if the company file was restored.  If so, then the customer list
        ' was totally reread and reset with the current values in the company file.
        If (Not CompanyFileRestored(lstCustomers, bError)) And _
            (Not bError) Then

            ' check to see if any customers have been deleted
            CheckDeletedCustomers(lstCustomers, bError)

            If (bError) Then
                Exit Sub
            End If

            ' check to see if any customers have been modified
            CheckModifiedCustomer(lstCustomers, bError)

        End If

        Exit Sub
Errs:
        MsgBox("HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    MsgBoxStyle.Critical, _
                    "Error in SyncCustomerList")
        bError = True
    End Sub

    Private Function CompanyFileRestored(ByRef lstCustomers As System.Windows.Forms.ListBox, _
                ByRef bError As Boolean) _
    As Boolean
        On Error GoTo Errs

        CompanyFileRestored = False

        ' make sure we do not have any old requests still defined
        msgSetRequest.ClearRequests()

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

        ' Add the CompanyActivityQuery request
        msgSetRequest.AppendCompanyActivityQueryRq()

        ' send the request to QB
        Dim msgSetResponse As IMsgSetResponse
        msgSetResponse = qbSessionManager.DoRequests(msgSetRequest)

        ' check to make sure we have objects to access first
        ' and that there are responses in the list
        If (msgSetResponse Is Nothing) Or _
            (msgSetResponse.ResponseList Is Nothing) Or _
            (msgSetResponse.ResponseList.Count <= 0) Then
            bError = True
            Exit Function
        End If

        ' Start parsing the response list
        Dim responseList As IResponseList
        responseList = msgSetResponse.ResponseList

        ' go thru each response and process the response.
        ' this example will only have one response in the list
        ' so we will look at index=0
        Dim response As IResponse
        response = responseList.GetAt(0)
        If (Not response Is Nothing) Then
            If response.StatusCode <> "0" Then
                'If the status is bad, report it to the user
                MsgBox("CompanyFileRestored unexpexcted Error - " & vbCrLf & "StatusCode = " & response.StatusCode & vbCrLf & vbCrLf & response.StatusMessage)
                bError = True
                Exit Function
            End If
        End If

        ' first make sure we have a response object to handle
        If (response Is Nothing) Or _
            (response.Type Is Nothing) Or _
            (response.Detail Is Nothing) Or _
            (response.Detail.Type Is Nothing) Then
            bError = True
            Exit Function
        End If

        ' make sure we are processing the CompanyActivityQueryRs and 
        ' the CompanyActivityRet responses in this response list
        Dim companyActivityRet As ICompanyActivityRet
        Dim responseType As ENResponseType
        Dim responseDetailType As ENObjectType
        responseType = response.Type.GetValue()
        responseDetailType = response.Detail.Type.GetValue()
        If (responseType = ENResponseType.rtCompanyActivityQueryRs) And _
            (responseDetailType = ENObjectType.otCompanyActivityRet) Then
            ' save the response detail in the appropriate object type
            ' since we have first verified the type of the response object
            companyActivityRet = response.Detail
        Else
            ' bail, we do not have the responses we were expecting
            bError = True
            Exit Function
        End If

        If (companyActivityRet Is Nothing) Or _
            (companyActivityRet.LastRestoreTime Is Nothing) Then
            bError = True
            Exit Function
        End If

        ' refill our customer list if the company file has been restored since 
        ' the last time we read the customer list.
        ' we will back date the sync time by 5 hours.
        ' see the readme for discussion about padding the sync time by 5 hours
        Dim tmpTime As Date
        Dim tmpTimeSpan As System.TimeSpan
        tmpTime = timeLastCustomerSync
        tmpTimeSpan = New System.TimeSpan(5, 0, 0) ' set to 5 hours
        tmpTime = tmpTime.Subtract(tmpTimeSpan)
        If (companyActivityRet.LastRestoreTime.GetValue().CompareTo(tmpTime) > 0) Then

            ' empty the list of old customer information
            lstCustomers.Items.Clear()

            ' refill our information with the current info in the company file
            GetCustomers(lstCustomers, bError)

            CompanyFileRestored = True
        End If

        Exit Function
Errs:
        MsgBox("HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    MsgBoxStyle.Critical, _
                    "Error in CompanyFileRestored")
        bError = True
    End Function

    Private Sub CheckDeletedCustomers(ByRef lstCustomers As System.Windows.Forms.ListBox, _
                ByRef bError As Boolean)
        On Error GoTo Errs

        ' make sure we do not have any old requests still defined
        msgSetRequest.ClearRequests()

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

        ' Add the ListDeletedQuery request
        Dim listDeletedQuery As IListDeletedQuery
        listDeletedQuery = msgSetRequest.AppendListDeletedQueryRq()

        ' Set we are interrested in the deleted customers
        listDeletedQuery.ListDelTypeList.Add(ENListDelType.ldtCustomer)

        'NOTE:  If you are using the Sample Company File, the deleted customers will have a date 
        ' off in the future.  Thus every deleted customer may be returned for every query

        ' set we want the deleted customers since the last time we sync'd customer list
        ' we will back date the sync time by 5 hours.
        ' see the readme for discussion about padding the sync time by 5 hours
        Dim tmpTime As Date
        tmpTime = timeLastCustomerSync
        tmpTime = tmpTime.AddHours(-5)
        listDeletedQuery.DeletedDateRangeFilter.FromDeletedDate.SetValue(tmpTime, False)

        ' send the request to QB
        Dim msgSetResponse As IMsgSetResponse
        msgSetResponse = qbSessionManager.DoRequests(msgSetRequest)

        'MsgBox(msgSetRequest.ToXMLString())
        'MsgBox(msgSetResponse.ToXMLString())

        ' check to make sure we have objects to access first
        ' and that there are responses in the list
        If (msgSetResponse Is Nothing) Or _
            (msgSetResponse.ResponseList Is Nothing) Or _
            (msgSetResponse.ResponseList.Count <= 0) Then
            bError = True
            Exit Sub
        End If

        ' Start parsing the response list
        Dim responseList As IResponseList
        responseList = msgSetResponse.ResponseList

        ' go thru each response and process the response.
        ' this example will only have one response in the list
        ' so we will look at index=0
        Dim response As IResponse
        response = responseList.GetAt(0)
        If (Not response Is Nothing) Then
            If response.StatusCode <> "0" Then
                If response.StatusCode <> "1" Then
                    'If the status is bad, report it to the user
                    MsgBox("CheckDeletedCustomer unexpexcted Error - " & vbCrLf & "StatusCode = " & response.StatusCode & vbCrLf & vbCrLf & response.StatusMessage)
                    bError = True
                End If
                Exit Sub
            End If
        End If

        ' first make sure we have a response object to handle
        If (response Is Nothing) Or _
            (response.Type Is Nothing) Or _
            (response.Detail Is Nothing) Or _
            (response.Detail.Type Is Nothing) Then
            bError = True
            Exit Sub
        End If

        ' make sure we are processing the ListDeletedQueryRs and 
        ' the ListDeletedRet responses in this response list
        Dim listDeletedRetList As IListDeletedRetList
        Dim responseType As ENResponseType
        Dim responseDetailType As ENObjectType
        responseType = response.Type.GetValue()
        responseDetailType = response.Detail.Type.GetValue()
        If (responseType = ENResponseType.rtListDeletedQueryRs) And _
            (responseDetailType = ENObjectType.otListDeletedRetList) Then
            ' save the response detail in the appropriate object type
            ' since we have first verified the type of the response object
            listDeletedRetList = response.Detail
        Else
            ' bail, we do not have the responses we were expecting
            bError = True
            Exit Sub
        End If

        'Parse the query response and delete the Customers to the Customer list box
        Dim count As Short
        Dim index As Short
        Dim listDeletedRet As IListDeletedRet
        Dim customerInfo As CustomerClass
        count = listDeletedRetList.Count
        For index = 0 To count - 1
            listDeletedRet = listDeletedRetList.GetAt(index)
            If (listDeletedRet Is Nothing) Or _
                (listDeletedRet.ListDelType Is Nothing) Or _
                (listDeletedRet.ListID Is Nothing) Or _
                (listDeletedRet.FullName Is Nothing) Then
                bError = True
                Exit Sub
            End If
            ' make sure this is a Customer which was deleted
            If (listDeletedRet.ListDelType.GetValue() = ENListDelType.ldtCustomer) Then

                ' go thru our list box and find the item which was deleted
                Dim i As Short
                Dim numCustomers As Short
                numCustomers = lstCustomers.Items.Count
                For i = 0 To numCustomers - 1
                    customerInfo = lstCustomers.Items.Item(i)
                    If (customerInfo.ListID = listDeletedRet.ListID.GetValue()) Then
                        ' we found the deleted customer so delete it from our list
                        lstCustomers.Items.RemoveAt(i)
                        Exit For
                    End If
                Next

            End If
        Next

        Exit Sub
Errs:
        MsgBox("HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    MsgBoxStyle.Critical, _
                    "Error in CheckDeletedCustomers")
        bError = True
    End Sub

    Private Sub CheckModifiedCustomer(ByRef lstCustomers As System.Windows.Forms.ListBox, _
                ByRef bError As Boolean)
        On Error GoTo Errs

        ' make sure we do not have any old requests still defined
        msgSetRequest.ClearRequests()

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

        ' Add the CustomerQuery request
        Dim customerQuery As ICustomerQuery
        customerQuery = msgSetRequest.AppendCustomerQueryRq()

        ' we want the list of modified customer which are active or inactive so we know
        ' if one customer was inactivated and we will delete this customer from our list
        customerQuery.ORCustomerListQuery.CustomerListFilter.ActiveStatus.SetValue(ENActiveStatus.asAll)

        ' set we want the modified customers since the last time we sync'd customer list
        ' we will back date the sync time by 5 hours.
        ' see the readme for discussion about padding the sync time by 5 hours
        Dim lastModifiedTime As System.DateTime
        lastModifiedTime = timeLastCustomerSync
        lastModifiedTime = lastModifiedTime.AddHours(-5)

        Dim bDone As Boolean = False
        Dim firstListID As String
        Dim numReturnedMultiplier As Short = 1

        Do While (Not bDone)

            ' only look for customer who where modified since our last sync or the last return value
            customerQuery.ORCustomerListQuery.CustomerListFilter.FromModifiedDate.SetValue(lastModifiedTime, False)

            ' we are going to set the number of entries returned to limit the size of the return structure
            customerQuery.ORCustomerListQuery.CustomerListFilter.MaxReturned.SetValue(MAX_RETURNED * numReturnedMultiplier)

            ' send the request to QB
            Dim msgSetResponse As IMsgSetResponse
            msgSetResponse = qbSessionManager.DoRequests(msgSetRequest)

            'MsgBox(msgSetRequest.ToXMLString())
            'MsgBox(msgSetResponse.ToXMLString())

            ' check to make sure we have objects to access first
            ' and that there are responses in the list
            If (msgSetResponse Is Nothing) Or _
                (msgSetResponse.ResponseList Is Nothing) Or _
                (msgSetResponse.ResponseList.Count <= 0) Then
                bError = True
                Exit Sub
            End If

            ' Start parsing the response list
            Dim responseList As IResponseList
            responseList = msgSetResponse.ResponseList

            ' go thru each response and process the response.
            ' this example will only have one response in the list
            ' so we will look at index=0
            Dim response As IResponse
            response = responseList.GetAt(0)
            If (Not response Is Nothing) Then
                If response.StatusCode <> "0" Then
                    If response.StatusCode <> "1" Then
                        'If the status is bad, report it to the user
                        MsgBox("CheckModifiedCustomer unexpexcted Error - " & vbCrLf & "StatusCode = " & response.StatusCode & vbCrLf & vbCrLf & response.StatusMessage)
                        bError = True
                    End If
                    Exit Sub
                End If
            End If

            ' first make sure we have a response object to handle
            If (response Is Nothing) Or _
                (response.Type Is Nothing) Or _
                (response.Detail Is Nothing) Or _
                (response.Detail.Type Is Nothing) Then
                bError = True
                Exit Sub
            End If

            ' make sure we are processing the CustomerQueryRs and 
            ' the CustomerRetList responses in this response list
            Dim customerRetList As ICustomerRetList
            Dim responseType As ENResponseType
            Dim responseDetailType As ENObjectType
            responseType = response.Type.GetValue()
            responseDetailType = response.Detail.Type.GetValue()
            If (responseType = ENResponseType.rtCustomerQueryRs) And _
                (responseDetailType = ENObjectType.otCustomerRetList) Then
                ' save the response detail in the appropriate object type
                ' since we have first verified the type of the response object
                customerRetList = response.Detail
            Else
                ' bail, we do not have the responses we were expecting
                bError = True
                Exit Sub
            End If

            'Parse the query response and add the Customers to the Customer list box
            Dim count As Short
            Dim index As Short
            Dim customerRet As ICustomerRet
            Dim customerInfo As CustomerClass
            count = customerRetList.Count

            ' if we have less than the number of MaxReturned we asked for
            ' we are done
            If (count < (MAX_RETURNED * numReturnedMultiplier)) Then
                bDone = True
            End If

            For index = 0 To count - 1
                customerRet = customerRetList.GetAt(index)
                If (customerRet Is Nothing) Or _
                    (customerRet.ListID Is Nothing) Or _
                    (customerRet.FullName Is Nothing) Then
                    bError = True
                    Exit Sub
                End If
                ' first see if we have the same exact list as before
                If (index = 0) Then
                    If (customerRet.ListID.GetValue() = firstListID) Then
                        ' ooops we have the same list as last time  so expand the 
                        ' size of our return so we will not end up in an endless loop
                        numReturnedMultiplier = numReturnedMultiplier + 1
                        Exit For
                    Else
                        firstListID = customerRet.ListID.GetValue()
                        numReturnedMultiplier = 1
                    End If

                End If

                ' go thru our list box and find the item which was modified
                Dim i As Short
                Dim numCustomers As Short
                Dim bFound As Boolean
                bFound = False
                numCustomers = lstCustomers.Items.Count
                For i = 0 To numCustomers - 1
                    customerInfo = lstCustomers.Items.Item(i)
                    If (customerInfo.ListID = customerRet.ListID.GetValue()) Then
                        ' if the customer is inactive, delete it from our list
                        If (Not customerRet.IsActive Is Nothing) Then
                            If (customerRet.IsActive.GetValue() = False) Then
                                lstCustomers.Items.RemoveAt(i)
                            End If
                        End If
                        ' we found the modified customer so modify it in our list
                        customerInfo.customerRet = customerRet
                        lstCustomers.Refresh()
                        bFound = True
                        Exit For
                    End If
                Next

                ' add the customer if we did not find the modified customer in our list box
                If (Not bFound) Then
                    ' we are saving the FullName and ListID pairs for each Customer
                    ' this is good practice since the FullName can change but the 
                    ' ListID will not change for an Customer.
                    customerInfo = New CustomerClass()
                    ' save all the field information for this customer
                    customerInfo.customerRet = customerRet
                    lstCustomers.Items.Add(customerInfo)
                End If

                If (customerRet.TimeModified Is Nothing) Then
                    ' cannot get out of the loop if we do not get a timeModified
                    ' so we should exit now.
                    bError = True
                    bDone = True
                End If

                ' this value will start the request for the next set of customer records
                ' we will get some overlap of records returned.
                lastModifiedTime = customerRet.TimeModified.GetValue()

            Next
        Loop

        Exit Sub
Errs:
        MsgBox("HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    MsgBoxStyle.Critical, _
                    "Error in CheckModifiedCustomer")
        bError = True
    End Sub

    Private Function FoundCustomerInListBox(ByRef lstCustomers As System.Windows.Forms.ListBox, _
                ByRef listID As String) As Boolean
        On Error GoTo Errs

        FoundCustomerInListBox = False

        Dim customerInfo As CustomerClass

        ' go thru our list box and find the item which was modified
        Dim i As Short
        Dim numCustomers As Short
        numCustomers = lstCustomers.Items.Count
        For i = 0 To numCustomers - 1
            customerInfo = lstCustomers.Items.Item(i)
            If (customerInfo.ListID = listID) Then
                FoundCustomerInListBox = True
                Exit For
            End If
        Next
        Exit Function
Errs:
        MsgBox("HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    MsgBoxStyle.Critical, _
                    "Error in FoundCustomerInListBox")

    End Function

    Public Function SetLastTimeSync(ByRef timeStamp As System.DateTime)
        timeLastCustomerSync = timeStamp
    End Function

End Module
