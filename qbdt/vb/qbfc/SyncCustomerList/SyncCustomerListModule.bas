Attribute VB_Name = "SyncCustomerListModule"
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
' Updated to QBFC 12.0 and fixed setting max QBXML version: 09/2012
'----------------------------------------------------------
    Dim booSessionBegun As Boolean
    Const MAX_RETURNED = 20

    'Module objects
    Dim QBSessionManager As QBSessionManager
    Dim msgSetRequest As IMsgSetRequest
    Dim timeLastCustomerSync As Date
    Dim customerCollection As Collection


    Public Sub OpenConnectionBeginSession()

        booSessionBegun = False

        On Error GoTo Errs

        Set QBSessionManager = New QBSessionManager

        QBSessionManager.OpenConnection "", "IDN Sync Customer List Sample - QBFC"

        QBSessionManager.BeginSession "", ENOpenMode.omDontCare
        booSessionBegun = True

        'Check to make sure the QuickBooks we're working with supports version 2 of the SDK
        Dim booSupports2dot0 As Boolean
        booSupports2dot0 = False
        Dim supportedVersion As String
        supportedVersion = Val(QBFCLatestVersion(QBSessionManager))
        If (supportedVersion >= 6#) Then
            booSupports2dot0 = True
            Set msgSetRequest = QBSessionManager.CreateMsgSetRequest("US", 6, 0)
        ElseIf (supportedVersion >= 5#) Then
            booSupports2dot0 = True
            Set msgSetRequest = QBSessionManager.CreateMsgSetRequest("US", 5, 0)
        ElseIf (supportedVersion >= 4#) Then
            booSupports2dot0 = True
            Set msgSetRequest = QBSessionManager.CreateMsgSetRequest("US", 4, 0)
        ElseIf (supportedVersion >= 3#) Then
            booSupports2dot0 = True
            Set msgSetRequest = QBSessionManager.CreateMsgSetRequest("US", 3, 0)
        ElseIf (supportedVersion >= 2#) Then
            booSupports2dot0 = True
            Set msgSetRequest = QBSessionManager.CreateMsgSetRequest("US", 2, 0)
        End If
        
        If Not booSupports2dot0 Then
            MsgBox "This program only runs against QuickBooks installations which support the 2.0 qbXML spec.  Your version of QuickBooks does not support qbXML 2.0"
            End
        End If
        Exit Sub

Errs:
        If Err.Number = &H80040416 Then
            MsgBox ("You must have QuickBooks running with the company" & vbCrLf & "file open to use this program.")
            QBSessionManager.CloseConnection
            End
        ElseIf Err.Number = &H80040422 Then
            MsgBox ("This QuickBooks company file is open in single user mode and" & vbCrLf & "another application is already accessing it.  Please exit the" & vbCrLf & "other application and run this application again.")
            QBSessionManager.CloseConnection
            End
        Else
            MsgBox "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, vbOKOnly, "Error in OpenConnectionBeginSession"

            If booSessionBegun Then
                QBSessionManager.EndSession
            End If

            QBSessionManager.CloseConnection
            End
        End If
    End Sub

    Public Sub EndSessionCloseConnection()
        On Error GoTo Errs
        
        If (QBSessionManager Is Nothing) Then
            Exit Sub
        End If

        If booSessionBegun Then
            QBSessionManager.EndSession
            QBSessionManager.CloseConnection
        End If

        Exit Sub
Errs:
        MsgBox "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    vbOKOnly, _
                    "Error in EndSessionCloseConnection"
    End Sub

    Public Sub GetCustomers(ByRef lstCustomers As ListBox, _
                        ByRef bError As Boolean)
        On Error GoTo Errs

        If (msgSetRequest Is Nothing) Then
            bError = True
            Exit Sub
        End If
        
        ' make sure we do not have any old requests still defined
        msgSetRequest.ClearRequests

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

        ' Add the CustomerQuery request
        Dim customerQuery As ICustomerQuery
        Set customerQuery = msgSetRequest.AppendCustomerQueryRq()

        ' in this example we are going to keep a list of only "active" status Customers
        ' the default is receiving only the active status customers, but we will
        ' set the value anyway for illustration.
        customerQuery.ORCustomerListQuery.CustomerListFilter.ActiveStatus.SetValue (ENActiveStatus.asActiveOnly)

        ' we are going to set the number of entries returned to limit the size of the return structure
        customerQuery.ORCustomerListQuery.CustomerListFilter.MaxReturned.SetValue (MAX_RETURNED)

        Dim bDone As Boolean
        bDone = False
        Dim firstFullName As String
        firstFullName = "!"

        'Clear the list box and collections
        lstCustomers.Clear
        If (customerCollection Is Nothing) Then
            Set customerCollection = New Collection
        End If
        Do While (customerCollection.count > 0)
            customerCollection.Remove (1)
        Loop

        Do While (Not bDone)

            ' start looking for customer next in the list
            ' we will have one overlap
            customerQuery.ORCustomerListQuery.CustomerListFilter.ORNameFilter.NameRangeFilter.FromName.SetValue (firstFullName)

            ' send the request to QB
            Dim msgSetResponse As IMsgSetResponse
            Set msgSetResponse = QBSessionManager.DoRequests(msgSetRequest)

            'MsgBox(msgSetRequest.ToXMLString())

            FillCustomerListBox msgSetResponse, lstCustomers, bDone, firstFullName, bError

        Loop

        Exit Sub

Errs:
        MsgBox "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    vbOKOnly, _
                    "Error in GetCustomers"
        bError = True
    End Sub

    Public Sub FillCustomerListBox(ByRef msgSetResponse As IMsgSetResponse, _
                    ByRef lstCustomers As ListBox, _
                    ByRef bDone As Boolean, _
                    ByRef lastFullName As String, _
                    ByRef bError As Boolean)
        On Error GoTo Errs

        ' check to make sure we have objects to access first
        ' and that there are responses in the list
        If (msgSetResponse Is Nothing) Then
            bDone = True
            bError = True
            Exit Sub
        End If
        If (msgSetResponse.responseList Is Nothing) Then
            bDone = True
            bError = True
            Exit Sub
        End If
        If (msgSetResponse.responseList.count <= 0) Then
            bDone = True
            bError = True
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
                MsgBox ("FillCustomerListBox unexpexcted Error - " & response.StatusMessage)
                bDone = True
                bError = True
                Exit Sub
            End If
        End If

        ' first make sure we have a response object to handle
        If (response Is Nothing) Then
            bDone = True
            bError = True
            Exit Sub
        End If
        If (response.Type Is Nothing) Or _
            (response.Detail Is Nothing) Then
            bDone = True
            bError = True
            Exit Sub
        End If
        If (response.Detail.Type Is Nothing) Then
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
            Set customerRetList = response.Detail
        Else
            ' bail, we do not have the responses we were expecting
            bDone = True
            bError = True
            Exit Sub
        End If

        'Parse the query response and add the Customers to the Customer list box
        Dim count As Integer
        Dim index As Integer
        Dim customerRet As ICustomerRet
        Dim customerInfo As CustomerClass
        count = customerRetList.count

        ' we are done with the customerQueries if we have not received the MaxReturned
        If (count < MAX_RETURNED) Then
            bDone = True
        End If

        For index = 0 To count - 1
            ' skip this customer if this is a repeat from the last query
            Set customerRet = customerRetList.GetAt(index)
            If (customerRet Is Nothing) Then
                bDone = True
                bError = True
                Exit Sub
            End If
            If (customerRet.FullName Is Nothing) Or _
                (customerRet.listID Is Nothing) Then
                bDone = True
                bError = True
                Exit Sub
            End If
            ' only the first customerRet should be repeating and then
            ' lets just check to make sure we do not have the customer
            ' just in case another app changed a customer right between our
            ' queries.
            If (index <> 0) Or (Not FoundCustomerInListBox(customerRet.listID.GetValue())) Then
                ' we are saving the FullName and ListID pairs for each Customer
                ' this is good practice since the FullName can change but the
                ' ListID will not change for an Customer.
                Set customerInfo = New CustomerClass
                ' save all the field information for this customer
                customerInfo.SetCustomerRet customerRet
                lstCustomers.AddItem customerInfo.GetFullName
                customerCollection.Add customerInfo, customerInfo.GetListID
            End If
            lastFullName = customerRet.FullName.GetValue()
        Next

        Exit Sub
Errs:
        MsgBox "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    vbOKOnly, _
                    "Error in FillCustomerListBox"
        bDone = True
        bError = True
    End Sub

    Public Sub SyncCustomerList(ByRef lstCustomers As ListBox, _
                ByRef bError As Boolean)
        On Error GoTo Errs

        Dim today, tmpDate As Date
        today = Now
        tmpDate = DateAdd("m", 3, timeLastCustomerSync)

        ' if last resync is more than 3 months ago, we need to grab the whole list
        ' again.  The ListDeletedQuery will only save the last 3 months worth of deletes
        If (tmpDate <= today) Then
            GetCustomers lstCustomers, bError
            Exit Sub
        End If

        ' see if the company file was restored.  If so, then the customer list
        ' was totally reread and reset with the current values in the company file.
        If (Not CompanyFileRestored(lstCustomers, bError)) And _
            (Not bError) Then

            ' check to see if any customers have been deleted
            CheckDeletedCustomers lstCustomers, bError

            If (bError) Then
                Exit Sub
            End If

            ' check to see if any customers have been modified
            CheckModifiedCustomer lstCustomers, bError

        End If

        Exit Sub
Errs:
        MsgBox "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    vbOKOnly, _
                    "Error in SyncCustomerList"
        bError = True
    End Sub

    Private Function CompanyFileRestored(ByRef lstCustomers As ListBox, _
                ByRef bError As Boolean) _
    As Boolean
        On Error GoTo Errs

        CompanyFileRestored = False

        If (msgSetRequest Is Nothing) Then
            bError = True
            Exit Function
        End If
        
        ' make sure we do not have any old requests still defined
        msgSetRequest.ClearRequests

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

        ' Add the CompanyActivityQuery request
       msgSetRequest.AppendCompanyActivityQueryRq

        ' send the request to QB
        Dim msgSetResponse As IMsgSetResponse
        Set msgSetResponse = QBSessionManager.DoRequests(msgSetRequest)

        ' check to make sure we have objects to access first
        ' and that there are responses in the list
        If (msgSetResponse Is Nothing) Then
            bError = True
            Exit Function
        End If
        If (msgSetResponse.responseList Is Nothing) Then
            bError = True
            Exit Function
        End If
        If (msgSetResponse.responseList.count <= 0) Then
            bError = True
            Exit Function
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
                MsgBox ("CompanyFileRestored unexpexcted Error - " & vbCrLf & "StatusCode = " & response.StatusCode & vbCrLf & vbCrLf & response.StatusMessage)
                bError = True
                Exit Function
            End If
        End If

        ' first make sure we have a response object to handle
        If (response Is Nothing) Then
            bError = True
            Exit Function
        End If
        If (response.Type Is Nothing) Or _
            (response.Detail Is Nothing) Then
            bError = True
            Exit Function
        End If
        If (response.Detail.Type Is Nothing) Then
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
            Set companyActivityRet = response.Detail
        Else
            ' bail, we do not have the responses we were expecting
            bError = True
            Exit Function
        End If

        If (companyActivityRet Is Nothing) Then
            bError = True
            Exit Function
        End If
        If (companyActivityRet.LastRestoreTime Is Nothing) Then
            bError = True
            Exit Function
        End If

        ' refill our customer list if the company file has been restored since
        ' the last time we read the customer list.
        ' we will back date the sync time by 5 hours.
        ' see the readme for discussion about padding the sync time by 5 hours
        Dim lastCustomerSync As Date
        lastCustomerSync = DateAdd("h", -5, timeLastCustomerSync)
        Dim lastCompanyRestore As Date
        lastCompanyRestore = companyActivityRet.LastRestoreTime.GetValue
        If (lastCompanyRestore > lastCustomerSync) Then

            ' refill our information with the current info in the company file
            GetCustomers lstCustomers, bError

            CompanyFileRestored = True
        End If

        Exit Function
Errs:
        MsgBox "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    vbOKOnly, _
                    "Error in CompanyFileRestored"
        bError = True
    End Function

    Private Sub CheckDeletedCustomers(ByRef lstCustomers As ListBox, _
                ByRef bError As Boolean)
        On Error GoTo Errs

        If (msgSetRequest Is Nothing) Then
            bError = True
            Exit Sub
        End If
        
        ' make sure we do not have any old requests still defined
        msgSetRequest.ClearRequests

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

        ' Add the ListDeletedQuery request
        Dim listDeletedQuery As IListDeletedQuery
        Set listDeletedQuery = msgSetRequest.AppendListDeletedQueryRq()

        ' Set we are interrested in the deleted customers
        listDeletedQuery.ListDelTypeList.Add (ENListDelType.ldtCustomer)

        'NOTE:  If you are using the Sample Company File, the deleted customers will have a date
        ' off in the future.  Since the "ToDeletedDate" is not specified, the QuickBooks
        ' uses the current date for ToDeletedDate.  As a result if you are using the
        ' sample company file, you may not detect any deleted customer.
        ' Suggest you do not use a sample company file to test this application.
        
        ' set we want the deleted customers since the last time we sync'd customer list
        ' we will back date the sync time by 5 hours.
        ' see the readme for discussion about padding the sync time by 5 hours
        Dim tmpTime As Date
        tmpTime = DateAdd("h", -5, timeLastCustomerSync)
        listDeletedQuery.DeletedDateRangeFilter.FromDeletedDate.SetValue tmpTime, False

        ' send the request to QB
        Dim msgSetResponse As IMsgSetResponse
        Set msgSetResponse = QBSessionManager.DoRequests(msgSetRequest)

        'MsgBox(msgSetRequest.ToXMLString())
        'MsgBox(msgSetResponse.ToXMLString())

        ' check to make sure we have objects to access first
        ' and that there are responses in the list
        If (msgSetResponse Is Nothing) Then
            bError = True
            Exit Sub
        End If
        If (msgSetResponse.responseList Is Nothing) Then
            bError = True
            Exit Sub
        End If
        If (msgSetResponse.responseList.count <= 0) Then
            bError = True
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
                If response.StatusCode <> "1" Then
                    'If the status is bad, report it to the user
                    MsgBox ("CheckDeletedCustomer unexpexcted Error - " & vbCrLf & "StatusCode = " & response.StatusCode & vbCrLf & vbCrLf & response.StatusMessage)
                    bError = True
                End If
                Exit Sub
            End If
        End If

        ' first make sure we have a response object to handle
        If (response Is Nothing) Then
            bError = True
            Exit Sub
        End If
        If (response.Type Is Nothing) Or _
            (response.Detail Is Nothing) Then
            bError = True
            Exit Sub
        End If
        If (response.Detail.Type Is Nothing) Then
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
            Set listDeletedRetList = response.Detail
        Else
            ' bail, we do not have the responses we were expecting
            bError = True
            Exit Sub
        End If

        'Parse the query response and delete the Customers to the Customer list box
        Dim count As Integer
        Dim index As Integer
        Dim listDeletedRet As IListDeletedRet
        Dim customerInfo As CustomerClass
        count = listDeletedRetList.count
        For index = 0 To count - 1
            Set listDeletedRet = listDeletedRetList.GetAt(index)
            If (listDeletedRet Is Nothing) Then
                bError = True
                Exit Sub
            End If
            If (listDeletedRet.ListDelType Is Nothing) Or _
                (listDeletedRet.listID Is Nothing) Or _
                (listDeletedRet.FullName Is Nothing) Then
                bError = True
                Exit Sub
            End If
            ' make sure this is a Customer which was deleted
            If (listDeletedRet.ListDelType.GetValue() = ENListDelType.ldtCustomer) Then

                ' go thru our list box and find the item which was deleted
                Dim i As Integer
                Dim numCustomers As Integer
                numCustomers = lstCustomers.ListCount
                For i = 0 To numCustomers - 1
                    Set customerInfo = customerCollection.Item(i + 1)
                    If (customerInfo.GetListID = listDeletedRet.listID.GetValue()) Then
                        ' we found the deleted customer so delete it from our list
                        lstCustomers.RemoveItem (i)
                        lstCustomers.Refresh
                        customerCollection.Remove (i + 1)
                        Exit For
                    End If
                Next

            End If
        Next

        Exit Sub
Errs:
        MsgBox "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    vbOKOnly, _
                    "Error in CheckDeletedCustomers"
        bError = True
    End Sub

    Private Sub CheckModifiedCustomer(ByRef lstCustomers As ListBox, _
                ByRef bError As Boolean)
        On Error GoTo Errs

        If (msgSetRequest Is Nothing) Then
            bError = True
            Exit Sub
        End If
        
        ' make sure we do not have any old requests still defined
        msgSetRequest.ClearRequests

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

        ' Add the CustomerQuery request
        Dim customerQuery As ICustomerQuery
        Set customerQuery = msgSetRequest.AppendCustomerQueryRq()

        ' we want the list of modified customer which are active or inactive so we know
        ' if one customer was inactivated and we will delete this customer from our list
        customerQuery.ORCustomerListQuery.CustomerListFilter.ActiveStatus.SetValue (ENActiveStatus.asAll)

        ' set we want the modified customers since the last time we sync'd customer list
        ' we will back date the sync time by 5 hours.
        ' see the readme for discussion about padding the sync time by 5 hours
        Dim lastModifiedTime As Date
        lastModifiedTime = DateAdd("h", -5, timeLastCustomerSync)

        Dim bDone As Boolean
        bDone = False
        Dim firstListID As String
        Dim numReturnedMultiplier As Integer
        numReturnedMultiplier = 1

        Do While (Not bDone)

            ' only look for customer who where modified since our last sync or the last return value
            customerQuery.ORCustomerListQuery.CustomerListFilter.FromModifiedDate.SetValue lastModifiedTime, False

            ' we are going to set the number of entries returned to limit the size of the return structure
            customerQuery.ORCustomerListQuery.CustomerListFilter.MaxReturned.SetValue (MAX_RETURNED * numReturnedMultiplier)

            ' send the request to QB
            Dim msgSetResponse As IMsgSetResponse
            Set msgSetResponse = QBSessionManager.DoRequests(msgSetRequest)

            'MsgBox(msgSetRequest.ToXMLString())
            'MsgBox(msgSetResponse.ToXMLString())

            ' check to make sure we have objects to access first
            ' and that there are responses in the list
            If (msgSetResponse Is Nothing) Then
                bError = True
                Exit Sub
            End If
            If (msgSetResponse.responseList Is Nothing) Then
                bError = True
                Exit Sub
            End If
            If (msgSetResponse.responseList.count <= 0) Then
                bError = True
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
                    If response.StatusCode <> "1" Then
                        'If the status is bad, report it to the user
                        MsgBox ("CheckModifiedCustomer unexpexcted Error - " & vbCrLf & "StatusCode = " & response.StatusCode & vbCrLf & vbCrLf & response.StatusMessage)
                        bError = True
                    End If
                    Exit Sub
                End If
            End If

            ' first make sure we have a response object to handle
            If (response Is Nothing) Then
                bError = True
                Exit Sub
            End If
            If (response.Type Is Nothing) Or _
                (response.Detail Is Nothing) Then
                bError = True
                Exit Sub
            End If
            If (response.Detail.Type Is Nothing) Then
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
                Set customerRetList = response.Detail
            Else
                ' bail, we do not have the responses we were expecting
                bError = True
                Exit Sub
            End If

            'Parse the query response and add the Customers to the Customer list box
            Dim count As Integer
            Dim index As Integer
            Dim customerRet As ICustomerRet
            Dim customerInfo As CustomerClass
            count = customerRetList.count

            ' if we have less than the number of MaxReturned we asked for
            ' we are done
            If (count < (MAX_RETURNED * numReturnedMultiplier)) Then
                bDone = True
            End If

            For index = 0 To count - 1
                Set customerRet = customerRetList.GetAt(index)
                If (customerRet Is Nothing) Then
                    bError = True
                    Exit Sub
                End If
                If (customerRet.listID Is Nothing) Or _
                    (customerRet.FullName Is Nothing) Then
                    bError = True
                    Exit Sub
                End If
                ' first see if we have the same exact list as before
                If (index = 0) Then
                    If (customerRet.listID.GetValue() = firstListID) Then
                        ' ooops we have the same list as last time  so expand the
                        ' size of our return so we will not end up in an endless loop
                        numReturnedMultiplier = numReturnedMultiplier + 1
                        Exit For
                    Else
                        firstListID = customerRet.listID.GetValue()
                        numReturnedMultiplier = 1
                    End If

                End If

                ' go thru our list box and find the item which was modified
                Dim i As Integer
                Dim numCustomers As Integer
                Dim bFound As Boolean
                bFound = False
                numCustomers = lstCustomers.ListCount
                For i = 0 To numCustomers - 1
                    Set customerInfo = customerCollection.Item(i + 1)
                    If (customerInfo.GetListID = customerRet.listID.GetValue()) Then
                        ' if the customer is inactive, delete it from our list
                        If (Not customerRet.IsActive Is Nothing) Then
                            If (customerRet.IsActive.GetValue() = False) Then
                                lstCustomers.RemoveItem (i)
                                lstCustomers.Refresh
                                customerCollection.Remove (i + 1)
                            End If
                        End If
                        ' we found the modified customer so modify it in our list
                        customerInfo.SetCustomerRet customerRet
                        lstCustomers.List(i) = customerInfo.GetFullName
                        lstCustomers.Refresh
                        bFound = True
                        Exit For
                    End If
                Next

                ' add the customer if we did not find the modified customer in our list box
                If (Not bFound) Then
                    ' we are saving the FullName and ListID pairs for each Customer
                    ' this is good practice since the FullName can change but the
                    ' ListID will not change for an Customer.
                    Set customerInfo = New CustomerClass
                    ' save all the field information for this customer
                    customerInfo.SetCustomerRet customerRet
                    lstCustomers.AddItem customerInfo.GetFullName
                    lstCustomers.Refresh
                    customerCollection.Add customerInfo, customerInfo.GetListID
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
        MsgBox "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    vbOKOnly, _
                    "Error in CheckModifiedCustomer"
        bError = True
    End Sub

    Private Function FoundCustomerInListBox(ByRef listID As String)
        On Error GoTo Errs

        Dim customerInfo As CustomerClass
        
        Set customerInfo = customerCollection.Item(listID)
        
        ' if the above line is successful, then the listID is in the list
        FoundCustomerInListBox = True

        Exit Function
Errs:
        FoundCustomerInListBox = False

    End Function

    Public Function SetLastTimeSync(ByRef timeStamp As Date)
        timeLastCustomerSync = timeStamp
    End Function

Function QBFCLatestVersion(SessionManager As QBSessionManager) As String
    Dim strXMLVersions() As String
    'Should be able to use this, but there appears to be a bug that may cause 2.0 to be returned
    'when it should not.
    'strXMLVersions = SessionManager.QBXMLVersionsForSession
    
    Dim msgset As IMsgSetRequest
    'Use oldest version to ensure that we work with any QuickBooks (US)
    Set msgset = SessionManager.CreateMsgSetRequest("US", 1, 0)
    msgset.AppendHostQueryRq
    Dim QueryResponse As IMsgSetResponse
    Set QueryResponse = SessionManager.DoRequests(msgset)
    Dim response As IResponse
    
    ' The response list contains only one response,
    ' which corresponds to our single HostQuery request
    Set response = QueryResponse.responseList.GetAt(0)
    Dim HostResponse As IHostRet
    Set HostResponse = response.Detail
    Dim supportedVersions As IBSTRList
    Set supportedVersions = HostResponse.SupportedQBXMLVersionList
    
    Dim i As Long
    Dim vers As Double
    Dim LastVers As Double
    LastVers = 0
    For i = 0 To supportedVersions.count - 1
        vers = Val(supportedVersions.GetAt(i))
        If (vers > LastVers) Then
            LastVers = vers
            QBFCLatestVersion = supportedVersions.GetAt(i)
        End If
    Next i
End Function

