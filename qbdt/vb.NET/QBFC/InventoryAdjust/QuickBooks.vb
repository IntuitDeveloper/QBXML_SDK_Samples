Option Strict Off
Option Explicit On
Module QuickBooks
	'----------------------------------------------------------
	' Module: QuickBoooksModule
	'
	' Description: This module contains the code which illustrates using
	'               the InventoryAdjust feature.
	'
	' Routines: qbConnect
	'             Opens a connection and begins a sesson with the
	'             currently open company file.  If a company isn't open,
	'             the routine will display a message and then exit the
	'             program.
	'
	'           EndSessionCloseConnection
	'             Calls EndSession and CloseConnection if the boolean
	'             booSessionBegun is true.
	'
	'           AdjustInventory
	'             This procedure will construct and send an InventoryAdjustmentAdd
	'             request.  Then the response will be parsed for the status.
	'
	'           ClassesEnabled
	'             This procedure sends a Preferences Query to determine
	'             if the company file has "classes" enbled to be used.
	'
	'           fillAccountList
	'             This procedure will construct the AccountQuery and fill
	'             the account combo box with the FullNames of the accounts.
	'
	'           fillClassComboBox
	'             This procedure will parse the results of a ClassQuery
	'             and fill in the Class combo box with the Names returned.
	'
	'           fillClassList
	'             This procedure will construct a ClassQuery and send the
	'             response to the fillClassComboBox procedure to fill in
	'             the Class combo box
	'
	'           fillCustomerComboBox
	'             This procedure will parse the results of the CustomerQuery
	'             and fill in the Customer combo box with the FullNames returned.
	'
	'           fillCustomerList
	'             This procedure will construct a CustomerQuery and send
	'             the response to the fillCustomerComboBox procedure to fill
	'             in the Customer combo box.
	'
	'           fillItemComboBox
	'             This procedure will parst the results of the ItemInventoryQuery
	'             and fill in the Item combo box with the FullNames returned.
	'
	'           fillItemList
	'             This procedure will construct a ItemInventoryQuery and
	'             send the response to the fillItemComboBox to fill in the
	'             Item combo box.
	'
	'           FoundInComboBox
	'             This procedure will go thru the elements of the combo
	'             box to determine if there is already an element with this
	'             name already in the list.
	'
	' Copyright © 2002-2020 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	' Updated to QBFC 12.0 and fixed setting max QBXML version: 09/2012
	'----------------------------------------------------------
	Const MAX_RETURNED As Short = 20
	Dim booSessionBegun As Boolean
	Dim booConnectionOpen As Boolean
    'UPGRADE_WARNING: Arrays in structure SessionManager may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
    Dim SessionManager As QBFC15Lib.QBSessionManager
    Dim msgSetRequest As QBFC15Lib.IMsgSetRequest

    Function qbConnect() As Boolean
        On Error GoTo errHandler

        If (booSessionBegun) Then
            qbConnect = True
            Exit Function
        End If

        ' create the new QBSessionManager object
        If (SessionManager Is Nothing) Then
            SessionManager = New QBFC15Lib.QBSessionManager
        End If

        ' open the connection to QuickBooks
        If (Not booConnectionOpen) Then
            SessionManager.OpenConnection("", "IDN Inventory Adjust Sample - QBFC")
            booConnectionOpen = True
        End If

        ' InventoryAdjust must be done in SingleUser mode
        ' Begin a session with QuickBooks
        SessionManager.BeginSession("", QBFC15Lib.ENOpenMode.omSingleUser)
        booSessionBegun = True

        'Check to make sure the QuickBooks we're working with supports version 2 of the SDK
        Dim booSupports2dot0 As Boolean
        booSupports2dot0 = False
        Dim supportedVersion As String
        supportedVersion = CStr(Val(QBFCLatestVersion(SessionManager)))
        If (supportedVersion >= CStr(6.0#)) Then
            booSupports2dot0 = True
            msgSetRequest = SessionManager.CreateMsgSetRequest("US", 6, 0)
        ElseIf (supportedVersion >= CStr(5.0#)) Then
            booSupports2dot0 = True
            msgSetRequest = SessionManager.CreateMsgSetRequest("US", 5, 0)
        ElseIf (supportedVersion >= CStr(4.0#)) Then
            booSupports2dot0 = True
            msgSetRequest = SessionManager.CreateMsgSetRequest("US", 4, 0)
        ElseIf (supportedVersion >= CStr(3.0#)) Then
            booSupports2dot0 = True
            msgSetRequest = SessionManager.CreateMsgSetRequest("US", 3, 0)
        ElseIf (supportedVersion >= CStr(2.0#)) Then
            booSupports2dot0 = True
            msgSetRequest = SessionManager.CreateMsgSetRequest("US", 2, 0)
        ElseIf (supportedVersion >= CStr(15.0#)) Then
            booSupports2dot0 = True
            msgSetRequest = SessionManager.CreateMsgSetRequest("US", 15, 0)
        End If

        If Not booSupports2dot0 Then
            MsgBox("This program only runs against QuickBooks installations which support the 15.0 qbXML spec.  Your version of QuickBooks does not support qbXML 15.0")
            SessionManager.EndSession()
            SessionManager.CloseConnection()
            End
        End If

        qbConnect = True
        Exit Function

errHandler:
        MsgBox("Error in qbConnect" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)

        If booSessionBegun Then
            SessionManager.EndSession()
        End If
        If booConnectionOpen Then
            SessionManager.CloseConnection()
        End If

        qbConnect = False
        End
    End Function
    Public Sub EndSessionCloseConnection()
        On Error GoTo Errs
        If booSessionBegun Then
            SessionManager.EndSession()
            SessionManager.CloseConnection()
        End If
        Exit Sub
Errs:
        MsgBox("Error in EndSessionCloseConnection" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)

    End Sub
    Function ClassesEnabled() As Boolean
        On Error GoTo Errs

        Dim msgSetResponse As QBFC15Lib.IMsgSetResponse

        ClassesEnabled = False
        msgSetRequest.ClearRequests()

        ' Add the PreferencesQueryRq request
        msgSetRequest.AppendPreferencesQueryRq()

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = QBFC15Lib.ENRqOnError.roeContinue

        'MsgBox msgSetRequest.ToXMLString()

        ' send the request to QuickBooks
        msgSetResponse = SessionManager.DoRequests(msgSetRequest)

        'MsgBox msgSetResponse.ToXMLString()

        ' check to make sure we have objects to access first
        ' and that there are responses in the list
        If (msgSetResponse Is Nothing) Then
            ClassesEnabled = False
            Exit Function
        End If
        If (msgSetResponse.ResponseList Is Nothing) Then
            ClassesEnabled = False
            Exit Function
        End If
        If (msgSetResponse.ResponseList.Count <= 0) Then
            ClassesEnabled = False
            Exit Function
        End If

        ' Start parsing the response list
        Dim responseList As QBFC15Lib.IResponseList
        responseList = msgSetResponse.ResponseList

        ' go thru each response and process the response.
        Dim response As QBFC15Lib.IResponse
        response = responseList.GetAt(0)
        If (response Is Nothing) Then
            ClassesEnabled = False
            Exit Function
        End If
        If response.StatusCode <> CDbl("0") Then
            'If the status is bad, report it to the user
            MsgBox("ClassesEnabled unexpected Error - " & response.StatusMessage)
            ClassesEnabled = False
            Exit Function
        End If

        ' first make sure we have a response object to handle
        If (response Is Nothing) Then
            ClassesEnabled = False
            Exit Function
        End If
        If (response.Type Is Nothing) Then
            ClassesEnabled = False
            Exit Function
        End If
        If (response.Detail Is Nothing) Then
            ClassesEnabled = False
            Exit Function
        End If
        If (response.Detail.Type Is Nothing) Then
            ClassesEnabled = False
            Exit Function
        End If

        ' make sure we are processing the PreferencesQueryRs and
        ' the PreferencesRet responses in this response list
        Dim preferencesRet As QBFC15Lib.IPreferencesRet
        Dim responseType As QBFC15Lib.ENResponseType
        Dim responseDetailType As QBFC15Lib.ENObjectType
        responseType = response.Type.GetValue()
        responseDetailType = response.Detail.Type.GetValue()
        If (responseType = QBFC15Lib.ENResponseType.rtPreferencesQueryRs) And (responseDetailType = QBFC15Lib.ENObjectType.otPreferencesRet) Then
            ' save the response detail in the appropriate object type
            ' since we have first verified the type of the response object
            preferencesRet = response.Detail
        Else
            ' bail, we do not have the responses we were expecting
            ClassesEnabled = False
            Exit Function
        End If

        'get the query response return the preferencesRet variable
        If (Not preferencesRet.AccountingPreferences Is Nothing) Then
            If (Not preferencesRet.AccountingPreferences.IsUsingClassTracking Is Nothing) Then
                ClassesEnabled = preferencesRet.AccountingPreferences.IsUsingClassTracking.GetValue()
                Exit Function
            End If
        End If

        ClassesEnabled = False
        Exit Function
Errs:
        MsgBox("Error in ClassesEnabled" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)

    End Function
    Sub fillAccountList(ByRef AccountList As System.Windows.Forms.ComboBox)
        On Error GoTo Errs

        Dim msgSetResponse As QBFC15Lib.IMsgSetResponse

        msgSetRequest.ClearRequests()

        ' Add the AccountQueryRq request
        Dim accountQuery As QBFC15Lib.IAccountQuery
        accountQuery = msgSetRequest.AppendAccountQueryRq()

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = QBFC15Lib.ENRqOnError.roeContinue

        'MsgBox msgSetRequest.ToXMLString()

        ' send the request to QuickBooks
        msgSetResponse = SessionManager.DoRequests(msgSetRequest)

        'MsgBox msgSetResponse.ToXMLString()

        ' check to make sure we have objects to access first
        ' and that there are responses in the list
        If (msgSetResponse Is Nothing) Then
            Exit Sub
        End If
        If (msgSetResponse.ResponseList Is Nothing) Then
            Exit Sub
        End If
        If (msgSetResponse.ResponseList.Count <= 0) Then
            Exit Sub
        End If

        ' Start parsing the response list
        Dim responseList As QBFC15Lib.IResponseList
        responseList = msgSetResponse.ResponseList

        ' go thru each response and process the response.
        Dim response As QBFC15Lib.IResponse
        ' we will only have one response from the one AccountQuery
        response = responseList.GetAt(0)
        If (response Is Nothing) Then
            Exit Sub
        End If
        If response.StatusCode <> CDbl("0") Then
            'If the status is bad, report it to the user
            MsgBox("fillAccountList unexpected Error - " & response.StatusMessage)
            Exit Sub
        End If

        ' first make sure we have a response object to handle
        If (response Is Nothing) Then
            Exit Sub
        End If
        If (response.Type Is Nothing) Then
            Exit Sub
        End If
        If (response.Detail Is Nothing) Then
            Exit Sub
        End If
        If (response.Detail.Type Is Nothing) Then
            Exit Sub
        End If

        ' make sure we are processing the AccountQueryRs and
        ' the AccountRetList responses in this response list
        Dim accountRetList As QBFC15Lib.IAccountRetList
        Dim responseType As QBFC15Lib.ENResponseType
        Dim responseDetailType As QBFC15Lib.ENObjectType
        responseType = response.Type.GetValue()
        responseDetailType = response.Detail.Type.GetValue()
        If (responseType = QBFC15Lib.ENResponseType.rtAccountQueryRs) And (responseDetailType = QBFC15Lib.ENObjectType.otAccountRetList) Then
            ' save the response detail in the appropriate object type
            ' since we have first verified the type of the response object
            accountRetList = response.Detail
        Else
            ' bail, we do not have the responses we were expecting
            Exit Sub
        End If

        'Parse the query response and add the Estimate to the Estimate list box
        Dim count As Short
        Dim index As Short
        Dim accountRet As QBFC15Lib.IAccountRet
        count = accountRetList.Count

        For index = 0 To count - 1
            accountRet = accountRetList.GetAt(index)
            If (accountRet Is Nothing) Then
                Exit Sub
            End If
            If (accountRet.FullName Is Nothing) Then
                Exit Sub
            End If
            AccountList.Items.Add((accountRet.FullName.GetValue()))
        Next

        Exit Sub
Errs:
        MsgBox("Error in fillAccountList" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)
    End Sub
    Sub fillItemList(ByRef ItemList As System.Windows.Forms.ComboBox)
        On Error GoTo Errs

        ' make sure we do not have any old requests still defined
        msgSetRequest.ClearRequests()

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = QBFC15Lib.ENRqOnError.roeContinue

        ' Add the ItemInventoryQuery request
        Dim itemInventoryQuery As QBFC15Lib.IItemInventoryQuery
        itemInventoryQuery = msgSetRequest.AppendItemInventoryQueryRq()

        ' we are going to set the number of entries returned to limit the size of the return structure
        itemInventoryQuery.ORListQueryWithOwnerIDAndClass.ListWithClassFilter.MaxReturned.SetValue((MAX_RETURNED))

        Dim bDone As Boolean
        bDone = False
        Dim firstFullName As String
        firstFullName = "!"

        'Clear the list box
        ItemList.Items.Clear()

        Dim msgSetResponse As QBFC15Lib.IMsgSetResponse
        Do While (Not bDone)

            ' start looking for itemInventory next in the list
            ' we will have one overlap
            itemInventoryQuery.ORListQueryWithOwnerIDAndClass.ListWithClassFilter.ORNameFilter.NameRangeFilter.FromName.SetValue((firstFullName))

            ' send the request to QB
            msgSetResponse = SessionManager.DoRequests(msgSetRequest)

            'MsgBox(msgSetRequest.ToXMLString())

            ' fill the ItemList box with the list of Item Inventory FullNames
            ' from the response returned from QuickBooks
            fillItemComboBox(msgSetResponse, ItemList, bDone, firstFullName)

        Loop

        Exit Sub

Errs:
        MsgBox("Error in fillItemList" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)
    End Sub
    Public Sub fillItemComboBox(ByRef msgSetResponse As QBFC15Lib.IMsgSetResponse, ByRef listComboBox As System.Windows.Forms.ComboBox, ByRef bDone As Boolean, ByRef lastFullName As String)
        On Error GoTo Errs

        ' check to make sure we have objects to access first
        ' and that there are responses in the list
        If (msgSetResponse Is Nothing) Then
            bDone = True
            Exit Sub
        End If
        If (msgSetResponse.ResponseList Is Nothing) Then
            bDone = True
            Exit Sub
        End If
        If (msgSetResponse.ResponseList.Count <= 0) Then
            bDone = True
            Exit Sub
        End If

        ' Start parsing the response list
        Dim responseList As QBFC15Lib.IResponseList
        responseList = msgSetResponse.ResponseList

        ' go thru each response and process the response.
        ' this example will only have one response in the list
        ' so we will look at index=0
        Dim response As QBFC15Lib.IResponse
        response = responseList.GetAt(0)
        If (Not response Is Nothing) Then
            If response.StatusCode <> CDbl("0") Then
                'If the status is bad, report it to the user
                MsgBox("fillItemComboBox unexpected Error - " & response.StatusMessage)
                bDone = True
                Exit Sub
            End If
        End If

        ' first make sure we have a response object to handle
        If (response Is Nothing) Then
            bDone = True
            Exit Sub
        End If
        If (response.Type Is Nothing) Or (response.Detail Is Nothing) Then
            bDone = True
            Exit Sub
        End If
        If (response.Detail.Type Is Nothing) Then
            bDone = True
            Exit Sub
        End If

        ' make sure we are processing the correct IItemInventoryQueryRs and
        ' the correct IItemInventoryRetList responses in this response list
        Dim itemInventoryRetList As QBFC15Lib.IItemInventoryRetList
        If (QBFC15Lib.ENResponseType.rtItemInventoryQueryRs = response.Type.GetValue()) And (QBFC15Lib.ENObjectType.otItemInventoryRetList = response.Detail.Type.GetValue()) Then
            ' save the response detail in the appropriate object type
            ' since we have first verified the type of the response object
            itemInventoryRetList = response.Detail
        Else
            ' bail, we do not have the responses we were expecting
            bDone = True
            Exit Sub
        End If

        'Parse the query response and add the names to the combo box
        Dim count As Short
        Dim index As Short
        Dim itemInventoryRet As QBFC15Lib.IItemInventoryRet
        count = itemInventoryRetList.Count

        ' we are done with the Queries if we have not received the MaxReturned
        If (count < MAX_RETURNED) Then
            bDone = True
        End If

        For index = 0 To count - 1
            ' skip this item if this is a repeat from the last query
            itemInventoryRet = itemInventoryRetList.GetAt(index)
            If (Not itemInventoryRet Is Nothing) Then
                If (Not itemInventoryRet.FullName Is Nothing) Then
                    ' only the first itemInventoryRet should be repeating and then
                    ' lets just check to make sure we do not have the item
                    ' just in case another app changed a item right between our
                    ' queries.
                    If (index <> 0) Or (Not FoundInComboBox(listComboBox, itemInventoryRet.FullName.GetValue())) Then
                        listComboBox.Items.Add((itemInventoryRet.FullName.GetValue()))
                    End If
                    lastFullName = itemInventoryRet.FullName.GetValue()
                Else
                    bDone = True
                End If
            Else
                bDone = True
            End If
        Next

        Exit Sub
Errs:
        MsgBox("Error in fillItemComboBox" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)
        bDone = True
    End Sub
    Sub fillCustomerList(ByRef CustList As System.Windows.Forms.ComboBox)
        On Error GoTo Errs

        ' make sure we do not have any old requests still defined
        msgSetRequest.ClearRequests()

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = QBFC15Lib.ENRqOnError.roeContinue

        ' Add the CustomerQuery request
        Dim customerQuery As QBFC15Lib.ICustomerQuery
        customerQuery = msgSetRequest.AppendCustomerQueryRq()

        ' we are going to set the number of entries returned to limit the size of the return structure
        customerQuery.ORCustomerListQuery.CustomerListFilter.MaxReturned.SetValue((MAX_RETURNED))

        Dim bDone As Boolean
        bDone = False
        Dim firstFullName As String
        firstFullName = "!"

        'Clear the list box
        CustList.Items.Clear()

        Dim msgSetResponse As QBFC15Lib.IMsgSetResponse
        Do While (Not bDone)

            ' start looking for customer next in the list
            ' we may have one overlap
            customerQuery.ORCustomerListQuery.CustomerListFilter.ORNameFilter.NameRangeFilter.FromName.SetValue((firstFullName))

            ' send the request to QB
            msgSetResponse = SessionManager.DoRequests(msgSetRequest)

            'MsgBox(msgSetRequest.ToXMLString())

            fillCustomerComboBox(msgSetResponse, CustList, bDone, firstFullName)

        Loop

        Exit Sub

Errs:
        MsgBox("Error in fillCustomerList" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)
    End Sub

    Public Sub fillCustomerComboBox(ByRef msgSetResponse As QBFC15Lib.IMsgSetResponse, ByRef listComboBox As System.Windows.Forms.ComboBox, ByRef bDone As Boolean, ByRef lastFullName As String)
        On Error GoTo Errs

        ' check to make sure we have objects to access first
        ' and that there are responses in the list
        If (msgSetResponse Is Nothing) Then
            bDone = True
            Exit Sub
        End If
        If (msgSetResponse.ResponseList Is Nothing) Then
            bDone = True
            Exit Sub
        End If
        If (msgSetResponse.ResponseList.Count <= 0) Then
            bDone = True
            Exit Sub
        End If

        ' Start parsing the response list
        Dim responseList As QBFC15Lib.IResponseList
        responseList = msgSetResponse.ResponseList

        ' go thru each response and process the response.
        ' this example will only have one response in the list
        ' so we will look at index=0
        Dim response As QBFC15Lib.IResponse
        response = responseList.GetAt(0)
        If (Not response Is Nothing) Then
            If response.StatusCode <> CDbl("0") Then
                'If the status is bad, report it to the user
                MsgBox("fillCustomerComboBox unexpected Error - " & response.StatusMessage)
                bDone = True
                Exit Sub
            End If
        End If

        ' first make sure we have a response object to handle
        If (response Is Nothing) Then
            bDone = True
            Exit Sub
        End If
        If (response.Type Is Nothing) Or (response.Detail Is Nothing) Then
            bDone = True
            Exit Sub
        End If
        If (response.Detail.Type Is Nothing) Then
            bDone = True
            Exit Sub
        End If

        ' make sure we are processing the correct ICustomerQueryRs and
        ' the correct ICustomerRetList responses in this response list
        Dim customerRetList As QBFC15Lib.ICustomerRetList
        If (QBFC15Lib.ENResponseType.rtCustomerQueryRs = response.Type.GetValue()) And (QBFC15Lib.ENObjectType.otCustomerRetList = response.Detail.Type.GetValue()) Then
            ' save the response detail in the appropriate object type
            ' since we have first verified the type of the response object
            customerRetList = response.Detail
        Else
            ' bail, we do not have the responses we were expecting
            bDone = True
            Exit Sub
        End If

        'Parse the query response and add the names to the combo box
        Dim count As Short
        Dim index As Short
        Dim customerRet As QBFC15Lib.ICustomerRet
        count = customerRetList.Count

        ' we are done with the customerQueries if we have not received the MaxReturned
        If (count < MAX_RETURNED) Then
            bDone = True
        End If

        For index = 0 To count - 1
            ' skip this customer if this is a repeat from the last query
            customerRet = customerRetList.GetAt(index)
            If (Not customerRet Is Nothing) Then
                If (Not customerRet.FullName Is Nothing) Then
                    ' only the first customerRet should be repeating and then
                    ' lets just check to make sure we do not have the customer
                    ' just in case another app changed a customer right between our
                    ' queries.
                    If (index <> 0) Or (Not FoundInComboBox(listComboBox, customerRet.FullName.GetValue())) Then
                        listComboBox.Items.Add((customerRet.FullName.GetValue()))
                    End If
                    lastFullName = customerRet.FullName.GetValue()
                Else
                    bDone = True
                End If
            Else
                bDone = True
            End If
        Next

        Exit Sub
Errs:
        MsgBox("Error in fillCustomerComboBox" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)
        bDone = True
    End Sub

    Private Function FoundInComboBox(ByRef comboBox As System.Windows.Forms.ComboBox, ByRef name As String) As Boolean
        On Error GoTo Errs

        FoundInComboBox = False

        ' go thru our combo box and find if the name exists
        Dim i As Short
        Dim numItems As Short
        numItems = comboBox.Items.Count
        ' our overlap should be from the last item added, so start from the end
        ' to do the compare
        For i = (numItems - 1) To 0 Step -1
            If (VB6.GetItemString(comboBox, i) = name) Then
                FoundInComboBox = True
                Exit For
            End If
        Next
        Exit Function
Errs:
        MsgBox("Error in FoundInComboBox" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)
        FoundInComboBox = False
    End Function
    Sub fillClassList(ByRef ClassList As System.Windows.Forms.ComboBox)
        On Error GoTo Errs

        ' make sure we do not have any old requests still defined
        msgSetRequest.ClearRequests()

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = QBFC15Lib.ENRqOnError.roeContinue

        ' Add the ClassQuery request
        Dim classQuery As QBFC15Lib.IClassQuery
        classQuery = msgSetRequest.AppendClassQueryRq()

        ' we are going to set the number of entries returned to limit the size of the return structure
        classQuery.ORListQuery.ListFilter.MaxReturned.SetValue((MAX_RETURNED))

        Dim bDone As Boolean
        bDone = False
        Dim firstFullName As String
        firstFullName = "!"

        'Clear the list box
        ClassList.Items.Clear()

        Dim msgSetResponse As QBFC15Lib.IMsgSetResponse
        Do While (Not bDone)

            ' start looking for customer next in the list
            ' we will have one overlap
            classQuery.ORListQuery.ListFilter.ORNameFilter.NameRangeFilter.FromName.SetValue((firstFullName))

            ' send the request to QB
            msgSetResponse = SessionManager.DoRequests(msgSetRequest)

            'MsgBox(msgSetRequest.ToXMLString())

            fillClassComboBox(msgSetResponse, ClassList, bDone, firstFullName)

        Loop

        Exit Sub

Errs:
        MsgBox("Error in fillClassList" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)
    End Sub
    Public Sub fillClassComboBox(ByRef msgSetResponse As QBFC15Lib.IMsgSetResponse, ByRef listComboBox As System.Windows.Forms.ComboBox, ByRef bDone As Boolean, ByRef lastFullName As String)
        On Error GoTo Errs

        ' check to make sure we have objects to access first
        ' and that there are responses in the list
        If (msgSetResponse Is Nothing) Then
            bDone = True
            Exit Sub
        End If
        If (msgSetResponse.ResponseList Is Nothing) Then
            bDone = True
            Exit Sub
        End If
        If (msgSetResponse.ResponseList.Count <= 0) Then
            bDone = True
            Exit Sub
        End If

        ' Start parsing the response list
        Dim responseList As QBFC15Lib.IResponseList
        responseList = msgSetResponse.ResponseList

        ' go thru each response and process the response.
        ' this example will only have one response in the list
        ' so we will look at index=0
        Dim response As QBFC15Lib.IResponse
        response = responseList.GetAt(0)
        If (Not response Is Nothing) Then
            If response.StatusCode <> CDbl("0") Then
                'If the status is bad, report it to the user
                MsgBox("fillClassComboBox unexpected Error - " & response.StatusMessage)
                bDone = True
                Exit Sub
            End If
        End If

        ' first make sure we have a response object to handle
        If (response Is Nothing) Then
            bDone = True
            Exit Sub
        End If
        If ((response.Type Is Nothing) Or (response.Detail Is Nothing)) Then
            bDone = True
            Exit Sub
        End If
        If (response.Detail.Type Is Nothing) Then
            bDone = True
            Exit Sub
        End If

        ' make sure we are processing the correct IClassQueryRs and
        ' the correct IClassRetList responses in this response list
        Dim classRetList As QBFC15Lib.IClassRetList
        If (QBFC15Lib.ENResponseType.rtClassQueryRs = response.Type.GetValue()) And (QBFC15Lib.ENObjectType.otClassRetList = response.Detail.Type.GetValue()) Then
            ' save the response detail in the appropriate object type
            ' since we have first verified the type of the response object
            classRetList = response.Detail
        Else
            ' bail, we do not have the responses we were expecting
            bDone = True
            Exit Sub
        End If

        'Parse the query response and add the names to the combo box
        Dim count As Short
        Dim index As Short
        Dim classRet As QBFC15Lib.IClassRet
        count = classRetList.Count

        ' we are done with the classQueries if we have not received the MaxReturned
        If (count < MAX_RETURNED) Then
            bDone = True
        End If

        For index = 0 To count - 1
            ' skip this class if this is a repeat from the last query
            classRet = classRetList.GetAt(index)
            If (Not classRet Is Nothing) Then
                If (Not classRet.Name Is Nothing) Then
                    ' only the first classRet should be repeating and then
                    ' lets just check to make sure we do not have the class
                    ' just in case another app changed a class right between our
                    ' queries.
                    If (index <> 0) Or (Not FoundInComboBox(listComboBox, classRet.Name.GetValue())) Then
                        listComboBox.Items.Add((classRet.Name.GetValue()))
                    End If
                    lastFullName = classRet.Name.GetValue()
                Else
                    bDone = True
                End If
            Else
                bDone = True
            End If
        Next

        Exit Sub
Errs:
        MsgBox("Error in fillClassComboBox" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)
        bDone = True
    End Sub

    'UPGRADE_NOTE: class was upgraded to class_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Function AdjustInventory(ByRef acct As String, ByRef cust As String, ByRef class_Renamed As String, ByRef Memo As String, ByRef LineItems As System.Windows.Forms.ListView) As String
        Dim i As Object
        On Error GoTo Errs

        Dim msgSetResponse As QBFC15Lib.IMsgSetResponse

        msgSetRequest.ClearRequests()

        ' Add the InventoryAdjustQueryRq request
        Dim inventoryAdjustAdd As QBFC15Lib.IInventoryAdjustmentAdd
        inventoryAdjustAdd = msgSetRequest.AppendInventoryAdjustmentAddRq()

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = QBFC15Lib.ENRqOnError.roeContinue

        ' set the AccountRef field
        If Not (acct = "") Then
            inventoryAdjustAdd.AccountRef.FullName.SetValue((acct))
        End If

        ' set the CustomerRef field
        If Not (cust = "") Then
            inventoryAdjustAdd.CustomerRef.FullName.SetValue((cust))
        End If

        ' set the ClassRef field
        If Not (class_Renamed = "") Then
            inventoryAdjustAdd.ClassRef.FullName.SetValue((class_Renamed))
        End If

        ' set the memo field
        If Not (Memo = "") Then
            inventoryAdjustAdd.Memo.SetValue((Memo))
        End If

        ' now add the line items
        Dim item As String
        Dim what As String
        Dim diff As String
        Dim quant As String
        Dim value As String
        Dim inventoryAdjustLineAdd As QBFC15Lib.IInventoryAdjustmentLineAdd
        'For i = 1 To LineItems.Items.Count
        For i = 0 To LineItems.Items.Count - 1
            inventoryAdjustLineAdd = inventoryAdjustAdd.InventoryAdjustmentLineAddList.Append
            'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Lower bound of collection LineItems.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
            item = LineItems.Items.Item(i).Text
            inventoryAdjustLineAdd.ItemRef.FullName.SetValue((item))
            'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Lower bound of collection LineItems.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
            'UPGRADE_WARNING: Lower bound of collection LineItems.ListItems.item() has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
            what = LineItems.Items.Item(i).SubItems(1).Text
            If "Quantity" = what Then
                'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Lower bound of collection LineItems.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                'UPGRADE_WARNING: Lower bound of collection LineItems.ListItems.item() has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                diff = LineItems.Items.Item(i).SubItems(2).Text
                'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Lower bound of collection LineItems.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                'UPGRADE_WARNING: Lower bound of collection LineItems.ListItems.item() has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                quant = LineItems.Items.Item(i).SubItems(3).Text
                If "Relative" = diff Then
                    inventoryAdjustLineAdd.ORTypeAdjustment.QuantityAdjustment.ORQuantityAdjustment.QuantityDifference.SetValue((CShort(quant)))
                Else
                    inventoryAdjustLineAdd.ORTypeAdjustment.QuantityAdjustment.ORQuantityAdjustment.NewQuantity.SetValue((CShort(quant)))
                End If
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Lower bound of collection LineItems.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                'UPGRADE_WARNING: Lower bound of collection LineItems.ListItems.item() has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                value = LineItems.Items.Item(i).SubItems(2).Text
                'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Lower bound of collection LineItems.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                'UPGRADE_WARNING: Lower bound of collection LineItems.ListItems.item() has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                quant = LineItems.Items.Item(i).SubItems(3).Text
                If Not "0" = quant Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object inventoryAdjustLineAdd.ORTypeAdjustment.ValueAdjustment.NewQuantity. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    inventoryAdjustLineAdd.ORTypeAdjustment.ValueAdjustment.NewQuantity.setValue(CShort(quant))
                End If
                'UPGRADE_WARNING: Couldn't resolve default property of object inventoryAdjustLineAdd.ORTypeAdjustment.ValueAdjustment.NewValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                inventoryAdjustLineAdd.ORTypeAdjustment.ValueAdjustment.NewValue.setValue(CDbl(value))
            End If
        Next i

        'MsgBox msgSetRequest.ToXMLString()

        ' send the request to QuickBooks
        msgSetResponse = SessionManager.DoRequests(msgSetRequest)

        'MsgBox msgSetResponse.ToXMLString()

        ' check to make sure we have objects to access first
        ' and that there are responses in the list
        If (msgSetResponse Is Nothing) Then
            Exit Function
        End If
        If (msgSetResponse.ResponseList Is Nothing) Then
            Exit Function
        End If
        If (msgSetResponse.ResponseList.Count <= 0) Then
            Exit Function
        End If

        ' Start parsing the response list
        Dim responseList As QBFC15Lib.IResponseList
        responseList = msgSetResponse.ResponseList

        ' go thru each response and process the response.
        Dim response As QBFC15Lib.IResponse
        ' we will only have one response from the one InventoryAdjusmentAdd
        response = responseList.GetAt(0)
        If (response Is Nothing) Then
            Exit Function
        End If

        AdjustInventory = response.StatusMessage & vbCrLf & msgSetResponse.ToXMLString()

        Exit Function
Errs:
        AdjustInventory = ("Error in AdjustInventory" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)
    End Function
    Function QBFCLatestVersion(ByRef SessionManager As QBFC15Lib.QBSessionManager) As String
        'Dim strXMLVersions() As String
        'Should be able to use this, but there appears to be a bug that may cause 2.0 to be returned
        'when it should not.
        'strXMLVersions = SessionManager.QBXMLVersionsForSession

        Dim msgset As QBFC15Lib.IMsgSetRequest
        'Use oldest version to ensure that we work with any QuickBooks (US)
        msgset = SessionManager.CreateMsgSetRequest("US", 1, 0)
        msgset.AppendHostQueryRq()
        Dim QueryResponse As QBFC15Lib.IMsgSetResponse
        QueryResponse = SessionManager.DoRequests(msgset)
        Dim response As QBFC15Lib.IResponse

        ' The response list contains only one response,
        ' which corresponds to our single HostQuery request
        response = QueryResponse.ResponseList.GetAt(0)
        Dim HostResponse As QBFC15Lib.IHostRet
        HostResponse = response.Detail
        Dim supportedVersions As QBFC15Lib.IBSTRList
        supportedVersions = HostResponse.SupportedQBXMLVersionList
		
		Dim i As Integer
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
End Module