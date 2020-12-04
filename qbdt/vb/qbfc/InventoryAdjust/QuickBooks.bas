Attribute VB_Name = "QuickBooks"
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
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
' Updated to QBFC 12.0 and fixed setting max QBXML version: 09/2012
'----------------------------------------------------------
Const MAX_RETURNED = 20
Dim booSessionBegun As Boolean
Dim booConnectionOpen As Boolean
Dim SessionManager As QBSessionManager
Dim msgSetRequest As IMsgSetRequest

Function qbConnect() As Boolean
    On Error GoTo errHandler
    
    If (booSessionBegun) Then
        qbConnect = True
        Exit Function
    End If
        
    ' create the new QBSessionManager object
    If (SessionManager Is Nothing) Then
        Set SessionManager = New QBSessionManager
    End If
    
    ' open the connection to QuickBooks
    If (Not booConnectionOpen) Then
        SessionManager.OpenConnection "", "IDN Inventory Adjust Sample - QBFC"
        booConnectionOpen = True
    End If
    
    ' InventoryAdjust must be done in SingleUser mode
    ' Begin a session with QuickBooks
    SessionManager.BeginSession "", ENOpenMode.omSingleUser
    booSessionBegun = True

    'Check to make sure the QuickBooks we're working with supports version 2 of the SDK
    Dim booSupports2dot0 As Boolean
    booSupports2dot0 = False
    Dim supportedVersion As String
    supportedVersion = Val(QBFCLatestVersion(SessionManager))
    If (supportedVersion >= 6#) Then
        booSupports2dot0 = True
        Set msgSetRequest = SessionManager.CreateMsgSetRequest("US", 6, 0)
    ElseIf (supportedVersion >= 5#) Then
        booSupports2dot0 = True
        Set msgSetRequest = SessionManager.CreateMsgSetRequest("US", 5, 0)
    ElseIf (supportedVersion >= 4#) Then
        booSupports2dot0 = True
        Set msgSetRequest = SessionManager.CreateMsgSetRequest("US", 4, 0)
    ElseIf (supportedVersion >= 3#) Then
        booSupports2dot0 = True
        Set msgSetRequest = SessionManager.CreateMsgSetRequest("US", 3, 0)
    ElseIf (supportedVersion >= 2#) Then
        booSupports2dot0 = True
        Set msgSetRequest = SessionManager.CreateMsgSetRequest("US", 2, 0)
    End If
    
    If Not booSupports2dot0 Then
        MsgBox "This program only runs against QuickBooks installations which support the 2.0 qbXML spec.  Your version of QuickBooks does not support qbXML 2.0"
        SessionManager.EndSession
        SessionManager.CloseConnection
        End
    End If
    
    qbConnect = True
    Exit Function
    
errHandler:
    MsgBox "Error in qbConnect" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description

    If booSessionBegun Then
        SessionManager.EndSession
    End If
    If booConnectionOpen Then
        SessionManager.CloseConnection
    End If
    
    qbConnect = False
    End
End Function
Public Sub EndSessionCloseConnection()
    On Error GoTo Errs
    If booSessionBegun Then
        SessionManager.EndSession
        SessionManager.CloseConnection
    End If
    Exit Sub
Errs:
    MsgBox "Error in EndSessionCloseConnection" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description

End Sub
Function ClassesEnabled() As Boolean
    On Error GoTo Errs
    
    Dim msgSetResponse As IMsgSetResponse
    
    ClassesEnabled = False
    msgSetRequest.ClearRequests
     
    ' Add the PreferencesQueryRq request
    msgSetRequest.AppendPreferencesQueryRq
    
    ' set the OnError attribute to continueOnError
    msgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

    'MsgBox msgSetRequest.ToXMLString()
    
    ' send the request to QuickBooks
    Set msgSetResponse = SessionManager.DoRequests(msgSetRequest)
    
    'MsgBox msgSetResponse.ToXMLString()
    
    ' check to make sure we have objects to access first
    ' and that there are responses in the list
    If (msgSetResponse Is Nothing) Then
        ClassesEnabled = False
        Exit Function
    End If
    If (msgSetResponse.responseList Is Nothing) Then
        ClassesEnabled = False
        Exit Function
    End If
    If (msgSetResponse.responseList.count <= 0) Then
        ClassesEnabled = False
        Exit Function
    End If

    ' Start parsing the response list
    Dim responseList As IResponseList
    Set responseList = msgSetResponse.responseList

    ' go thru each response and process the response.
    Dim response As IResponse
    Set response = responseList.GetAt(0)
    If (response Is Nothing) Then
        ClassesEnabled = False
        Exit Function
    End If
    If response.StatusCode <> "0" Then
        'If the status is bad, report it to the user
        MsgBox ("ClassesEnabled unexpected Error - " & response.StatusMessage)
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
    Dim preferencesRet As IPreferencesRet
    Dim responseType As ENResponseType
    Dim responseDetailType As ENObjectType
    responseType = response.Type.getValue()
    responseDetailType = response.Detail.Type.getValue()
    If (responseType = ENResponseType.rtPreferencesQueryRs) And _
        (responseDetailType = ENObjectType.otPreferencesRet) Then
        ' save the response detail in the appropriate object type
        ' since we have first verified the type of the response object
        Set preferencesRet = response.Detail
    Else
        ' bail, we do not have the responses we were expecting
        ClassesEnabled = False
        Exit Function
    End If

    'get the query response return the preferencesRet variable
    If (Not preferencesRet.AccountingPreferences Is Nothing) Then
        If (Not preferencesRet.AccountingPreferences.IsUsingClassTracking Is Nothing) Then
            ClassesEnabled = preferencesRet.AccountingPreferences.IsUsingClassTracking.getValue()
            Exit Function
        End If
    End If
        
    ClassesEnabled = False
    Exit Function
Errs:
    MsgBox ("Error in ClassesEnabled" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)
    
End Function
Sub fillAccountList(AccountList As comboBox)
    On Error GoTo Errs
    
    Dim msgSetResponse As IMsgSetResponse
    
    msgSetRequest.ClearRequests
     
    ' Add the AccountQueryRq request
    Dim accountQuery As IAccountQuery
    Set accountQuery = msgSetRequest.AppendAccountQueryRq()
    
    ' set the OnError attribute to continueOnError
    msgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

    'MsgBox msgSetRequest.ToXMLString()
    
    ' send the request to QuickBooks
    Set msgSetResponse = SessionManager.DoRequests(msgSetRequest)
    
    'MsgBox msgSetResponse.ToXMLString()
    
    ' check to make sure we have objects to access first
    ' and that there are responses in the list
    If (msgSetResponse Is Nothing) Then
        Exit Sub
    End If
    If (msgSetResponse.responseList Is Nothing) Then
        Exit Sub
    End If
    If (msgSetResponse.responseList.count <= 0) Then
        Exit Sub
    End If

    ' Start parsing the response list
    Dim responseList As IResponseList
    Set responseList = msgSetResponse.responseList

    ' go thru each response and process the response.
    Dim response As IResponse
    ' we will only have one response from the one AccountQuery
    Set response = responseList.GetAt(0)
    If (response Is Nothing) Then
        Exit Sub
    End If
    If response.StatusCode <> "0" Then
        'If the status is bad, report it to the user
        MsgBox ("fillAccountList unexpected Error - " & response.StatusMessage)
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
    Dim accountRetList As IAccountRetList
    Dim responseType As ENResponseType
    Dim responseDetailType As ENObjectType
    responseType = response.Type.getValue()
    responseDetailType = response.Detail.Type.getValue()
    If (responseType = ENResponseType.rtAccountQueryRs) And _
        (responseDetailType = ENObjectType.otAccountRetList) Then
        ' save the response detail in the appropriate object type
        ' since we have first verified the type of the response object
        Set accountRetList = response.Detail
    Else
        ' bail, we do not have the responses we were expecting
        Exit Sub
    End If

    'Parse the query response and add the Estimate to the Estimate list box
    Dim count As Integer
    Dim index As Integer
    Dim accountRet As IAccountRet
    count = accountRetList.count

    For index = 0 To count - 1
        Set accountRet = accountRetList.GetAt(index)
        If (accountRet Is Nothing) Then
            Exit Sub
        End If
        If (accountRet.FullName Is Nothing) Then
            Exit Sub
        End If
        AccountList.AddItem (accountRet.FullName.getValue())
    Next
        
    Exit Sub
Errs:
    MsgBox ("Error in fillAccountList" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)
End Sub
Sub fillItemList(ItemList As comboBox)
    On Error GoTo Errs

    ' make sure we do not have any old requests still defined
    msgSetRequest.ClearRequests

    ' set the OnError attribute to continueOnError
    msgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

    ' Add the ItemInventoryQuery request
    Dim itemInventoryQuery As IItemInventoryQuery
    Set itemInventoryQuery = msgSetRequest.AppendItemInventoryQueryRq()

    ' we are going to set the number of entries returned to limit the size of the return structure
    itemInventoryQuery.ORListQueryWithOwnerIDAndClass.ListWithClassFilter.MaxReturned.setValue (MAX_RETURNED)

    Dim bDone As Boolean
    bDone = False
    Dim firstFullName As String
    firstFullName = "!"

    'Clear the list box
    ItemList.Clear

    Do While (Not bDone)

        ' start looking for itemInventory next in the list
        ' we will have one overlap
        itemInventoryQuery.ORListQueryWithOwnerIDAndClass.ListWithClassFilter.ORNameFilter.NameRangeFilter.FromName.setValue (firstFullName)

        ' send the request to QB
        Dim msgSetResponse As IMsgSetResponse
        Set msgSetResponse = SessionManager.DoRequests(msgSetRequest)

        'MsgBox(msgSetRequest.ToXMLString())

        ' fill the ItemList box with the list of Item Inventory FullNames
        ' from the response returned from QuickBooks
        fillItemComboBox msgSetResponse, ItemList, bDone, firstFullName

    Loop

    Exit Sub

Errs:
    MsgBox ("Error in fillItemList" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)
End Sub
Public Sub fillItemComboBox(ByRef msgSetResponse As IMsgSetResponse, _
                ByRef listComboBox As comboBox, _
                ByRef bDone As Boolean, _
                ByRef lastFullName As String)
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
            MsgBox ("fillItemComboBox unexpected Error - " & response.StatusMessage)
            bDone = True
            Exit Sub
        End If
    End If

    ' first make sure we have a response object to handle
    If (response Is Nothing) Then
        bDone = True
        Exit Sub
    End If
    If (response.Type Is Nothing) Or _
        (response.Detail Is Nothing) Then
        bDone = True
        Exit Sub
    End If
    If (response.Detail.Type Is Nothing) Then
        bDone = True
        Exit Sub
    End If

    ' make sure we are processing the correct IItemInventoryQueryRs and
    ' the correct IItemInventoryRetList responses in this response list
    Dim itemInventoryRetList As IItemInventoryRetList
    If (rtItemInventoryQueryRs = response.Type.getValue()) And _
        (otItemInventoryRetList = response.Detail.Type.getValue()) Then
        ' save the response detail in the appropriate object type
        ' since we have first verified the type of the response object
        Set itemInventoryRetList = response.Detail
    Else
        ' bail, we do not have the responses we were expecting
        bDone = True
        Exit Sub
    End If

    'Parse the query response and add the names to the combo box
    Dim count As Integer
    Dim index As Integer
    Dim itemInventoryRet As IItemInventoryRet
    count = itemInventoryRetList.count

    ' we are done with the Queries if we have not received the MaxReturned
    If (count < MAX_RETURNED) Then
        bDone = True
    End If

    For index = 0 To count - 1
        ' skip this item if this is a repeat from the last query
        Set itemInventoryRet = itemInventoryRetList.GetAt(index)
        If (Not itemInventoryRet Is Nothing) Then
            If (Not itemInventoryRet.FullName Is Nothing) Then
                ' only the first itemInventoryRet should be repeating and then
                ' lets just check to make sure we do not have the item
                ' just in case another app changed a item right between our
                ' queries.
                If (index <> 0) Or (Not FoundInComboBox(listComboBox, itemInventoryRet.FullName.getValue())) Then
                    listComboBox.AddItem (itemInventoryRet.FullName.getValue())
                End If
                lastFullName = itemInventoryRet.FullName.getValue()
            Else
                bDone = True
            End If
        Else
            bDone = True
        End If
    Next

    Exit Sub
Errs:
    MsgBox ("Error in fillItemComboBox" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)
    bDone = True
End Sub
Sub fillCustomerList(CustList As comboBox)
    On Error GoTo Errs

    ' make sure we do not have any old requests still defined
    msgSetRequest.ClearRequests

    ' set the OnError attribute to continueOnError
    msgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

    ' Add the CustomerQuery request
    Dim customerQuery As ICustomerQuery
    Set customerQuery = msgSetRequest.AppendCustomerQueryRq()

    ' we are going to set the number of entries returned to limit the size of the return structure
    customerQuery.ORCustomerListQuery.CustomerListFilter.MaxReturned.setValue (MAX_RETURNED)

    Dim bDone As Boolean
    bDone = False
    Dim firstFullName As String
    firstFullName = "!"

    'Clear the list box
    CustList.Clear

    Do While (Not bDone)

        ' start looking for customer next in the list
        ' we may have one overlap
        customerQuery.ORCustomerListQuery.CustomerListFilter.ORNameFilter.NameRangeFilter.FromName.setValue (firstFullName)

        ' send the request to QB
        Dim msgSetResponse As IMsgSetResponse
        Set msgSetResponse = SessionManager.DoRequests(msgSetRequest)

        'MsgBox(msgSetRequest.ToXMLString())

        fillCustomerComboBox msgSetResponse, CustList, bDone, firstFullName

    Loop

    Exit Sub

Errs:
    MsgBox ("Error in fillCustomerList" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)
End Sub

Public Sub fillCustomerComboBox(ByRef msgSetResponse As IMsgSetResponse, _
                ByRef listComboBox As comboBox, _
                ByRef bDone As Boolean, _
                ByRef lastFullName As String)
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
            MsgBox ("fillCustomerComboBox unexpected Error - " & response.StatusMessage)
            bDone = True
            Exit Sub
        End If
    End If

    ' first make sure we have a response object to handle
    If (response Is Nothing) Then
        bDone = True
        Exit Sub
    End If
    If (response.Type Is Nothing) Or _
        (response.Detail Is Nothing) Then
        bDone = True
        Exit Sub
    End If
    If (response.Detail.Type Is Nothing) Then
        bDone = True
        Exit Sub
    End If

    ' make sure we are processing the correct ICustomerQueryRs and
    ' the correct ICustomerRetList responses in this response list
    Dim customerRetList As ICustomerRetList
    If (rtCustomerQueryRs = response.Type.getValue()) And _
        (otCustomerRetList = response.Detail.Type.getValue()) Then
        ' save the response detail in the appropriate object type
        ' since we have first verified the type of the response object
        Set customerRetList = response.Detail
    Else
        ' bail, we do not have the responses we were expecting
        bDone = True
        Exit Sub
    End If

    'Parse the query response and add the names to the combo box
    Dim count As Integer
    Dim index As Integer
    Dim customerRet As ICustomerRet
    count = customerRetList.count

    ' we are done with the customerQueries if we have not received the MaxReturned
    If (count < MAX_RETURNED) Then
        bDone = True
    End If

    For index = 0 To count - 1
        ' skip this customer if this is a repeat from the last query
        Set customerRet = customerRetList.GetAt(index)
        If (Not customerRet Is Nothing) Then
            If (Not customerRet.FullName Is Nothing) Then
                ' only the first customerRet should be repeating and then
                ' lets just check to make sure we do not have the customer
                ' just in case another app changed a customer right between our
                ' queries.
                If (index <> 0) Or (Not FoundInComboBox(listComboBox, customerRet.FullName.getValue())) Then
                    listComboBox.AddItem (customerRet.FullName.getValue())
                End If
                lastFullName = customerRet.FullName.getValue()
            Else
                bDone = True
            End If
        Else
            bDone = True
        End If
    Next

    Exit Sub
Errs:
    MsgBox ("Error in fillCustomerComboBox" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)
    bDone = True
End Sub

Private Function FoundInComboBox(ByRef comboBox As comboBox, _
            ByRef name As String) As Boolean
    On Error GoTo Errs

    FoundInComboBox = False

    ' go thru our combo box and find if the name exists
    Dim i As Integer
    Dim numItems As Integer
    numItems = comboBox.ListCount
    ' our overlap should be from the last item added, so start from the end
    ' to do the compare
    For i = (numItems - 1) To 0 Step -1
        If (comboBox.List(i) = name) Then
            FoundInComboBox = True
            Exit For
        End If
    Next
    Exit Function
Errs:
    MsgBox ("Error in FoundInComboBox" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)
    FoundInComboBox = False
    End Function
Sub fillClassList(ClassList As comboBox)
    On Error GoTo Errs

    ' make sure we do not have any old requests still defined
    msgSetRequest.ClearRequests

    ' set the OnError attribute to continueOnError
    msgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

    ' Add the ClassQuery request
    Dim classQuery As IClassQuery
    Set classQuery = msgSetRequest.AppendClassQueryRq()

    ' we are going to set the number of entries returned to limit the size of the return structure
    classQuery.ORListQuery.ListFilter.MaxReturned.setValue (MAX_RETURNED)

    Dim bDone As Boolean
    bDone = False
    Dim firstFullName As String
    firstFullName = "!"

    'Clear the list box
    ClassList.Clear

    Do While (Not bDone)

        ' start looking for customer next in the list
        ' we will have one overlap
        classQuery.ORListQuery.ListFilter.ORNameFilter.NameRangeFilter.FromName.setValue (firstFullName)

        ' send the request to QB
        Dim msgSetResponse As IMsgSetResponse
        Set msgSetResponse = SessionManager.DoRequests(msgSetRequest)

        'MsgBox(msgSetRequest.ToXMLString())

        fillClassComboBox msgSetResponse, ClassList, bDone, firstFullName

    Loop

    Exit Sub

Errs:
    MsgBox ("Error in fillClassList" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)
End Sub
Public Sub fillClassComboBox(ByRef msgSetResponse As IMsgSetResponse, _
                ByRef listComboBox As comboBox, _
                ByRef bDone As Boolean, _
                ByRef lastFullName As String)
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
            MsgBox ("fillClassComboBox unexpected Error - " & response.StatusMessage)
            bDone = True
            Exit Sub
        End If
    End If

    ' first make sure we have a response object to handle
    If (response Is Nothing) Then
        bDone = True
        Exit Sub
    End If
    If ((response.Type Is Nothing) Or _
        (response.Detail Is Nothing)) Then
        bDone = True
        Exit Sub
    End If
    If (response.Detail.Type Is Nothing) Then
        bDone = True
        Exit Sub
    End If

    ' make sure we are processing the correct IClassQueryRs and
    ' the correct IClassRetList responses in this response list
    Dim classRetList As IClassRetList
    If (rtClassQueryRs = response.Type.getValue()) And _
        (otClassRetList = response.Detail.Type.getValue()) Then
        ' save the response detail in the appropriate object type
        ' since we have first verified the type of the response object
        Set classRetList = response.Detail
    Else
        ' bail, we do not have the responses we were expecting
        bDone = True
        Exit Sub
    End If

    'Parse the query response and add the names to the combo box
    Dim count As Integer
    Dim index As Integer
    Dim classRet As IClassRet
    count = classRetList.count

    ' we are done with the classQueries if we have not received the MaxReturned
    If (count < MAX_RETURNED) Then
        bDone = True
    End If

    For index = 0 To count - 1
        ' skip this class if this is a repeat from the last query
        Set classRet = classRetList.GetAt(index)
        If (Not classRet Is Nothing) Then
            If (Not classRet.name Is Nothing) Then
                ' only the first classRet should be repeating and then
                ' lets just check to make sure we do not have the class
                ' just in case another app changed a class right between our
                ' queries.
                If (index <> 0) Or (Not FoundInComboBox(listComboBox, classRet.name.getValue())) Then
                    listComboBox.AddItem (classRet.name.getValue())
                End If
                lastFullName = classRet.name.getValue()
            Else
                bDone = True
            End If
        Else
            bDone = True
        End If
    Next

    Exit Sub
Errs:
    MsgBox ("Error in fillClassComboBox" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)
    bDone = True
End Sub

Function AdjustInventory(acct As String, cust As String, class As String, Memo As String, LineItems As ListView) As String
    On Error GoTo Errs
    
    Dim msgSetResponse As IMsgSetResponse
    
    msgSetRequest.ClearRequests
     
    ' Add the InventoryAdjustQueryRq request
    Dim inventoryAdjustAdd As IInventoryAdjustmentAdd
    Set inventoryAdjustAdd = msgSetRequest.AppendInventoryAdjustmentAddRq()
    
    ' set the OnError attribute to continueOnError
    msgSetRequest.Attributes.OnError = ENRqOnError.roeContinue
    
    ' set the AccountRef field
    If Not (acct = "") Then
        inventoryAdjustAdd.AccountRef.FullName.setValue (acct)
    End If
    
    ' set the CustomerRef field
    If Not (cust = "") Then
        inventoryAdjustAdd.CustomerRef.FullName.setValue (cust)
    End If
    
    ' set the ClassRef field
    If Not (class = "") Then
        inventoryAdjustAdd.ClassRef.FullName.setValue (class)
    End If
    
    ' set the memo field
    If Not (Memo = "") Then
        inventoryAdjustAdd.Memo.setValue (Memo)
    End If
    
    ' now add the line items
    For i = 1 To LineItems.ListItems.count
        Dim item As String
        Dim what As String
        Dim diff As String
        Dim quant As String
        Dim value As String
        Dim inventoryAdjustLineAdd As IInventoryAdjustmentLineAdd
        Set inventoryAdjustLineAdd = inventoryAdjustAdd.InventoryAdjustmentLineAddList.Append
        item = LineItems.ListItems.item(i).Text
        inventoryAdjustLineAdd.ItemRef.FullName.setValue (item)
        what = LineItems.ListItems.item(i).SubItems(1)
        If "Quantity" = what Then
            diff = LineItems.ListItems.item(i).SubItems(2)
            quant = LineItems.ListItems.item(i).SubItems(3)
            If "Relative" = diff Then
                inventoryAdjustLineAdd.ORTypeAdjustment.QuantityAdjustment.ORQuantityAdjustment.QuantityDifference.setValue (CInt(quant))
            Else
                inventoryAdjustLineAdd.ORTypeAdjustment.QuantityAdjustment.ORQuantityAdjustment.NewQuantity.setValue (CInt(quant))
            End If
        Else
            value = LineItems.ListItems.item(i).SubItems(2)
            quant = LineItems.ListItems.item(i).SubItems(3)
            If Not "0" = quant Then
                inventoryAdjustLineAdd.ORTypeAdjustment.ValueAdjustment.NewQuantity.setValue (CInt(quant))
            End If
            inventoryAdjustLineAdd.ORTypeAdjustment.ValueAdjustment.NewValue.setValue (CDbl(value))
        End If
    Next i
    
    'MsgBox msgSetRequest.ToXMLString()
    
    ' send the request to QuickBooks
    Set msgSetResponse = SessionManager.DoRequests(msgSetRequest)
    
    'MsgBox msgSetResponse.ToXMLString()
    
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

    ' go thru each response and process the response.
    Dim response As IResponse
    ' we will only have one response from the one InventoryAdjusmentAdd
    Set response = responseList.GetAt(0)
    If (response Is Nothing) Then
        Exit Function
    End If
    
    AdjustInventory = response.StatusMessage & vbCrLf & msgSetResponse.ToXMLString()
        
    Exit Function
Errs:
    AdjustInventory = ("Error in AdjustInventory" & vbCrLf & "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)
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
