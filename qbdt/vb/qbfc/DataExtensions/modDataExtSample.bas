Attribute VB_Name = "modDataExtSample"
    '----------------------------------------------------------
    ' Module: modDataExtSample
    '
    ' Description: This module contains the code which creates QBFC
    '              messages, exchanges them with QuickBooks, interprets
    '              the responses and loads information into form objects.
    '
    '              The routines in this module are hardcoded to use the
    '              OwnerID {E09C86CF-9D6E-4EF2-BCBE-4D66B6B0F754}.  The
    '              string constant cstrOwnerID is defined with this value.
    '
    ' Routines: OpenConnectionBeginSession
    '             Opens a connection and begins a sesson with the
    '             currently open company file.  If a company isn't open,
    '             the routine will display a message and then exit the
    '             program.
    '
    '           GetDataExts
    '             Builds a DataExtQueryRq request, sends it to QuickBooks
    '             and passes the reply to FillDataExts
    '
    '           GetCustomFields
    '             Builds a DataExtQueryRq request, sends it to QuickBooks
    '             and passes the reply to FillDataExts
    '
    '           FillDataExts
    '             Parses a DataExtDefQueryRs response message and fills
    '             a passed text box with the data extension definitions
    '             returned by the query
    '
    '           AddDataExtDef
    '             Builds a DataExtDefAddRq message, sends it to
    '             QuickBooks, parsed the response and if successful calls
    '             GetDataExts to add the new data extension definition to
    '             the main form, frmDataExtSample
    '
    '           GetCustomers
    '             Builds a CustomerQueryRq message, sends it to
    '             QuickBooks and calls FillCustomerListBox
    '
    '           FillCustomerListBox
    '             Parses the passed CustomerQueryRs response and adds
    '             sublevel 0 (top level customers, not jobs) to a
    '             the listbox passed to the routine.
    '
    '           CustomersHaveDataExts
    '             Calls GetDataExts to fill a text box with the defined
    '             data extensions, searches to see if any of them contain
    '             the word "Customer" (indicating they are assigned to
    '             customer records) and returns True or False depending
    '             on the result of the search.
    '
    '           FillUnusedDataExts
    '             Calls GetDataExts to get a list of all data extensions
    '             available and makes a list of those assigned for
    '             customers records, calls GetUsedDataExts to get the
    '             list of data extensions already in use for the passed
    '             in customer, compares the lists and puts the unused
    '             data extension names and types into the passed in
    '             listbox
    '
    '           GetUsedDataExts
    '             Builds a CustomerQueryRq for the passed in customer with
    '             OwnerID set to the constant for this module causing the
    '             return of data extensions for that customer record to
    '             be returned.  It then passes the response to
    '             FillUsedDataExts
    '
    '           FillUsedDataExts
    '             Parses a CustomerQueryRq response for a single customer,
    '             extracts the returned data extensions (if any) and
    '             adds them to a passed in list box with or without their
    '             values depending on a passed in boolean
    '
    '           AddDataExt
    '             Using values passed in adds a data extension to a
    '             customer record, building the request, parsing the
    '             response and displaying a message box with the result
    '             and returning a boolean indicating the add success
    '
    '           ModDataExt
    '             Using values passed in modifies a data extension to a
    '             customer record, building the request, parsing the
    '             response and displaying a message box with the result
    '             and returning a boolean indicating the modify success
    '
    '           EndSessionCloseConnection
    '             Calls EndSession and CloseConnection if the boolean
    '             booSessionBegun is true.
    '
    ' Copyright © 2002-2013 Intuit Inc. All rights reserved.
    ' Use is subject to the terms specified at:
    '      http://developer.intuit.com/legal/devsite_tos.html
    ' Updated to QBFC 12.0 and fixed setting max QBXML version: 09/2012
    ' 
    '----------------------------------------------------------
    
    Const cstrOwnerID As String = "{E09C86CF-9D6E-4EF2-BCBE-4D66B6B0F754}"
    
    Dim booSessionBegun As Boolean
    
    'Module objects
    Dim QBSessionManager As QBSessionManager
    Dim msgSetRequest As IMsgSetRequest

    Public Sub OpenConnectionBeginSession()

        booSessionBegun = False

        On Error GoTo Errs

        Set QBSessionManager = New QBSessionManager

        QBSessionManager.OpenConnection "", "IDN Data Extension Sample - QBFC"

        QBSessionManager.BeginSession "", ENOpenMode.omDontCare
        booSessionBegun = True

        'Check to make sure the QuickBooks we're working with supports version 2 of the SDK
        Dim supportedVersion As Double
		supportedVersion = Val(QBFCLatestVersion(QBSessionManager))
        
        Dim booSupports2dot0 As Boolean
        booSupports2dot0 = False
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


    Public Sub GetDataExts(ByRef txtDataExts As TextBox)
        On Error GoTo Errs

        '    "<?xml version=""1.0"" ?>" & _
        '    "<?qbxml version=""2.0""?>" & _
        '    "<QBXML>" & _
        '    "<QBXMLMsgsRq onError = ""continueOnError"">" & _
        '    "<DataExtDefQueryRq requestID = ""0"">" & _
        '    "<OwnerID>" & cstrOwnerID & "</OwnerID>" & _
        '    "</DataExtDefQueryRq>" & _
        '    "</QBXMLMsgsRq>" & _
        '    "</QBXML>"

        If (msgSetRequest Is Nothing) Then
            Exit Sub
        End If

        ' make sure we do not have any old requests still defined
        msgSetRequest.ClearRequests

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

        ' Add the DataExtDefQuery request
        Dim dataExtDefQuery As IDataExtDefQuery
        Set dataExtDefQuery = msgSetRequest.AppendDataExtDefQueryRq()

        ' set the OwnerID for the Data Ext we want returned
        dataExtDefQuery.ORDataExtDefQuery.OwnerIDList.Add (cstrOwnerID)

        ' send the request to QB
        Dim msgSetResponse As IMsgSetResponse
        Set msgSetResponse = QBSessionManager.DoRequests(msgSetRequest)

        ' fill the response info into the form
        FillDataExts msgSetResponse, txtDataExts
        Exit Sub
Errs:
        MsgBox "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    vbOKOnly, _
                    "Error in GetDataExts"
    End Sub


    Public Sub GetCustomFields(ByRef txtCustomFields As TextBox)
        On Error GoTo Errs

        '    "<?xml version=""1.0"" ?>" & _
        '    "<?qbxml version=""2.0""?>" & _
        '    "<QBXML>" & _
        '    "<QBXMLMsgsRq onError = ""continueOnError"">" & _
        '    "<DataExtDefQueryRq requestID = ""0"">" & _
        '    "<OwnerID>0</OwnerID>" & _
        '    "</DataExtDefQueryRq>" & _
        '    "</QBXMLMsgsRq>" & _
        '    "</QBXML>"

        If (msgSetRequest Is Nothing) Then
            Exit Sub
        End If

        ' make sure we do not have any old requests still defined
        msgSetRequest.ClearRequests

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

        ' Add the DataExtDefQuery request
        Dim dataExtDefQuery As IDataExtDefQuery
        Set dataExtDefQuery = msgSetRequest.AppendDataExtDefQueryRq()

        ' set the OwnerID for the Data Ext we want returned
        ' specify an ownerID of "0" for custom fields
        dataExtDefQuery.ORDataExtDefQuery.OwnerIDList.Add ("0")

        ' send the request to QB
        Dim msgSetResponse As IMsgSetResponse
        Set msgSetResponse = QBSessionManager.DoRequests(msgSetRequest)

        ' fill the response info into the form
        FillDataExts msgSetResponse, txtCustomFields
        Exit Sub
Errs:
        MsgBox "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    vbOKOnly, _
                    "Error in GetCustomFields"
    End Sub


    Public Sub FillDataExts(ByRef msgSetResponse As IMsgSetResponse, ByRef txtDataExts As TextBox)
        On Error GoTo Errs

        'Clear the text box
        txtDataExts.Text = ""
        txtDataExts.Refresh

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
        ' this example will only have one response in the list
        ' so we will look at index=0
        Dim response As IResponse
        Set response = responseList.GetAt(0)
        If (Not response Is Nothing) Then
            If response.StatusCode <> "0" Then
                If response.StatusCode = "1" Then
                    txtDataExts.Text = "NO DATA EXTENSIONS DEFINED FOR THIS COMPANY FILE"
                Else
                    MsgBox ("FillDataExts unexpexcted Error - " & response.StatusMessage)
                End If
                Exit Sub
            End If
        End If

        ' first make sure we have a response object to handle
        If (response Is Nothing) Then
            Exit Sub
        End If
        If (response.Type Is Nothing) Or _
            (response.Detail Is Nothing) Then
            Exit Sub
        End If
        If (response.Detail.Type Is Nothing) Then
            Exit Sub
        End If

        '<DataExtDefQueryRs requestID = "UUIDTYPE" statusCode = "INTTYPE" statusSeverity = "STRTYPE" statusMessage = "STRTYPE">
        '  <DataExtDefRet>                                       <!-- opt, may rep -->
        '    <OwnerID>GUIDTYPE</OwnerID>                         <!-- opt -->
        '    <DataExtName>STRTYPE</DataExtName>                  <!-- QBO max = 31 -->
        '    <!-- DataExtType may have one of the following values: IntType, AmtType, PriceType, QuanType, PercentType, DateTimeType, StrType, BlobType -->
        '    <DataExtType>ENUMTYPE</DataExtType>
        '    <!-- AssignToObject may have one of the following values: Customer, Vendor, Employee, OtherName, Item, Account, Bill, BillPaymentCheck, BillPaymentCreditCard, Charge, Check, CreditCardCharge, CreditCardCredit, CreditMemo, Deposit, Estimate, InventoryAdjustment, Invoice, JournalEntry, PurchaseOrder, ReceivePayment, SalesReceipt, SalesTaxPaymentCheck, VendorCredit -->
        '    <AssignToObject>ENUMTYPE</AssignToObject>           <!-- opt, may rep -->
        '  </DataExtDefRet>
        '</DataExtDefQueryRs>

        ' make sure we are processing the DataExtDefQueryRs and
        ' the DataExtDefRetList responses in this response list
        Dim dataExtDefRetList As IDataExtDefRetList
        Dim responseType As ENResponseType
        Dim responseDetailType As ENObjectType
        responseType = response.Type.getValue()
        responseDetailType = response.Detail.Type.getValue()
        If (responseType = ENResponseType.rtDataExtDefQueryRs) And _
            (responseDetailType = ENObjectType.otDataExtDefRetList) Then
            ' save the response detail in the appropriate object type
            ' since we have first verified the type of the response object
            Set dataExtDefRetList = response.Detail
        Else
            ' bail, we do not have the responses we were expecting
            Exit Sub
        End If

        Dim j As Integer
        Dim numObjects As Integer
        Dim strDisplayLine As String

        Dim count As Integer
        Dim index As Integer
        Dim dataExtDefRet As IDataExtDefRet
        Dim assignToObjectList As IENAssignToObjectList
        Dim assignToObjectString As String
        count = dataExtDefRetList.count
        For index = 0 To count - 1
            Set dataExtDefRet = dataExtDefRetList.GetAt(index)
            If (Not dataExtDefRet Is Nothing) Then
                strDisplayLine = ""
                If (Not dataExtDefRet.DataExtName Is Nothing) Then
                    strDisplayLine = dataExtDefRet.DataExtName.getValue()
                End If
                If (Not dataExtDefRet.DataExtType Is Nothing) Then
                    strDisplayLine = strDisplayLine & "  |  " & dataExtDefRet.DataExtType.GetAsString() & "  |  "
                End If

                'Parse the Data Extension return object Assigned Objects to the DataExtDef text box
                Set assignToObjectList = dataExtDefRet.assignToObjectList
                If (Not assignToObjectList Is Nothing) Then
                    numObjects = assignToObjectList.count
                    For j = 0 To numObjects - 1
                        assignToObjectString = assignToObjectList.GetAtAsString(j)
                        If j > 0 Then
                            strDisplayLine = strDisplayLine & ", " & assignToObjectString
                        Else
                            strDisplayLine = strDisplayLine & assignToObjectString
                        End If
                    Next
                End If
            End If

            txtDataExts.Text = txtDataExts.Text & strDisplayLine & vbCrLf
        Next

        Exit Sub
Errs:
        MsgBox "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    vbOKOnly, _
                    "Error in FillDataExts"
    End Sub


    Public Sub AddDataExtDef(ByRef strDataExtName As String, ByRef strDataExtType As String, ByRef strObjects As String)
        On Error GoTo Errs

        If (msgSetRequest Is Nothing) Then
            Exit Sub
        End If

        ' make sure we do not have any old requests still defined
        msgSetRequest.ClearRequests

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = ENRqOnError.roeStop

        ' Add the DataExtDefAddRq request
        Dim dataExtDefAdd As IDataExtDefAdd
        Set dataExtDefAdd = msgSetRequest.AppendDataExtDefAddRq()

        'Add the OwnerID element
        dataExtDefAdd.OwnerID.setValue (cstrOwnerID)

        'Add the DataExtName element
        dataExtDefAdd.DataExtName.setValue (strDataExtName)

        'Add the DataExtType element
        dataExtDefAdd.DataExtType.SetAsString (strDataExtType)

        'Add the AssignToObject elements
        Dim strAssignToObject() As String
        Dim i As Integer
        strAssignToObject = Split(strObjects)
        For i = LBound(strAssignToObject) To UBound(strAssignToObject)
            dataExtDefAdd.assignToObjectList.AddAsString (strAssignToObject(i))
        Next

        ' send the request to QB
        Dim msgSetResponse As IMsgSetResponse
        Set msgSetResponse = QBSessionManager.DoRequests(msgSetRequest)

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
        ' this example will only have one response in the list
        ' so we will look at index=0
        Dim response As IResponse
        Set response = responseList.GetAt(0)
        If (Not response Is Nothing) Then
            If response.StatusCode <> "0" Then
                'If the status is bad, report it to the user
                MsgBox ("AddDataExtDef unexpexcted Error - " & response.StatusMessage)
                Exit Sub
            End If
        End If

        'Since we added a new data extension, update the data extensions text box in the main form
        GetDataExts frmDataExtSample.txtDataExts

        Exit Sub

Errs:
        MsgBox "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    vbOKOnly, _
                    "Error in AddDataExtDef"
    End Sub


    Public Sub GetCustomers(ByRef lstCustomers As ListBox)
        On Error GoTo Errs

        '    "<?xml version=""1.0"" ?>" & _
        '    "<?qbxml version=""2.0""?>" & _
        '    "<QBXML>" & _
        '    "<QBXMLMsgsRq onError = ""continueOnError"">" & _
        '    "<CustomerQueryRq requestID = ""0"">" & _
        '    "</CustomerQueryRq>" & _
        '    "</QBXMLMsgsRq>" & _
        '    "</QBXML>"

        If (msgSetRequest Is Nothing) Then
            Exit Sub
        End If

        ' make sure we do not have any old requests still defined
        msgSetRequest.ClearRequests

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

        ' Add the DataExtDefQuery request
        Dim customerQuery As ICustomerQuery
        Set customerQuery = msgSetRequest.AppendCustomerQueryRq()

        ' send the request to QB
        Dim msgSetResponse As IMsgSetResponse
        Set msgSetResponse = QBSessionManager.DoRequests(msgSetRequest)

        FillCustomerListBox msgSetResponse, lstCustomers
        Exit Sub

Errs:
        MsgBox "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    vbOKOnly, _
                    "Error in AddDataExtDef"
    End Sub

    Public Sub FillCustomerListBox(ByRef msgSetResponse As IMsgSetResponse, ByRef lstCustomers As ListBox)
        On Error GoTo Errs

        'Clear the list box
        lstCustomers.Clear

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
        ' this example will only have one response in the list
        ' so we will look at index=0
        Dim response As IResponse
        Set response = responseList.GetAt(0)
        If (Not response Is Nothing) Then
            If response.StatusCode <> "0" Then
                'If the status is bad, report it to the user
                MsgBox ("FillCustomerListBox unexpexcted Error - " & response.StatusMessage)
                Exit Sub
            End If
        End If

        ' first make sure we have a response object to handle
        If (response Is Nothing) Then
            Exit Sub
        End If
        If (response.Type Is Nothing) Or _
            (response.Detail Is Nothing) Then
            Exit Sub
        End If
        If (response.Detail.Type Is Nothing) Then
            Exit Sub
        End If

        ' make sure we are processing the CustomerQueryRs and
        ' the CustomerRetList responses in this response list
        Dim customerRetList As ICustomerRetList
        Dim responseType As ENResponseType
        Dim responseDetailType As ENObjectType
        responseType = response.Type.getValue()
        responseDetailType = response.Detail.Type.getValue()
        If (responseType = ENResponseType.rtCustomerQueryRs) And _
            (responseDetailType = ENObjectType.otCustomerRetList) Then
            ' save the response detail in the appropriate object type
            ' since we have first verified the type of the response object
            Set customerRetList = response.Detail
        Else
            ' bail, we do not have the responses we were expecting
            Exit Sub
        End If

        'Parse the query response and add the Customers to the Customer list box
        Dim count As Integer
        Dim index As Integer
        Dim CustomerRet As ICustomerRet
        count = customerRetList.count
        For index = 0 To count - 1
            Set CustomerRet = customerRetList.GetAt(index)
            If (Not CustomerRet Is Nothing) Then
                ' do not add any customer:job to the list, just level "0"
                If (CustomerRet.Sublevel.getValue() = 0) Then
                    lstCustomers.AddItem CustomerRet.FullName.getValue()
                End If
            End If
        Next

        Exit Sub
Errs:
        MsgBox "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    vbOKOnly, _
                    "Error in FillCustomerListBox"
    End Sub


    Public Function CustomersHaveDataExts() As Boolean

        GetDataExts frmDataExtSample.txtUtility
        If InStr(frmDataExtSample.txtUtility.Text, "Customer") Then
            CustomersHaveDataExts = True
        Else
            CustomersHaveDataExts = False
        End If
    End Function


    Public Sub FillUnusedDataExts(ByRef strCustomer As String, ByRef lstUnusedDataExtDefs As ListBox)

        Dim strDataExts() As String
        Dim i As Integer
        Dim dummy As String

        lstUnusedDataExtDefs.Clear
        frmAddDataExtension.lstAvailableDataExts.Clear
        frmAddDataExtension.lstUsedDataExts.Clear

        ' fill all the available Data Extensions for "Customer" in the lstAvailableDataExts list box
        GetDataExts frmDataExtSample.txtUtility
        strDataExts = Split(frmDataExtSample.txtUtility.Text, vbCrLf)
        For i = LBound(strDataExts) To UBound(strDataExts) - 1
            If InStr(strDataExts(i), "Customer") > 0 Then
                frmAddDataExtension.lstAvailableDataExts.AddItem (Left(strDataExts(i), InStr(strDataExts(i), "|  Customer") - 3))
            End If
        Next

        ' fill all used data extension for this customer "strCustomer" in the lstUsedDataExts
        GetUsedCustomerDataExts strCustomer, frmAddDataExtension.lstUsedDataExts, False

        'If the list counts are equal then the given customer has values
        'assigned for all of their possible data extensions.
        If frmAddDataExtension.lstUsedDataExts.ListCount = frmAddDataExtension.lstAvailableDataExts.ListCount Then
            lstUnusedDataExtDefs.AddItem ("All Data Extensions for " & strCustomer & " are already assigned values")
            Exit Sub
        End If

        'If none are used, they are all available
        If frmAddDataExtension.lstUsedDataExts.ListCount = 0 Then
            For i = 0 To frmAddDataExtension.lstAvailableDataExts.ListCount - 1
                lstUnusedDataExtDefs.AddItem frmAddDataExtension.lstAvailableDataExts.List(i)
            Next
            Exit Sub
        End If

        'Some are used, so lets figure out which ones
        Dim j As Integer

        j = 0
        Dim usedDataExtCount As Integer
        Dim availableDataExtCount As Integer
        Dim usedDataExtString As String
        Dim availableDataExtString As String
        Dim bFound As Boolean
        usedDataExtCount = frmAddDataExtension.lstUsedDataExts.ListCount
        availableDataExtCount = frmAddDataExtension.lstAvailableDataExts.ListCount
        For i = 0 To availableDataExtCount - 1
            bFound = False
            availableDataExtString = Left(frmAddDataExtension.lstAvailableDataExts.List(i), InStr(frmAddDataExtension.lstAvailableDataExts.List(i), "  |") - 1)
            For j = 0 To usedDataExtCount - 1
                usedDataExtString = frmAddDataExtension.lstUsedDataExts.List(j)
                If (usedDataExtString = availableDataExtString) Then
                    bFound = True
                    Exit For
                End If
            Next
            If Not bFound Then
                lstUnusedDataExtDefs.AddItem (frmAddDataExtension.lstAvailableDataExts.List(i))
            End If
        Next

    End Sub


    Public Sub GetUsedCustomerDataExts(ByRef strCustomer As String, ByRef lstUsedDataExts As ListBox, ByRef booIncludeValues As Boolean)
        On Error GoTo Errs

        '"<?xml version=""1.0"" ?>" & _
        '"<?qbxml version=""2.0""?>" & _
        '"<QBXML>" & _
        '"<QBXMLMsgsRq onError = ""continueOnError"">" & _
        '"<CustomerQueryRq requestID = ""0"">" & _
        '"<FullName>" & strCustomer & "</FullName>" & _
        '"<OwnerID>" & cstrOwnerID & "</OwnerID>" & _
        '"</CustomerQueryRq>" & _
        '"</QBXMLMsgsRq>" & _
        '"</QBXML>"

        If (msgSetRequest Is Nothing) Then
            Exit Sub
        End If

        ' make sure we do not have any old requests still defined
        msgSetRequest.ClearRequests

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

        ' Add the CustomerQuery request
        Dim customerQuery As ICustomerQuery
        Set customerQuery = msgSetRequest.AppendCustomerQueryRq()

        ' set the OwnerID
        customerQuery.OwnerIDList.Add (cstrOwnerID)

        ' set the FullName of the Customer
        customerQuery.ORCustomerListQuery.FullNameList.Add (strCustomer)

        ' send the request to QB
        Dim msgSetResponse As IMsgSetResponse
        Set msgSetResponse = QBSessionManager.DoRequests(msgSetRequest)

        FillUsedDataExts msgSetResponse, lstUsedDataExts, booIncludeValues
        Exit Sub
Errs:
        MsgBox "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    vbOKOnly, _
                    "Error in GetUsedCustomerDataExts"
    End Sub


    Public Sub FillUsedDataExts(ByRef msgSetResponse As IMsgSetResponse, _
                    ByRef lstUsedDataExts As ListBox, _
                    ByRef booIncludeValues As Boolean)
        On Error GoTo Errs

        'Clear the list box
        lstUsedDataExts.Clear

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
        ' this example will only have one response in the list
        ' so we will look at index=0
        Dim response As IResponse
        Set response = responseList.GetAt(0)
        If (Not response Is Nothing) Then
            'Get the status for our query request
            If response.StatusCode <> "0" Then
                'If the status is bad, report it to the user
                MsgBox ("FullUsedDataExts unexpexcted Error - " & response.StatusMessage)
                Exit Sub
            End If
        End If

        ' first make sure we have a response object to handle
        If (response Is Nothing) Then
            Exit Sub
        End If
        If (response.Type Is Nothing) Or _
            (response.Detail Is Nothing) Then
            Exit Sub
        End If
        If (response.Detail.Type Is Nothing) Then
            Exit Sub
        End If

        ' make sure we are processing the CustomerQueryRs and
        ' the CustomerRetList responses in this response list
        Dim customerRetList As ICustomerRetList
        Dim responseType As ENResponseType
        Dim responseDetailType As ENObjectType
        responseType = response.Type.getValue()
        responseDetailType = response.Detail.Type.getValue()
        If (responseType = ENResponseType.rtCustomerQueryRs) And _
            (responseDetailType = ENObjectType.otCustomerRetList) Then
            ' save the response detail in the appropriate object type
            ' since we have first verified the type of the response object
            Set customerRetList = response.Detail
        Else
            ' bail, we do not have the responses we were expecting
            Exit Sub
        End If

        ' <DataExtRet>                                        <!-- opt, may rep -->
        '  <OwnerID>GUIDTYPE</OwnerID>                       <!-- opt -->
        '  <DataExtName>STRTYPE</DataExtName>                <!-- QBO max = 31 -->
        '  <!-- DataExtType may have one of the following values: IntType, AmtType, PriceType, QuanType, PercentType, DateTimeType, StrType, BlobType -->
        '  <DataExtType>ENUMTYPE</DataExtType>
        '  <DataExtValue>STRTYPE</DataExtValue>
        '</DataExtRet>

        'Parse the query response and add the Customers to the Customer list box
        Dim count As Integer
        Dim index As Integer
        Dim strDisplayLine As String
        Dim CustomerRet As ICustomerRet
        Dim dataExtRetList As IDataExtRetList
        Dim dataExtRet As IDataExtRet
        count = customerRetList.count
        For index = 0 To count - 1
            Set CustomerRet = customerRetList.GetAt(index)
            If (Not CustomerRet Is Nothing) Then
                strDisplayLine = ""
                Set dataExtRetList = CustomerRet.dataExtRetList
                If (Not dataExtRetList Is Nothing) Then
                    Dim i As Integer
                    Dim maxI As Integer
                    maxI = dataExtRetList.count
                    For i = 0 To maxI - 1
                        Set dataExtRet = dataExtRetList.GetAt(i)
                        If (Not dataExtRet Is Nothing) Then
                            If (Not dataExtRet.DataExtName Is Nothing) Then
                                strDisplayLine = dataExtRet.DataExtName.getValue()
                            End If
                            If booIncludeValues Then
                                If (Not dataExtRet.DataExtValue Is Nothing) Then
                                    strDisplayLine = strDisplayLine & "  =  " & dataExtRet.DataExtValue.getValue()
                                End If
                            End If
                            lstUsedDataExts.AddItem (strDisplayLine)
                        End If
                    Next

                End If
            End If
        Next
        Exit Sub
Errs:
        MsgBox "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    vbOKOnly, _
                    "Error in FillUsedDataExts"

    End Sub


    Public Function AddDataExt(ByRef strCustomer As String, ByRef strDataExtName As String, ByRef strDataExtValue As String) As Boolean
        On Error GoTo Errs

        If (msgSetRequest Is Nothing) Then
            Exit Function
        End If

        ' make sure we do not have any old requests still defined
        msgSetRequest.ClearRequests

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = ENRqOnError.roeStop

        ' Add the DataExtAddRq request
        Dim dataExtAdd As IDataExtAdd
        Set dataExtAdd = msgSetRequest.AppendDataExtAddRq()

        'Add the OwnerID element
        dataExtAdd.OwnerID.setValue (cstrOwnerID)

        'Add the DataExtName element
        dataExtAdd.DataExtName.setValue (strDataExtName)

        'Add the ListDataExtType element
        dataExtAdd.ORListTxnWithMacro.ListDataExt.ListDataExtType.SetAsString ("Customer")

        'Add the FullName element
        dataExtAdd.ORListTxnWithMacro.ListDataExt.ListObjRef.FullName.setValue (strCustomer)

        'Add the DataExtValue element
        dataExtAdd.DataExtValue.setValue (strDataExtValue)

        ' send the request to QB
        Dim msgSetResponse As IMsgSetResponse
        Set msgSetResponse = QBSessionManager.DoRequests(msgSetRequest)

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
        ' this example will only have one response in the list
        ' so we will look at index=0
        Dim response As IResponse
        Set response = responseList.GetAt(0)
        If (Not response Is Nothing) Then
            If response.StatusCode <> "0" Then
                'If the status is bad, report it to the user
                MsgBox ("AddDataExt unexpexcted Error - " & response.StatusMessage)
                Exit Function
            End If
        End If

        MsgBox ("Successfully added Data Extension " & strDataExtName & " with value " & strDataExtValue & " to customer " & strCustomer)
        AddDataExt = True
        Exit Function

Errs:
        MsgBox "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    vbOKOnly, _
                    "Error in AddDataExt"
        AddDataExt = False
    End Function


    Public Function ModDataExt(ByRef strCustomer As String, ByRef strDataExtName As String, ByRef strDataExtValue As String) As Boolean
        On Error GoTo Errs

        If (msgSetRequest Is Nothing) Then
            Exit Function
        End If

        ' make sure we do not have any old requests still defined
        msgSetRequest.ClearRequests

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = ENRqOnError.roeStop

        ' Add the DataExtAddRq request
        Dim dataExtMod As IDataExtMod
        Set dataExtMod = msgSetRequest.AppendDataExtModRq()

        'Add the OwnerID element
        dataExtMod.OwnerID.setValue (cstrOwnerID)

        'Add the DataExtName element
        dataExtMod.DataExtName.setValue (strDataExtName)

        'Add the ListDataExtType element
        dataExtMod.ORListTxn.ListDataExt.ListDataExtType.SetAsString ("Customer")

        'Add the FullName element
        dataExtMod.ORListTxn.ListDataExt.ListObjRef.FullName.setValue (strCustomer)

        'Add the DataExtValue element
        dataExtMod.DataExtValue.setValue (strDataExtValue)

        ' send the request to QB
        Dim msgSetResponse As IMsgSetResponse
        Set msgSetResponse = QBSessionManager.DoRequests(msgSetRequest)

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
        ' this example will only have one response in the list
        ' so we will look at index=0
        Dim response As IResponse
        Set response = responseList.GetAt(0)
        If (Not response Is Nothing) Then
            If response.StatusCode <> "0" Then
                'If the status is bad, report it to the user
                MsgBox ("ModDataExt unexpexcted Error - " & response.StatusMessage)
                Exit Function
            End If
        End If

        MsgBox ("Successfully modified Data Extension " & strDataExtName & " with value " & strDataExtValue & " to customer " & strCustomer)
        ModDataExt = True
        Exit Function

Errs:
        MsgBox "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    vbOKOnly, _
                    "Error in ModDataExt"
        ModDataExt = False
    End Function


    Public Sub EndSessionCloseConnection()
        On Error GoTo Errs
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
