Option Strict Off
Option Explicit On 
Imports Interop.QBFC13

Module modDataExtSample
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
	'
	'----------------------------------------------------------
	
	Const cstrOwnerID As String = "{E09C86CF-9D6E-4EF2-BCBE-4D66B6B0F754}"
	
	Dim booSessionBegun As Boolean
	
    'Module objects
    Dim qbSessionManager As qbSessionManager
    Dim msgSetRequest As IMsgSetRequest

    Public Sub OpenConnectionBeginSession()

        booSessionBegun = False

        On Error GoTo Errs

        qbSessionManager = New QBSessionManager()

        qbSessionManager.OpenConnection("", "IDN Data Extension Sample - QBFC")

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


    Public Sub GetDataExts(ByRef txtDataExts As System.Windows.Forms.TextBox)
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
        msgSetRequest.ClearRequests()

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

        ' Add the DataExtDefQuery request
        Dim dataExtDefQuery As IDataExtDefQuery
        dataExtDefQuery = msgSetRequest.AppendDataExtDefQueryRq()

        ' set the OwnerID for the Data Ext we want returned
        dataExtDefQuery.ORDataExtDefQuery.OwnerIDList.Add(cstrOwnerID)

        ' send the request to QB
        Dim msgSetResponse As IMsgSetResponse
        msgSetResponse = qbSessionManager.DoRequests(msgSetRequest)

        ' fill the response info into the form
        FillDataExts(msgSetResponse, txtDataExts)
        Exit Sub
Errs:
        MsgBox("HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    MsgBoxStyle.Critical, _
                    "Error in GetDataExts")
    End Sub


    Public Sub GetCustomFields(ByRef txtCustomFields As System.Windows.Forms.TextBox)
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

        ' make sure we do not have any old requests still defined
        msgSetRequest.ClearRequests()

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

        ' Add the DataExtDefQuery request
        Dim dataExtDefQuery As IDataExtDefQuery
        dataExtDefQuery = msgSetRequest.AppendDataExtDefQueryRq()

        ' set the OwnerID for the Data Ext we want returned
        ' specify an ownerID of "0" for custom fields
        dataExtDefQuery.ORDataExtDefQuery.OwnerIDList.Add("0")

        ' send the request to QB
        Dim msgSetResponse As IMsgSetResponse
        msgSetResponse = qbSessionManager.DoRequests(msgSetRequest)

        ' fill the response info into the form
        FillDataExts(msgSetResponse, txtCustomFields)
        Exit Sub
Errs:
        MsgBox("HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    MsgBoxStyle.Critical, _
                    "Error in GetCustomFields")
    End Sub


    Public Sub FillDataExts(ByRef msgSetResponse As IMsgSetResponse, ByRef txtDataExts As System.Windows.Forms.TextBox)
        On Error GoTo Errs

        'Clear the text box
        txtDataExts.Text = ""
        txtDataExts.Refresh()

        ' check to make sure we have objects to access first
        ' and that there are responses in the list
        If (msgSetResponse Is Nothing) Or _
            (msgSetResponse.ResponseList Is Nothing) Or _
            (msgSetResponse.ResponseList.Count <= 0) Then
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
                If response.StatusCode = "1" Then
                    txtDataExts.Text = "NO DATA EXTENSIONS DEFINED FOR THIS COMPANY FILE"
                Else
                    MsgBox("FillDataExts unexpexcted Error - " & response.StatusMessage)
                End If
                Exit Sub
            End If
        End If

        ' first make sure we have a response object to handle
        If (response Is Nothing) Or _
            (response.Type Is Nothing) Or _
            (response.Detail Is Nothing) Or _
            (response.Detail.Type Is Nothing) Then
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
        responseType = response.Type.GetValue()
        responseDetailType = response.Detail.Type.GetValue()
        If (responseType = ENResponseType.rtDataExtDefQueryRs) And _
            (responseDetailType = ENObjectType.otDataExtDefRetList) Then
            ' save the response detail in the appropriate object type
            ' since we have first verified the type of the response object
            dataExtDefRetList = response.Detail
        Else
            ' bail, we do not have the responses we were expecting
            Exit Sub
        End If

        Dim j As Short
        Dim numObjects As Integer
        Dim strDisplayLine As String

        Dim count As Short
        Dim index As Short
        Dim dataExtDefRet As IDataExtDefRet
        Dim assignToObjectList As IENAssignToObjectList
        Dim assignToObjectString As String
        count = dataExtDefRetList.Count
        For index = 0 To count - 1
            dataExtDefRet = dataExtDefRetList.GetAt(index)
            If (Not dataExtDefRet Is Nothing) Then
                strDisplayLine = ""
                If (Not dataExtDefRet.DataExtName Is Nothing) Then
                    strDisplayLine = dataExtDefRet.DataExtName.GetValue()
                End If
                If (Not dataExtDefRet.DataExtType Is Nothing) Then
                    strDisplayLine = strDisplayLine & "  |  " & dataExtDefRet.DataExtType.GetAsString() & "  |  "
                End If

                'Parse the Data Extension return object Assigned Objects to the DataExtDef text box
                assignToObjectList = dataExtDefRet.AssignToObjectList
                If (Not assignToObjectList Is Nothing) Then
                    numObjects = assignToObjectList.Count
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
        MsgBox("HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    MsgBoxStyle.Critical, _
                    "Error in FillDataExts")
    End Sub


    Public Sub AddDataExtDef(ByRef strDataExtName As String, ByRef strDataExtType As String, ByRef strObjects As String)
        On Error GoTo Errs

        If (msgSetRequest Is Nothing) Then
            Exit Sub
        End If

        ' make sure we do not have any old requests still defined
        msgSetRequest.ClearRequests()

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = ENRqOnError.roeStop

        ' Add the DataExtDefAddRq request
        Dim dataExtDefAdd As IDataExtDefAdd
        dataExtDefAdd = msgSetRequest.AppendDataExtDefAddRq()

        'Add the OwnerID element
        dataExtDefAdd.OwnerID.SetValue(cstrOwnerID)

        'Add the DataExtName element
        dataExtDefAdd.DataExtName.SetValue(strDataExtName)

        'Add the DataExtType element
        dataExtDefAdd.DataExtType.SetAsString(strDataExtType)

        'Add the AssignToObject elements
        Dim strAssignToObject() As String
        Dim i As Short
        strAssignToObject = Split(strObjects)
        For i = LBound(strAssignToObject) To UBound(strAssignToObject)
            dataExtDefAdd.AssignToObjectList.AddAsString(strAssignToObject(i))
        Next

        ' send the request to QB
        Dim msgSetResponse As IMsgSetResponse
        msgSetResponse = qbSessionManager.DoRequests(msgSetRequest)

        ' check to make sure we have objects to access first
        ' and that there are responses in the list
        If (msgSetResponse Is Nothing) Or _
            (msgSetResponse.ResponseList Is Nothing) Or _
            (msgSetResponse.ResponseList.Count <= 0) Then
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
                MsgBox("AddDataExtDef unexpexcted Error - " & response.StatusMessage)
                Exit Sub
            End If
        End If

        'Since we added a new data extension, update the data extensions text box in the main form
        GetDataExts((frmDataExtSample.DefInstance.txtDataExts))

        Exit Sub

Errs:
        MsgBox("HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    MsgBoxStyle.Critical, _
                    "Error in AddDataExtDef")
    End Sub


    Public Sub GetCustomers(ByRef lstCustomers As System.Windows.Forms.ListBox)
        On Error GoTo Errs

        '    "<?xml version=""1.0"" ?>" & _
        '    "<?qbxml version=""2.0""?>" & _
        '    "<QBXML>" & _
        '    "<QBXMLMsgsRq onError = ""continueOnError"">" & _
        '    "<CustomerQueryRq requestID = ""0"">" & _
        '    "</CustomerQueryRq>" & _
        '    "</QBXMLMsgsRq>" & _
        '    "</QBXML>"

        ' make sure we do not have any old requests still defined
        msgSetRequest.ClearRequests()

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

        ' Add the DataExtDefQuery request
        Dim customerQuery As ICustomerQuery
        customerQuery = msgSetRequest.AppendCustomerQueryRq()

        ' send the request to QB
        Dim msgSetResponse As IMsgSetResponse
        msgSetResponse = qbSessionManager.DoRequests(msgSetRequest)

        FillCustomerListBox(msgSetResponse, lstCustomers)
        Exit Sub

Errs:
        MsgBox("HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    MsgBoxStyle.Critical, _
                    "Error in AddDataExtDef")
    End Sub

    Public Sub FillCustomerListBox(ByRef msgSetResponse As IMsgSetResponse, ByRef lstCustomers As System.Windows.Forms.ListBox)
        On Error GoTo Errs

        'Clear the list box
        lstCustomers.Items.Clear()

        ' check to make sure we have objects to access first
        ' and that there are responses in the list
        If (msgSetResponse Is Nothing) Or _
            (msgSetResponse.ResponseList Is Nothing) Or _
            (msgSetResponse.ResponseList.Count <= 0) Then
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
                Exit Sub
            End If
        End If

        ' first make sure we have a response object to handle
        If (response Is Nothing) Or _
            (response.Type Is Nothing) Or _
            (response.Detail Is Nothing) Or _
            (response.Detail.Type Is Nothing) Then
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
            Exit Sub
        End If

        'Parse the query response and add the Customers to the Customer list box
        Dim count As Short
        Dim index As Short
        Dim CustomerRet As ICustomerRet
        count = customerRetList.Count
        For index = 0 To count - 1
            CustomerRet = customerRetList.GetAt(index)
            If (Not CustomerRet Is Nothing) Then
                ' do not add any customer:job to the list, just level "0"
                If (CustomerRet.Sublevel.GetValue() = 0) Then
                    lstCustomers.Items.Add(CustomerRet.FullName.GetValue())
                End If
            End If
        Next

        Exit Sub
Errs:
        MsgBox("HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    MsgBoxStyle.Critical, _
                    "Error in FillCustomerListBox")
    End Sub


    Public Function CustomersHaveDataExts() As Boolean

        GetDataExts((frmDataExtSample.DefInstance.txtUtility))
        If InStr(frmDataExtSample.DefInstance.txtUtility.Text, "Customer") Then
            CustomersHaveDataExts = True
        Else
            CustomersHaveDataExts = False
        End If
    End Function


    Public Sub FillUnusedDataExts(ByRef strCustomer As String, ByRef lstUnusedDataExtDefs As System.Windows.Forms.ListBox)

        Dim strDataExts() As String
        Dim i As Short
        Dim dummy As String

        lstUnusedDataExtDefs.Items.Clear()
        frmAddDataExtension.DefInstance.lstAvailableDataExts.Items.Clear()
        frmAddDataExtension.DefInstance.lstUsedDataExts.Items.Clear()

        ' fill all the available Data Extensions for "Customer" in the lstAvailableDataExts list box
        GetDataExts((frmDataExtSample.DefInstance.txtUtility))
        strDataExts = Split(frmDataExtSample.DefInstance.txtUtility.Text, vbCrLf)
        For i = LBound(strDataExts) To UBound(strDataExts) - 1
            If InStr(strDataExts(i), "Customer") > 0 Then
                frmAddDataExtension.DefInstance.lstAvailableDataExts.Items.Add(Left(strDataExts(i), InStr(strDataExts(i), "|  Customer") - 3))
            End If
        Next

        ' fill all used data extension for this customer "strCustomer" in the lstUsedDataExts
        GetUsedCustomerDataExts(strCustomer, (frmAddDataExtension.DefInstance.lstUsedDataExts), False)

        'If the list counts are equal then the given customer has values
        'assigned for all of their possible data extensions.
        If frmAddDataExtension.DefInstance.lstUsedDataExts.Items.Count = frmAddDataExtension.DefInstance.lstAvailableDataExts.Items.Count Then
            lstUnusedDataExtDefs.Items.Add("All Data Extensions for " & strCustomer & " are already assigned values")
            Exit Sub
        End If

        'If none are used, they are all available
        If frmAddDataExtension.DefInstance.lstUsedDataExts.Items.Count = 0 Then
            For i = 0 To frmAddDataExtension.DefInstance.lstAvailableDataExts.Items.Count - 1
                lstUnusedDataExtDefs.Items.Add(VB6.GetItemString(frmAddDataExtension.DefInstance.lstAvailableDataExts, i))
            Next
            Exit Sub
        End If

        'Some are used, so lets figure out which ones
        Dim j As Short

        j = 0
        Dim usedDataExtCount As Short
        Dim availableDataExtCount As Short
        Dim usedDataExtString As String
        Dim availableDataExtString As String
        Dim bFound As Boolean
        usedDataExtCount = frmAddDataExtension.DefInstance.lstUsedDataExts.Items.Count
        availableDataExtCount = frmAddDataExtension.DefInstance.lstAvailableDataExts.Items.Count
        For i = 0 To availableDataExtCount - 1
            bFound = False
            availableDataExtString = Left(VB6.GetItemString(frmAddDataExtension.DefInstance.lstAvailableDataExts, i), InStr(VB6.GetItemString(frmAddDataExtension.DefInstance.lstAvailableDataExts, i), "  |") - 1)
            For j = 0 To usedDataExtCount - 1
                usedDataExtString = VB6.GetItemString(frmAddDataExtension.DefInstance.lstUsedDataExts, j)
                If (usedDataExtString = availableDataExtString) Then
                    bFound = True
                    Exit For
                End If
            Next
            If Not bFound Then
                lstUnusedDataExtDefs.Items.Add(VB6.GetItemString(frmAddDataExtension.DefInstance.lstAvailableDataExts, i))
            End If
        Next

    End Sub


    Public Sub GetUsedCustomerDataExts(ByRef strCustomer As String, ByRef lstUsedDataExts As System.Windows.Forms.ListBox, ByRef booIncludeValues As Boolean)
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
        msgSetRequest.ClearRequests()

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

        ' Add the CustomerQuery request
        Dim customerQuery As ICustomerQuery
        customerQuery = msgSetRequest.AppendCustomerQueryRq()

        ' set the OwnerID 
        customerQuery.OwnerIDList.Add(cstrOwnerID)

        ' set the FullName of the Customer
        customerQuery.ORCustomerListQuery.FullNameList.Add(strCustomer)

        ' send the request to QB
        Dim msgSetResponse As IMsgSetResponse
        msgSetResponse = qbSessionManager.DoRequests(msgSetRequest)

        FillUsedDataExts(msgSetResponse, lstUsedDataExts, booIncludeValues)
        Exit Sub
Errs:
        MsgBox("HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    MsgBoxStyle.Critical, _
                    "Error in GetUsedCustomerDataExts")
    End Sub


    Public Sub FillUsedDataExts(ByRef msgSetResponse As IMsgSetResponse, _
                    ByRef lstUsedDataExts As System.Windows.Forms.ListBox, _
                    ByRef booIncludeValues As Boolean)
        On Error GoTo Errs

        'Clear the list box
        lstUsedDataExts.Items.Clear()

        ' check to make sure we have objects to access first
        ' and that there are responses in the list
        If (msgSetResponse Is Nothing) Or _
            (msgSetResponse.ResponseList Is Nothing) Or _
            (msgSetResponse.ResponseList.Count <= 0) Then
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
            'Get the status for our query request
            If response.StatusCode <> "0" Then
                'If the status is bad, report it to the user
                MsgBox("FullUsedDataExts unexpexcted Error - " & response.StatusMessage)
                Exit Sub
            End If
        End If

        ' first make sure we have a response object to handle
        If (response Is Nothing) Or _
            (response.Type Is Nothing) Or _
            (response.Detail Is Nothing) Or _
            (response.Detail.Type Is Nothing) Then
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
        Dim count As Short
        Dim index As Short
        Dim strDisplayLine As String
        Dim CustomerRet As ICustomerRet
        Dim dataExtRetList As IDataExtRetList
        Dim dataExtRet As IDataExtRet
        count = customerRetList.Count
        For index = 0 To count - 1
            CustomerRet = customerRetList.GetAt(index)
            If (Not CustomerRet Is Nothing) Then
                strDisplayLine = ""
                dataExtRetList = CustomerRet.DataExtRetList
                If (Not dataExtRetList Is Nothing) Then
                    Dim i As Short
                    Dim maxI As Short
                    maxI = dataExtRetList.Count
                    For i = 0 To maxI - 1
                        dataExtRet = dataExtRetList.GetAt(i)
                        If (Not dataExtRet Is Nothing) Then
                            If (Not dataExtRet.DataExtName Is Nothing) Then
                                strDisplayLine = dataExtRet.DataExtName.GetValue()
                            End If
                            If booIncludeValues Then
                                If (Not dataExtRet.DataExtValue Is Nothing) Then
                                    strDisplayLine = strDisplayLine & "  =  " & dataExtRet.DataExtValue.GetValue()
                                End If
                            End If
                            lstUsedDataExts.Items.Add(strDisplayLine)
                        End If
                    Next

                End If
            End If
        Next
        Exit Sub
Errs:
        MsgBox("HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    MsgBoxStyle.Critical, _
                    "Error in FillUsedDataExts")

    End Sub


    Public Function AddDataExt(ByRef strCustomer As String, ByRef strDataExtName As String, ByRef strDataExtValue As String) As Boolean
        On Error GoTo Errs

        If (msgSetRequest Is Nothing) Then
            Exit Function
        End If

        ' make sure we do not have any old requests still defined
        msgSetRequest.ClearRequests()

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = ENRqOnError.roeStop

        ' Add the DataExtAddRq request
        Dim dataExtAdd As IDataExtAdd
        dataExtAdd = msgSetRequest.AppendDataExtAddRq()

        'Add the OwnerID element
        dataExtAdd.OwnerID.SetValue(cstrOwnerID)

        'Add the DataExtName element
        dataExtAdd.DataExtName.SetValue(strDataExtName)

        'Add the ListDataExtType element
        dataExtAdd.ORListTxnWithMacro.ListDataExt.ListDataExtType.SetAsString("Customer")

        'Add the FullName element
        dataExtAdd.ORListTxnWithMacro.ListDataExt.ListObjRef.FullName.SetValue(strCustomer)

        'Add the DataExtValue element
        dataExtAdd.DataExtValue.SetValue(strDataExtValue)

        ' send the request to QB
        Dim msgSetResponse As IMsgSetResponse
        msgSetResponse = qbSessionManager.DoRequests(msgSetRequest)

        ' check to make sure we have objects to access first
        ' and that there are responses in the list
        If (msgSetResponse Is Nothing) Or _
            (msgSetResponse.ResponseList Is Nothing) Or _
            (msgSetResponse.ResponseList.Count <= 0) Then
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
                MsgBox("AddDataExt unexpexcted Error - " & response.StatusMessage)
                Exit Function
            End If
        End If

        MsgBox("Successfully added Data Extension " & strDataExtName & " with value " & strDataExtValue & " to customer " & strCustomer)
        AddDataExt = True
        Exit Function

Errs:
        MsgBox("HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    MsgBoxStyle.Critical, _
                    "Error in AddDataExt")
        AddDataExt = False
    End Function


    Public Function ModDataExt(ByRef strCustomer As String, ByRef strDataExtName As String, ByRef strDataExtValue As String) As Boolean
        On Error GoTo Errs

        If (msgSetRequest Is Nothing) Then
            Exit Function
        End If

        ' make sure we do not have any old requests still defined
        msgSetRequest.ClearRequests()

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = ENRqOnError.roeStop

        ' Add the DataExtAddRq request
        Dim dataExtMod As IDataExtMod
        dataExtMod = msgSetRequest.AppendDataExtModRq()

        'Add the OwnerID element
        dataExtMod.OwnerID.SetValue(cstrOwnerID)

        'Add the DataExtName element
        dataExtMod.DataExtName.SetValue(strDataExtName)

        'Add the ListDataExtType element
        dataExtMod.ORListTxn.ListDataExt.ListDataExtType.SetAsString("Customer")

        'Add the FullName element
        dataExtMod.ORListTxn.ListDataExt.ListObjRef.FullName.SetValue(strCustomer)

        'Add the DataExtValue element
        dataExtMod.DataExtValue.SetValue(strDataExtValue)

        ' send the request to QB
        Dim msgSetResponse As IMsgSetResponse
        msgSetResponse = qbSessionManager.DoRequests(msgSetRequest)

        ' check to make sure we have objects to access first
        ' and that there are responses in the list
        If (msgSetResponse Is Nothing) Or _
            (msgSetResponse.ResponseList Is Nothing) Or _
            (msgSetResponse.ResponseList.Count <= 0) Then
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
                MsgBox("ModDataExt unexpexcted Error - " & response.StatusMessage)
                Exit Function
            End If
        End If

        MsgBox("Successfully modified Data Extension " & strDataExtName & " with value " & strDataExtValue & " to customer " & strCustomer)
        ModDataExt = True
        Exit Function

Errs:
        MsgBox("HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    MsgBoxStyle.Critical, _
                    "Error in ModDataExt")
        ModDataExt = False
    End Function


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
                    "Error in EndSessionClosConnection")
    End Sub
End Module
