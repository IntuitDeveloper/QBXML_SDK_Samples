Option Strict Off
Option Explicit On
Imports MSXML2
Imports Interop.QBFC15


Module modQBFC
    '----------------------------------------------------------
    ' Copyright © 2003-2013 Intuit Inc. All rights reserved.
    ' Use is subject to the terms specified at:
    '      http://developer.intuit.com/rdmgr/?ID=100
    '
    ' Updated to QBFC 12.0 and fixed setting max QBXML version: 09/2012
    '----------------------------------------------------------


    Dim objSessionManager As New QBSessionManager

    'Set the max versions based on the version of QBFC we're using.  If we
    'were to run against QB 2004 the supported version would be 3.0, but
    'QBFC2_1 can only understand up to version 2.1, so we use these constants
    'to make sure we don't try to create version 3.x messages with a QBFC that
    'can't create them
    Const MaxMajorVersion As Short = 2
    Const MaxMinorVersion As Short = 1

    Dim booSessionBegun As Boolean

    Dim intMajorVersion As Short
    Dim intMinorVersion As Short

    'Declare global variables for keeping the Invoice Query Response object
    'and the Invoice Ret object so we can access them later
    Dim objInvoiceMsgSetResponse As IMsgSetResponse
    Dim objSavedInvoiceRet As IInvoiceRet

    Dim strSavedRequestCode As String


    Public Sub QBFC_OpenConnectionBeginSession()

        On Error GoTo Errs

        objSessionManager.OpenConnection("", "IDN Invoice Modify Sample Application")
        objSessionManager.BeginSession("", ENOpenMode.omDontCare)
        booSessionBegun = True
        Exit Sub

Errs:
        If Err.Number = &H80040416 Then
            MsgBox("You must have QuickBooks running with a company" & vbCrLf & "file open to use this program.")
            objSessionManager.CloseConnection()
            End
        ElseIf Err.Number = &H80040422 Then
            MsgBox("This QuickBooks company file is open in single user mode and" & vbCrLf & "another application is already accessing it.  Please exit the" & vbCrLf & "other application and run this application again.")
            objSessionManager.CloseConnection()
            End
        ElseIf Err.Number = &H80040308 Then
            MsgBox("It appears that the qbXML Request Processor has not" & vbCrLf & "been installed, indicating QuickBooks 2002 or later" & vbCrLf & "may not have been installed.  Please run this sample" & vbCrLf & "after installing QuickBooks 2003 and running the Upgrade.")
        ElseIf Err.Number = &H8007007E Then
            '    If QBFC2CA_IsntInstalled Then
            MsgBox("QBFC2 isn't installed.  You need QBFC2 or QBFC2CA installed to" & vbCrLf & "use QBFC with this sample program.")
            End
            '    Else
            '      QBFCCA_OpenConnectionBeginSession
            '      SetImplementation "QBFCCA"
            '    End If
        Else
            MsgBox("QBFC_OpenConnectionBeginSession" & vbCrLf & Err.Number & vbCrLf & Hex(Err.Number) & vbCrLf & Err.Description)
            End
        End If
    End Sub


    Public Sub QBFC_EndSessionCloseConnection()

        If booSessionBegun Then
            objSessionManager.EndSession()
            objSessionManager.CloseConnection()
        End If
    End Sub

    Function QBFC_MaxVersionSupported() As String

        Dim strVersions() As String

        strVersions = VB6.CopyArray(objSessionManager.QBXMLVersionsForSession)
        QBFC_MaxVersionSupported = strVersions(UBound(strVersions))

        If InStr(1, strVersions(UBound(strVersions)), "CA") Then
            'We're in the rare situation where QBFC2 and QBFC2CA are both
            'installed and we're running against QBCA.  End session, close our
            'connection and open and begin with QBFCCA
            '    QBFC_EndSessionCloseConnection
            '    SetImplementation "QBFCCA"
            '    QBFCCA_OpenConnectionBeginSession
            '    QBFC_MaxVersionSupported = QBFCCA_MaxVersionSupported
            '    Exit Function

            ' For now exit the program if we're dealing with the Canadian version of
            ' QB
            MsgBox("The Canadian version of QBFC does not support version 2.1 " & "messages.  Exiting.")
            End
        End If

        'Now make sure that the version of QBFC installed supports the
        'maximum version of qbXML that QuickBooks can handle
        Dim intQBFCMajorVersion As Short
        Dim intQBFCMinorVersion As Short
        Dim enumReleaseLevel As ENReleaseLevel
        Dim intReleaseNumber As Short

        objSessionManager.GetVersion(intQBFCMajorVersion, intQBFCMinorVersion, enumReleaseLevel, intReleaseNumber)

        Dim versionArray() As String
        versionArray = Split(strVersions(UBound(strVersions)), ".")

        intMajorVersion = Int(CDbl(versionArray(0)))
        intMinorVersion = Int(CDbl(versionArray(1)))

        If intMajorVersion > MaxMajorVersion Then
            intMajorVersion = MaxMajorVersion
            intMinorVersion = MaxMinorVersion
        End If

        QBFC_MaxVersionSupported = Trim(Str(intMajorVersion)) & "." & Trim(Str(intMinorVersion))
    End Function


    Public Sub QBFC_FillComboBox(ByRef cmbComboBox As System.Windows.Forms.ComboBox, ByRef strQueryType As String, ByRef strNameElement As String, ByRef strFilter As String, ByRef booMarkGroupItems As Boolean)

        'Clear the combo box
        cmbComboBox.Items.Clear()

        Dim strSplits() As String
        strSplits = Split(strQueryType, ",")

        Dim strNameElementSplits() As String
        strNameElementSplits = Split(strNameElement, ",")

        Dim objMsgSetRequest As IMsgSetRequest
        Dim objMsgSetResponse As IMsgSetResponse
        Dim objResponse As IResponse

        Dim objQuery As Object
        Dim objRetList As Object
        Dim objRet As Object

        Dim numItems As Short

        Dim i As Short
        Dim j As Short
        Dim strTemp As String
        Dim strStartTag As String
        Dim strEndTag As String
        Dim strValue As String
        Dim intTagLength As Short
        For i = 0 To UBound(strSplits)

            objMsgSetRequest = objSessionManager.CreateMsgSetRequest("US", intMajorVersion, intMinorVersion)
            objMsgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

            Select Case strSplits(i)
                Case Is = "Account"
                    objQuery = objMsgSetRequest.AppendAccountQueryRq
                Case Is = "Class"
                    objQuery = objMsgSetRequest.AppendClassQueryRq
                Case Is = "Customer"
                    objQuery = objMsgSetRequest.AppendCustomerQueryRq
                Case Is = "CustomerMsg"
                    objQuery = objMsgSetRequest.AppendCustomerMsgQueryRq
                Case Is = "ItemService"
                    objQuery = objMsgSetRequest.AppendItemServiceQueryRq
                Case Is = "ItemInventory"
                    objQuery = objMsgSetRequest.AppendItemInventoryQueryRq
                Case Is = "ItemInventoryAssembly"
                    objQuery = objMsgSetRequest.AppendItemInventoryAssemblyQueryRq
                Case Is = "ItemNonInventory"
                    objQuery = objMsgSetRequest.AppendItemNonInventoryQueryRq
                Case Is = "ItemOtherCharge"
                    objQuery = objMsgSetRequest.AppendItemOtherChargeQueryRq
                Case Is = "ItemSubtotal"
                    objQuery = objMsgSetRequest.AppendItemSubtotalQueryRq
                Case Is = "ItemGroup"
                    objQuery = objMsgSetRequest.AppendItemGroupQueryRq
                Case Is = "ItemDiscount"
                    objQuery = objMsgSetRequest.AppendItemDiscountQueryRq
                Case Is = "ItemPayment"
                    objQuery = objMsgSetRequest.AppendItemPaymentQueryRq
                Case Is = "ItemSalesTax"
                    objQuery = objMsgSetRequest.AppendItemSalesTaxQueryRq
                Case Is = "ItemSalesTaxGroup"
                    objQuery = objMsgSetRequest.AppendItemSalesTaxGroupQueryRq
                Case Is = "SalesRep"
                    objQuery = objMsgSetRequest.AppendSalesRepQueryRq
                Case Is = "SalesTaxCode"
                    objQuery = objMsgSetRequest.AppendSalesTaxCodeQueryRq
                Case Is = "ShipMethod"
                    objQuery = objMsgSetRequest.AppendShipMethodQueryRq
                Case Is = "StandardTerms"
                    objQuery = objMsgSetRequest.AppendStandardTermsQueryRq
                Case Else
                    MsgBox("Unknown type " & strSplits(i) & " passed to QBFC_FillComboBox")
            End Select

            'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            If Not String.IsNullOrEmpty(strFilter) Then

                strTemp = strFilter
                'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                Do While Not String.IsNullOrEmpty(strTemp)
                    GetTags(strTemp, strStartTag, strEndTag, intTagLength)

                    strValue = Left(strTemp, InStr(1, strTemp, strEndTag) - 1)
                    strValue = Right(strValue, Len(strValue) - intTagLength)
                    strTemp = Right(strTemp, Len(strTemp) - (InStr(1, strTemp, strEndTag) + intTagLength))

                    Select Case strStartTag
                        Case Is = "<AccountType>"
                            'UPGRADE_WARNING: Couldn't resolve default property of object objQuery.ORAccountListQuery. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            objQuery.ORAccountListQuery.AccountListFilter.AccountTypeList.AddAsString(strValue)
                        Case Else
                            MsgBox("Unknown filter " & strStartTag & " in QBFC_FillComboBox")
                    End Select
                Loop
            End If

            objMsgSetResponse = objSessionManager.DoRequests(objMsgSetRequest)

            objResponse = objMsgSetResponse.ResponseList.GetAt(0)

            If objResponse.StatusCode <> 0 Then
                If objResponse.StatusCode <> 1 Then
                    MsgBox("Status Code " & objResponse.StatusCode & " on call to QBFC_FillComboBox" & vbCrLf & " for " & strSplits(i) & " list items")
                End If
            Else

                objRetList = objResponse.Detail
                'UPGRADE_WARNING: Couldn't resolve default property of object objRetList.Count. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                numItems = objRetList.Count

                If UBound(strNameElementSplits) > 0 Then
                    strNameElement = strNameElementSplits(i)
                End If

                For j = 0 To numItems - 1
                    'UPGRADE_WARNING: Couldn't resolve default property of object objRetList.GetAt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    objRet = objRetList.GetAt(j)
                    If strSplits(i) = "ItemGroup" And booMarkGroupItems Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object objRet.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        cmbComboBox.Items.Add(objRet.Name.getValue & " - Group Item")
                    Else
                        If strNameElement = "FullName" Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object objRet.FullName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            cmbComboBox.Items.Add(objRet.FullName.getValue)
                        Else
                            'UPGRADE_WARNING: Couldn't resolve default property of object objRet.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            cmbComboBox.Items.Add(objRet.Name.getValue)
                        End If
                    End If
                Next j
            End If ' If objResponse.StatusCode <> 0
        Next i
    End Sub


    Public Sub QBFC_FillInvoiceList(ByRef lstInvoices As System.Windows.Forms.ListBox, ByRef strRefNumber As String, ByRef strFromDateTime As String, ByRef strToDateTime As String, ByRef strDateQueryType As String, ByRef strDateMacro As String, ByRef strCustomerJob As String, ByRef booCustomerWithChildren As Boolean, ByRef strAccount As String, ByRef booAccountWithChildren As Boolean, ByRef strFromRefNumberRange As String, ByRef strToRefNumberRange As String, ByRef strRefNumberPiece As String, ByRef strRefNumberCriteria As String, ByRef strPaidStatus As String)

        Dim strTimeIncluded As String

        lstInvoices.Items.Clear()
        lstInvoices.Refresh()

        Dim objMsgSetRequest As IMsgSetRequest
        objMsgSetRequest = objSessionManager.CreateMsgSetRequest("US", intMajorVersion, intMinorVersion)
        objMsgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

        strSavedRequestCode = "  Dim objMsgSetRequest As IMsgSetRequest" & vbCrLf & "  Set objMsgSetRequest = objSessionManager.CreateMsgSetRequest(""US""," & Str(intMajorVersion) & ", " & Str(intMinorVersion) & ")" & vbCrLf & "  objMsgSetRequest.Attributes.OnError = roeContinue" & vbCrLf & vbCrLf

        Dim objInvoiceQuery As IInvoiceQuery
        objInvoiceQuery = objMsgSetRequest.AppendInvoiceQueryRq

        strSavedRequestCode = strSavedRequestCode & "  Dim objInvoiceQuery As IInvoiceQuery" & vbCrLf & "  Set objInvoiceQuery = objMsgSetRequest.AppendInvoiceQueryRq" & vbCrLf

        objInvoiceQuery.IncludeLineItems.SetValue(True)
        If strRefNumber <> "" Then
            objInvoiceQuery.ORInvoiceQuery.RefNumberList.Add(strRefNumber)
            strSavedRequestCode = strSavedRequestCode & vbCrLf & "  objInvoiceQuery.ORInvoiceQuery.RefNumberList.Add " & strRefNumber
        Else
            'Get the invoice lines so we can put the line count in the invoice information
            strSavedRequestCode = strSavedRequestCode & vbCrLf & "  objInvoiceQuery.IncludeLineItems.SetValue True"

            'Use With statements to reduce the size of our lines
            With objInvoiceQuery.ORInvoiceQuery.InvoiceFilter

                'We're limiting ourselves to the first 30 invoices to avoid too much info
                .MaxReturned.SetValue(30)
                strSavedRequestCode = strSavedRequestCode & vbCrLf & "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.MaxReturned.SetValue 30" & vbCrLf

                If strFromDateTime <> "" Or strToDateTime <> "" Then
                    'We'll either be using a Modified date or a Txn date for filtering
                    If strDateQueryType = "ModifiedDateRangeFilter" Then
                        With .ORDateRangeFilter.ModifiedDateRangeFilter
                            If strFromDateTime <> "" Then
                                If InStr(1, strFromDateTime, ":") Then
                                    strFromDateTime = Replace(strFromDateTime, "T", " ")
                                    .FromModifiedDate.SetValue(CDate(strFromDateTime), True)
                                    strTimeIncluded = "True"
                                Else
                                    .FromModifiedDate.SetValue(CDate(strFromDateTime), False)
                                    strTimeIncluded = "False"
                                End If
                                strSavedRequestCode = strSavedRequestCode & vbCrLf & "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.ORDateRangeFilter.ModifiedDateRangeFilter.FromModifiedDate.SetValue CDate(""" & strFromDateTime & """), " & strTimeIncluded
                            End If

                            If strToDateTime <> "" Then
                                If InStr(1, strToDateTime, ":") Then
                                    strToDateTime = Replace(strToDateTime, "T", " ")
                                    .ToModifiedDate.SetValue(CDate(strToDateTime), True)
                                    strTimeIncluded = "True"
                                Else
                                    .ToModifiedDate.SetValue(CDate(strToDateTime), False)
                                    strTimeIncluded = "False"
                                End If
                                strSavedRequestCode = strSavedRequestCode & vbCrLf & "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.ORDateRangeFilter.ModifiedDateRangeFilter.ToModifiedDate.SetValue CDate(""" & strToDateTime & """), " & strTimeIncluded
                            End If
                        End With '.ORDateRangeFilter.ModifiedDateRangeFilter

                        'Since the to or from date string isn't blank and the date
                        'query type wasn't modified that mean's were using the Txn date filter
                    Else 'strDateQueryType = "ModifiedDateRangeFilter"
                        With .ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter.TxnDateFilter
                            If strFromDateTime <> "" Then
                                .FromTxnDate.SetValue(CDate(strFromDateTime))
                                strSavedRequestCode = strSavedRequestCode & vbCrLf & "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter.TxnDateFilter.FromTxnDate.SetValue CDate(""" & strFromDateTime & """)"
                            End If

                            If strToDateTime <> "" Then
                                .ToTxnDate.SetValue(CDate(strToDateTime))
                                strSavedRequestCode = strSavedRequestCode & vbCrLf & "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter.TxnDateFilter.ToTxnDate.SetValue CDate(""" & strToDateTime & """)"
                            End If
                        End With
                    End If 'strDateQueryType = "ModifiedDate"
                End If 'strFromDateTime <> "" Or strToDateTime <> ""

                If strDateMacro <> "" Then
                    With .ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter
                        .DateMacro.SetAsString(strDateMacro)
                    End With
                    strSavedRequestCode = strSavedRequestCode & vbCrLf & "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter.DateMacro.SetAsString " & strDateMacro
                End If

                If strCustomerJob <> "" Then
                    If booCustomerWithChildren Then
                        .EntityFilter.OREntityFilter.FullNameWithChildren.SetValue(strCustomerJob)
                        strSavedRequestCode = strSavedRequestCode & vbCrLf & "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.EntityFilter.OREntityFilter.FullNameWithChildren.SetValue " & strCustomerJob
                    Else
                        .EntityFilter.OREntityFilter.FullNameList.Add(strCustomerJob)
                        strSavedRequestCode = strSavedRequestCode & vbCrLf & "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.EntityFilter.OREntityFilter.FullNameList.Add " & strCustomerJob
                    End If
                End If

                If strAccount <> "" Then
                    If booAccountWithChildren Then
                        .AccountFilter.ORAccountFilter.FullNameWithChildren.SetValue(strAccount)
                        strSavedRequestCode = strSavedRequestCode & vbCrLf & "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.AccountFilter.ORAccountFilter.FullNameWithChildren.SetValue " & strAccount
                    Else
                        .AccountFilter.ORAccountFilter.FullNameList.Add(strAccount)
                        strSavedRequestCode = strSavedRequestCode & vbCrLf & "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.AccountFilter.ORAccountFilter.FullNameList.Add " & strAccount
                    End If
                End If

                If strFromRefNumberRange <> "" Then
                    .ORRefNumberFilter.RefNumberRangeFilter.FromRefNumber.SetValue(strFromRefNumberRange)
                    strSavedRequestCode = strSavedRequestCode & vbCrLf & "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.ORRefNumberFilter.RefNumberRangeFilter.FromRefNumber.SetValue " & strFromRefNumberRange
                End If

                If strToRefNumberRange <> "" Then
                    .ORRefNumberFilter.RefNumberRangeFilter.ToRefNumber.SetValue(strToRefNumberRange)
                    strSavedRequestCode = strSavedRequestCode & vbCrLf & "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.ORRefNumberFilter.RefNumberRangeFilter.ToRefNumber.SetValue " & strToRefNumberRange
                End If

                If strRefNumberPiece <> "" Then
                    .ORRefNumberFilter.RefNumberFilter.RefNumber.SetValue(strRefNumberPiece)
                    .ORRefNumberFilter.RefNumberFilter.MatchCriterion.SetAsString(strRefNumberCriteria)
                    strSavedRequestCode = strSavedRequestCode & vbCrLf & "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.ORRefNumberFilter.RefNumberFilter.RefNumber.SetValue " & strRefNumberPiece
                    strSavedRequestCode = strSavedRequestCode & vbCrLf & "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.ORRefNumberFilter.RefNumberFilter.MatchCriterion.SetAsString " & strRefNumberCriteria
                End If

                If strPaidStatus <> "" Then
                    .PaidStatus.SetAsString(strPaidStatus)
                    strSavedRequestCode = strSavedRequestCode & vbCrLf & "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.PaidStatus.SetAsString " & strPaidStatus
                End If
            End With 'objInvoiceQuery.ORInvoiceQuery.InvoiceFilter
        End If 'strRefNumber <> ""

        Dim objMsgSetResponse As IMsgSetResponse
        objMsgSetResponse = objSessionManager.DoRequests(objMsgSetRequest)

        Dim objResponse As IResponse
        objResponse = objMsgSetResponse.ResponseList.GetAt(0)

        If objResponse.StatusCode = 1 Then
            lstInvoices.Items.Add("No invoices match the query filter used")
            Exit Sub
        End If

        Dim objInvoiceRetList As IInvoiceRetList
        objInvoiceRetList = objResponse.Detail

        Dim objInvoiceRet As IInvoiceRet
        Dim strItems As String
        Dim strReturnedRefNumber As String
        Dim i As Short
        For i = 0 To objInvoiceRetList.Count - 1
            objInvoiceRet = objInvoiceRetList.GetAt(i)
            If objInvoiceRet.RefNumber Is Nothing Then
                strReturnedRefNumber = "Un-numbered "
            Else
                strReturnedRefNumber = "Invoice " & objInvoiceRet.RefNumber.GetValue
            End If
            strItems = Str(objInvoiceRet.ORInvoiceLineRetList.Count)
            If Len(strItems) = 1 Then strItems = "  " & strItems
            If Len(strItems) = 2 Then strItems = " " & strItems
            lstInvoices.Items.Add(strReturnedRefNumber & "     " & objInvoiceRet.TxnDate.GetValue & "     " & strItems & " items     " & objInvoiceRet.CustomerRef.FullName.GetValue & "     Balance " & objInvoiceRet.BalanceRemaining.GetAsString & "     " & objInvoiceRet.TxnID.GetValue)
        Next

    End Sub


    Public Sub QBFC_GetInvoice(ByRef TxnID As String)

        Dim objMsgSetRequest As IMsgSetRequest
        objMsgSetRequest = objSessionManager.CreateMsgSetRequest("US", intMajorVersion, intMinorVersion)

        Dim objInvoiceQuery As IInvoiceQuery
        objInvoiceQuery = objMsgSetRequest.AppendInvoiceQueryRq

        objInvoiceQuery.ORInvoiceQuery.TxnIDList.Add(TxnID)
        objInvoiceQuery.IncludeLineItems.SetValue(True)

        objInvoiceMsgSetResponse = objSessionManager.DoRequests(objMsgSetRequest)

        Dim objResponse As IResponse
        objResponse = objInvoiceMsgSetResponse.ResponseList.GetAt(0)

        Dim objInvoiceRetList As IInvoiceRetList
        objInvoiceRetList = objResponse.Detail

        objSavedInvoiceRet = objInvoiceRetList.GetAt(0)
    End Sub


    Public Sub QBFC_LoadInvoiceModifyForm()

        frmInvoiceModify.txtTxnID.Text = objSavedInvoiceRet.TxnID.GetValue

        frmInvoiceModify.txtEditSequence.Text = objSavedInvoiceRet.EditSequence.GetValue

        If Not objSavedInvoiceRet.RefNumber Is Nothing Then
            frmInvoiceModify.txtRefNumber.Text = objSavedInvoiceRet.RefNumber.GetValue
        End If

        frmInvoiceModify.txtTxnDate.Text = CStr(objSavedInvoiceRet.TxnDate.GetValue)

        If Not objSavedInvoiceRet.IsPending Is Nothing Then
            If objSavedInvoiceRet.IsPending.GetValue = True Then
                frmInvoiceModify.chkPending.CheckState = System.Windows.Forms.CheckState.Checked 'Checked
            End If
        End If

        If Not objSavedInvoiceRet.IsToBePrinted Is Nothing Then
            If objSavedInvoiceRet.IsToBePrinted.GetValue = True Then
                frmInvoiceModify.chkToBePrinted.CheckState = System.Windows.Forms.CheckState.Checked 'Checked
            End If
        End If

        frmInvoiceModify.cmbCustomer.Text = objSavedInvoiceRet.CustomerRef.FullName.GetValue

        If Not objSavedInvoiceRet.ClassRef Is Nothing Then
            frmInvoiceModify.cmbClass.Text = objSavedInvoiceRet.ClassRef.FullName.GetValue
        End If

        If Not objSavedInvoiceRet.BillAddress Is Nothing Then
            If Not objSavedInvoiceRet.BillAddress.Addr1 Is Nothing Then
                frmInvoiceModify.txtBillAddr1.Text = objSavedInvoiceRet.BillAddress.Addr1.GetValue
            End If

            If Not objSavedInvoiceRet.BillAddress.Addr2 Is Nothing Then
                frmInvoiceModify.txtBillAddr2.Text = objSavedInvoiceRet.BillAddress.Addr2.GetValue
            End If

            If Not objSavedInvoiceRet.BillAddress.Addr3 Is Nothing Then
                frmInvoiceModify.txtBillAddr3.Text = objSavedInvoiceRet.BillAddress.Addr3.GetValue
            End If

            If Not objSavedInvoiceRet.BillAddress.Addr4 Is Nothing Then
                frmInvoiceModify.txtBillAddr4.Text = objSavedInvoiceRet.BillAddress.Addr4.GetValue
            End If

            If Not objSavedInvoiceRet.BillAddress.City Is Nothing Then
                frmInvoiceModify.txtBillCity.Text = objSavedInvoiceRet.BillAddress.City.GetValue
            End If

            If Not objSavedInvoiceRet.BillAddress.State Is Nothing Then
                frmInvoiceModify.txtBillState.Text = objSavedInvoiceRet.BillAddress.State.GetValue
            End If

            If Not objSavedInvoiceRet.BillAddress.PostalCode Is Nothing Then
                frmInvoiceModify.txtBillPostalCode.Text = objSavedInvoiceRet.BillAddress.PostalCode.GetValue
            End If

            If Not objSavedInvoiceRet.BillAddress.Country Is Nothing Then
                frmInvoiceModify.txtBillCountry.Text = objSavedInvoiceRet.BillAddress.Country.GetValue
            End If
        End If 'If Not objSavedInvoiceRet.BillAddress is Nothing

        If Not objSavedInvoiceRet.ShipAddress Is Nothing Then
            If Not objSavedInvoiceRet.ShipAddress.Addr1 Is Nothing Then
                frmInvoiceModify.txtShipAddr1.Text = objSavedInvoiceRet.ShipAddress.Addr1.GetValue
            End If

            If Not objSavedInvoiceRet.ShipAddress.Addr2 Is Nothing Then
                frmInvoiceModify.txtShipAddr2.Text = objSavedInvoiceRet.ShipAddress.Addr2.GetValue
            End If

            If Not objSavedInvoiceRet.ShipAddress.Addr3 Is Nothing Then
                frmInvoiceModify.txtShipAddr3.Text = objSavedInvoiceRet.ShipAddress.Addr3.GetValue
            End If

            If Not objSavedInvoiceRet.ShipAddress.Addr4 Is Nothing Then
                frmInvoiceModify.txtShipAddr4.Text = objSavedInvoiceRet.ShipAddress.Addr4.GetValue
            End If

            If Not objSavedInvoiceRet.ShipAddress.City Is Nothing Then
                frmInvoiceModify.txtShipCity.Text = objSavedInvoiceRet.ShipAddress.City.GetValue
            End If

            If Not objSavedInvoiceRet.ShipAddress.State Is Nothing Then
                frmInvoiceModify.txtShipState.Text = objSavedInvoiceRet.ShipAddress.State.GetValue
            End If

            If Not objSavedInvoiceRet.ShipAddress.PostalCode Is Nothing Then
                frmInvoiceModify.txtShipPostalCode.Text = objSavedInvoiceRet.ShipAddress.PostalCode.GetValue
            End If

            If Not objSavedInvoiceRet.ShipAddress.Country Is Nothing Then
                frmInvoiceModify.txtShipCountry.Text = objSavedInvoiceRet.ShipAddress.Country.GetValue
            End If
        End If 'If Not objSavedInvoiceRet.ShipAddress is Nothing

        If Not objSavedInvoiceRet.ARAccountRef Is Nothing Then
            frmInvoiceModify.cmbARAccount.Text = objSavedInvoiceRet.ARAccountRef.FullName.GetValue
        End If

        If Not objSavedInvoiceRet.TermsRef Is Nothing Then
            frmInvoiceModify.cmbTerms.Text = objSavedInvoiceRet.TermsRef.FullName.GetValue
        End If

        If Not objSavedInvoiceRet.PONumber Is Nothing Then
            frmInvoiceModify.txtPONumber.Text = objSavedInvoiceRet.PONumber.GetValue
        End If

        If Not objSavedInvoiceRet.DueDate Is Nothing Then
            frmInvoiceModify.txtDueDate.Text = CStr(objSavedInvoiceRet.DueDate.GetValue)
        End If

        If Not objSavedInvoiceRet.ShipDate Is Nothing Then
            frmInvoiceModify.txtShipDate.Text = CStr(objSavedInvoiceRet.ShipDate.GetValue)
        End If

        If Not objSavedInvoiceRet.FOB Is Nothing Then
            frmInvoiceModify.txtFOB.Text = objSavedInvoiceRet.FOB.GetValue
        End If

        If Not objSavedInvoiceRet.SalesRepRef Is Nothing Then
            frmInvoiceModify.cmbSalesRep.Text = objSavedInvoiceRet.SalesRepRef.FullName.GetValue
        End If

        If Not objSavedInvoiceRet.ShipMethodRef Is Nothing Then
            frmInvoiceModify.cmbShipMethod.Text = objSavedInvoiceRet.ShipMethodRef.FullName.GetValue
        End If

        If Not objSavedInvoiceRet.ItemSalesTaxRef Is Nothing Then
            frmInvoiceModify.cmbItemSalesTax.Text = objSavedInvoiceRet.ItemSalesTaxRef.FullName.GetValue
        End If

        If Not objSavedInvoiceRet.CustomerSalesTaxCodeRef Is Nothing Then
            frmInvoiceModify.cmbCustTaxCode.Text = objSavedInvoiceRet.CustomerSalesTaxCodeRef.FullName.GetValue
        End If

        If Not objSavedInvoiceRet.CustomerMsgRef Is Nothing Then
            frmInvoiceModify.cmbCustomerMsg.Text = objSavedInvoiceRet.CustomerMsgRef.FullName.GetValue
        End If

        If Not objSavedInvoiceRet.Memo Is Nothing Then
            frmInvoiceModify.txtMemo.Text = objSavedInvoiceRet.Memo.GetValue
        End If
    End Sub


    Public Sub QBFC_LoadInvoiceLineArray(ByRef strLineArray() As String)

        Dim objInvoiceLineRetList As IORInvoiceLineRetList
        objInvoiceLineRetList = objSavedInvoiceRet.ORInvoiceLineRetList


        Dim objLine As IInvoiceLineRet
        Dim objGroupLine As IInvoiceLineGroupRet
        Dim objGroupLines As IInvoiceLineRetList

        Dim i As Short
        Dim j As Short
        Dim k As Short

        j = 0
        For i = 0 To objInvoiceLineRetList.Count - 1

            j = j + 1
            If objInvoiceLineRetList.GetAt(i).InvoiceLineGroupRet Is Nothing Then
                strLineArray(j) = ItemLineInfo((objInvoiceLineRetList.GetAt(i).InvoiceLineRet)) & "Item<spliter>Original"
            Else
                strLineArray(j) = GroupLineInfo((objInvoiceLineRetList.GetAt(i).InvoiceLineGroupRet)) & "Group<spliter>Original"

                objGroupLine = objInvoiceLineRetList.GetAt(i).InvoiceLineGroupRet
                objGroupLines = objGroupLine.InvoiceLineRetList
                If objGroupLines.Count > 0 Then
                    For k = 0 To objGroupLines.Count - 1
                        j = j + 1
                        objLine = objGroupLines.GetAt(k)
                        strLineArray(j) = ItemLineInfo(objLine) & "SubItem<spliter>Original"
                    Next k
                End If 'objGroupLines.length > 0
            End If 'objNode.nodeName = "InvoiceLineRet"
        Next i
    End Sub


    Public Sub QBFC_ModifyInvoice(ByRef strInvoiceChangeString As String)

        PrettyPrintXMLToFile(strInvoiceChangeString, "C:\InvoiceChangeString.txt")

        Dim objMsgSetRequest As IMsgSetRequest
        objMsgSetRequest = objSessionManager.CreateMsgSetRequest("US", intMajorVersion, intMinorVersion)

        Dim objInvoiceMod As IInvoiceMod
        objInvoiceMod = objMsgSetRequest.AppendInvoiceModRq

        Dim objInvoiceLineMod As IInvoiceLineMod
        Dim objInvoiceLineGroupMod As IInvoiceLineGroupMod

        Dim strTemp As String
        strTemp = strInvoiceChangeString
        strTemp = Right(strTemp, Len(strTemp) - 7)
        objInvoiceMod.TxnID.SetValue(Left(strTemp, InStr(1, strTemp, "</") - 1))

        strSavedRequestCode = "  Dim objMsgSetRequest As IMsgSetRequest" & vbCrLf & "  Dim objInvoiceLineMod As IInvoiceLineMod" & vbCrLf & "  Dim objInvoiceLineGroupMod As IInvoiceLineGroupMod" & vbCrLf & vbCrLf & "  Set objMsgSetRequest = objSessionManager.CreateMsgSetRequest(""US""," & Str(intMajorVersion) & ", " & Str(intMinorVersion) & ")" & vbCrLf & vbCrLf & "  objInvoiceMod.TxnID.SetValue " & Left(strTemp, InStr(1, strTemp, "</") - 1) & vbCrLf

        strTemp = Right(strTemp, Len(strTemp) - (InStr(1, strTemp, "<EditS") + 13))
        objInvoiceMod.EditSequence.SetValue(Left(strTemp, InStr(1, strTemp, "</") - 1))

        strSavedRequestCode = strSavedRequestCode & "  objInvoiceMod.EditSequence.SetValue " & Left(strTemp, InStr(1, strTemp, "</") - 1) & vbCrLf
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

            AddElementOrAggregate(strSegment, strStartTag, intTagLength, "", objInvoiceMod, objInvoiceLineMod, objInvoiceLineGroupMod)

            'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        Loop Until String.IsNullOrEmpty(strTemp)

        Dim objMsgSetResponse As IMsgSetResponse
        Dim str2 As String
        str2 = objMsgSetRequest.ToXMLString()

        objMsgSetResponse = objSessionManager.DoRequests(objMsgSetRequest)
        Dim str1 As String
        str1 = objMsgSetResponse.ToXMLString()


        Dim objResponse As IResponse
        objResponse = objMsgSetResponse.ResponseList.GetAt(0)

        If objResponse.StatusCode <> 0 Then
            MsgBox("Error modifying invoice" & vbCrLf & vbCrLf & "Error = " & objResponse.StatusCode & vbCrLf & vbCrLf & "Message = " & objResponse.StatusMessage)
        Else
            MsgBox("Successfully modified invoice")
        End If
    End Sub


    Public Sub QBFC_GetItemInfo(ByRef strItemFullName As String, ByRef strDesc As String, ByRef strRate As String, ByRef strRateOrPercent As String, ByRef strSalesTaxCode As String)

        strDesc = ""
        strRate = ""
        strRateOrPercent = "Rate"
        strSalesTaxCode = ""

        Dim objMsgSetRequest As IMsgSetRequest
        objMsgSetRequest = objSessionManager.CreateMsgSetRequest("US", intMajorVersion, intMinorVersion)
        objMsgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

        Dim objItemQuery As IItemQuery
        objItemQuery = objMsgSetRequest.AppendItemQueryRq
        objItemQuery.ORListQuery.FullNameList.Add(strItemFullName)

        Dim objMsgSetResponse As IMsgSetResponse
        objMsgSetResponse = objSessionManager.DoRequests(objMsgSetRequest)

        Dim objResponse As IResponse
        objResponse = objMsgSetResponse.ResponseList.GetAt(0)

        Dim objItemRetList As IORItemRetList
        objItemRetList = objResponse.Detail

        Dim enItemType As ENORItemRet
        enItemType = objItemRetList.GetAt(0).ortype

        Dim objItemRet As Object
        Select Case enItemType
            Case Is = ENORItemRet.orirItemServiceRet
                objItemRet = objItemRetList.GetAt(0).ItemServiceRet
            Case Is = ENORItemRet.orirItemNonInventoryRet
                objItemRet = objItemRetList.GetAt(0).ItemNonInventoryRet
            Case Is = ENORItemRet.orirItemOtherChargeRet
                objItemRet = objItemRetList.GetAt(0).ItemOtherChargeRet
            Case Is = ENORItemRet.orirItemInventoryRet
                objItemRet = objItemRetList.GetAt(0).ItemInventoryRet
            Case Is = ENORItemRet.orirItemInventoryAssemblyRet
                objItemRet = objItemRetList.GetAt(0).ItemInventoryAssemblyRet
            Case Is = ENORItemRet.orirItemSubtotalRet
                objItemRet = objItemRetList.GetAt(0).ItemSubtotalRet
            Case Is = ENORItemRet.orirItemDiscountRet
                objItemRet = objItemRetList.GetAt(0).ItemDiscountRet
            Case Is = ENORItemRet.orirItemPaymentRet
                objItemRet = objItemRetList.GetAt(0).ItemPaymentRet
            Case Is = ENORItemRet.orirItemSalesTaxRet
                objItemRet = objItemRetList.GetAt(0).ItemSalesTaxRet
            Case Is = ENORItemRet.orirItemGroupRet
                objItemRet = objItemRetList.GetAt(0).ItemGroupRet
        End Select

        If enItemType = ENORItemRet.orirItemServiceRet Or enItemType = ENORItemRet.orirItemNonInventoryRet Or enItemType = ENORItemRet.orirItemOtherChargeRet Then
            'UPGRADE_WARNING: Couldn't resolve default property of object objItemRet.ORSalesPurchase. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If Not objItemRet.ORSalesPurchase.SalesOrPurchase Is Nothing Then
                'UPGRADE_WARNING: Couldn't resolve default property of object objItemRet.ORSalesPurchase. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If Not objItemRet.ORSalesPurchase.SalesOrPurchase.Desc Is Nothing Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object objItemRet.ORSalesPurchase. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    strDesc = objItemRet.ORSalesPurchase.SalesOrPurchase.Desc.getValue
                End If

                'UPGRADE_WARNING: Couldn't resolve default property of object objItemRet.ORSalesPurchase. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If Not objItemRet.ORSalesPurchase.SalesOrPurchase.ORPrice Is Nothing Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object objItemRet.ORSalesPurchase. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If Not objItemRet.ORSalesPurchase.SalesOrPurchase.ORPrice.Price Is Nothing Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object objItemRet.ORSalesPurchase. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        strRate = objItemRet.ORSalesPurchase.SalesOrPurchase.ORPrice.Price.GetAsString
                    Else
                        'UPGRADE_WARNING: Couldn't resolve default property of object objItemRet.ORSalesPurchase. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        strRate = objItemRet.ORSalesPurchase.SalesOrPurchase.ORPrice.PricePercent.GetAsString
                        strRateOrPercent = "RatePercent"
                    End If
                End If
            Else ' Since it isn't SalesOrPurchase it must be SalesAndPurchase
                'UPGRADE_WARNING: Couldn't resolve default property of object objItemRet.ORSalesPurchase. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If Not objItemRet.ORSalesPurchase.SalesAndPurchase.SalesDesc Is Nothing Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object objItemRet.ORSalesPurchase. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    strDesc = objItemRet.ORSalesPurchase.SalesAndPurchase.SalesDesc.getValue
                End If

                'UPGRADE_WARNING: Couldn't resolve default property of object objItemRet.ORSalesPurchase. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If Not objItemRet.ORSalesPurchase.SalesAndPurchase.SalesPrice Is Nothing Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object objItemRet.ORSalesPurchase. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    strRate = objItemRet.ORSalesPurchase.SalesAndPurchase.SalesPrice.GetAsString
                End If
            End If
        ElseIf enItemType = ENORItemRet.orirItemInventoryRet Or enItemType = ENORItemRet.orirItemInventoryAssemblyRet Then
            'UPGRADE_WARNING: Couldn't resolve default property of object objItemRet.SalesDesc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If Not objItemRet.SalesDesc Is Nothing Then
                'UPGRADE_WARNING: Couldn't resolve default property of object objItemRet.SalesDesc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strDesc = objItemRet.SalesDesc.getValue
            End If

            'UPGRADE_WARNING: Couldn't resolve default property of object objItemRet.SalesPrice. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If Not objItemRet.SalesPrice Is Nothing Then
                'UPGRADE_WARNING: Couldn't resolve default property of object objItemRet.SalesPrice. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strRate = objItemRet.SalesPrice.GetAsString
            End If
        Else
            'UPGRADE_WARNING: Couldn't resolve default property of object objItemRet.ItemDesc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If Not objItemRet.ItemDesc Is Nothing Then
                'UPGRADE_WARNING: Couldn't resolve default property of object objItemRet.ItemDesc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strDesc = objItemRet.ItemDesc.getValue
            End If
        End If

        If Not (enItemType = ENORItemRet.orirItemSubtotalRet Or enItemType = ENORItemRet.orirItemPaymentRet Or enItemType = ENORItemRet.orirItemSalesTaxRet Or enItemType = ENORItemRet.orirItemGroupRet) Then
            'UPGRADE_WARNING: Couldn't resolve default property of object objItemRet.SalesTaxCodeRef. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If Not objItemRet.SalesTaxCodeRef Is Nothing Then
                'UPGRADE_WARNING: Couldn't resolve default property of object objItemRet.SalesTaxCodeRef. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strSalesTaxCode = objItemRet.SalesTaxCodeRef.FullName.getValue
            End If
        End If
    End Sub


    Public Sub QBFC_LoadRequest(ByRef strRequestText As String)
        strRequestText = strSavedRequestCode
    End Sub


    Private Function ItemLineInfo(ByRef objInvoiceLine As IInvoiceLineRet) As String

        Dim strLineInfo As String
        Dim strRateOrPercent As String

        strLineInfo = objInvoiceLine.TxnLineID.GetValue & "<spliter>"

        If Not objInvoiceLine.Quantity Is Nothing Then
            strLineInfo = strLineInfo & objInvoiceLine.Quantity.GetAsString
        End If
        strLineInfo = strLineInfo & "<spliter>"

        If Not objInvoiceLine.ItemRef Is Nothing Then
            strLineInfo = strLineInfo & objInvoiceLine.ItemRef.FullName.GetValue
        End If
        strLineInfo = strLineInfo & "<spliter>"

        If Not objInvoiceLine.Desc Is Nothing Then
            strLineInfo = strLineInfo & objInvoiceLine.Desc.GetValue
        End If
        strLineInfo = strLineInfo & "<spliter>"

        If Not objInvoiceLine.ORRate Is Nothing Then
            If Not objInvoiceLine.ORRate.Rate Is Nothing Then
                strLineInfo = strLineInfo & objInvoiceLine.ORRate.Rate.GetAsString
                strRateOrPercent = "Rate"
            Else
                strLineInfo = strLineInfo & objInvoiceLine.ORRate.RatePercent.GetAsString
                strRateOrPercent = "RatePercent"
            End If
        Else
            strRateOrPercent = "Neither"
        End If
        strLineInfo = strLineInfo & "<spliter>"

        If Not objInvoiceLine.Amount Is Nothing Then
            strLineInfo = strLineInfo & objInvoiceLine.Amount.GetAsString
        End If
        strLineInfo = strLineInfo & "<spliter>"

        If Not objInvoiceLine.ClassRef Is Nothing Then
            strLineInfo = strLineInfo & objInvoiceLine.ClassRef.FullName.GetValue
        End If
        strLineInfo = strLineInfo & "<spliter>"

        If Not objInvoiceLine.ServiceDate Is Nothing Then
            strLineInfo = strLineInfo & VB6.Format(objInvoiceLine.ServiceDate.GetValue, "YYYY-MM-DD")
        End If
        strLineInfo = strLineInfo & "<spliter>"

        If Not objInvoiceLine.SalesTaxCodeRef Is Nothing Then
            strLineInfo = strLineInfo & objInvoiceLine.SalesTaxCodeRef.FullName.GetValue
        End If
        strLineInfo = strLineInfo & "<spliter>" & strRateOrPercent & "<spliter><spliter>"


        ItemLineInfo = strLineInfo
    End Function


    Private Function GroupLineInfo(ByRef objInvoiceGroupLine As IInvoiceLineGroupRet) As String

        Dim strLineInfo As String
        Dim strRateOrPercent As String

        strLineInfo = objInvoiceGroupLine.TxnLineID.GetValue & "<spliter>"

        If Not objInvoiceGroupLine.Quantity Is Nothing Then
            strLineInfo = strLineInfo & objInvoiceGroupLine.Quantity.GetAsString
        End If
        strLineInfo = strLineInfo & "<spliter>"

        If Not objInvoiceGroupLine.ItemGroupRef Is Nothing Then
            strLineInfo = strLineInfo & objInvoiceGroupLine.ItemGroupRef.FullName.GetValue
        End If
        strLineInfo = strLineInfo & "<spliter>"

        If Not objInvoiceGroupLine.Desc Is Nothing Then
            strLineInfo = strLineInfo & objInvoiceGroupLine.Desc.GetValue
        End If
        strLineInfo = strLineInfo & "<spliter><spliter><spliter><spliter>"

        If Not objInvoiceGroupLine.ServiceDate Is Nothing Then
            strLineInfo = strLineInfo & VB6.Format(objInvoiceGroupLine.ServiceDate.GetValue, "YYYY-MM-DD")
        End If
        strLineInfo = strLineInfo & "<spliter><spliter><spliter><spliter>"

        GroupLineInfo = strLineInfo
    End Function


    Private Sub AddElementOrAggregate(ByRef strElOrAggString As String, ByRef strStartTag As String, ByRef intTagLength As Short, ByRef strAggregateType As String, ByRef objInvoiceMod As IInvoiceMod, ByRef objCurrentModLine As IInvoiceLineMod, ByRef objCurrentModGroupLine As IInvoiceLineGroupMod)

        Dim objInvoiceLineMod As IInvoiceLineMod
        Dim objInvoiceLineGroupMod As IInvoiceLineGroupMod
        Dim strSegment As String
        Dim strInnerValue As String

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

        If String.IsNullOrEmpty(strInnerStartTag) Then
            Select Case strAggregateType
                Case Is = ""
                    Select Case strTagName
                        Case Is = "TxnID"
                            objInvoiceMod.TxnID.SetValue(strValue)
                            strSavedRequestCode = strSavedRequestCode &
              "  objInvoiceMod.TxnID.SetValue " & strValue & vbCrLf
                        Case Is = "EditSequence"
                            objInvoiceMod.EditSequence.SetValue(strValue)
                            strSavedRequestCode = strSavedRequestCode &
              "  objInvoiceMod.EditSequence.SetValue " & strValue & vbCrLf
                        Case Is = "TxnDate"
                            objInvoiceMod.TxnDate.SetValue(CDate(strValue))
                            strSavedRequestCode = strSavedRequestCode &
              "  objInvoiceMod.TxnDate.SetValue CDate(""" & strValue & """)" & vbCrLf
                        Case Is = "RefNumber"
                            objInvoiceMod.RefNumber.SetValue(strValue)
                            strSavedRequestCode = strSavedRequestCode &
              "  objInvoiceMod.RefNumber.SetValue " & strValue & vbCrLf
                        Case Is = "IsPending"
                            objInvoiceMod.IsPending.SetValue(strValue = "1")
                            strSavedRequestCode = strSavedRequestCode &
                              "  objInvoiceMod.IsPending.SetValue " & (strValue = "1") & vbCrLf
                        Case Is = "PONumber"
                            objInvoiceMod.PONumber.SetValue(strValue)
                            strSavedRequestCode = strSavedRequestCode &
              "  objInvoiceMod.PONumber.SetValue " & strValue & vbCrLf
                        Case Is = "DueDate"
                            objInvoiceMod.DueDate.SetValue(CDate(strValue))
                            strSavedRequestCode = strSavedRequestCode &
              "  objInvoiceMod.DueDate.SetValue CDate(""" & strValue & """)" & vbCrLf
                        Case Is = "FOB"
                            objInvoiceMod.FOB.SetValue(strValue)
                            strSavedRequestCode = strSavedRequestCode &
              "  objInvoiceMod.FOB.SetValue " & strValue & vbCrLf
                        Case Is = "ShipDate"
                            objInvoiceMod.ShipDate.SetValue(CDate(strValue))
                            strSavedRequestCode = strSavedRequestCode &
              "  objInvoiceMod.ShipDate.SetValue CDate(""" & strValue & """)" & vbCrLf
                        Case Is = "Memo"
                            objInvoiceMod.Memo.SetValue(strValue)
                            strSavedRequestCode = strSavedRequestCode &
              "  objInvoiceMod.Memo.SetValue " & strValue & vbCrLf
                        Case Is = "IsToBePrinted"
                            objInvoiceMod.IsToBePrinted.SetValue(strValue = "1")
                            strSavedRequestCode = strSavedRequestCode &
                              "  objInvoiceMod.IsToBePrinted.SetValue " & (strValue = "1") & vbCrLf
                    End Select
                Case Is = "BillAddress"
                    Select Case strTagName
                        Case Is = "Addr1"
                            objInvoiceMod.BillAddress.Addr1.SetValue(strValue)
                            strSavedRequestCode = strSavedRequestCode &
              "  objInvoiceMod.BillAddress.Addr1.SetValue " & strValue & vbCrLf
                        Case Is = "Addr2"
                            objInvoiceMod.BillAddress.Addr2.SetValue(strValue)
                            strSavedRequestCode = strSavedRequestCode &
              "  objInvoiceMod.BillAddress.Addr2.SetValue " & strValue & vbCrLf
                        Case Is = "Addr3"
                            objInvoiceMod.BillAddress.Addr3.SetValue(strValue)
                            strSavedRequestCode = strSavedRequestCode &
              "  objInvoiceMod.BillAddress.Addr3.SetValue " & strValue & vbCrLf
                        Case Is = "Addr4"
                            objInvoiceMod.BillAddress.Addr4.SetValue(strValue)
                            strSavedRequestCode = strSavedRequestCode &
              "  objInvoiceMod.BillAddress.Addr4.SetValue " & strValue & vbCrLf
                        Case Is = "City"
                            objInvoiceMod.BillAddress.City.SetValue(strValue)
                            strSavedRequestCode = strSavedRequestCode &
              "  objInvoiceMod.BillAddress.City.SetValue " & strValue & vbCrLf
                        Case Is = "State"
                            objInvoiceMod.BillAddress.State.SetValue(strValue)
                            strSavedRequestCode = strSavedRequestCode &
              "  objInvoiceMod.BillAddress.State.SetValue " & strValue & vbCrLf
                        Case Is = "PostalCode"
                            objInvoiceMod.BillAddress.PostalCode.SetValue(strValue)
                            strSavedRequestCode = strSavedRequestCode &
              "  objInvoiceMod.BillAddress.PostalCode.SetValue " & strValue & vbCrLf
                        Case Is = "Country"
                            objInvoiceMod.BillAddress.Country.SetValue(strValue)
                            strSavedRequestCode = strSavedRequestCode &
              "  objInvoiceMod.BillAddress.Country.SetValue " & strValue & vbCrLf
                    End Select
                Case Is = "ShipAddress"
                    Select Case strTagName
                        Case Is = "Addr1"
                            objInvoiceMod.ShipAddress.Addr1.SetValue(strValue)
                            strSavedRequestCode = strSavedRequestCode &
              "  objInvoiceMod.ShipAddress.Addr1.SetValue " & strValue & vbCrLf
                        Case Is = "Addr2"
                            objInvoiceMod.ShipAddress.Addr2.SetValue(strValue)
                            strSavedRequestCode = strSavedRequestCode &
              "  objInvoiceMod.ShipAddress.Addr2.SetValue " & strValue & vbCrLf
                        Case Is = "Addr3"
                            objInvoiceMod.ShipAddress.Addr3.SetValue(strValue)
                            strSavedRequestCode = strSavedRequestCode &
              "  objInvoiceMod.ShipAddress.Addr3.SetValue " & strValue & vbCrLf
                        Case Is = "Addr4"
                            objInvoiceMod.ShipAddress.Addr4.SetValue(strValue)
                            strSavedRequestCode = strSavedRequestCode &
              "  objInvoiceMod.ShipAddress.Addr4.SetValue " & strValue & vbCrLf
                        Case Is = "City"
                            objInvoiceMod.ShipAddress.City.SetValue(strValue)
                            strSavedRequestCode = strSavedRequestCode &
              "  objInvoiceMod.ShipAddress.City.SetValue " & strValue & vbCrLf
                        Case Is = "State"
                            objInvoiceMod.ShipAddress.State.SetValue(strValue)
                            strSavedRequestCode = strSavedRequestCode &
              "  objInvoiceMod.ShipAddress.State.SetValue " & strValue & vbCrLf
                        Case Is = "PostalCode"
                            objInvoiceMod.ShipAddress.PostalCode.SetValue(strValue)
                            strSavedRequestCode = strSavedRequestCode &
              "  objInvoiceMod.ShipAddress.PostalCode.SetValue " & strValue & vbCrLf
                        Case Is = "Country"
                            objInvoiceMod.ShipAddress.Country.SetValue(strValue)
                            strSavedRequestCode = strSavedRequestCode &
              "  objInvoiceMod.ShipAddress.Country.SetValue " & strValue & vbCrLf
                    End Select
                Case Is = "InvoiceLineGroupMod"
                    Select Case strTagName
                        Case Is = "TxnLineID"
                            objCurrentModGroupLine.TxnLineID.SetValue(strValue)
                            strSavedRequestCode = strSavedRequestCode &
              "  objInvoiceLineGroupMod.TxnLineID.SetValue " & strValue & vbCrLf
                        Case Is = "Quantity"
                            objCurrentModGroupLine.Quantity.SetAsString(strValue)
                            strSavedRequestCode = strSavedRequestCode &
              "  objInvoiceLineGroupMod.Quantity.SetAsString " & strValue & vbCrLf
                        Case Is = "ServiceDate"
                            objCurrentModGroupLine.ServiceDate.setValue(CDate(strValue))
                            strSavedRequestCode = strSavedRequestCode &
              "  objInvoiceLineGroupMod.ServiceDate.SetValue CDate(""" & strValue & """)" & vbCrLf
                    End Select
                Case Else
                    If strAggregateType = "InvoiceLineMod" Or strAggregateType = "SubItem" Then

                        Select Case strTagName
                            Case Is = "TxnLineID"
                                objCurrentModLine.TxnLineID.SetValue(strValue)
                                strSavedRequestCode = strSavedRequestCode &
                "  objInvoiceLineMod.TxnLineID.SetValue " & strValue & vbCrLf
                            Case Is = "Desc"
                                objCurrentModLine.Desc.SetValue(strValue)
                                strSavedRequestCode = strSavedRequestCode &
                "  objInvoiceLineMod.Desc.SetValue " & strValue & vbCrLf
                            Case Is = "Quantity"
                                objCurrentModLine.Quantity.SetAsString(strValue)
                                strSavedRequestCode = strSavedRequestCode &
                "  objInvoiceLineMod.Quantity.SetAsString " & strValue & vbCrLf
                            Case Is = "RatePercent"
                                objCurrentModLine.ORRatePriceLevel.RatePercent.SetAsString(strValue)
                                strSavedRequestCode = strSavedRequestCode &
                "  objInvoiceLineMod.RatePercent.SetAsString " & strValue & vbCrLf
                            Case Is = "Rate"
                                objCurrentModLine.ORRatePriceLevel.Rate.SetAsString(strValue)
                                strSavedRequestCode = strSavedRequestCode &
                "  objInvoiceLineMod.Rate.SetAsString " & strValue & vbCrLf
                            Case Is = "Amount"
                                objCurrentModLine.Amount.SetAsString(strValue)
                                strSavedRequestCode = strSavedRequestCode &
                "  objInvoiceLineMod.Amount.SetAsString " & strValue & vbCrLf
                            Case Is = "ServiceDate"
                                objCurrentModLine.ServiceDate.SetValue(CDate(strValue))
                                strSavedRequestCode = strSavedRequestCode &
                "  objInvoiceLineMod.ServiceDate.SetValue CDate(""" & strValue & """)" & vbCrLf
                        End Select
                    End If
            End Select
        Else 'For If strInnerStartTag = Empty Then
            strInnerValue = Left(strValue, Len(strValue) - (intInnerTagLength + 1))
            strInnerValue = Right(strInnerValue, Len(strInnerValue) - intInnerTagLength)

            Select Case strAggregateType
                Case Is = ""
                    'We'll get this for the aggregates in the InvoiceMod, so look for
                    'InvoiceLineMod, InvoiceLineGroupMod or a reference aggretate and
                    'use what's between the <FullName> and </FullName> tags

                    Select Case strTagName
                        Case Is = "InvoiceLineMod"
                            objCurrentModLine = objInvoiceMod.ORInvoiceLineModList.Append.InvoiceLineMod
                            strSavedRequestCode = strSavedRequestCode & vbCrLf &
              "  Set objInvoiceLineMod = objInvoiceMod.ORInvoiceLineModList.Append.InvoiceLineMod" & vbCrLf
                        Case Is = "InvoiceLineGroupMod"
                            objCurrentModGroupLine = objInvoiceMod.ORInvoiceLineModList.Append.InvoiceLineGroupMod
                            strSavedRequestCode = strSavedRequestCode & vbCrLf &
              "  Set objInvoiceLineGroupMod = objInvoiceMod.ORInvoiceLineModList.Append.InvoiceLineGroupMod" & vbCrLf
                        Case Is = "CustomerRef"
                            objInvoiceMod.CustomerRef.FullName.SetValue(strInnerValue)
                            strSavedRequestCode = strSavedRequestCode &
              "  objInvoiceMod.CustomerRef.FullName.SetValue " & strInnerValue & vbCrLf
                        Case Is = "ClassRef"
                            objInvoiceMod.ClassRef.FullName.SetValue(strInnerValue)
                            strSavedRequestCode = strSavedRequestCode &
              "  objInvoiceMod.ClassRef.FullName.SetValue " & strInnerValue & vbCrLf
                        Case Is = "ARAccountRef"
                            objInvoiceMod.ARAccountRef.FullName.SetValue(strInnerValue)
                            strSavedRequestCode = strSavedRequestCode &
              "  objInvoiceMod.ARAccountRef.FullName.SetValue " & strInnerValue & vbCrLf
                        Case Is = "TermsRef"
                            objInvoiceMod.TermsRef.FullName.SetValue(strInnerValue)
                            strSavedRequestCode = strSavedRequestCode &
              "  objInvoiceMod.TermsRef.FullName.SetValue " & strInnerValue & vbCrLf
                        Case Is = "SalesRepRef"
                            objInvoiceMod.SalesRepRef.FullName.SetValue(strInnerValue)
                            strSavedRequestCode = strSavedRequestCode &
              "  objInvoiceMod.SalesRepRef.FullName.SetValue " & strInnerValue & vbCrLf
                        Case Is = "ShipMethodRef"
                            objInvoiceMod.ShipMethodRef.FullName.SetValue(strInnerValue)
                            strSavedRequestCode = strSavedRequestCode &
              "  objInvoiceMod.ShipMethodRef.FullName.SetValue " & strInnerValue & vbCrLf
                        Case Is = "ItemSalesTaxRef"
                            objInvoiceMod.ItemSalesTaxRef.FullName.SetValue(strInnerValue)
                            strSavedRequestCode = strSavedRequestCode &
              "  objInvoiceMod.ItemSalesTaxRef.FullName.SetValue " & strInnerValue & vbCrLf
                        Case Is = "CustomerMsgRef"
                            objInvoiceMod.CustomerMsgRef.FullName.SetValue(strInnerValue)
                            strSavedRequestCode = strSavedRequestCode &
              "  objInvoiceMod.CustomerMsgRef.FullName.SetValue " & strInnerValue & vbCrLf
                        Case Is = "CustomerSalesTaxCodeRef"
                            objInvoiceMod.CustomerSalesTaxCodeRef.FullName.SetValue(strInnerValue)
                            strSavedRequestCode = strSavedRequestCode &
              "  objInvoiceMod.CustomerSalesTaxCodeRef.FullName.SetValue " & strInnerValue & vbCrLf
                    End Select
                Case Is = "InvoiceLineGroupMod"
                    Select Case strTagName
                        Case Is = "ItemGroupRef"
                            objCurrentModGroupLine.ItemGroupRef.FullName.SetValue(strInnerValue)
                            strSavedRequestCode = strSavedRequestCode &
              "  objInvoiceLineGroupMod.ItemGroupRef.FullName.SetValue " & strInnerValue & vbCrLf
                        Case Is = "InvoiceLineMod"
                            objCurrentModLine = objCurrentModGroupLine.InvoiceLineModList.Append
                            strSavedRequestCode = strSavedRequestCode & vbCrLf &
              "  Set objInvoiceLineMod = objInvoiceLineGroupMod.InvoiceLineModList.Append " & vbCrLf

                            Do While String.IsNullOrEmpty(strValue)
                                strSegment = Left(strValue, InStr(1, strValue, strInnerEndTag) + intInnerTagLength)
                                strValue = Right(strValue, Len(strValue) - Len(strSegment))

                                AddElementOrAggregate(strSegment, strInnerStartTag, intInnerTagLength, "SubItem", objInvoiceMod, objCurrentModLine, objCurrentModGroupLine)

                                If String.IsNullOrEmpty(strValue) Then GetTags(strValue, strInnerStartTag, strInnerEndTag, intInnerTagLength)
                            Loop
                    End Select
                Case Else

                    If strAggregateType = "InvoiceLineMod" Or strAggregateType = "SubItem" Then
                        Select Case strTagName
                            Case Is = "ItemRef"
                                objCurrentModLine.ItemRef.FullName.SetValue(strInnerValue)
                                strSavedRequestCode = strSavedRequestCode &
                "  objInvoiceLineMod.ItemRef.FullName.SetValue " & strInnerValue & vbCrLf
                            Case Is = "ClassRef"
                                objCurrentModLine.ClassRef.FullName.SetValue(strInnerValue)
                                strSavedRequestCode = strSavedRequestCode &
                "  objInvoiceLineMod.ClassRef.FullName.SetValue " & strInnerValue & vbCrLf
                            Case Is = "OverrideAccountRef"
                                objCurrentModLine.OverrideAccountRef.FullName.setValue(strInnerValue)
                                strSavedRequestCode = strSavedRequestCode &
                "  objInvoiceLineMod.OverrideAccountRef.FullName.SetValue " & strInnerValue & vbCrLf
                            Case Is = "SalesTaxCodeRef"
                                objCurrentModLine.SalesTaxCodeRef.FullName.SetValue(strInnerValue)
                                strSavedRequestCode = strSavedRequestCode &
                "  objInvoiceLineMod.SalesTaxCodeRef.FullName.SetValue " & strInnerValue & vbCrLf
                        End Select
                    End If
            End Select

            'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            Do While Not String.IsNullOrEmpty(strValue)
                strSegment = Left(strValue, InStr(1, strValue, strInnerEndTag) + intInnerTagLength)
                strValue = Right(strValue, Len(strValue) - Len(strSegment))

                AddElementOrAggregate(strSegment, strInnerStartTag, intInnerTagLength, strTagName, objInvoiceMod, objCurrentModLine, objCurrentModGroupLine)

                'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                If Not String.IsNullOrEmpty(strValue) Then GetTags(strValue, strInnerStartTag, strInnerEndTag, intInnerTagLength)
            Loop
        End If
    End Sub
End Module