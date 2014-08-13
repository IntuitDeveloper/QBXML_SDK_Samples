Attribute VB_Name = "modQBFC"
'----------------------------------------------------------
' Copyright © 2003-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/rdmgr/?ID=100
'
' Updated to QBFC 12.0 and fixed setting max QBXML version: 09/2012
'----------------------------------------------------------

Option Explicit

Dim objSessionManager As New QBSessionManager

'Set the max versions based on the version of QBFC we're using.  If we
'were to run against QB 2004 the supported version would be 3.0, but
'QBFC2_1 can only understand up to version 2.1, so we use these constants
'to make sure we don't try to create version 3.x messages with a QBFC that
'can't create them
Const MaxMajorVersion = 2
Const MaxMinorVersion = 1

Dim booSessionBegun As Boolean

Dim intMajorVersion As Integer
Dim intMinorVersion As Integer

'Declare global variables for keeping the Invoice Query Response object
'and the Invoice Ret object so we can access them later
Dim objInvoiceMsgSetResponse As IMsgSetResponse
Dim objSavedInvoiceRet As IInvoiceRet

Dim strSavedRequestCode As String


Public Sub QBFC_OpenConnectionBeginSession()

  On Error GoTo Errs
  
  objSessionManager.OpenConnection "", "IDN Invoice Modify Sample Application"
  objSessionManager.BeginSession "", omDontCare
  booSessionBegun = True
  Exit Sub

Errs:
  If Err.Number = &H80040416 Then
    MsgBox "You must have QuickBooks running with a company" & vbCrLf & _
           "file open to use this program."
    objSessionManager.CloseConnection
    End
  ElseIf Err.Number = &H80040422 Then
    MsgBox "This QuickBooks company file is open in single user mode and" & vbCrLf & _
           "another application is already accessing it.  Please exit the" & vbCrLf & _
           "other application and run this application again."
    objSessionManager.CloseConnection
    End
  ElseIf Err.Number = &H80040308 Then
    MsgBox _
      "It appears that the qbXML Request Processor has not" & vbCrLf & _
      "been installed, indicating QuickBooks 2002 or later" & vbCrLf & _
      "may not have been installed.  Please run this sample" & vbCrLf & _
      "after installing QuickBooks 2003 and running the Upgrade."
  ElseIf Err.Number = &H8007007E Then
'    If QBFC2CA_IsntInstalled Then
      MsgBox _
        "QBFC2 isn't installed.  You need QBFC2 or QBFC2CA installed to" & vbCrLf & _
        "use QBFC with this sample program."
      End
'    Else
'      QBFCCA_OpenConnectionBeginSession
'      SetImplementation "QBFCCA"
'    End If
  Else
    MsgBox "QBFC_OpenConnectionBeginSession" & vbCrLf & _
      Err.Number & vbCrLf & Hex(Err.Number) & vbCrLf & _
      Err.Description
    End
  End If
End Sub


Public Sub QBFC_EndSessionCloseConnection()

  If booSessionBegun Then
    objSessionManager.EndSession
    objSessionManager.CloseConnection
  End If
End Sub

Function QBFC_MaxVersionSupported() As String

Dim strVersions() As String

  strVersions = objSessionManager.QBXMLVersionsForSession
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
    MsgBox "The Canadian version of QBFC does not support version 2.1 " & _
      "messages.  Exiting."
    End
  End If
  
  'Now make sure that the version of QBFC installed supports the
  'maximum version of qbXML that QuickBooks can handle
  Dim intQBFCMajorVersion As Integer
  Dim intQBFCMinorVersion As Integer
  Dim enumReleaseLevel As ENReleaseLevel
  Dim intReleaseNumber As Integer
  
  objSessionManager.GetVersion _
    intQBFCMajorVersion, intQBFCMinorVersion, _
    enumReleaseLevel, intReleaseNumber
  
  Dim versionArray() As String
  versionArray = Split(strVersions(UBound(strVersions)), ".")
  
  intMajorVersion = Int(versionArray(0))
  intMinorVersion = Int(versionArray(1))
  
  If intMajorVersion > MaxMajorVersion Then
    intMajorVersion = MaxMajorVersion
    intMinorVersion = MaxMinorVersion
  End If
  
  QBFC_MaxVersionSupported = Trim(Str(intMajorVersion)) & "." & _
                             Trim(Str(intMinorVersion))
End Function


Public Sub QBFC_FillComboBox(cmbComboBox As ComboBox, _
                             strQueryType As String, _
                             strNameElement As String, _
                             strFilter As String, _
                             booMarkGroupItems As Boolean)

  'Clear the combo box
  cmbComboBox.Clear
  
  Dim strSplits() As String
  strSplits = Split(strQueryType, ",")
  
  Dim strNameElementSplits() As String
  strNameElementSplits = Split(strNameElement, ",")
  
  Dim objMsgSetRequest As IMsgSetRequest
  Dim objMsgSetResponse As IMsgSetResponse
  Dim objResponse As IResponse
  
  Dim objQuery
  Dim objRetList
  Dim objRet
      
  Dim numItems As Integer
  
  Dim i As Integer
  Dim j As Integer
  For i = 0 To UBound(strSplits)
  
    Set objMsgSetRequest = objSessionManager.CreateMsgSetRequest("US", intMajorVersion, intMinorVersion)
    objMsgSetRequest.Attributes.OnError = roeContinue
    
    Select Case strSplits(i)
      Case Is = "Account"
        Set objQuery = objMsgSetRequest.AppendAccountQueryRq
      Case Is = "Class"
        Set objQuery = objMsgSetRequest.AppendClassQueryRq
      Case Is = "Customer"
        Set objQuery = objMsgSetRequest.AppendCustomerQueryRq
      Case Is = "CustomerMsg"
        Set objQuery = objMsgSetRequest.AppendCustomerMsgQueryRq
      Case Is = "ItemService"
        Set objQuery = objMsgSetRequest.AppendItemServiceQueryRq
      Case Is = "ItemInventory"
        Set objQuery = objMsgSetRequest.AppendItemInventoryQueryRq
      Case Is = "ItemInventoryAssembly"
        Set objQuery = objMsgSetRequest.AppendItemInventoryAssemblyQueryRq
      Case Is = "ItemNonInventory"
        Set objQuery = objMsgSetRequest.AppendItemNonInventoryQueryRq
      Case Is = "ItemOtherCharge"
        Set objQuery = objMsgSetRequest.AppendItemOtherChargeQueryRq
      Case Is = "ItemSubtotal"
        Set objQuery = objMsgSetRequest.AppendItemSubtotalQueryRq
      Case Is = "ItemGroup"
        Set objQuery = objMsgSetRequest.AppendItemGroupQueryRq
      Case Is = "ItemDiscount"
        Set objQuery = objMsgSetRequest.AppendItemDiscountQueryRq
      Case Is = "ItemPayment"
        Set objQuery = objMsgSetRequest.AppendItemPaymentQueryRq
      Case Is = "ItemSalesTax"
        Set objQuery = objMsgSetRequest.AppendItemSalesTaxQueryRq
      Case Is = "ItemSalesTaxGroup"
        Set objQuery = objMsgSetRequest.AppendItemSalesTaxGroupQueryRq
      Case Is = "SalesRep"
        Set objQuery = objMsgSetRequest.AppendSalesRepQueryRq
      Case Is = "SalesTaxCode"
        Set objQuery = objMsgSetRequest.AppendSalesTaxCodeQueryRq
      Case Is = "ShipMethod"
        Set objQuery = objMsgSetRequest.AppendShipMethodQueryRq
      Case Is = "StandardTerms"
        Set objQuery = objMsgSetRequest.AppendStandardTermsQueryRq
      Case Else
        MsgBox "Unknown type " & strSplits(i) & " passed to QBFC_FillComboBox"
    End Select
    
    If strFilter <> Empty Then
      Dim strTemp As String
      Dim strStartTag As String
      Dim strEndTag As String
      Dim strValue As String
      Dim intTagLength As Integer
      
      strTemp = strFilter
      Do While strTemp <> Empty
        GetTags strTemp, strStartTag, strEndTag, intTagLength
        
        strValue = Left(strTemp, InStr(1, strTemp, strEndTag) - 1)
        strValue = Right(strValue, Len(strValue) - intTagLength)
        strTemp = Right(strTemp, Len(strTemp) - (InStr(1, strTemp, strEndTag) + intTagLength))
      
        Select Case strStartTag
          Case Is = "<AccountType>"
            objQuery.ORAccountListQuery.AccountListFilter.AccountTypeList.AddAsString strValue
          Case Else
            MsgBox "Unknown filter " & strStartTag & " in QBFC_FillComboBox"
        End Select
      Loop
    End If
    
    Set objMsgSetResponse = objSessionManager.DoRequests(objMsgSetRequest)
    
    Set objResponse = objMsgSetResponse.ResponseList.GetAt(0)
    
    If objResponse.StatusCode <> 0 Then
      If objResponse.StatusCode <> 1 Then
        MsgBox "Status Code " & objResponse.StatusCode & _
               " on call to QBFC_FillComboBox" & vbCrLf & _
               " for " & strSplits(i) & " list items"
      End If
    Else
    
      Set objRetList = objResponse.Detail
      numItems = objRetList.Count
  
      If UBound(strNameElementSplits) > 0 Then
        strNameElement = strNameElementSplits(i)
      End If
    
      For j = 0 To numItems - 1
        Set objRet = objRetList.GetAt(j)
        If strSplits(i) = "ItemGroup" And booMarkGroupItems Then
          cmbComboBox.AddItem objRet.Name.getValue & _
            " - Group Item"
        Else
          If strNameElement = "FullName" Then
            cmbComboBox.AddItem objRet.FullName.getValue
          Else
            cmbComboBox.AddItem objRet.Name.getValue
          End If
        End If
      Next j
    End If ' If objResponse.StatusCode <> 0
  Next i
End Sub


Public Sub QBFC_FillInvoiceList(lstInvoices As ListBox, _
                                strRefNumber As String, _
                                strFromDateTime As String, _
                                strToDateTime As String, _
                                strDateQueryType As String, _
                                strDateMacro As String, _
                                strCustomerJob As String, _
                                booCustomerWithChildren As Boolean, _
                                strAccount As String, _
                                booAccountWithChildren As Boolean, _
                                strFromRefNumberRange As String, _
                                strToRefNumberRange As String, _
                                strRefNumberPiece As String, _
                                strRefNumberCriteria As String, _
                                strPaidStatus As String)

  Dim strTimeIncluded As String
  
  lstInvoices.Clear
  lstInvoices.Refresh
  
  Dim objMsgSetRequest As IMsgSetRequest
  Set objMsgSetRequest = objSessionManager.CreateMsgSetRequest("US", intMajorVersion, intMinorVersion)
  objMsgSetRequest.Attributes.OnError = roeContinue
  
  strSavedRequestCode = _
    "  Dim objMsgSetRequest As IMsgSetRequest" & vbCrLf & _
    "  Set objMsgSetRequest = objSessionManager.CreateMsgSetRequest(""US""," & Str(intMajorVersion) & ", " & Str(intMinorVersion) & ")" & vbCrLf & _
    "  objMsgSetRequest.Attributes.OnError = roeContinue" & vbCrLf & vbCrLf
  
  Dim objInvoiceQuery As IInvoiceQuery
  Set objInvoiceQuery = objMsgSetRequest.AppendInvoiceQueryRq
  
  strSavedRequestCode = strSavedRequestCode & _
    "  Dim objInvoiceQuery As IInvoiceQuery" & vbCrLf & _
    "  Set objInvoiceQuery = objMsgSetRequest.AppendInvoiceQueryRq" & vbCrLf

  objInvoiceQuery.IncludeLineItems.setValue True
  If strRefNumber <> "" Then
    objInvoiceQuery.ORInvoiceQuery.RefNumberList.Add strRefNumber
    strSavedRequestCode = strSavedRequestCode & vbCrLf & _
      "  objInvoiceQuery.ORInvoiceQuery.RefNumberList.Add " & strRefNumber
  Else
    'Get the invoice lines so we can put the line count in the invoice information
    strSavedRequestCode = strSavedRequestCode & vbCrLf & _
      "  objInvoiceQuery.IncludeLineItems.SetValue True"
      
    'Use With statements to reduce the size of our lines
    With objInvoiceQuery.ORInvoiceQuery.InvoiceFilter
    
      'We're limiting ourselves to the first 30 invoices to avoid too much info
      .MaxReturned.setValue 30
    strSavedRequestCode = strSavedRequestCode & vbCrLf & _
      "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.MaxReturned.SetValue 30" & vbCrLf
      
      If strFromDateTime <> "" Or strToDateTime <> "" Then
        'We'll either be using a Modified date or a Txn date for filtering
        If strDateQueryType = "ModifiedDateRangeFilter" Then
          With .ORDateRangeFilter.ModifiedDateRangeFilter
            If strFromDateTime <> "" Then
              If InStr(1, strFromDateTime, ":") Then
                strFromDateTime = Replace(strFromDateTime, "T", " ")
                .FromModifiedDate.setValue CDate(strFromDateTime), True
                strTimeIncluded = "True"
              Else
                .FromModifiedDate.setValue CDate(strFromDateTime), False
                strTimeIncluded = "False"
              End If
              strSavedRequestCode = strSavedRequestCode & vbCrLf & _
                "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.ORDateRangeFilter.ModifiedDateRangeFilter.FromModifiedDate.SetValue CDate(""" & strFromDateTime & """), " & strTimeIncluded
            End If
      
            If strToDateTime <> "" Then
              If InStr(1, strToDateTime, ":") Then
                strToDateTime = Replace(strToDateTime, "T", " ")
                .ToModifiedDate.setValue CDate(strToDateTime), True
                strTimeIncluded = "True"
              Else
                .ToModifiedDate.setValue CDate(strToDateTime), False
                strTimeIncluded = "False"
              End If
              strSavedRequestCode = strSavedRequestCode & vbCrLf & _
                "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.ORDateRangeFilter.ModifiedDateRangeFilter.ToModifiedDate.SetValue CDate(""" & strToDateTime & """), " & strTimeIncluded
            End If
          End With '.ORDateRangeFilter.ModifiedDateRangeFilter

        'Since the to or from date string isn't blank and the date
        'query type wasn't modified that mean's were using the Txn date filter
        Else 'strDateQueryType = "ModifiedDateRangeFilter"
          With .ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter.TxnDateFilter
            If strFromDateTime <> "" Then
              .FromTxnDate.setValue CDate(strFromDateTime)
              strSavedRequestCode = strSavedRequestCode & vbCrLf & _
                "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter.TxnDateFilter.FromTxnDate.SetValue CDate(""" & strFromDateTime & """)"
            End If
      
            If strToDateTime <> "" Then
              .ToTxnDate.setValue CDate(strToDateTime)
              strSavedRequestCode = strSavedRequestCode & vbCrLf & _
                "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter.TxnDateFilter.ToTxnDate.SetValue CDate(""" & strToDateTime & """)"
            End If
          End With
        End If 'strDateQueryType = "ModifiedDate"
      End If 'strFromDateTime <> "" Or strToDateTime <> ""
    
      If strDateMacro <> "" Then
        With .ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter
          .DateMacro.SetAsString strDateMacro
        End With
        strSavedRequestCode = strSavedRequestCode & vbCrLf & _
          "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter.DateMacro.SetAsString " & strDateMacro
      End If
    
      If strCustomerJob <> "" Then
        If booCustomerWithChildren Then
          .EntityFilter.OREntityFilter.FullNameWithChildren.setValue strCustomerJob
          strSavedRequestCode = strSavedRequestCode & vbCrLf & _
            "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.EntityFilter.OREntityFilter.FullNameWithChildren.SetValue " & strCustomerJob
        Else
          .EntityFilter.OREntityFilter.FullNameList.Add strCustomerJob
          strSavedRequestCode = strSavedRequestCode & vbCrLf & _
            "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.EntityFilter.OREntityFilter.FullNameList.Add " & strCustomerJob
        End If
      End If
    
      If strAccount <> "" Then
        If booAccountWithChildren Then
          .AccountFilter.ORAccountFilter.FullNameWithChildren.setValue strAccount
          strSavedRequestCode = strSavedRequestCode & vbCrLf & _
            "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.AccountFilter.ORAccountFilter.FullNameWithChildren.SetValue " & strAccount
        Else
          .AccountFilter.ORAccountFilter.FullNameList.Add strAccount
          strSavedRequestCode = strSavedRequestCode & vbCrLf & _
            "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.AccountFilter.ORAccountFilter.FullNameList.Add " & strAccount
        End If
      End If
    
      If strFromRefNumberRange <> "" Then
        .ORRefNumberFilter.RefNumberRangeFilter.FromRefNumber.setValue strFromRefNumberRange
        strSavedRequestCode = strSavedRequestCode & vbCrLf & _
          "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.ORRefNumberFilter.RefNumberRangeFilter.FromRefNumber.SetValue " & strFromRefNumberRange
      End If
  
      If strToRefNumberRange <> "" Then
        .ORRefNumberFilter.RefNumberRangeFilter.ToRefNumber.setValue strToRefNumberRange
        strSavedRequestCode = strSavedRequestCode & vbCrLf & _
          "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.ORRefNumberFilter.RefNumberRangeFilter.ToRefNumber.SetValue " & strToRefNumberRange
      End If
    
      If strRefNumberPiece <> "" Then
        .ORRefNumberFilter.RefNumberFilter.RefNumber.setValue strRefNumberPiece
        .ORRefNumberFilter.RefNumberFilter.MatchCriterion.SetAsString strRefNumberCriteria
        strSavedRequestCode = strSavedRequestCode & vbCrLf & _
          "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.ORRefNumberFilter.RefNumberFilter.RefNumber.SetValue " & strRefNumberPiece
        strSavedRequestCode = strSavedRequestCode & vbCrLf & _
          "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.ORRefNumberFilter.RefNumberFilter.MatchCriterion.SetAsString " & strRefNumberCriteria
      End If
    
      If strPaidStatus <> "" Then
        .PaidStatus.SetAsString strPaidStatus
        strSavedRequestCode = strSavedRequestCode & vbCrLf & _
          "  objInvoiceQuery.ORInvoiceQuery.InvoiceFilter.PaidStatus.SetAsString " & strPaidStatus
      End If
    End With 'objInvoiceQuery.ORInvoiceQuery.InvoiceFilter
  End If 'strRefNumber <> ""
  
  Dim objMsgSetResponse As IMsgSetResponse
  Set objMsgSetResponse = objSessionManager.DoRequests(objMsgSetRequest)
  
  Dim objResponse As IResponse
  Set objResponse = objMsgSetResponse.ResponseList.GetAt(0)
  
  If objResponse.StatusCode = 1 Then
    lstInvoices.AddItem "No invoices match the query filter used"
    Exit Sub
  End If
  
  Dim objInvoiceRetList As IInvoiceRetList
  Set objInvoiceRetList = objResponse.Detail
  
  Dim objInvoiceRet As IInvoiceRet
  Dim strItems As String
  Dim strReturnedRefNumber As String
  Dim i As Integer
  For i = 0 To objInvoiceRetList.Count - 1
    Set objInvoiceRet = objInvoiceRetList.GetAt(i)
    If objInvoiceRet.RefNumber Is Nothing Then
      strReturnedRefNumber = "Un-numbered "
    Else
      strReturnedRefNumber = "Invoice " & objInvoiceRet.RefNumber.getValue
    End If
    strItems = Str(objInvoiceRet.ORInvoiceLineRetList.Count)
    If Len(strItems) = 1 Then strItems = "  " & strItems
    If Len(strItems) = 2 Then strItems = " " & strItems
    lstInvoices.AddItem _
      strReturnedRefNumber & "     " & objInvoiceRet.TxnDate.getValue & _
      "     " & strItems & " items     " & _
      objInvoiceRet.CustomerRef.FullName.getValue & _
      "     Balance " & objInvoiceRet.BalanceRemaining.GetAsString & _
      "     " & objInvoiceRet.TxnID.getValue
  Next
  
End Sub


Public Sub QBFC_GetInvoice(TxnID As String)

  Dim objMsgSetRequest As IMsgSetRequest
  Set objMsgSetRequest = objSessionManager.CreateMsgSetRequest("US", intMajorVersion, intMinorVersion)
  
  Dim objInvoiceQuery As IInvoiceQuery
  Set objInvoiceQuery = objMsgSetRequest.AppendInvoiceQueryRq
  
  objInvoiceQuery.ORInvoiceQuery.TxnIDList.Add TxnID
  objInvoiceQuery.IncludeLineItems.setValue True
  
  Set objInvoiceMsgSetResponse = objSessionManager.DoRequests(objMsgSetRequest)
  
  Dim objResponse As IResponse
  Set objResponse = objInvoiceMsgSetResponse.ResponseList.GetAt(0)
  
  Dim objInvoiceRetList As IInvoiceRetList
  Set objInvoiceRetList = objResponse.Detail
  
  Set objSavedInvoiceRet = objInvoiceRetList.GetAt(0)
End Sub


Public Sub QBFC_LoadInvoiceModifyForm()

  frmInvoiceModify.txtTxnID.Text = objSavedInvoiceRet.TxnID.getValue
  
  frmInvoiceModify.txtEditSequence.Text = _
    objSavedInvoiceRet.EditSequence.getValue
  
  If Not objSavedInvoiceRet.RefNumber Is Nothing Then
    frmInvoiceModify.txtRefNumber.Text = objSavedInvoiceRet.RefNumber.getValue
  End If
  
  frmInvoiceModify.txtTxnDate.Text = objSavedInvoiceRet.TxnDate.getValue
  
  If Not objSavedInvoiceRet.IsPending Is Nothing Then
    If objSavedInvoiceRet.IsPending.getValue = True Then
      frmInvoiceModify.chkPending.Value = 1 'Checked
    End If
  End If
  
  If Not objSavedInvoiceRet.IsToBePrinted Is Nothing Then
    If objSavedInvoiceRet.IsToBePrinted.getValue = True Then
      frmInvoiceModify.chkToBePrinted.Value = 1 'Checked
    End If
  End If
  
  frmInvoiceModify.cmbCustomer.Text = _
    objSavedInvoiceRet.CustomerRef.FullName.getValue
  
  If Not objSavedInvoiceRet.ClassRef Is Nothing Then
    frmInvoiceModify.cmbClass.Text = _
      objSavedInvoiceRet.ClassRef.FullName.getValue
  End If

  If Not objSavedInvoiceRet.BillAddress Is Nothing Then
    If Not objSavedInvoiceRet.BillAddress.Addr1 Is Nothing Then
      frmInvoiceModify.txtBillAddr1.Text = _
        objSavedInvoiceRet.BillAddress.Addr1.getValue
    End If

    If Not objSavedInvoiceRet.BillAddress.Addr2 Is Nothing Then
      frmInvoiceModify.txtBillAddr2.Text = _
        objSavedInvoiceRet.BillAddress.Addr2.getValue
    End If

    If Not objSavedInvoiceRet.BillAddress.Addr3 Is Nothing Then
      frmInvoiceModify.txtBillAddr3.Text = _
        objSavedInvoiceRet.BillAddress.Addr3.getValue
    End If

    If Not objSavedInvoiceRet.BillAddress.Addr4 Is Nothing Then
      frmInvoiceModify.txtBillAddr4.Text = _
        objSavedInvoiceRet.BillAddress.Addr4.getValue
    End If

    If Not objSavedInvoiceRet.BillAddress.City Is Nothing Then
      frmInvoiceModify.txtBillCity.Text = _
        objSavedInvoiceRet.BillAddress.City.getValue
    End If

    If Not objSavedInvoiceRet.BillAddress.State Is Nothing Then
      frmInvoiceModify.txtBillState.Text = _
        objSavedInvoiceRet.BillAddress.State.getValue
    End If

    If Not objSavedInvoiceRet.BillAddress.PostalCode Is Nothing Then
      frmInvoiceModify.txtBillPostalCode.Text = _
        objSavedInvoiceRet.BillAddress.PostalCode.getValue
    End If

    If Not objSavedInvoiceRet.BillAddress.Country Is Nothing Then
      frmInvoiceModify.txtBillCountry.Text = _
        objSavedInvoiceRet.BillAddress.Country.getValue
    End If
  End If 'If Not objSavedInvoiceRet.BillAddress is Nothing

  If Not objSavedInvoiceRet.ShipAddress Is Nothing Then
    If Not objSavedInvoiceRet.ShipAddress.Addr1 Is Nothing Then
      frmInvoiceModify.txtShipAddr1.Text = _
        objSavedInvoiceRet.ShipAddress.Addr1.getValue
    End If

    If Not objSavedInvoiceRet.ShipAddress.Addr2 Is Nothing Then
      frmInvoiceModify.txtShipAddr2.Text = _
        objSavedInvoiceRet.ShipAddress.Addr2.getValue
    End If

    If Not objSavedInvoiceRet.ShipAddress.Addr3 Is Nothing Then
      frmInvoiceModify.txtShipAddr3.Text = _
        objSavedInvoiceRet.ShipAddress.Addr3.getValue
    End If

    If Not objSavedInvoiceRet.ShipAddress.Addr4 Is Nothing Then
      frmInvoiceModify.txtShipAddr4.Text = _
        objSavedInvoiceRet.ShipAddress.Addr4.getValue
    End If

    If Not objSavedInvoiceRet.ShipAddress.City Is Nothing Then
      frmInvoiceModify.txtShipCity.Text = _
        objSavedInvoiceRet.ShipAddress.City.getValue
    End If

    If Not objSavedInvoiceRet.ShipAddress.State Is Nothing Then
      frmInvoiceModify.txtShipState.Text = _
        objSavedInvoiceRet.ShipAddress.State.getValue
    End If

    If Not objSavedInvoiceRet.ShipAddress.PostalCode Is Nothing Then
      frmInvoiceModify.txtShipPostalCode.Text = _
        objSavedInvoiceRet.ShipAddress.PostalCode.getValue
    End If

    If Not objSavedInvoiceRet.ShipAddress.Country Is Nothing Then
      frmInvoiceModify.txtShipCountry.Text = _
        objSavedInvoiceRet.ShipAddress.Country.getValue
    End If
  End If 'If Not objSavedInvoiceRet.ShipAddress is Nothing

  If Not objSavedInvoiceRet.ARAccountRef Is Nothing Then
    frmInvoiceModify.cmbARAccount.Text = _
      objSavedInvoiceRet.ARAccountRef.FullName.getValue
  End If

  If Not objSavedInvoiceRet.TermsRef Is Nothing Then
    frmInvoiceModify.cmbTerms.Text = _
      objSavedInvoiceRet.TermsRef.FullName.getValue
  End If

  If Not objSavedInvoiceRet.PONumber Is Nothing Then
    frmInvoiceModify.txtPONumber.Text = _
      objSavedInvoiceRet.PONumber.getValue
  End If

  If Not objSavedInvoiceRet.DueDate Is Nothing Then
    frmInvoiceModify.txtDueDate.Text = _
      objSavedInvoiceRet.DueDate.getValue
  End If

  If Not objSavedInvoiceRet.ShipDate Is Nothing Then
    frmInvoiceModify.txtShipDate.Text = _
      objSavedInvoiceRet.ShipDate.getValue
  End If

  If Not objSavedInvoiceRet.FOB Is Nothing Then
    frmInvoiceModify.txtFOB.Text = _
      objSavedInvoiceRet.FOB.getValue
  End If

  If Not objSavedInvoiceRet.SalesRepRef Is Nothing Then
    frmInvoiceModify.cmbSalesRep.Text = _
      objSavedInvoiceRet.SalesRepRef.FullName.getValue
  End If

  If Not objSavedInvoiceRet.ShipMethodRef Is Nothing Then
    frmInvoiceModify.cmbShipMethod.Text = _
      objSavedInvoiceRet.ShipMethodRef.FullName.getValue
  End If

  If Not objSavedInvoiceRet.ItemSalesTaxRef Is Nothing Then
    frmInvoiceModify.cmbItemSalesTax.Text = _
      objSavedInvoiceRet.ItemSalesTaxRef.FullName.getValue
  End If

  If Not objSavedInvoiceRet.CustomerSalesTaxCodeRef Is Nothing Then
    frmInvoiceModify.cmbCustTaxCode.Text = _
      objSavedInvoiceRet.CustomerSalesTaxCodeRef.FullName.getValue
  End If

  If Not objSavedInvoiceRet.CustomerMsgRef Is Nothing Then
    frmInvoiceModify.cmbCustomerMsg.Text = _
      objSavedInvoiceRet.CustomerMsgRef.FullName.getValue
  End If

  If Not objSavedInvoiceRet.Memo Is Nothing Then
    frmInvoiceModify.txtMemo.Text = _
      objSavedInvoiceRet.Memo.getValue
  End If
End Sub


Public Sub QBFC_LoadInvoiceLineArray(strLineArray() As String)

  Dim objInvoiceLineRetList As IORInvoiceLineRetList
  Set objInvoiceLineRetList = objSavedInvoiceRet.ORInvoiceLineRetList
  
  
  Dim objLine As IInvoiceLineRet
  Dim objGroupLine As IInvoiceLineGroupRet
  Dim objGroupLines As IInvoiceLineRetList
  
  Dim i As Integer
  Dim j As Integer
  Dim k As Integer
  
  j = 0
  For i = 0 To objInvoiceLineRetList.Count - 1
  
    j = j + 1
    If objInvoiceLineRetList.GetAt(i).InvoiceLineGroupRet Is Nothing Then
      strLineArray(j) = ItemLineInfo(objInvoiceLineRetList.GetAt(i).InvoiceLineRet) & _
                        "Item<spliter>Original"
    Else
      strLineArray(j) = GroupLineInfo(objInvoiceLineRetList.GetAt(i).InvoiceLineGroupRet) & _
                        "Group<spliter>Original"
      
      Set objGroupLine = objInvoiceLineRetList.GetAt(i).InvoiceLineGroupRet
      Set objGroupLines = objGroupLine.InvoiceLineRetList
      If objGroupLines.Count > 0 Then
        For k = 0 To objGroupLines.Count - 1
          j = j + 1
          Set objLine = objGroupLines.GetAt(k)
          strLineArray(j) = ItemLineInfo(objLine) & "SubItem<spliter>Original"
        Next k
      End If 'objGroupLines.length > 0
    End If 'objNode.nodeName = "InvoiceLineRet"
  Next i
End Sub


Public Sub QBFC_ModifyInvoice(strInvoiceChangeString As String)

  PrettyPrintXMLToFile strInvoiceChangeString, "C:\InvoiceChangeString.txt"
  
  Dim objMsgSetRequest As IMsgSetRequest
  Set objMsgSetRequest = objSessionManager.CreateMsgSetRequest("US", intMajorVersion, intMinorVersion)
  
  Dim objInvoiceMod As IInvoiceMod
  Set objInvoiceMod = objMsgSetRequest.AppendInvoiceModRq
  
  Dim objInvoiceLineMod As IInvoiceLineMod
  Dim objInvoiceLineGroupMod As IInvoiceLineGroupMod
  
  Dim strTemp As String
  strTemp = strInvoiceChangeString
  strTemp = Right(strTemp, Len(strTemp) - 7)
  objInvoiceMod.TxnID.setValue Left(strTemp, InStr(1, strTemp, "</") - 1)

  strSavedRequestCode = _
    "  Dim objMsgSetRequest As IMsgSetRequest" & vbCrLf & _
    "  Dim objInvoiceLineMod As IInvoiceLineMod" & vbCrLf & _
    "  Dim objInvoiceLineGroupMod As IInvoiceLineGroupMod" & vbCrLf & vbCrLf & _
    "  Set objMsgSetRequest = objSessionManager.CreateMsgSetRequest(""US""," & Str(intMajorVersion) & ", " & Str(intMinorVersion) & ")" & vbCrLf & vbCrLf & _
    "  objInvoiceMod.TxnID.SetValue " & Left(strTemp, InStr(1, strTemp, "</") - 1) & vbCrLf
    
  strTemp = Right(strTemp, Len(strTemp) - (InStr(1, strTemp, "<EditS") + 13))
  objInvoiceMod.EditSequence.setValue Left(strTemp, InStr(1, strTemp, "</") - 1)

  strSavedRequestCode = strSavedRequestCode & _
    "  objInvoiceMod.EditSequence.SetValue " & Left(strTemp, InStr(1, strTemp, "</") - 1) & vbCrLf
  strTemp = Right(strTemp, Len(strTemp) - (InStr(1, strTemp, "</Edit") + 14))
  
  'The rest of our invoice change string is either going to be elements to
  'add to the InvoiceModRq aggregate or aggregates to add to the aggregate.
  'We'll pull each one out of the invoice change string and treat them the
  'same letting the recursive procedure we call figure it out.

  Dim strStartTag As String
  Dim strEndTag As String
  Dim strSegment As String
  Dim intTagLength As Integer
  
  Dim objNode As IXMLDOMNode
  
  Do
    GetTags strTemp, strStartTag, strEndTag, intTagLength
    If intTagLength = 0 Then
      MsgBox "Error processing invoice change string, exiting"
      End
    End If
    
    strSegment = Left(strTemp, InStr(1, strTemp, strEndTag) + intTagLength)
    strTemp = Right(strTemp, Len(strTemp) - Len(strSegment))
    
    AddElementOrAggregate _
      strSegment, strStartTag, intTagLength, "", objInvoiceMod, _
      objInvoiceLineMod, objInvoiceLineGroupMod

  Loop Until strTemp = Empty
  
  Dim objMsgSetResponse As IMsgSetResponse
  Set objMsgSetResponse = objSessionManager.DoRequests(objMsgSetRequest)
  
  Dim objResponse As IResponse
  Set objResponse = objMsgSetResponse.ResponseList.GetAt(0)
  
  If objResponse.StatusCode <> 0 Then
    MsgBox "Error modifying invoice" & vbCrLf & vbCrLf & _
      "Error = " & objResponse.StatusCode & vbCrLf & vbCrLf & _
      "Message = " & objResponse.StatusMessage
  Else
    MsgBox "Successfully modified invoice"
  End If
End Sub


Public Sub QBFC_GetItemInfo(strItemFullName As String, _
                            strDesc As String, _
                            strRate As String, _
                            strRateOrPercent As String, _
                            strSalesTaxCode As String)

  strDesc = ""
  strRate = ""
  strRateOrPercent = "Rate"
  strSalesTaxCode = ""
  
  Dim objMsgSetRequest As IMsgSetRequest
  Set objMsgSetRequest = objSessionManager.CreateMsgSetRequest("US", intMajorVersion, intMinorVersion)
  objMsgSetRequest.Attributes.OnError = roeContinue
  
  Dim objItemQuery As IItemQuery
  Set objItemQuery = objMsgSetRequest.AppendItemQueryRq
  objItemQuery.ORListQuery.FullNameList.Add strItemFullName
  
  Dim objMsgSetResponse As IMsgSetResponse
  Set objMsgSetResponse = objSessionManager.DoRequests(objMsgSetRequest)
  
  Dim objResponse As IResponse
  Set objResponse = objMsgSetResponse.ResponseList.GetAt(0)
  
  Dim objItemRetList As IORItemRetList
  Set objItemRetList = objResponse.Detail
  
  Dim enItemType As ENORItemRet
  enItemType = objItemRetList.GetAt(0).ortype
  
  Dim objItemRet
  Select Case enItemType
    Case Is = orirItemServiceRet
      Set objItemRet = objItemRetList.GetAt(0).ItemServiceRet
    Case Is = orirItemNonInventoryRet
      Set objItemRet = objItemRetList.GetAt(0).ItemNonInventoryRet
    Case Is = orirItemOtherChargeRet
      Set objItemRet = objItemRetList.GetAt(0).ItemOtherChargeRet
    Case Is = orirItemInventoryRet
      Set objItemRet = objItemRetList.GetAt(0).ItemInventoryRet
    Case Is = orirItemInventoryAssemblyRet
      Set objItemRet = objItemRetList.GetAt(0).ItemInventoryAssemblyRet
    Case Is = orirItemSubtotalRet
      Set objItemRet = objItemRetList.GetAt(0).ItemSubtotalRet
    Case Is = orirItemDiscountRet
      Set objItemRet = objItemRetList.GetAt(0).ItemDiscountRet
    Case Is = orirItemPaymentRet
      Set objItemRet = objItemRetList.GetAt(0).ItemPaymentRet
    Case Is = orirItemSalesTaxRet
      Set objItemRet = objItemRetList.GetAt(0).ItemSalesTaxRet
    Case Is = orirItemGroupRet
      Set objItemRet = objItemRetList.GetAt(0).ItemGroupRet
  End Select
  
  If enItemType = orirItemServiceRet Or enItemType = orirItemNonInventoryRet Or _
     enItemType = orirItemOtherChargeRet Then
    If Not objItemRet.ORSalesPurchase.SalesOrPurchase Is Nothing Then
      If Not objItemRet.ORSalesPurchase.SalesOrPurchase.Desc Is Nothing Then
        strDesc = objItemRet.ORSalesPurchase.SalesOrPurchase.Desc.getValue
      End If
      
      If Not objItemRet.ORSalesPurchase.SalesOrPurchase.ORPrice Is Nothing Then
        If Not objItemRet.ORSalesPurchase.SalesOrPurchase.ORPrice.Price Is Nothing Then
          strRate = objItemRet.ORSalesPurchase.SalesOrPurchase.ORPrice.Price.GetAsString
        Else
          strRate = objItemRet.ORSalesPurchase.SalesOrPurchase.ORPrice.PricePercent.GetAsString
          strRateOrPercent = "RatePercent"
        End If
      End If
    Else ' Since it isn't SalesOrPurchase it must be SalesAndPurchase
      If Not objItemRet.ORSalesPurchase.SalesAndPurchase.SalesDesc Is Nothing Then
        strDesc = objItemRet.ORSalesPurchase.SalesAndPurchase.SalesDesc.getValue
      End If
      
      If Not objItemRet.ORSalesPurchase.SalesAndPurchase.SalesPrice Is Nothing Then
        strRate = objItemRet.ORSalesPurchase.SalesAndPurchase.SalesPrice.GetAsString
      End If
    End If
  ElseIf enItemType = orirItemInventoryRet Or _
         enItemType = orirItemInventoryAssemblyRet Then
    If Not objItemRet.SalesDesc Is Nothing Then
      strDesc = objItemRet.SalesDesc.getValue
    End If
  
    If Not objItemRet.SalesPrice Is Nothing Then
      strRate = objItemRet.SalesPrice.GetAsString
    End If
  Else
    If Not objItemRet.ItemDesc Is Nothing Then
      strDesc = objItemRet.ItemDesc.getValue
    End If
  End If
  
  If Not (enItemType = orirItemSubtotalRet Or _
          enItemType = orirItemPaymentRet Or _
          enItemType = orirItemSalesTaxRet Or _
          enItemType = orirItemGroupRet) Then
    If Not objItemRet.SalesTaxCodeRef Is Nothing Then
      strSalesTaxCode = objItemRet.SalesTaxCodeRef.FullName.getValue
    End If
  End If
End Sub


Public Sub QBFC_LoadRequest(strRequestText As String)
  strRequestText = strSavedRequestCode
End Sub


Private Function ItemLineInfo(objInvoiceLine As IInvoiceLineRet) As String

  Dim strLineInfo As String
  Dim strRateOrPercent As String

  strLineInfo = objInvoiceLine.TxnLineID.getValue & "<spliter>"
  
  If Not objInvoiceLine.Quantity Is Nothing Then
    strLineInfo = strLineInfo & objInvoiceLine.Quantity.GetAsString
  End If
  strLineInfo = strLineInfo & "<spliter>"
  
  If Not objInvoiceLine.ItemRef Is Nothing Then
    strLineInfo = strLineInfo & _
      objInvoiceLine.ItemRef.FullName.getValue
  End If
  strLineInfo = strLineInfo & "<spliter>"
  
  If Not objInvoiceLine.Desc Is Nothing Then
    strLineInfo = strLineInfo & objInvoiceLine.Desc.getValue
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
    strLineInfo = strLineInfo & _
      objInvoiceLine.ClassRef.FullName.getValue
  End If
  strLineInfo = strLineInfo & "<spliter>"
  
  If Not objInvoiceLine.ServiceDate Is Nothing Then
    strLineInfo = strLineInfo & _
      Format(objInvoiceLine.ServiceDate.getValue, "YYYY-MM-DD")
  End If
  strLineInfo = strLineInfo & "<spliter>"
    
  If Not objInvoiceLine.SalesTaxCodeRef Is Nothing Then
    strLineInfo = strLineInfo & _
      objInvoiceLine.SalesTaxCodeRef.FullName.getValue
  End If
  strLineInfo = strLineInfo & "<spliter>" & strRateOrPercent & _
    "<spliter><spliter>"
  
  
  ItemLineInfo = strLineInfo
End Function


Private Function GroupLineInfo(objInvoiceGroupLine As IInvoiceLineGroupRet) As String

  Dim strLineInfo As String
  Dim strRateOrPercent As String

  strLineInfo = objInvoiceGroupLine.TxnLineID.getValue & "<spliter>"
  
  If Not objInvoiceGroupLine.Quantity Is Nothing Then
    strLineInfo = strLineInfo & objInvoiceGroupLine.Quantity.GetAsString
  End If
  strLineInfo = strLineInfo & "<spliter>"
  
  If Not objInvoiceGroupLine.ItemGroupRef Is Nothing Then
    strLineInfo = strLineInfo & _
      objInvoiceGroupLine.ItemGroupRef.FullName.getValue
  End If
  strLineInfo = strLineInfo & "<spliter>"
  
  If Not objInvoiceGroupLine.Desc Is Nothing Then
    strLineInfo = strLineInfo & objInvoiceGroupLine.Desc.getValue
  End If
  strLineInfo = strLineInfo & "<spliter><spliter><spliter><spliter>"

  If Not objInvoiceGroupLine.ServiceDate Is Nothing Then
    strLineInfo = strLineInfo & _
      Format(objInvoiceGroupLine.ServiceDate.getValue, "YYYY-MM-DD")
  End If
  strLineInfo = strLineInfo & "<spliter><spliter><spliter><spliter>"
  
  GroupLineInfo = strLineInfo
End Function


Private Sub AddElementOrAggregate(strElOrAggString As String, _
                                  strStartTag As String, _
                                  intTagLength As Integer, _
                                  strAggregateType As String, _
                                  objInvoiceMod As IInvoiceMod, _
                                  objCurrentModLine As IInvoiceLineMod, _
                                  objCurrentModGroupLine As IInvoiceLineGroupMod)
  
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
  Dim intInnerTagLength As Integer

  GetTags strValue, strInnerStartTag, strInnerEndTag, intInnerTagLength
  
  If strInnerStartTag = Empty Then
    Select Case strAggregateType
      Case Is = Empty
        Select Case strTagName
          Case Is = "TxnID"
            objInvoiceMod.TxnID.setValue strValue
            strSavedRequestCode = strSavedRequestCode & _
              "  objInvoiceMod.TxnID.SetValue " & strValue & vbCrLf
          Case Is = "EditSequence"
            objInvoiceMod.EditSequence.setValue strValue
            strSavedRequestCode = strSavedRequestCode & _
              "  objInvoiceMod.EditSequence.SetValue " & strValue & vbCrLf
          Case Is = "TxnDate"
            objInvoiceMod.TxnDate.setValue CDate(strValue)
            strSavedRequestCode = strSavedRequestCode & _
              "  objInvoiceMod.TxnDate.SetValue CDate(""" & strValue & """)" & vbCrLf
          Case Is = "RefNumber"
            objInvoiceMod.RefNumber.setValue strValue
            strSavedRequestCode = strSavedRequestCode & _
              "  objInvoiceMod.RefNumber.SetValue " & strValue & vbCrLf
          Case Is = "IsPending"
            objInvoiceMod.IsPending.setValue (strValue = "1")
            strSavedRequestCode = strSavedRequestCode & _
              "  objInvoiceMod.IsPending.SetValue " & (strValue = "1") & vbCrLf
          Case Is = "PONumber"
            objInvoiceMod.PONumber.setValue strValue
            strSavedRequestCode = strSavedRequestCode & _
              "  objInvoiceMod.PONumber.SetValue " & strValue & vbCrLf
          Case Is = "DueDate"
            objInvoiceMod.DueDate.setValue CDate(strValue)
            strSavedRequestCode = strSavedRequestCode & _
              "  objInvoiceMod.DueDate.SetValue CDate(""" & strValue & """)" & vbCrLf
          Case Is = "FOB"
            objInvoiceMod.FOB.setValue strValue
            strSavedRequestCode = strSavedRequestCode & _
              "  objInvoiceMod.FOB.SetValue " & strValue & vbCrLf
          Case Is = "ShipDate"
            objInvoiceMod.ShipDate.setValue CDate(strValue)
            strSavedRequestCode = strSavedRequestCode & _
              "  objInvoiceMod.ShipDate.SetValue CDate(""" & strValue & """)" & vbCrLf
          Case Is = "Memo"
            objInvoiceMod.Memo.setValue strValue
            strSavedRequestCode = strSavedRequestCode & _
              "  objInvoiceMod.Memo.SetValue " & strValue & vbCrLf
          Case Is = "IsToBePrinted"
            objInvoiceMod.IsToBePrinted.setValue (strValue = "1")
            strSavedRequestCode = strSavedRequestCode & _
              "  objInvoiceMod.IsToBePrinted.SetValue " & (strValue = "1") & vbCrLf
        End Select
      Case Is = "BillAddress"
        Select Case strTagName
          Case Is = "Addr1"
            objInvoiceMod.BillAddress.Addr1.setValue strValue
            strSavedRequestCode = strSavedRequestCode & _
              "  objInvoiceMod.BillAddress.Addr1.SetValue " & strValue & vbCrLf
          Case Is = "Addr2"
            objInvoiceMod.BillAddress.Addr2.setValue strValue
            strSavedRequestCode = strSavedRequestCode & _
              "  objInvoiceMod.BillAddress.Addr2.SetValue " & strValue & vbCrLf
          Case Is = "Addr3"
            objInvoiceMod.BillAddress.Addr3.setValue strValue
            strSavedRequestCode = strSavedRequestCode & _
              "  objInvoiceMod.BillAddress.Addr3.SetValue " & strValue & vbCrLf
          Case Is = "Addr4"
            objInvoiceMod.BillAddress.Addr4.setValue strValue
            strSavedRequestCode = strSavedRequestCode & _
              "  objInvoiceMod.BillAddress.Addr4.SetValue " & strValue & vbCrLf
          Case Is = "City"
            objInvoiceMod.BillAddress.City.setValue strValue
            strSavedRequestCode = strSavedRequestCode & _
              "  objInvoiceMod.BillAddress.City.SetValue " & strValue & vbCrLf
          Case Is = "State"
            objInvoiceMod.BillAddress.State.setValue strValue
            strSavedRequestCode = strSavedRequestCode & _
              "  objInvoiceMod.BillAddress.State.SetValue " & strValue & vbCrLf
          Case Is = "PostalCode"
            objInvoiceMod.BillAddress.PostalCode.setValue strValue
            strSavedRequestCode = strSavedRequestCode & _
              "  objInvoiceMod.BillAddress.PostalCode.SetValue " & strValue & vbCrLf
          Case Is = "Country"
            objInvoiceMod.BillAddress.Country.setValue strValue
            strSavedRequestCode = strSavedRequestCode & _
              "  objInvoiceMod.BillAddress.Country.SetValue " & strValue & vbCrLf
        End Select
      Case Is = "ShipAddress"
        Select Case strTagName
          Case Is = "Addr1"
            objInvoiceMod.ShipAddress.Addr1.setValue strValue
            strSavedRequestCode = strSavedRequestCode & _
              "  objInvoiceMod.ShipAddress.Addr1.SetValue " & strValue & vbCrLf
          Case Is = "Addr2"
            objInvoiceMod.ShipAddress.Addr2.setValue strValue
            strSavedRequestCode = strSavedRequestCode & _
              "  objInvoiceMod.ShipAddress.Addr2.SetValue " & strValue & vbCrLf
          Case Is = "Addr3"
            objInvoiceMod.ShipAddress.Addr3.setValue strValue
            strSavedRequestCode = strSavedRequestCode & _
              "  objInvoiceMod.ShipAddress.Addr3.SetValue " & strValue & vbCrLf
          Case Is = "Addr4"
            objInvoiceMod.ShipAddress.Addr4.setValue strValue
            strSavedRequestCode = strSavedRequestCode & _
              "  objInvoiceMod.ShipAddress.Addr4.SetValue " & strValue & vbCrLf
          Case Is = "City"
            objInvoiceMod.ShipAddress.City.setValue strValue
            strSavedRequestCode = strSavedRequestCode & _
              "  objInvoiceMod.ShipAddress.City.SetValue " & strValue & vbCrLf
          Case Is = "State"
            objInvoiceMod.ShipAddress.State.setValue strValue
            strSavedRequestCode = strSavedRequestCode & _
              "  objInvoiceMod.ShipAddress.State.SetValue " & strValue & vbCrLf
          Case Is = "PostalCode"
            objInvoiceMod.ShipAddress.PostalCode.setValue strValue
            strSavedRequestCode = strSavedRequestCode & _
              "  objInvoiceMod.ShipAddress.PostalCode.SetValue " & strValue & vbCrLf
          Case Is = "Country"
            objInvoiceMod.ShipAddress.Country.setValue strValue
            strSavedRequestCode = strSavedRequestCode & _
              "  objInvoiceMod.ShipAddress.Country.SetValue " & strValue & vbCrLf
        End Select
      Case Is = "InvoiceLineGroupMod"
        Select Case strTagName
          Case Is = "TxnLineID"
            objCurrentModGroupLine.TxnLineID.setValue strValue
            strSavedRequestCode = strSavedRequestCode & _
              "  objInvoiceLineGroupMod.TxnLineID.SetValue " & strValue & vbCrLf
          Case Is = "Quantity"
            objCurrentModGroupLine.Quantity.SetAsString strValue
            strSavedRequestCode = strSavedRequestCode & _
              "  objInvoiceLineGroupMod.Quantity.SetAsString " & strValue & vbCrLf
          Case Is = "ServiceDate"
            objCurrentModGroupLine.ServiceDate.setValue CDate(strValue)
            strSavedRequestCode = strSavedRequestCode & _
              "  objInvoiceLineGroupMod.ServiceDate.SetValue CDate(""" & strValue & """)" & vbCrLf
        End Select
      Case Else
        If strAggregateType = "InvoiceLineMod" Or strAggregateType = "SubItem" Then

          Select Case strTagName
            Case Is = "TxnLineID"
              objCurrentModLine.TxnLineID.setValue strValue
              strSavedRequestCode = strSavedRequestCode & _
                "  objInvoiceLineMod.TxnLineID.SetValue " & strValue & vbCrLf
            Case Is = "Desc"
              objCurrentModLine.Desc.setValue strValue
              strSavedRequestCode = strSavedRequestCode & _
                "  objInvoiceLineMod.Desc.SetValue " & strValue & vbCrLf
            Case Is = "Quantity"
              objCurrentModLine.Quantity.SetAsString strValue
              strSavedRequestCode = strSavedRequestCode & _
                "  objInvoiceLineMod.Quantity.SetAsString " & strValue & vbCrLf
            Case Is = "RatePercent"
              objCurrentModLine.ORRatePriceLevel.RatePercent.SetAsString strValue
              strSavedRequestCode = strSavedRequestCode & _
                "  objInvoiceLineMod.RatePercent.SetAsString " & strValue & vbCrLf
            Case Is = "Rate"
              objCurrentModLine.ORRatePriceLevel.Rate.SetAsString strValue
              strSavedRequestCode = strSavedRequestCode & _
                "  objInvoiceLineMod.Rate.SetAsString " & strValue & vbCrLf
            Case Is = "Amount"
              objCurrentModLine.Amount.SetAsString strValue
              strSavedRequestCode = strSavedRequestCode & _
                "  objInvoiceLineMod.Amount.SetAsString " & strValue & vbCrLf
            Case Is = "ServiceDate"
              objCurrentModLine.ServiceDate.setValue CDate(strValue)
              strSavedRequestCode = strSavedRequestCode & _
                "  objInvoiceLineMod.ServiceDate.SetValue CDate(""" & strValue & """)" & vbCrLf
          End Select
        End If
    End Select
  Else 'For If strInnerStartTag = Empty Then
    strInnerValue = Left(strValue, Len(strValue) - (intInnerTagLength + 1))
    strInnerValue = Right(strInnerValue, Len(strInnerValue) - intInnerTagLength)
        
    Select Case strAggregateType
      Case Is = Empty
        'We'll get this for the aggregates in the InvoiceMod, so look for
        'InvoiceLineMod, InvoiceLineGroupMod or a reference aggretate and
        'use what's between the <FullName> and </FullName> tags
        
        Select Case strTagName
          Case Is = "InvoiceLineMod"
            Set objCurrentModLine = objInvoiceMod.ORInvoiceLineModList.Append.InvoiceLineMod
            strSavedRequestCode = strSavedRequestCode & vbCrLf & _
              "  Set objInvoiceLineMod = objInvoiceMod.ORInvoiceLineModList.Append.InvoiceLineMod" & vbCrLf
          Case Is = "InvoiceLineGroupMod"
            Set objCurrentModGroupLine = objInvoiceMod.ORInvoiceLineModList.Append.InvoiceLineGroupMod
            strSavedRequestCode = strSavedRequestCode & vbCrLf & _
              "  Set objInvoiceLineGroupMod = objInvoiceMod.ORInvoiceLineModList.Append.InvoiceLineGroupMod" & vbCrLf
          Case Is = "CustomerRef"
            objInvoiceMod.CustomerRef.FullName.setValue strInnerValue
            strSavedRequestCode = strSavedRequestCode & _
              "  objInvoiceMod.CustomerRef.FullName.SetValue " & strInnerValue & vbCrLf
          Case Is = "ClassRef"
            objInvoiceMod.ClassRef.FullName.setValue strInnerValue
            strSavedRequestCode = strSavedRequestCode & _
              "  objInvoiceMod.ClassRef.FullName.SetValue " & strInnerValue & vbCrLf
          Case Is = "ARAccountRef"
            objInvoiceMod.ARAccountRef.FullName.setValue strInnerValue
            strSavedRequestCode = strSavedRequestCode & _
              "  objInvoiceMod.ARAccountRef.FullName.SetValue " & strInnerValue & vbCrLf
          Case Is = "TermsRef"
            objInvoiceMod.TermsRef.FullName.setValue strInnerValue
            strSavedRequestCode = strSavedRequestCode & _
              "  objInvoiceMod.TermsRef.FullName.SetValue " & strInnerValue & vbCrLf
          Case Is = "SalesRepRef"
            objInvoiceMod.SalesRepRef.FullName.setValue strInnerValue
            strSavedRequestCode = strSavedRequestCode & _
              "  objInvoiceMod.SalesRepRef.FullName.SetValue " & strInnerValue & vbCrLf
          Case Is = "ShipMethodRef"
            objInvoiceMod.ShipMethodRef.FullName.setValue strInnerValue
            strSavedRequestCode = strSavedRequestCode & _
              "  objInvoiceMod.ShipMethodRef.FullName.SetValue " & strInnerValue & vbCrLf
          Case Is = "ItemSalesTaxRef"
            objInvoiceMod.ItemSalesTaxRef.FullName.setValue strInnerValue
            strSavedRequestCode = strSavedRequestCode & _
              "  objInvoiceMod.ItemSalesTaxRef.FullName.SetValue " & strInnerValue & vbCrLf
          Case Is = "CustomerMsgRef"
            objInvoiceMod.CustomerMsgRef.FullName.setValue strInnerValue
            strSavedRequestCode = strSavedRequestCode & _
              "  objInvoiceMod.CustomerMsgRef.FullName.SetValue " & strInnerValue & vbCrLf
          Case Is = "CustomerSalesTaxCodeRef"
            objInvoiceMod.CustomerSalesTaxCodeRef.FullName.setValue strInnerValue
            strSavedRequestCode = strSavedRequestCode & _
              "  objInvoiceMod.CustomerSalesTaxCodeRef.FullName.SetValue " & strInnerValue & vbCrLf
        End Select
      Case Is = "InvoiceLineGroupMod"
        Select Case strTagName
          Case Is = "ItemGroupRef"
            objCurrentModGroupLine.ItemGroupRef.FullName.setValue strInnerValue
            strSavedRequestCode = strSavedRequestCode & _
              "  objInvoiceLineGroupMod.ItemGroupRef.FullName.SetValue " & strInnerValue & vbCrLf
          Case Is = "InvoiceLineMod"
            Set objCurrentModLine = objCurrentModGroupLine.InvoiceLineModList.Append
            strSavedRequestCode = strSavedRequestCode & vbCrLf & _
              "  Set objInvoiceLineMod = objInvoiceLineGroupMod.InvoiceLineModList.Append " & vbCrLf
    
            Do While strValue <> Empty
              strSegment = Left(strValue, InStr(1, strValue, strInnerEndTag) + intInnerTagLength)
              strValue = Right(strValue, Len(strValue) - Len(strSegment))

              AddElementOrAggregate _
                strSegment, strInnerStartTag, intInnerTagLength, "SubItem", _
                objInvoiceMod, objCurrentModLine, objCurrentModGroupLine
    
              If strValue <> Empty Then GetTags strValue, strInnerStartTag, strInnerEndTag, intInnerTagLength
            Loop
        End Select
      Case Else
        
        If strAggregateType = "InvoiceLineMod" Or strAggregateType = "SubItem" Then
          Select Case strTagName
            Case Is = "ItemRef"
              objCurrentModLine.ItemRef.FullName.setValue strInnerValue
              strSavedRequestCode = strSavedRequestCode & _
                "  objInvoiceLineMod.ItemRef.FullName.SetValue " & strInnerValue & vbCrLf
            Case Is = "ClassRef"
              objCurrentModLine.ClassRef.FullName.setValue strInnerValue
              strSavedRequestCode = strSavedRequestCode & _
                "  objInvoiceLineMod.ClassRef.FullName.SetValue " & strInnerValue & vbCrLf
            Case Is = "OverrideAccountRef"
              objCurrentModLine.OverrideAccountRef.FullName.setValue strInnerValue
              strSavedRequestCode = strSavedRequestCode & _
                "  objInvoiceLineMod.OverrideAccountRef.FullName.SetValue " & strInnerValue & vbCrLf
            Case Is = "SalesTaxCodeRef"
              objCurrentModLine.SalesTaxCodeRef.FullName.setValue strInnerValue
              strSavedRequestCode = strSavedRequestCode & _
                "  objInvoiceLineMod.SalesTaxCodeRef.FullName.SetValue " & strInnerValue & vbCrLf
          End Select
        End If
    End Select
    
    Do While strValue <> Empty
      strSegment = Left(strValue, InStr(1, strValue, strInnerEndTag) + intInnerTagLength)
      strValue = Right(strValue, Len(strValue) - Len(strSegment))

      AddElementOrAggregate _
        strSegment, strInnerStartTag, intInnerTagLength, strTagName, _
        objInvoiceMod, objCurrentModLine, objCurrentModGroupLine
    
      If strValue <> Empty Then GetTags strValue, strInnerStartTag, strInnerEndTag, intInnerTagLength
    Loop
  End If
End Sub


