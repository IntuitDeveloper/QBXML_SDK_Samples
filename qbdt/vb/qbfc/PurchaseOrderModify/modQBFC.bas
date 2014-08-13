Attribute VB_Name = "modQBFC"
'----------------------------------------------------------
' Copyright © 2003-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
' Added values for intMajorVersion and intMinorVersion
'----------------------------------------------------------

Option Explicit

Dim objSessionManager As New QBSessionManager

Dim booSessionBegun As Boolean

Dim intMajorVersion As Integer
Dim intMinorVersion As Integer

Public Sub QBFC_OpenConnectionBeginSession()

  On Error GoTo Errs
  
  objSessionManager.OpenConnection "", "IDN PO Modify Sample Application"
  objSessionManager.BeginSession "", omDontCare
  booSessionBegun = True
  intMajorVersion = 3
  intMinorVersion = 0
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




Public Sub QBFC_LoadPOInfoArray(strPOInfo() As String, _
                                dateFrom As Date, _
                                dateTo As Date, _
                                booGiveWarning As Boolean)

  Dim objMsgSetRequest As IMsgSetRequest
  Set objMsgSetRequest = objSessionManager.CreateMsgSetRequest("US", 3, 0)
  
  Dim objGeneralDetailReportQuery As IGeneralDetailReportQuery
  Set objGeneralDetailReportQuery = objMsgSetRequest.AppendGeneralDetailReportQueryRq
  
  objGeneralDetailReportQuery.ORReportPeriod.ReportPeriod.FromReportDate.setValue dateFrom
  objGeneralDetailReportQuery.ORReportPeriod.ReportPeriod.ToReportDate.setValue dateTo
  objGeneralDetailReportQuery.GeneralDetailReportType.setValue gdrtOpenPOs
  
  objGeneralDetailReportQuery.IncludeColumnList.Add icRefNumber
  objGeneralDetailReportQuery.IncludeColumnList.Add icDate
  objGeneralDetailReportQuery.IncludeColumnList.Add icName
  objGeneralDetailReportQuery.IncludeColumnList.Add icDeliveryDate
  objGeneralDetailReportQuery.IncludeColumnList.Add icAmount
  
  If SupportsModify Then
    objGeneralDetailReportQuery.IncludeColumnList.Add icTxnID
  End If

  Dim objMsgSetResponse As IMsgSetResponse
  Set objMsgSetResponse = objSessionManager.DoRequests(objMsgSetRequest)
  
  Dim objResponse As IResponse
  Set objResponse = objMsgSetResponse.ResponseList.GetAt(0)
  
  If objResponse.StatusCode <> 0 Then
    MsgBox "Error getting Open Purchase Orders" & vbCrLf & vbCrLf & _
      "Error = " & Hex(objResponse.StatusCode) & vbCrLf & vbCrLf & _
      objResponse.StatusMessage
    strPOInfo(0) = "<spliter><spliter>There were no open purchase orders returned<spliter><spliter>"
    Exit Sub
  End If
  
  Dim objGeneralDetailReport As IReportRet
  Set objGeneralDetailReport = objResponse.Detail
  
  If objGeneralDetailReport.NumRows.getValue < 3 Then
    strPOInfo(0) = "<spliter><spliter>There were no open purchase orders returned<spliter><spliter>"
    Exit Sub
  End If
  
  Dim i As Integer
  Dim j As Integer
  If objGeneralDetailReport.NumRows.getValue < 32 Then
    j = objGeneralDetailReport.NumRows.getValue - 2
  Else
    j = 30
  End If
  
  Dim objReportDataList As IORReportDataList
  Set objReportDataList = objGeneralDetailReport.ReportData.ORReportDataList
  
  Dim objReportRow As IORReportData
  Dim objDataRow As IDataRow
  Dim objColDataList As IColDataList
  Dim objColData As IColData
  
  For i = 1 To j
    Set objReportRow = objReportDataList.GetAt(i)
    Set objDataRow = objReportRow.DataRow
    Set objColDataList = objDataRow.ColDataList
    
    strPOInfo(i) = _
      objColDataList.GetAt(0).Value.getValue & "<spliter>" & _
      objColDataList.GetAt(1).Value.getValue & "<spliter>" & _
      objColDataList.GetAt(2).Value.getValue & "<spliter>" & _
      objColDataList.GetAt(3).Value.getValue & "<spliter>" & _
      objColDataList.GetAt(4).Value.getValue & "<spliter>"
    
    If SupportsModify Then
      strPOInfo(i) = strPOInfo(i) & objColDataList.GetAt(5).Value.getValue
    End If
  Next
End Sub


Public Sub QBFC_GetPOInfo(strTxnID As String, _
                          strEditSequence As String, _
                          strRefNumber As String, _
                          strTxnDate As String, _
                          strVendor As String, _
                          strPOLines() As String, _
                          strSelectedPOInfo As String)
  
  Dim strPOInfoSplits() As String
  strPOInfoSplits = Split(strSelectedPOInfo, "<spliter>")
  
  Dim objMsgSetRequest As IMsgSetRequest
  Set objMsgSetRequest = objSessionManager.CreateMsgSetRequest("US", intMajorVersion, intMinorVersion)
  objMsgSetRequest.Attributes.OnError = roeContinue
  
  Dim objPOQuery As IPurchaseOrderQuery
  Set objPOQuery = objMsgSetRequest.AppendPurchaseOrderQueryRq
  
  objPOQuery.IncludeLineItems.setValue True
  objPOQuery.ORTxnQuery.TxnFilter.EntityFilter.OREntityFilter.FullNameList.Add strPOInfoSplits(2)
  objPOQuery.ORTxnQuery.TxnFilter.ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter.TxnDateFilter.FromTxnDate.setValue CDate(strPOInfoSplits(1))
  objPOQuery.ORTxnQuery.TxnFilter.ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter.TxnDateFilter.ToTxnDate.setValue CDate(strPOInfoSplits(1))
  objPOQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberRangeFilter.FromRefNumber.setValue strPOInfoSplits(0)
  objPOQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberRangeFilter.ToRefNumber.setValue strPOInfoSplits(0)
  
  Dim objMsgSetResponse As IMsgSetResponse
  Set objMsgSetResponse = objSessionManager.DoRequests(objMsgSetRequest)
  
  Dim objResponse As IResponse
  Set objResponse = objMsgSetResponse.ResponseList.GetAt(0)
  
  If objResponse.StatusCode <> 0 Then
    MsgBox "Unexpected Error getting purchase order information " & _
      vbCrLf & vbCrLf & "Status Code = " & objResponse.StatusCode & _
      vbCrLf & vbCrLf & "Status Message = " & objResponse.StatusMessage
    strPOLines(1) = Empty
    Exit Sub
  End If
  
  Dim objPORetList As IPurchaseOrderRetList
  Set objPORetList = objResponse.Detail
  
  Dim objPORet As IPurchaseOrderRet
  Set objPORet = objPORetList.GetAt(0)
  
  If objPORet.ORPurchaseOrderLineRetList Is Nothing Then
    strPOLines(1) = Empty
    Exit Sub
  End If
  
  strTxnID = objPORet.TxnID.getValue
  strEditSequence = objPORet.EditSequence.getValue
  strTxnDate = Format(objPORet.TxnDate.getValue, "YYYY-MM-DD")
  
  If Not objPORet.RefNumber Is Nothing Then
    strRefNumber = objPORet.RefNumber.getValue
  End If
  
  If Not objPORet.VendorRef Is Nothing Then
    strVendor = _
      objPORet.VendorRef.FullName.getValue
  End If
  
  Dim objPOLineRetList As IORPurchaseOrderLineRetList
  Set objPOLineRetList = objPORet.ORPurchaseOrderLineRetList
  
  
  Dim booDone As Boolean
  Dim i As Integer
  Dim j As Integer
  Dim k As Integer
  booDone = False
  
  Dim objPOLine As IPurchaseOrderLineRet
  Dim objPOGroupLine As IPurchaseOrderLineGroupRet
  
  k = 1
  For i = 0 To objPOLineRetList.Count - 1
    If objPOLineRetList.GetAt(i).ortype = orpolrPurchaseOrderLineRet Then
      Set objPOLine = objPOLineRetList.GetAt(i).PurchaseOrderLineRet
      strPOLines(k) = POLineString(objPOLine, "")
    Else
      strPOLines(k) = POGroupLineString(objPOGroupLine, "")
      If Not objPOGroupLine.PurchaseOrderLineRetList Is Nothing Then
        For j = 0 To objPOGroupLine.PurchaseOrderLineRetList.Count - 1
          k = k + 1
          strPOLines(k) = POLineString(objPOGroupLine.PurchaseOrderLineRetList.GetAt(j), "GroupSubItem")
        Next j
      End If
    End If
    
    k = k + 1
  Next i
End Sub


Private Function POLineString(objPOLineRet As IPurchaseOrderLineRet, _
                              strLineType As String)

  Dim strOrdered As String
  Dim strReceived As String

  Dim strOutput As String
  
  If Not objPOLineRet.ItemRef Is Nothing Then
      strOutput = objPOLineRet.ItemRef.FullName.getValue
  End If
  strOutput = strOutput & "<spliter>"
  
  If Not objPOLineRet.Desc Is Nothing Then
    strOutput = strOutput & objPOLineRet.Desc.getValue
  End If
  strOutput = strOutput & "<spliter>"
  
  If Not objPOLineRet.Quantity Is Nothing Then
    strOrdered = objPOLineRet.Quantity.GetAsString
    strOutput = strOutput & strOrdered
  Else
    strOrdered = Empty
  End If
  strOutput = strOutput & "<spliter>"
  
  If Not objPOLineRet.Rate Is Nothing Then
    strOutput = strOutput & objPOLineRet.Rate.GetAsString
  End If
  strOutput = strOutput & "<spliter>"
  
  If Not objPOLineRet.CustomerRef Is Nothing Then
    strOutput = strOutput & objPOLineRet.CustomerRef.FullName.getValue
  End If
  strOutput = strOutput & "<spliter>"
  
  If Not objPOLineRet.Amount Is Nothing Then
    strOutput = strOutput & objPOLineRet.Amount.GetAsString
  End If
  strOutput = strOutput & "<spliter>"
  
  If Not objPOLineRet.ReceivedQuantity Is Nothing Then
    strReceived = objPOLineRet.ReceivedQuantity.GetAsString
    strOutput = strOutput & strReceived
  Else
    strReceived = Empty
  End If
  strOutput = strOutput & "<spliter>"
  
  If Not objPOLineRet.IsManuallyClosed Is Nothing Then
    If objPOLineRet.IsManuallyClosed.getValue = "true" Or _
       (strOrdered <> Empty And strOrdered = strReceived) Or _
       (strOrdered = "0") Then
      strOutput = strOutput & "X"
    End If
  Else
    If (strOrdered <> Empty And strOrdered = strReceived) Or _
       (strOrdered = "0") Then
      strOutput = strOutput & "X"
    End If
  End If
  strOutput = strOutput & "<spliter>"
  
  strOutput = strOutput & _
    objPOLineRet.TxnLineID.getValue & "<spliter>"
  
  If strLineType = Empty Then
    strOutput = strOutput & "Item"
  Else
    strOutput = strOutput & strLineType
  End If
  strOutput = strOutput & "<spliter>"
    
  If Not objPOLineRet.IsManuallyClosed Is Nothing Then
    strOutput = strOutput & _
      objPOLineRet.IsManuallyClosed.getValue
  End If
  strOutput = strOutput & "<spliter>"
  
  POLineString = strOutput
End Function


Private Function POGroupLineString(objPOGroupLineRet As IPurchaseOrderLineGroupRet, _
                                   strLineType As String)

  Dim strOutput As String

  If Not objPOGroupLineRet.ItemGroupRef Is Nothing Then
      strOutput = objPOGroupLineRet.ItemGroupRef.FullName.getValue
  End If
  strOutput = strOutput & "<spliter>"
  
  If Not objPOGroupLineRet.Desc Is Nothing Then
    strOutput = strOutput & objPOGroupLineRet.Desc.getValue
  End If
  strOutput = strOutput & "<spliter>"
  
  If Not objPOGroupLineRet.Quantity Is Nothing Then
    strOutput = strOutput & objPOGroupLineRet.Quantity.GetAsString
  End If
  strOutput = strOutput & "<spliter><spliter><spliter><spliter><spliter><spliter>"
  
  strOutput = strOutput & _
    objPOGroupLineRet.TxnLineID.getValue & "<spliter>"
  
  strOutput = strOutput & strLineType & "<spliter><spliter>"
  
  POGroupLineString = strOutput
End Function


Public Sub QBFC_ClosePO(strTxnID As String, _
                        strEditSequence As String)

  Dim objMsgSetRequest As IMsgSetRequest
  Set objMsgSetRequest = objSessionManager.CreateMsgSetRequest("US", intMajorVersion, intMinorVersion)
  objMsgSetRequest.Attributes.OnError = roeContinue
  
  Dim objPOModify As IPurchaseOrderMod
  Set objPOModify = objMsgSetRequest.AppendPurchaseOrderModRq
  
  objPOModify.TxnID.setValue strTxnID
  objPOModify.EditSequence.setValue strEditSequence
  objPOModify.IsManuallyClosed.setValue True
  
  Dim objMsgSetResponse As IMsgSetResponse
  Set objMsgSetResponse = objSessionManager.DoRequests(objMsgSetRequest)
  
  Dim objResponse As IResponse
  Set objResponse = objMsgSetResponse.ResponseList.GetAt(0)
  
  If objResponse.StatusCode <> 0 Then
    MsgBox "Unexpected Error closing purchase order" & _
      vbCrLf & vbCrLf & "Status Code = " & objResponse.StatusCode & _
      vbCrLf & vbCrLf & "Status Message = " & objResponse.StatusMessage
  Else
    MsgBox "Purchase Order successfully closed"
  End If
End Sub


Public Sub QBFC_ChangePOLine(strAction As String, _
                             strTxnID As String, _
                             strEditSequence As String, _
                             strPOLines() As String, _
                             intSelectedPOLine As Integer, _
                             intGroupLine As Integer)
  
  Dim objMsgSetRequest As IMsgSetRequest
  Set objMsgSetRequest = objSessionManager.CreateMsgSetRequest("US", intMajorVersion, intMinorVersion)
  objMsgSetRequest.Attributes.OnError = roeContinue
  
  Dim objPOModify As IPurchaseOrderMod
  Set objPOModify = objMsgSetRequest.AppendPurchaseOrderModRq
  
  objPOModify.TxnID.setValue strTxnID
  objPOModify.EditSequence.setValue strEditSequence
  
  Dim objPOLine As IPurchaseOrderLineMod
  Dim objPOGroupLine As IPurchaseOrderLineGroupMod
  
  Dim Splits() As String
  
  Dim i As Integer
  Dim booDone As Boolean
  Dim booInGroup As Boolean
  Dim booInSelectedGroup As Boolean
  booInGroup = False
  booInSelectedGroup = False
  booDone = False
  i = 1
  Do
    Splits = Split(strPOLines(i), "<spliter>")

    If Splits(9) = "GroupItem" Then
      Set objPOGroupLine = objPOModify.ORPurchaseOrderLineModList.Append.PurchaseOrderLineGroupMod
      objPOGroupLine.TxnLineID.setValue Splits(8)
      booInGroup = True
    ElseIf (Splits(9) = "GroupSubItem" And _
           (booInSelectedGroup Or strAction = "ReceiveAll")) Then
      Set objPOLine = objPOGroupLine.PurchaseOrderLineModList.Append
      objPOLine.TxnLineID.setValue Splits(8)
    ElseIf Splits(9) = "Item" Then
      Set objPOLine = objPOModify.ORPurchaseOrderLineModList.Append.PurchaseOrderLineMod
      objPOLine.TxnLineID.setValue Splits(8)
    End If
    
    If Not booInSelectedGroup Then
      If i = intGroupLine Then
        booInSelectedGroup = True
      End If
    ElseIf Splits(9) <> "GroupSubItem" Then
      booInSelectedGroup = False
    End If
    
    If i = intSelectedPOLine Or strAction = "ReceiveAll" Then
      If strAction = "Close" Then
        objPOLine.IsManuallyClosed.setValue True
      ElseIf Splits(2) <> Empty Then
        'This is where code for ReceiveAll would have been handled, but
        'it was dropped during implementation only "Close" is used.
      End If
    End If
    
    If i = UBound(strPOLines) Then
      booDone = True
    ElseIf strPOLines(i + 1) = Empty Then
      booDone = True
    End If
    
    i = i + 1
  Loop Until booDone
  
  Dim objMsgSetResponse As IMsgSetResponse
  Set objMsgSetResponse = objSessionManager.DoRequests(objMsgSetRequest)
  
  Dim objResponse As IResponse
  Set objResponse = objMsgSetResponse.ResponseList.GetAt(0)
  
  If objResponse.StatusCode <> 0 Then
    MsgBox "Unexpected Error closing purchase order" & _
      vbCrLf & vbCrLf & "Status Code = " & objResponse.StatusCode & _
      vbCrLf & vbCrLf & "Status Message = " & objResponse.StatusMessage
  Else
    If strAction <> "ReceiveAll" Then
      MsgBox "Purchase order line successfully closed"
    Else
      MsgBox "Purchase order line successfully modified"
    End If
  End If
End Sub


Public Sub QBFC_SetQuantitiesAndBillForRemainingItems( _
             strTxnID As String, _
             strEditSequence As String, _
             strVendor As String, _
             strRefNumber As String, _
             strTxnDate As String, _
             strPOLines() As String, _
             intPOLine As Integer)

  'We're creating a standard request message set node setting the
  'newMessageSetID because this includes a modify and an add request and
  'setting the onError attribute to stopOnError so we don't create a bill
  'if our purchase order modify fails
  Dim objMsgSetRequest As IMsgSetRequest
  Set objMsgSetRequest = objSessionManager.CreateMsgSetRequest("US", intMajorVersion, intMinorVersion)
  objMsgSetRequest.Attributes.OnError = roeStop
  
  Dim objPOModify As IPurchaseOrderMod
  Set objPOModify = objMsgSetRequest.AppendPurchaseOrderModRq
  
  objPOModify.TxnID.setValue strTxnID
  objPOModify.EditSequence.setValue strEditSequence
  
  'Loop through the PO lines and change the ordered quantity to match the
  'received quantity
  Dim objPOLine As IPurchaseOrderLineMod
  Dim objPOGroupLine As IPurchaseOrderLineGroupMod
  
  Dim i As Integer
  Dim booDone As Boolean
  Dim booPOModified As Boolean
  Dim strSplits() As String
  
  booPOModified = False
  i = 1
  Do
    strSplits = Split(strPOLines(i), "<spliter>")
    
    'The quantity ordered or received can be empty, set them to zero if they're empty
    If strSplits(6) = "" Then strSplits(6) = "0"
    If strSplits(2) = "" Then strSplits(2) = "0"
    
    If strSplits(9) <> "GroupSubItem" Then
      Set objPOLine = objPOModify.ORPurchaseOrderLineModList.Append.PurchaseOrderLineMod
    Else
      Set objPOLine = objPOGroupLine.PurchaseOrderLineModList.Append
    End If
    
    If strSplits(9) = "GroupItem" Then
      Set objPOGroupLine = objPOModify.ORPurchaseOrderLineModList.Append.PurchaseOrderLineGroupMod
      objPOGroupLine.TxnLineID.setValue strSplits(8)
    Else
      objPOLine.TxnLineID.setValue strSplits(8)
      
      'If the PO line hasn't been manually closed and the number of items
      'received is less than the number ordered change the number ordered
      'to the number received
      If strSplits(7) <> "X" And Int(strSplits(6)) < Int(strSplits(2)) And _
         ((intPOLine = 0) Or (intPOLine = i)) Then
        objPOLine.Quantity.SetAsString strSplits(6)
        booPOModified = True
      End If
    End If
    
    If i = UBound(strPOLines) Then
      booDone = True
    ElseIf strPOLines(i + 1) = Empty Then
      booDone = True
    Else
      i = i + 1
    End If
  Loop Until booDone
  
  If Not booPOModified Then
    MsgBox "The purchase order selected has either has all lines " & _
      "manually closed or all lines received.  Ending processing " & _
      "on the existing purchase order and not creating a new bill " & _
      "for missing items."
    Exit Sub
  End If
  
  'Now add the bill for the items we modified in the purchase order
  Dim objBillAdd As IBillAdd
  Set objBillAdd = objMsgSetRequest.AppendBillAddRq
  
  objBillAdd.VendorRef.FullName.setValue strVendor
  
  objBillAdd.Memo.setValue _
    "Created by Purchase Order Modify Sample for modified PO " & strRefNumber & _
    "; TxnDate " & strTxnDate & "; TxnID " & strTxnID & _
    " .  Quantity ordered " & _
    "reduced to quantity recieved for one or all lines, differences " & _
    "reflected in this bill."

  Dim objBillLine As IItemLineAdd
  Dim j As Integer
  For j = 1 To i
    strSplits = Split(strPOLines(j), "<spliter>")
    
    'The quantity ordered or received can be empty, set then to zero if they're empty
    If strSplits(6) = "" Then strSplits(6) = "0"
    If strSplits(2) = "" Then strSplits(2) = "0"
    
    If strSplits(9) <> "GroupItem" And strSplits(7) <> "X" And _
       Int(strSplits(6)) < Int(strSplits(2)) And _
       ((intPOLine = 0) Or (intPOLine = j)) Then

      Set objBillLine = objBillAdd.ORItemLineAddList.Append.ItemLineAdd

      objBillLine.ItemRef.FullName.setValue strSplits(0)
      objBillLine.Quantity.setValue Int(strSplits(2)) - Int(strSplits(6))
    
      If strSplits(3) <> Empty Then
        objBillLine.Cost.SetAsString strSplits(3)
      ElseIf strSplits(5) <> Empty Then
        objBillLine.Amount.SetAsString strSplits(5)
      End If
      
      If strSplits(4) <> Empty Then
        objBillLine.CustomerRef.FullName.setValue strSplits(4)
      End If
    End If
  Next j
  
  Dim objMsgSetResponse As IMsgSetResponse
  Set objMsgSetResponse = objSessionManager.DoRequests(objMsgSetRequest)
  
  Dim objResponse As IResponse
  Set objResponse = objMsgSetResponse.ResponseList.GetAt(0)
  
  If objResponse.StatusCode <> 0 Then
    MsgBox "Unexpected error modifying purchase order" & _
      vbCrLf & vbCrLf & "Status Code = " & objResponse.StatusCode & _
      vbCrLf & vbCrLf & "Status Message = " & objResponse.StatusMessage
    Exit Sub
  End If
  
  Set objResponse = objMsgSetResponse.ResponseList.GetAt(1)
  
  If objResponse.StatusCode <> 0 Then
    MsgBox "Unexpected error adding bill" & _
      vbCrLf & vbCrLf & "Status Code = " & objResponse.StatusCode & _
      vbCrLf & vbCrLf & "Status Message = " & objResponse.StatusMessage
    Exit Sub
  Else
    If intPOLine = 0 Then
      MsgBox "Successfully set PO lines and billed for items"
    Else
      MsgBox "Successfully set PO line and billed for items"
    End If
  End If

  Set objResponse = objMsgSetResponse.ResponseList.GetAt(0)
  
  'Get the new edit sequence and Memo for updating the Memo field of the PO
  Dim objPORet As IPurchaseOrderRet
  Set objPORet = objResponse.Detail
  
  Dim strNewEditSequence As String
  strNewEditSequence = objPORet.EditSequence.getValue
  
  Dim strMemo As String
  If Not objPORet.Memo Is Nothing Then
    strMemo = objPORet.Memo.getValue
  Else
    strMemo = ""
  End If
  
  Set objResponse = objMsgSetResponse.ResponseList.GetAt(1)
  
  Dim objBillRet As IBillRet
  Set objBillRet = objResponse.Detail
    
  'Now let's update the PO Memo to include the manipulation we performed
  Dim strBillTxnID As String
  Dim strBillRefNumber As String
  Dim strBillTxnDate As String
  
  strBillTxnID = objBillRet.TxnID.getValue
  strBillTxnDate = Format(objBillRet.TxnDate.getValue, "YYYY-MM-DD")
  
  If Not objBillRet.RefNumber Is Nothing Then
    strBillRefNumber = objBillRet.RefNumber.getValue
  End If
  
  strMemo = strMemo & _
    " >>> Purchase order modified by ""Purchase Order Modify Sample"": " & _
    "Quantity ordered changed to quantity received value "
  If intPOLine = 0 Then
    strMemo = strMemo & "on all lines not closed or fully recieved.  "
  Else
    strMemo = strMemo & "on one line.  "
  End If
  strMemo = strMemo & "Created bill " & strBillRefNumber & _
    " with TxnDate " & strBillTxnDate & "; TxnID " & strBillTxnID & _
    " for these items."

  Set objMsgSetRequest = objSessionManager.CreateMsgSetRequest("US", intMajorVersion, intMinorVersion)
  objMsgSetRequest.Attributes.OnError = roeStop

  Set objPOModify = objMsgSetRequest.AppendPurchaseOrderModRq
  
  objPOModify.TxnID.setValue strTxnID
  objPOModify.EditSequence.setValue strNewEditSequence
  objPOModify.Memo.setValue strMemo

  Set objMsgSetResponse = objSessionManager.DoRequests(objMsgSetRequest)
  
  Set objResponse = objMsgSetResponse.ResponseList.GetAt(0)
  
  If objResponse.StatusCode <> 0 Then
    MsgBox "Unexpected error modifying purchase order memo field" & _
      vbCrLf & vbCrLf & "Status Code = " & objResponse.StatusCode & _
      vbCrLf & vbCrLf & "Status Message = " & objResponse.StatusMessage
    Exit Sub
  End If

End Sub




