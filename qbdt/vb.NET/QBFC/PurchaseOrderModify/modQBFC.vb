Option Strict Off
Option Explicit On
Imports Interop.QBFC15

Module modQBFC
    '----------------------------------------------------------
    ' Copyright © 2003-2013 Intuit Inc. All rights reserved.
    ' Use is subject to the terms specified at:
    '      http://developer.intuit.com/legal/devsite_tos.html
    '
    ' Added values for intMajorVersion and intMinorVersion
    '----------------------------------------------------------


    Dim objSessionManager As New QBSessionManager

    Dim booSessionBegun As Boolean
	
	Dim intMajorVersion As Short
	Dim intMinorVersion As Short
	
	Public Sub QBFC_OpenConnectionBeginSession()
		
		On Error GoTo Errs
		
		objSessionManager.OpenConnection("", "IDN PO Modify Sample Application")
        objSessionManager.BeginSession("", ENOpenMode.omDontCare)
        booSessionBegun = True
		intMajorVersion = 3
		intMinorVersion = 0
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
	
	
	
	
	Public Sub QBFC_LoadPOInfoArray(ByRef strPOInfo() As String, ByRef dateFrom As Date, ByRef dateTo As Date, ByRef booGiveWarning As Boolean)

        Dim objMsgSetRequest As IMsgSetRequest
        objMsgSetRequest = objSessionManager.CreateMsgSetRequest("US", 3, 0)

        Dim objGeneralDetailReportQuery As IGeneralDetailReportQuery
        objGeneralDetailReportQuery = objMsgSetRequest.AppendGeneralDetailReportQueryRq

        objGeneralDetailReportQuery.ORReportPeriod.ReportPeriod.FromReportDate.SetValue(dateFrom)
        objGeneralDetailReportQuery.ORReportPeriod.ReportPeriod.ToReportDate.SetValue(dateTo)
        objGeneralDetailReportQuery.GeneralDetailReportType.SetValue(ENGeneralDetailReportType.gdrtOpenPOs)

        objGeneralDetailReportQuery.IncludeColumnList.Add(ENIncludeColumn.icRefNumber)
        objGeneralDetailReportQuery.IncludeColumnList.Add(ENIncludeColumn.icDate)
        objGeneralDetailReportQuery.IncludeColumnList.Add(ENIncludeColumn.icName)
        objGeneralDetailReportQuery.IncludeColumnList.Add(ENIncludeColumn.icDeliveryDate)
        objGeneralDetailReportQuery.IncludeColumnList.Add(ENIncludeColumn.icAmount)

        If SupportsModify() Then
            objGeneralDetailReportQuery.IncludeColumnList.Add(ENIncludeColumn.icTxnID)
        End If

        Dim objMsgSetResponse As IMsgSetResponse
        objMsgSetResponse = objSessionManager.DoRequests(objMsgSetRequest)

        Dim objResponse As IResponse
        objResponse = objMsgSetResponse.ResponseList.GetAt(0)

        If objResponse.StatusCode <> 0 Then
            MsgBox("Error getting Open Purchase Orders" & vbCrLf & vbCrLf & "Error = " & Hex(objResponse.StatusCode) & vbCrLf & vbCrLf & objResponse.StatusMessage)
            strPOInfo(0) = "<spliter><spliter>There were no open purchase orders returned<spliter><spliter>"
            Exit Sub
        End If

        Dim objGeneralDetailReport As IReportRet
        objGeneralDetailReport = objResponse.Detail

        If objGeneralDetailReport.NumRows.GetValue < 3 Then
            strPOInfo(0) = "<spliter><spliter>There were no open purchase orders returned<spliter><spliter>"
            Exit Sub
        End If

        Dim i As Short
        Dim j As Short
        If objGeneralDetailReport.NumRows.GetValue < 32 Then
            j = objGeneralDetailReport.NumRows.GetValue - 2
        Else
            j = 30
        End If

        Dim objReportDataList As IORReportDataList
        objReportDataList = objGeneralDetailReport.ReportData.ORReportDataList

        Dim objReportRow As IORReportData
        Dim objDataRow As IDataRow
        Dim objColDataList As IColDataList
        Dim objColData As IColData

        For i = 1 To j
            objReportRow = objReportDataList.GetAt(i)
            objDataRow = objReportRow.DataRow
            objColDataList = objDataRow.ColDataList

            strPOInfo(i) = objColDataList.GetAt(0).value.GetValue & "<spliter>" & objColDataList.GetAt(1).value.GetValue & "<spliter>" & objColDataList.GetAt(2).value.GetValue & "<spliter>" & objColDataList.GetAt(3).value.GetValue & "<spliter>" & objColDataList.GetAt(4).value.GetValue & "<spliter>"

            If SupportsModify() Then
                strPOInfo(i) = strPOInfo(i) & objColDataList.GetAt(5).value.GetValue
            End If
        Next
    End Sub


    Public Sub QBFC_GetPOInfo(ByRef strTxnID As String, ByRef strEditSequence As String, ByRef strRefNumber As String, ByRef strTxnDate As String, ByRef strVendor As String, ByRef strPOLines() As String, ByRef strSelectedPOInfo As String)

        Dim strPOInfoSplits() As String
        strPOInfoSplits = Split(strSelectedPOInfo, "<spliter>")

        Dim objMsgSetRequest As IMsgSetRequest
        objMsgSetRequest = objSessionManager.CreateMsgSetRequest("US", intMajorVersion, intMinorVersion)
        objMsgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

        Dim objPOQuery As IPurchaseOrderQuery
        objPOQuery = objMsgSetRequest.AppendPurchaseOrderQueryRq

        objPOQuery.IncludeLineItems.SetValue(True)
        objPOQuery.ORTxnQuery.TxnFilter.EntityFilter.OREntityFilter.FullNameList.Add(strPOInfoSplits(2))
        objPOQuery.ORTxnQuery.TxnFilter.ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter.TxnDateFilter.FromTxnDate.SetValue(CDate(strPOInfoSplits(1)))
        objPOQuery.ORTxnQuery.TxnFilter.ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter.TxnDateFilter.ToTxnDate.SetValue(CDate(strPOInfoSplits(1)))
        objPOQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberRangeFilter.FromRefNumber.SetValue(strPOInfoSplits(0))
        objPOQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberRangeFilter.ToRefNumber.SetValue(strPOInfoSplits(0))

        Dim objMsgSetResponse As IMsgSetResponse
        objMsgSetResponse = objSessionManager.DoRequests(objMsgSetRequest)

        Dim objResponse As IResponse
        objResponse = objMsgSetResponse.ResponseList.GetAt(0)

        If objResponse.StatusCode <> 0 Then
            MsgBox("Unexpected Error getting purchase order information " & vbCrLf & vbCrLf & "Status Code = " & objResponse.StatusCode & vbCrLf & vbCrLf & "Status Message = " & objResponse.StatusMessage)
            strPOLines(1) = CStr(Nothing)
            Exit Sub
        End If

        Dim objPORetList As IPurchaseOrderRetList
        objPORetList = objResponse.Detail

        Dim objPORet As IPurchaseOrderRet
        objPORet = objPORetList.GetAt(0)

        If objPORet.ORPurchaseOrderLineRetList Is Nothing Then
            strPOLines(1) = CStr(Nothing)
            Exit Sub
        End If

        strTxnID = objPORet.TxnID.GetValue
        strEditSequence = objPORet.EditSequence.GetValue
        strTxnDate = VB6.Format(objPORet.TxnDate.GetValue, "YYYY-MM-DD")

        If Not objPORet.RefNumber Is Nothing Then
            strRefNumber = objPORet.RefNumber.GetValue
        End If

        If Not objPORet.VendorRef Is Nothing Then
            strVendor = objPORet.VendorRef.FullName.GetValue
        End If

        Dim objPOLineRetList As IORPurchaseOrderLineRetList
        objPOLineRetList = objPORet.ORPurchaseOrderLineRetList


        Dim booDone As Boolean
        Dim i As Short
        Dim j As Short
        Dim k As Short
        booDone = False

        Dim objPOLine As IPurchaseOrderLineRet
        Dim objPOGroupLine As IPurchaseOrderLineGroupRet

        k = 1
        For i = 0 To objPOLineRetList.Count - 1
            If objPOLineRetList.GetAt(i).ortype = ENORPurchaseOrderLineRet.orpolrPurchaseOrderLineRet Then
                objPOLine = objPOLineRetList.GetAt(i).PurchaseOrderLineRet
                'UPGRADE_WARNING: Couldn't resolve default property of object POLineString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strPOLines(k) = POLineString(objPOLine, "")
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object POGroupLineString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strPOLines(k) = POGroupLineString(objPOGroupLine, "")
                If Not objPOGroupLine.PurchaseOrderLineRetList Is Nothing Then
                    For j = 0 To objPOGroupLine.PurchaseOrderLineRetList.Count - 1
                        k = k + 1
                        'UPGRADE_WARNING: Couldn't resolve default property of object POLineString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        strPOLines(k) = POLineString(objPOGroupLine.PurchaseOrderLineRetList.GetAt(j), "GroupSubItem")
                    Next j
                End If
            End If

            k = k + 1
        Next i
    End Sub


    Private Function POLineString(ByRef objPOLineRet As IPurchaseOrderLineRet, ByRef strLineType As String) As Object

        Dim strOrdered As String
        Dim strReceived As String

        Dim strOutput As String

        If Not objPOLineRet.ItemRef Is Nothing Then
            strOutput = objPOLineRet.ItemRef.FullName.GetValue
        End If
        strOutput = strOutput & "<spliter>"

        If Not objPOLineRet.Desc Is Nothing Then
            strOutput = strOutput & objPOLineRet.Desc.GetValue
        End If
        strOutput = strOutput & "<spliter>"

        If Not objPOLineRet.Quantity Is Nothing Then
            strOrdered = objPOLineRet.Quantity.GetAsString
            strOutput = strOutput & strOrdered
        Else
            strOrdered = CStr(Nothing)
        End If
        strOutput = strOutput & "<spliter>"

        If Not objPOLineRet.Rate Is Nothing Then
            strOutput = strOutput & objPOLineRet.Rate.GetAsString
        End If
        strOutput = strOutput & "<spliter>"

        If Not objPOLineRet.CustomerRef Is Nothing Then
            strOutput = strOutput & objPOLineRet.CustomerRef.FullName.GetValue
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
            strReceived = CStr(Nothing)
        End If
        strOutput = strOutput & "<spliter>"

        If Not objPOLineRet.IsManuallyClosed Is Nothing Then
            'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            If objPOLineRet.IsManuallyClosed.GetValue = CBool("true") Or (Not IsNothing(strOrdered) And strOrdered = strReceived) Or (strOrdered = "0") Then
                strOutput = strOutput & "X"
            End If
        Else
            'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            If (Not IsNothing(strOrdered) And strOrdered = strReceived) Or (strOrdered = "0") Then
                strOutput = strOutput & "X"
            End If
        End If
        strOutput = strOutput & "<spliter>"

        strOutput = strOutput & objPOLineRet.TxnLineID.GetValue & "<spliter>"

        'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If IsNothing(strLineType) Then
            strOutput = strOutput & "Item"
        Else
            strOutput = strOutput & strLineType
        End If
        strOutput = strOutput & "<spliter>"

        If Not objPOLineRet.IsManuallyClosed Is Nothing Then
            strOutput = strOutput & objPOLineRet.IsManuallyClosed.GetValue
        End If
        strOutput = strOutput & "<spliter>"

        'UPGRADE_WARNING: Couldn't resolve default property of object POLineString. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        POLineString = strOutput
    End Function


    Private Function POGroupLineString(ByRef objPOGroupLineRet As IPurchaseOrderLineGroupRet, ByRef strLineType As String) As Object

        Dim strOutput As String

        If Not objPOGroupLineRet.ItemGroupRef Is Nothing Then
            strOutput = objPOGroupLineRet.ItemGroupRef.FullName.GetValue
        End If
        strOutput = strOutput & "<spliter>"

        If Not objPOGroupLineRet.Desc Is Nothing Then
            strOutput = strOutput & objPOGroupLineRet.Desc.GetValue
        End If
        strOutput = strOutput & "<spliter>"

        If Not objPOGroupLineRet.Quantity Is Nothing Then
            strOutput = strOutput & objPOGroupLineRet.Quantity.GetAsString
        End If
        strOutput = strOutput & "<spliter><spliter><spliter><spliter><spliter><spliter>"

        strOutput = strOutput & objPOGroupLineRet.TxnLineID.GetValue & "<spliter>"

        strOutput = strOutput & strLineType & "<spliter><spliter>"

        'UPGRADE_WARNING: Couldn't resolve default property of object POGroupLineString. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        POGroupLineString = strOutput
    End Function


    Public Sub QBFC_ClosePO(ByRef strTxnID As String, ByRef strEditSequence As String)

        Dim objMsgSetRequest As IMsgSetRequest
        objMsgSetRequest = objSessionManager.CreateMsgSetRequest("US", intMajorVersion, intMinorVersion)
        objMsgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

        Dim objPOModify As IPurchaseOrderMod
        objPOModify = objMsgSetRequest.AppendPurchaseOrderModRq

        objPOModify.TxnID.SetValue(strTxnID)
        objPOModify.EditSequence.SetValue(strEditSequence)
        objPOModify.IsManuallyClosed.SetValue(True)

        Dim objMsgSetResponse As IMsgSetResponse
        objMsgSetResponse = objSessionManager.DoRequests(objMsgSetRequest)

        Dim objResponse As IResponse
        objResponse = objMsgSetResponse.ResponseList.GetAt(0)

        If objResponse.StatusCode <> 0 Then
            MsgBox("Unexpected Error closing purchase order" & vbCrLf & vbCrLf & "Status Code = " & objResponse.StatusCode & vbCrLf & vbCrLf & "Status Message = " & objResponse.StatusMessage)
        Else
            MsgBox("Purchase Order successfully closed")
        End If
    End Sub


    Public Sub QBFC_ChangePOLine(ByRef strAction As String, ByRef strTxnID As String, ByRef strEditSequence As String, ByRef strPOLines() As String, ByRef intSelectedPOLine As Short, ByRef intGroupLine As Short)

        Dim objMsgSetRequest As IMsgSetRequest
        objMsgSetRequest = objSessionManager.CreateMsgSetRequest("US", intMajorVersion, intMinorVersion)
        objMsgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

        Dim objPOModify As IPurchaseOrderMod
        objPOModify = objMsgSetRequest.AppendPurchaseOrderModRq

        objPOModify.TxnID.SetValue(strTxnID)
        objPOModify.EditSequence.SetValue(strEditSequence)

        Dim objPOLine As IPurchaseOrderLineMod
        Dim objPOGroupLine As IPurchaseOrderLineGroupMod

        Dim Splits() As String

        Dim i As Short
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
                objPOGroupLine = objPOModify.ORPurchaseOrderLineModList.Append.PurchaseOrderLineGroupMod
                objPOGroupLine.TxnLineID.SetValue(Splits(8))
                booInGroup = True
            ElseIf (Splits(9) = "GroupSubItem" And (booInSelectedGroup Or strAction = "ReceiveAll")) Then
                objPOLine = objPOGroupLine.PurchaseOrderLineModList.Append
                objPOLine.TxnLineID.SetValue(Splits(8))
            ElseIf Splits(9) = "Item" Then
                objPOLine = objPOModify.ORPurchaseOrderLineModList.Append.PurchaseOrderLineMod
                objPOLine.TxnLineID.SetValue(Splits(8))
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
                    objPOLine.IsManuallyClosed.SetValue(True)
                    'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                ElseIf Not IsNothing(Splits(2)) Then
                    'This is where code for ReceiveAll would have been handled, but
                    'it was dropped during implementation only "Close" is used.
                End If
            End If

            If i = UBound(strPOLines) Then
                booDone = True
                'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            ElseIf IsNothing(strPOLines(i + 1)) Then
                booDone = True
            End If

            i = i + 1
        Loop Until booDone

        Dim objMsgSetResponse As IMsgSetResponse
        objMsgSetResponse = objSessionManager.DoRequests(objMsgSetRequest)

        Dim objResponse As IResponse
        objResponse = objMsgSetResponse.ResponseList.GetAt(0)

        If objResponse.StatusCode <> 0 Then
            MsgBox("Unexpected Error closing purchase order" & vbCrLf & vbCrLf & "Status Code = " & objResponse.StatusCode & vbCrLf & vbCrLf & "Status Message = " & objResponse.StatusMessage)
        Else
            If strAction <> "ReceiveAll" Then
                MsgBox("Purchase order line successfully closed")
            Else
                MsgBox("Purchase order line successfully modified")
            End If
        End If
    End Sub


    Public Sub QBFC_SetQuantitiesAndBillForRemainingItems(ByRef strTxnID As String, ByRef strEditSequence As String, ByRef strVendor As String, ByRef strRefNumber As String, ByRef strTxnDate As String, ByRef strPOLines() As String, ByRef intPOLine As Short)

        'We're creating a standard request message set node setting the
        'newMessageSetID because this includes a modify and an add request and
        'setting the onError attribute to stopOnError so we don't create a bill
        'if our purchase order modify fails
        Dim objMsgSetRequest As IMsgSetRequest
        objMsgSetRequest = objSessionManager.CreateMsgSetRequest("US", intMajorVersion, intMinorVersion)
        objMsgSetRequest.Attributes.OnError = ENRqOnError.roeStop

        Dim objPOModify As IPurchaseOrderMod
        objPOModify = objMsgSetRequest.AppendPurchaseOrderModRq

        objPOModify.TxnID.SetValue(strTxnID)
        objPOModify.EditSequence.SetValue(strEditSequence)

        'Loop through the PO lines and change the ordered quantity to match the
        'received quantity
        Dim objPOLine As IPurchaseOrderLineMod
        Dim objPOGroupLine As IPurchaseOrderLineGroupMod

        Dim i As Short
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
                objPOLine = objPOModify.ORPurchaseOrderLineModList.Append.PurchaseOrderLineMod
            Else
                objPOLine = objPOGroupLine.PurchaseOrderLineModList.Append
            End If

            If strSplits(9) = "GroupItem" Then
                objPOGroupLine = objPOModify.ORPurchaseOrderLineModList.Append.PurchaseOrderLineGroupMod
                objPOGroupLine.TxnLineID.SetValue(strSplits(8))
            Else
                objPOLine.TxnLineID.SetValue(strSplits(8))

                'If the PO line hasn't been manually closed and the number of items
                'received is less than the number ordered change the number ordered
                'to the number received
                If strSplits(7) <> "X" And Int(CDbl(strSplits(6))) < Int(CDbl(strSplits(2))) And ((intPOLine = 0) Or (intPOLine = i)) Then
                    objPOLine.Quantity.SetAsString(strSplits(6))
                    booPOModified = True
                End If
            End If

            If i = UBound(strPOLines) Then
                booDone = True
                'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            ElseIf String.IsNullOrEmpty(strPOLines(i + 1)) Then
                booDone = True
            Else
                i = i + 1
            End If
        Loop Until booDone

        If Not booPOModified Then
            MsgBox("The purchase order selected has either has all lines " & "manually closed or all lines received.  Ending processing " & "on the existing purchase order and not creating a new bill " & "for missing items.")
            Exit Sub
        End If

        'Now add the bill for the items we modified in the purchase order
        Dim objBillAdd As IBillAdd
        objBillAdd = objMsgSetRequest.AppendBillAddRq

        objBillAdd.VendorRef.FullName.SetValue(strVendor)

        objBillAdd.Memo.SetValue("Created by Purchase Order Modify Sample for modified PO " & strRefNumber & "; TxnDate " & strTxnDate & "; TxnID " & strTxnID & " .  Quantity ordered " & "reduced to quantity recieved for one or all lines, differences " & "reflected in this bill.")

        Dim objBillLine As IItemLineAdd
        Dim j As Short
        For j = 1 To i
            strSplits = Split(strPOLines(j), "<spliter>")

            'The quantity ordered or received can be empty, set then to zero if they're empty
            If strSplits(6) = "" Then strSplits(6) = "0"
            If strSplits(2) = "" Then strSplits(2) = "0"

            If strSplits(9) <> "GroupItem" And strSplits(7) <> "X" And Int(CDbl(strSplits(6))) < Int(CDbl(strSplits(2))) And ((intPOLine = 0) Or (intPOLine = j)) Then

                objBillLine = objBillAdd.ORItemLineAddList.Append.ItemLineAdd

                objBillLine.ItemRef.FullName.SetValue(strSplits(0))
                objBillLine.Quantity.SetValue(Int(CDbl(strSplits(2))) - Int(CDbl(strSplits(6))))

                'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                If Not String.IsNullOrEmpty(strSplits(3)) Then
                    objBillLine.Cost.SetAsString(strSplits(3))
                    'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                ElseIf Not String.IsNullOrEmpty(strSplits(5)) Then
                    objBillLine.Amount.SetAsString(strSplits(5))
                End If

                'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                If Not String.IsNullOrEmpty(strSplits(4)) Then
                    objBillLine.CustomerRef.FullName.SetValue(strSplits(4))
                End If
            End If
        Next j

        Dim objMsgSetResponse As IMsgSetResponse
        objMsgSetResponse = objSessionManager.DoRequests(objMsgSetRequest)

        Dim objResponse As IResponse
        objResponse = objMsgSetResponse.ResponseList.GetAt(0)

        If objResponse.StatusCode <> 0 Then
            MsgBox("Unexpected error modifying purchase order" & vbCrLf & vbCrLf & "Status Code = " & objResponse.StatusCode & vbCrLf & vbCrLf & "Status Message = " & objResponse.StatusMessage)
            Exit Sub
        End If

        objResponse = objMsgSetResponse.ResponseList.GetAt(1)

        If objResponse.StatusCode <> 0 Then
            MsgBox("Unexpected error adding bill" & vbCrLf & vbCrLf & "Status Code = " & objResponse.StatusCode & vbCrLf & vbCrLf & "Status Message = " & objResponse.StatusMessage)
            Exit Sub
        Else
            If intPOLine = 0 Then
                MsgBox("Successfully set PO lines and billed for items")
            Else
                MsgBox("Successfully set PO line and billed for items")
            End If
        End If

        objResponse = objMsgSetResponse.ResponseList.GetAt(0)

        'Get the new edit sequence and Memo for updating the Memo field of the PO
        Dim objPORet As IPurchaseOrderRet
        objPORet = objResponse.Detail

        Dim strNewEditSequence As String
        strNewEditSequence = objPORet.EditSequence.GetValue

        Dim strMemo As String
        If Not objPORet.Memo Is Nothing Then
            strMemo = objPORet.Memo.GetValue
        Else
            strMemo = ""
        End If

        objResponse = objMsgSetResponse.ResponseList.GetAt(1)

        Dim objBillRet As IBillRet
        objBillRet = objResponse.Detail

        'Now let's update the PO Memo to include the manipulation we performed
        Dim strBillTxnID As String
        Dim strBillRefNumber As String
        Dim strBillTxnDate As String

        strBillTxnID = objBillRet.TxnID.GetValue
        strBillTxnDate = VB6.Format(objBillRet.TxnDate.GetValue, "YYYY-MM-DD")

        If Not objBillRet.RefNumber Is Nothing Then
            strBillRefNumber = objBillRet.RefNumber.GetValue
        End If

        strMemo = strMemo & " >>> Purchase order modified by ""Purchase Order Modify Sample"": " & "Quantity ordered changed to quantity received value "
        If intPOLine = 0 Then
            strMemo = strMemo & "on all lines not closed or fully recieved.  "
        Else
            strMemo = strMemo & "on one line.  "
        End If
        strMemo = strMemo & "Created bill " & strBillRefNumber & " with TxnDate " & strBillTxnDate & "; TxnID " & strBillTxnID & " for these items."

        objMsgSetRequest = objSessionManager.CreateMsgSetRequest("US", intMajorVersion, intMinorVersion)
        objMsgSetRequest.Attributes.OnError = ENRqOnError.roeStop

        objPOModify = objMsgSetRequest.AppendPurchaseOrderModRq
		
		objPOModify.TxnID.setValue(strTxnID)
		objPOModify.EditSequence.setValue(strNewEditSequence)
		objPOModify.Memo.setValue(strMemo)
		
		objMsgSetResponse = objSessionManager.DoRequests(objMsgSetRequest)
		
		objResponse = objMsgSetResponse.ResponseList.GetAt(0)
		
		If objResponse.StatusCode <> 0 Then
			MsgBox("Unexpected error modifying purchase order memo field" & vbCrLf & vbCrLf & "Status Code = " & objResponse.StatusCode & vbCrLf & vbCrLf & "Status Message = " & objResponse.StatusMessage)
			Exit Sub
		End If
		
	End Sub
End Module