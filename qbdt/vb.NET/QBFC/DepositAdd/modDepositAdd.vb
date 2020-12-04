Option Strict Off
Option Explicit On 
Imports Interop.QBFC13

Module modDepositAdd
	'-----------------------------------------------------------
	' Form Module: modDepositAdd
	'
    ' Description: this module contains the code which creates QBFC
	'              messages, exchanges them with QuickBooks, interprets
	'              the responses and loads information into form objects.
	'
	' Copyright © 2002-2013 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	'----------------------------------------------------------
	
	Dim booConnected As Boolean
	
    'Module objects
    Dim qbSessionManager As qbSessionManager
    Dim msgSetRequest As IMsgSetRequest

    Public Sub Connect()

        On Error GoTo Errs
        qbSessionManager = New QBSessionManager()

        qbSessionManager.OpenConnection("", "IDN Deposit Add")

        qbSessionManager.BeginSession("", ENOpenMode.omDontCare)
        booConnected = True

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
            MsgBox("You must have QuickBooks running to run this program.")
            End
        ElseIf Err.Number = &H80040422 Then
            MsgBox("The company file is open in single user mode and there is already another application accessing it.  You need to have the other application end its session with the company file before this program can be run successfully.")
            End
        Else
            MsgBox("HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                        MsgBoxStyle.Critical, _
                        "Error in Connect")
        End If
    End Sub

    Public Function CreateReceivePaymentForDepositQuery() As String
        On Error GoTo Errs

        If (msgSetRequest Is Nothing) Then
            Exit Function
        End If

        ' make sure we do not have any old requests still defined
        msgSetRequest.ClearRequests()

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = ENRqOnError.roeStop

        ' Add the ReceivePaymentToDepositQuery request
        msgSetRequest.AppendReceivePaymentToDepositQueryRq()

        Exit Function

Errs:
        MsgBox("HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    MsgBoxStyle.Critical, _
                    "Error in CreateReceivePaymentForDepositQuery")

    End Function

    Public Sub GetFundsForDeposit(ByRef lstFundsForDeposit As System.Windows.Forms.ListBox)

        On Error GoTo Errs

        Dim msgSetResponse As IMsgSetResponse

        CreateReceivePaymentForDepositQuery()

        ' send the request to QB
        msgSetResponse = qbSessionManager.DoRequests(msgSetRequest)

        FillFundsList(lstFundsForDeposit, msgSetResponse)

        Exit Sub
Errs:
        If Err.Number = &H80040416 Then
            MsgBox("You must have QuickBooks running to run this program.")
            End
        Else
            MsgBox("HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                        MsgBoxStyle.Critical, _
                        "Error in GetFundsForDeposit")
        End If
    End Sub

    Public Sub EndSessionCloseConnection()
        On Error GoTo Errs

        If booConnected Then
            qbSessionManager.EndSession()
            qbSessionManager.CloseConnection()
        End If

        Exit Sub
Errs:
        MsgBox("HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    MsgBoxStyle.Critical, _
                    "Error in EndSessionClosConnection")
    End Sub

    Public Sub FillFundsList(ByRef lstListBox As System.Windows.Forms.ListBox, ByRef msgSetResponse As IMsgSetResponse)
        On Error GoTo Errs

        'Clear the list box
        lstListBox.Items.Clear()

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
                MsgBox("FillFundsList unexpexcted Error - " & vbCrLf & "StatusCode = " & response.StatusCode & vbCrLf & vbCrLf & response.StatusMessage)
            End If
        End If

        ' first make sure we have a response object to handle
        If (response Is Nothing) Or _
            (response.Type Is Nothing) Or _
            (response.Detail Is Nothing) Or _
            (response.Detail.Type Is Nothing) Then
            Exit Sub
        End If

        ' make sure we are processing the ReceivePaymentToDepositQueryRs and 
        ' the ReceivePaymentToDepositRetList responses in this response list
        Dim receivePaymentToDepositRetList As IReceivePaymentToDepositRetList
        Dim responseType As ENResponseType
        Dim responseDetailType As ENObjectType
        responseType = response.Type.GetValue()
        responseDetailType = response.Detail.Type.GetValue()
        If (responseType = ENResponseType.rtReceivePaymentToDepositQueryRs) And _
            (responseDetailType = ENObjectType.otReceivePaymentToDepositRetList) Then
            ' save the response detail in the appropriate object type
            ' since we have first verified the type of the response object
            receivePaymentToDepositRetList = response.Detail
        Else
            ' bail, we do not have the responses we were expecting
            Exit Sub
        End If

        Dim strDisplayLine As String
        Dim count As Short
        Dim index As Short
        Dim receivePaymentToDepositRet As IReceivePaymentToDepositRet
        Dim assignToObjectString As String
        count = receivePaymentToDepositRetList.Count
        For index = 0 To count - 1
            receivePaymentToDepositRet = receivePaymentToDepositRetList.GetAt(index)
            If (Not receivePaymentToDepositRet Is Nothing) Then
                strDisplayLine = ""
                '    "TxnDate"
                If (Not receivePaymentToDepositRet.TxnDate Is Nothing) Then
                    strDisplayLine = receivePaymentToDepositRet.TxnDate.GetValue()
                End If
                '    "Amount"
                If (Not receivePaymentToDepositRet.Amount Is Nothing) Then
                    strDisplayLine = strDisplayLine & "    " & receivePaymentToDepositRet.Amount.GetAsString()
                End If
                '    "TxnType"
                If (Not receivePaymentToDepositRet.TxnType Is Nothing) Then
                    strDisplayLine = strDisplayLine & "    " & receivePaymentToDepositRet.TxnType.GetAsString()
                End If
                '    "TxnID"
                If (Not receivePaymentToDepositRet.TxnID Is Nothing) Then
                    strDisplayLine = strDisplayLine & "    " & receivePaymentToDepositRet.TxnID.GetValue()
                End If
                lstListBox.Items.Add(strDisplayLine)
            End If

        Next

        Exit Sub

Errs:
        MsgBox("HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    MsgBoxStyle.Critical, _
                    "Error in FillFundsList")

    End Sub

    Public Sub DepositFunds(ByRef strFundsInfo As String)
        On Error GoTo Errs

        Dim strTxnID As String

        strTxnID = Right(strFundsInfo, Len(strFundsInfo) - InStrRev(strFundsInfo, " "))

        If (msgSetRequest Is Nothing) Then
            Exit Sub
        End If

        ' make sure we do not have any old requests still defined
        msgSetRequest.ClearRequests()

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = ENRqOnError.roeStop

        ' Add the DepositAdd request
        Dim depositAdd As IDepositAdd
        depositAdd = msgSetRequest.AppendDepositAddRq()

        'Add the FullName element
        depositAdd.DepositToAccountRef.FullName.SetValue("Checking")

        'Add the PaymentTxnID element
        Dim depositLineAdd As IDepositLineAdd
        depositLineAdd = depositAdd.DepositLineAddList.Append()
        depositLineAdd.ORDepositLineAdd.PaymentLine.PaymentTxnID.SetValue(strTxnID)

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
                MsgBox("DepositFunds unexpexcted Error - " & vbCrLf & "StatusCode = " & response.StatusCode & vbCrLf & vbCrLf & response.StatusMessage)
            Else
                MsgBox("The funds were successfully deposited in Checking")
            End If
        End If

        Exit Sub

Errs:
        MsgBox("HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    MsgBoxStyle.Critical, _
                    "Error in DepositFunds")
    End Sub
End Module
