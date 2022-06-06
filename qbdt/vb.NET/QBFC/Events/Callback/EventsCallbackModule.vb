Option Strict Off
Option Explicit On
Module QBFCEventsCallbackModule
	Const cAppName As String = "Desktop VB QBFC EventsCallback"
	Const cAppID As String = ""

    ' Queries QuickBooks for all the customers and displays
    Public Sub GetAllCustomers()
		Dim customerFullName As Object
		Dim i As Object
        'Clear the query result text box
        If frm.InvokeRequired Then
                frm.Invoke(Sub()
                               frm.QueryResult.Text = ""
                           End Sub)
            Else
                frm.QueryResult.Text = ""
            End If


            Dim result As String
            result = "QuickBooks list of customers: " & vbCrLf

            'Start a session with QuickBooks.
            ' Create the session manager object
            Dim sessionManager As New QBFC15Lib.QBSessionManager
            Dim bConnectionOpen As Boolean
            Dim bSessionOpen As Boolean
        Try
            ' Connect to QuickBooks and begin a session.
            sessionManager.OpenConnection(cAppID, cAppName)
            bConnectionOpen = True
            sessionManager.BeginSession("", QBFC15Lib.ENOpenMode.omDontCare)
            bSessionOpen = True

            ' Create the message set request object for 3.0 version messages.
            Dim requestMsgSet As QBFC15Lib.IMsgSetRequest
            requestMsgSet = sessionManager.CreateMsgSetRequest("US", 3, 0)
            requestMsgSet.Attributes.OnError = QBFC15Lib.ENRqOnError.roeContinue

            'Add the request to the message set request object.
            Dim customerQuery As QBFC15Lib.ICustomerQuery
            customerQuery = requestMsgSet.AppendCustomerQueryRq

            ' Perform the request and obtain a response from QuickBooks.
            Dim responseMsgSet As QBFC15Lib.IMsgSetResponse
            responseMsgSet = sessionManager.DoRequests(requestMsgSet)

            ' Close the session and connection with QuickBooks.
            If (bConnectionOpen) Then
                sessionManager.EndSession()
                bSessionOpen = False
                sessionManager.CloseConnection()
                bConnectionOpen = False
            End If


            Dim responseList As QBFC15Lib.IResponseList
            responseList = responseMsgSet.ResponseList
            If (responseList Is Nothing) Then
                Exit Sub
            End If

            ' Should only expect 1 response
            Dim response As QBFC15Lib.IResponse
            response = responseList.GetAt(0)

            ' Iterate through the list of Customers


            ' Check the status returned for the response, which will be a CustomerRet.
            Dim customerRetList As QBFC15Lib.ICustomerRetList
            Dim customerRet As QBFC15Lib.ICustomerRet
            If (response.StatusCode = 0) Then
                customerRetList = response.Detail
                ' Should only be 1 CustomerRet object returned
                For i = 0 To customerRetList.Count - 1
                    'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    customerRet = customerRetList.GetAt(i)
                    'Get the fullName
                    'UPGRADE_WARNING: Couldn't resolve default property of object customerFullName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    customerFullName = customerRet.FullName.GetValue
                    'UPGRADE_WARNING: Couldn't resolve default property of object customerFullName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    result = result + customerFullName & vbCrLf
                Next i
            End If

            'Write out the customer list
            If frm.InvokeRequired Then
                frm.Invoke(Sub()
                               frm.QueryResult.Text = result
                               frm.ErrorMsg.Text = ""
                           End Sub)
            Else
                frm.QueryResult.Text = result
                frm.ErrorMsg.Text = ""
            End If
            Exit Sub

        Catch ex As Exception
            ' Close the session and connection with QuickBooks.
            If (bConnectionOpen) Then
                sessionManager.EndSession()
                bSessionOpen = False
                sessionManager.CloseConnection()
                bConnectionOpen = False
            End If

            ' Write error to ErrorMsg Text Box
            If frm.InvokeRequired Then
                frm.Invoke(Sub()
                               frm.ErrorMsg.Text = "HRESULT = " & ex.HResult & vbCrLf & vbCrLf & ex.Message
                           End Sub)
            Else
                frm.ErrorMsg.Text = "HRESULT = " & ex.HResult & vbCrLf & vbCrLf & ex.Message
            End If
        End Try
    End Sub
End Module