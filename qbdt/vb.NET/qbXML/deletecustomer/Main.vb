Option Strict Off
Option Explicit On
Friend Class Mainfrm
    Inherits System.Windows.Forms.Form
    ' Main.frm
    '
    ' This is the main form of the Delete Customer sample application which
    ' uses QuickBooks SDK 2.0.  This form contains most of the data for the User
    ' Interface. The code for communicating with QuickBooks is handled in a
    ' separate module qbModule.bas. Likewise, the XML generating and parsing code
    ' is handled by the separate module XMLBuilder.
    '
    ' This sample application illustrates how to use the DeleteCustomer
    ' functionality of the SDK.  It shows, first, how to delete
    ' a customer who is inactive.  Next it allows the user to attempt to delete
    ' a customer who is currently active and view the errors reported by QuickBooks
    ' in this case.
    '
    '
    ' Created: February, 2002
    ' Updated to SDK 2.0 August, 2002
    '
    ' Copyright © 2021-2022 Intuit Inc. All rights reserved.
    ' Use is subject to the terms specified at:
    '      http://developer.intuit.com/legal/devsite_tos.html
    '
    '----------------------------------------------------------


    ' Declare variables for Customer info
    Dim custName As String
    Dim ListID As String

    ' Declare variables for Response info
    Dim resCustName As String
    Dim resCustFullName As String
    Dim resListID As String


    ' Submit button for adding a customer
    Private Sub Comm_Submit_Add_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Comm_Submit_Add.Click

        On Error GoTo ErrHandler

        ' Initialize variables
        requestXML = ""
        responseXML = ""

        ' Get input data
        If Not CollectFormData("Add") Then
            Exit Sub
        End If

        If Not blnIsOpenConnection Then
            'Call OpenConnection to establish a QuickBooks connection and session
            If Not OpenConnection() Then
                Exit Sub
            End If
        End If

        ' Build the request XML to add a customer
        requestXML = CreateCustomerAddRq()


        ' Send request to QuickBooks
        DoRequest()


        ' Parse response
        If Not ParseResponseXML("CustomerAddRs") Then
            CloseConnection()
            Exit Sub
        End If


        Text_ListID.Text = resListID

        ' Display the result
        MsgBox("Customer " & custName & " has been successfully created." & vbCr & vbCr & "  ListID = " & resListID & vbCr & "  FullName = " & resCustFullName & vbCr & vbCr & "The ListID test box has been automatically populated with this value.", MsgBoxStyle.OkOnly, "Success")

        Exit Sub

ErrHandler:
        MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
        CloseConnection()
        Exit Sub

    End Sub


    ' Query QuickBooks for open Invoices, pull out a customer ListID from
    ' one of them, and fill in the ListID box with this value.
    Private Sub Comm_Submit_Find_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Comm_Submit_Find.Click

        On Error GoTo ErrHandler

        ' Initialize
        requestXML = ""
        responseXML = ""


        If Not blnIsOpenConnection Then
            'Call OpenConnection
            If Not OpenConnection() Then
                Exit Sub
            End If
        End If

        ' Build the request XML to query QuickBooks for open Invoices
        requestXML = CreateInvoiceQueryRq()

        ' Send request to QuickBooks
        DoRequest()

        ' Parse response
        If Not ParseResponseXML("InvoiceQueryRs") Then
            Exit Sub
        End If

        Me.Text_ListID.Text = resListID
        Me.Text_CustomerName.Text = ""

        ' Display the result
        If resListID <> "" Then
            MsgBox("There is an invoice open for customer with ListID " & resListID & ". The ListID test box has been automatically populated with this value.", MsgBoxStyle.OkOnly, "Success")
        Else
            MsgBox("No open invoices were found in QuickBooks data file.", MsgBoxStyle.OkOnly, "Success")
        End If

        Exit Sub

ErrHandler:
        MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
        CloseConnection()
        Exit Sub

    End Sub


    ' Submit button for deleting a customer
    Private Sub Comm_Submit_Remove_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Comm_Submit_Remove.Click
        On Error GoTo ErrHandler

        ' Get input data
        If Not CollectFormData("Delete") Then
            Exit Sub
        End If

        If Not blnIsOpenConnection Then
            'Call OpenConnection
            If Not OpenConnection() Then
                Exit Sub
            End If
        End If

        ' Build the request XML to delete a customer
        requestXML = CreateCustomerDelRq()

        ' Send request to QuickBooks
        DoRequest()


        ' Parse response
        If Not ParseResponseXML("ListDelRs") Then
            Exit Sub
        End If

        ' Display the result
        MsgBox("Customer " & custName & " has been successfully deleted.",  , "Success")

        Exit Sub

ErrHandler:
        MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
        CloseConnection()
        Exit Sub

    End Sub


    ' View qbXML Request button
    Private Sub Comm_View_Req_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Comm_View_Req.Click

        On Error GoTo ErrHandler

        Dim reqFrm As DisplayXML
        If requestXML <> "" Then

            reqFrm = New DisplayXML

            reqFrm.Text_Content.Text = requestXML
            reqFrm.Text = "Request XML"
            reqFrm.Show()

        Else
            MsgBox("Request is empty.  Please add a customer first", MsgBoxStyle.Information, "Request XML")
        End If
        Exit Sub


ErrHandler:
        MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
        Exit Sub

    End Sub

    ' View Response Button
    Private Sub Comm_View_Res_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Comm_View_Res.Click

        On Error GoTo ErrHandler

        Dim resFrm As New DisplayXML
        Dim tmpResponseXML As String
        If responseXML <> "" Then


            ' replace Lf to CrLf, this is for display only
            tmpResponseXML = Replace(responseXML, vbLf, vbCrLf)
            resFrm.Text_Content.Text = tmpResponseXML
            resFrm.Text = "Response XML"
            resFrm.Show()

        Else
            MsgBox("Response is empty.  Please add a customer first", MsgBoxStyle.Information, "Response XML")
        End If

        Exit Sub

ErrHandler:
        MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
        Exit Sub

    End Sub

    ' Exit button
    Private Sub Comm_Exit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Comm_Exit.Click
        Me.Close() ' close the window
    End Sub


    ' Create CustomerAdd request
    Private Function CreateCustomerAddRq() As String

        Dim doc As New MSXML2.DOMDocument60
        Dim QBXML As MSXML2.IXMLDOMNode
        Dim MsgsRq As MSXML2.IXMLDOMElement
        Dim CustomerAdd As MSXML2.IXMLDOMElement
        Dim Name_Renamed As MSXML2.IXMLDOMElement

        QBXML = doc.appendChild(doc.createElement("QBXML"))
        MsgsRq = QBXML.appendChild(doc.createElement("QBXMLMsgsRq"))
        MsgsRq.setAttribute("onError", "continueOnError")
        CustomerAdd = MsgsRq.appendChild(doc.createElement("CustomerAddRq"))
        CustomerAdd = CustomerAdd.appendChild(doc.createElement("CustomerAdd"))
        Name_Renamed = CustomerAdd.appendChild(doc.createElement("Name"))
        Name_Renamed.appendChild(doc.createTextNode(custName))
        CreateCustomerAddRq = doc.xml

    End Function

    ' Create an invoice query request
    Private Function CreateInvoiceQueryRq() As String

        Dim doc As New MSXML2.DOMDocument60
        Dim QBXML As MSXML2.IXMLDOMNode
        Dim MsgsRq As MSXML2.IXMLDOMElement
        Dim InvoiceQuery As MSXML2.IXMLDOMElement
        Dim PaidStatus As MSXML2.IXMLDOMElement
        Dim ListID As MSXML2.IXMLDOMElement

        QBXML = doc.appendChild(doc.createElement("QBXML"))
        MsgsRq = QBXML.appendChild(doc.createElement("QBXMLMsgsRq"))
        MsgsRq.setAttribute("onError", "continueOnError")
        InvoiceQuery = MsgsRq.appendChild(doc.createElement("InvoiceQueryRq"))
        PaidStatus = InvoiceQuery.appendChild(doc.createElement("PaidStatus"))
        PaidStatus.appendChild(doc.createTextNode("NotPaidOnly"))
        CreateInvoiceQueryRq = doc.xml
    End Function


    ' Create Customer Delete request
    Private Function CreateCustomerDelRq() As String

        Dim doc As New MSXML2.DOMDocument60
        Dim QBXML As MSXML2.IXMLDOMNode
        Dim MsgsRq As MSXML2.IXMLDOMElement
        Dim ListDel As MSXML2.IXMLDOMElement
        Dim dataElement As MSXML2.IXMLDOMElement

        QBXML = doc.appendChild(doc.createElement("QBXML"))
        MsgsRq = QBXML.appendChild(doc.createElement("QBXMLMsgsRq"))
        MsgsRq.setAttribute("onError", "continueOnError")
        ListDel = MsgsRq.appendChild(doc.createElement("ListDelRq"))
        dataElement = ListDel.appendChild(doc.createElement("ListDelType"))
        dataElement.appendChild(doc.createTextNode("Customer"))
        dataElement = ListDel.appendChild(doc.createElement("ListID"))
        dataElement.appendChild(doc.createTextNode(ListID))
        CreateCustomerDelRq = doc.xml

    End Function



    ' Parse response XML from QuickBooks using MSXML parser
    Public Function ParseResponseXML(ByRef elementName As String) As Boolean

        On Error GoTo ErrHandler

        resListID = ""

        Dim retStatusCode As String
        Dim retStatusMessage As String
        Dim retStatusSeverity As String

        ' Create xmlDoc Obj

        ' DOM Document Object
        Dim xmlDoc As New MSXML2.DOMDocument60

        ' DOM Node list object for looping through
        Dim objNodeList As MSXML2.IXMLDOMNodeList

        ' Node objects
        Dim objChild As MSXML2.IXMLDOMNode
        Dim custChildNode As MSXML2.IXMLDOMNode
        Dim invoiceChildNode As MSXML2.IXMLDOMNode

        ' Attributes Name Mapping
        Dim attrNamedNodeMap As MSXML2.IXMLDOMNamedNodeMap

        Dim i As Short
        Dim ret As Boolean
        Dim errorMsg As String

        errorMsg = ""

        ' Load xml doc
        ret = xmlDoc.loadXML(responseXML)
        If Not ret Then
            errorMsg = "loadXML failed, reason: " & xmlDoc.parseError.reason
            GoTo ErrHandler
        End If

        ' Get nodes list
        objNodeList = xmlDoc.getElementsByTagName(elementName)

        ' Loop through each node
        ' Since we have only one request, we should only have one
        ' response.  The loop is actually unnecessary, but it
        ' is a good programming practice
        For i = 0 To (objNodeList.length - 1)

            ' Get the CustomerRetRs
            attrNamedNodeMap = objNodeList.item(i).attributes

            ' Get the status Code, info and Severity
            '
            retStatusCode = attrNamedNodeMap.getNamedItem("statusCode").nodeValue
            retStatusSeverity = attrNamedNodeMap.getNamedItem("statusSeverity").nodeValue
            retStatusMessage = attrNamedNodeMap.getNamedItem("statusMessage").nodeValue

            ' Check status code to see if there is error or warning
            If retStatusCode <> "0" Then
                ' Checking for Warning is a good practice, although unlikely to happen
                ' on an add request.
                If retStatusSeverity = "Warning" Then
                    ' Show the warning, then continue normal processing
                    MsgBox(retStatusMessage, MsgBoxStyle.Exclamation, "Warning from QuickBooks")
                ElseIf retStatusSeverity = "Error" Then
                    MsgBox(retStatusMessage, MsgBoxStyle.Exclamation, "Error from QuickBooks")
                    ' We only have one response thus we will exit.  If we have multiple
                    ' responses, then we may want to continue with the loop.
                    ParseResponseXML = False
                    Exit Function
                End If
            End If

            ' Look at the child nodes
            For Each objChild In objNodeList.item(i).childNodes

                ' Get the CustomerRet block if we were adding a customer
                If objChild.nodeName = "CustomerRet" Then

                    ' Get the elements in this block
                    For Each custChildNode In objChild.childNodes
                        If custChildNode.nodeName = "ListID" Then
                            resListID = custChildNode.text
                        ElseIf custChildNode.nodeName = "Name" Then
                            resCustName = custChildNode.text
                        ElseIf custChildNode.nodeName = "FullName" Then
                            resCustFullName = custChildNode.text
                        End If
                    Next custChildNode

                End If ' End of customerRet


                ' Get the "InvoiceQueryRet" if we were looking for
                ' an in-use customer by querying for invoices -- if
                ' we find one of these, we'll have all the information
                ' we need and can then break from the function.
                If objChild.nodeName = "InvoiceRet" Then

                    For Each invoiceChildNode In objChild.childNodes
                        If invoiceChildNode.nodeName = "CustomerRef" Then

                            ' Get the elements in this block
                            For Each custChildNode In invoiceChildNode.childNodes
                                If custChildNode.nodeName = "ListID" Then
                                    resListID = custChildNode.text
                                ElseIf custChildNode.nodeName = "Name" Then
                                    resCustName = custChildNode.text
                                ElseIf custChildNode.nodeName = "FullName" Then
                                    resCustFullName = custChildNode.text
                                End If
                            Next custChildNode

                            ' If we get here, we have all the information we need
                            GoTo BreakPoint

                        End If
                    Next invoiceChildNode

                End If ' end of InvoiceQueryRet


                ' Get the Customer Name if we were deleting a customer
                If elementName = "ListDelRs" And objChild.nodeName = "FullName" Then
                    custName = objChild.text
                End If ' end of ListDelType


            Next objChild
        Next

BreakPoint:
        ParseResponseXML = True
        Exit Function

ErrHandler:
        If errorMsg <> "" Then
            MsgBox(errorMsg, MsgBoxStyle.Exclamation, "Error")
        Else
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
        End If
        ParseResponseXML = False
        Exit Function

    End Function

    ' Form Data collection and verification
    '
    Private Function CollectFormData(ByRef operation As String) As Boolean

        On Error GoTo ErrHandler

        ' Get data from the form
        custName = Text_CustomerName.Text
        ListID = Text_ListID.Text

        ' Customer Name is required for an Add
        If (custName = "") And ("Add" = operation) Then
            MsgBox("Customer Name is empty", MsgBoxStyle.OkOnly, "Error")
            CollectFormData = False
            GoTo ExitProc
        End If

        If (ListID = "") And ("Delete" = operation) Then
            MsgBox("ListID is empty", MsgBoxStyle.OkOnly, "Error")
            CollectFormData = False
            GoTo ExitProc
        End If

        CollectFormData = True

ExitProc:
        Exit Function

ErrHandler:
        CollectFormData = False
        MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
        Exit Function

    End Function



    Private Sub Mainfrm_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        If blnIsOpenConnection Then
            'Call CloseConnection
            CloseConnection()
        End If

    End Sub

    Private Sub Mainfrm_Paint(sender As Object, e As PaintEventArgs) Handles MyBase.Paint
        Dim myGraphics As Graphics = Me.CreateGraphics
        Dim myPen As Pen
        myPen = New Pen(Brushes.Black, 1)
        e.Graphics.DrawLine(myPen, 40, 192, 464, 192)
        e.Graphics.DrawLine(myPen, 32, 336, 456, 336)
        e.Graphics.DrawLine(myPen, 44, 416, 456, 416)
        e.Graphics.DrawLine(myPen, 488, 64, 488, 440)
        e.Graphics.DrawLine(myPen, 56, 56, 640, 56)

    End Sub
End Class