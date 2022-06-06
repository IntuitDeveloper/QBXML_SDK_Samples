Option Strict Off
Option Explicit On
Friend Class Main_Renamed
    Inherits System.Windows.Forms.Form
    '-----------------------------------------------------------
    ' Form Module: Main
    '
    ' Description:  This sample program demonstrates the following:
    '               -   Use of DOMXMLBuilder to build a request
    '               -   Parsing the response
    '               -   Creation of qbXML query and modify requests
    '               -   Sending a request to QuickBooks
    '
    ' Updated to SDK 2.0: 08/05/2002
    '
    ' Copyright © 2021-2022 Intuit Inc. All rights reserved.
    ' Use is subject to the terms specified at:
    '      http://developer.intuit.com/legal/devsite_tos.html
    '
    '----------------------------------------------------------

    '----------------------------------------------------------
    ' Data Declaration
    '
    '----------------------------------------------------------


    Const cAppID As String = "1234"
    Const cAppName As String = "VB Modify Customer Sample App"

    Dim requestXML As String
    Dim responseXML As String

    ' Collection object, to store the list of customers
    Public customerCollection As Collection
    Public companyFile As String
    ' Make rpWrapper public so it can be used by CustomerMod
    ' We want to have a single session to QuickBooks
    Public rpWrapper As qbXMLRPWrapper



    Private Sub Comm_Browse_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Comm_Browse.Click

        On Error GoTo ErrHandler

        ' Set CancelError is True
        ' CommDlg_Browse.CancelError = True

        ' Set flags
        CommDlg_BrowseOpen.ShowReadOnly = False
        ' Set filters
        CommDlg_BrowseOpen.Filter = "All Files (*.*)|*.*|QuickBooks data file" & "(*.qbw)|*.qbw"
        ' Specify default filter
        CommDlg_BrowseOpen.FilterIndex = 2
        ' Display the Open dialog box
        CommDlg_BrowseOpen.ShowDialog()
        ' Display name of selected file
        CompanyFileNameInput.Text = CommDlg_BrowseOpen.FileName

        Exit Sub

ErrHandler:
        'User pressed the Cancel button
        Exit Sub

    End Sub

    ' Submit buttom
    '
    Private Sub Comm_Submit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Comm_Submit.Click


        On Error GoTo ErrorHandler

        ' Get QB data file name
        If Not CollectFormData() Then
            Exit Sub
        End If

        ' Build the request XML
        BuildXML()

        ' Send request to QuickBooks
        Dim success As Boolean
        Dim errNumber As Integer
        Dim errMsg As String

        rpWrapper = New qbXMLRPWrapper

        ' Start session with QuickBooks
        success = rpWrapper.Start(cAppID, cAppName, companyFile)
        If Not success Then
            rpWrapper.GetErrorInfo(errNumber, errMsg)
            Err.Raise(Number:=errNumber, Description:=errMsg)
        End If


        'Send request
        success = rpWrapper.DoRequest(requestXML, responseXML)
        If Not success Then
            rpWrapper.GetErrorInfo(errNumber, errMsg)
            Err.Raise(Number:=errNumber, Description:=errMsg)
        End If
        customerCollection = Nothing ' Reset previous collection
        customerCollection = New Collection

        If Not ParseResponseXML() Then
            Exit Sub
        End If

        ' Load CustomerList Form
        Dim CustListForm As New CustomerList
        VB6.ShowForm(CustListForm, VB6.FormShowConstants.Modal, Me)

        ' Close session with QuickBooks
        success = rpWrapper.Finish()
        If Not success Then
            rpWrapper.GetErrorInfo(errNumber, errMsg)
            Err.Raise(Number:=errNumber, Description:=errMsg)
        End If
        rpWrapper = Nothing

        Exit Sub

ErrorHandler:
        MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
        Exit Sub

    End Sub

    ' exit button
    '
    Private Sub Comm_Exit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Comm_Exit.Click
        ' close windows
        Me.Close()
    End Sub


    ' Form Data Collcetion
    '
    Private Function CollectFormData() As Boolean

        companyFile = CompanyFileNameInput.Text

        ' Verify file existence
        If Dir(companyFile) = "" Then
            MsgBox("The company file doesn't exist!" & vbCr & companyFile, MsgBoxStyle.Exclamation, "Error")
            CollectFormData = False
            Exit Function
        End If

        CollectFormData = True
        Exit Function

    End Function



    Private Sub Main_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        On Error GoTo ErrHandler

        Dim appPath As Object
        appPath = My.Application.Info.DirectoryPath
        Image_QBBANNER.Image = System.Drawing.Image.FromFile(appPath & "/qbbanner.bmp")

        Exit Sub

ErrHandler:
        MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error")
        Exit Sub

    End Sub

    '----------------------------------------------------------
    'Build XML request
    '----------------------------------------------------------
    Private Sub BuildXML()

        Dim doc As New MSXML2.DOMDocument60
        Dim QBXML As MSXML2.IXMLDOMNode
        Dim MsgsRq As MSXML2.IXMLDOMElement
        Dim CustomerQuery As MSXML2.IXMLDOMElement

        QBXML = doc.appendChild(doc.createElement("QBXML"))
        MsgsRq = QBXML.appendChild(doc.createElement("QBXMLMsgsRq"))
        MsgsRq.setAttribute("onError", "continueOnError")
        CustomerQuery = MsgsRq.appendChild(doc.createElement("CustomerQueryRq"))

        requestXML = doc.xml

    End Sub
    '
    ' ParseResponseXML: To Use MSXML to parse response XML
    '
    Private Function ParseResponseXML() As Boolean

        On Error GoTo ErrHandler

        ' DOM Document object
        Dim xmlDoc As New MSXML2.DOMDocument60

        ' Node List
        Dim objNodeList As MSXML2.IXMLDOMNodeList
        Dim objChild As MSXML2.IXMLDOMNode
        Dim custChildNode As MSXML2.IXMLDOMNode

        xmlDoc.async = False

        ' Attributes Name Mapping
        Dim attrNamedNodeMap As MSXML2.IXMLDOMNamedNodeMap

        Dim i As Short

        ' QBXML status values

        Dim retStatusCode As String
        Dim retStatusMessage As String
        Dim retStatusSeverity As String

        Dim ret As Boolean
        Dim errorMsg As String

        On Error GoTo ErrHandler

        ' Load xml document
        ret = xmlDoc.loadXML(responseXML)

        If Not ret Then
            errorMsg = "loadXML failed, reason: " & xmlDoc.parseError.reason
            GoTo ErrHandler
        End If

        ' Get CustomerQueryRs Node list
        objNodeList = xmlDoc.getElementsByTagName("CustomerQueryRs")

        ' The following loop is actually unnecessary for this case.
        ' We have only one request, so we should only have one response.
        Dim cust As New Customer
        For i = 0 To (objNodeList.length - 1)

            ' Get the CustomerRetRs
            attrNamedNodeMap = objNodeList.item(i).attributes

            retStatusCode = attrNamedNodeMap.getNamedItem("statusCode").nodeValue
            retStatusSeverity = attrNamedNodeMap.getNamedItem("statusSeverity").nodeValue
            retStatusMessage = attrNamedNodeMap.getNamedItem("statusMessage").nodeValue

            ' Check status code to see if there is an error
            If retStatusCode <> "0" Then
                ' For a query response, we want to have special check for status code = 1
                If retStatusCode = "1" Then ' Status code for no record found
                    MsgBox("No customer is found", MsgBoxStyle.Information, "Message from QuickBooks")
                    ParseResponseXML = False
                    Exit Function
                Else
                    ' Error or warning
                    ' retStatusSeverity can be used to distinguish error from warning
                    MsgBox(retStatusMessage, MsgBoxStyle.Exclamation, "Error from QuickBooks")
                    ' We have only one response thus we will exit.  If we have multiple
                    ' responses, then we need to continue looping.
                    ParseResponseXML = False
                    Exit Function
                End If
            End If

            ' Walk through the child nodes to get the detail Customer info
            For Each objChild In objNodeList.item(i).childNodes

                ' Get the CustomerRet block
                If objChild.nodeName = "CustomerRet" Then


                    ' Get the elements in this block
                    With cust
                        For Each custChildNode In objChild.childNodes
                            If custChildNode.nodeName = "ListID" Then
                                .ListID = custChildNode.text
                            ElseIf custChildNode.nodeName = "FullName" Then
                                .FullName = custChildNode.text
                            ElseIf custChildNode.nodeName = "TimeCreated" Then
                                .TimeCreated = custChildNode.text
                            ElseIf custChildNode.nodeName = "TimeModified" Then
                                .TimeModified = custChildNode.text
                            ElseIf custChildNode.nodeName = "Balance" Then
                                .Balance = custChildNode.text
                            ElseIf custChildNode.nodeName = "Phone" Then
                                .Phone = custChildNode.text
                            ElseIf custChildNode.nodeName = "FirstName" Then
                                .FirstName = custChildNode.text
                            ElseIf custChildNode.nodeName = "LastName" Then
                                .LastName = custChildNode.text
                            ElseIf custChildNode.nodeName = "Name" Then
                                .Name = custChildNode.text
                            ElseIf custChildNode.nodeName = "CompanyName" Then
                                .CompanyName = custChildNode.text
                            ElseIf custChildNode.nodeName = "EditSequence" Then
                                .EditSequence = custChildNode.text
                            ElseIf custChildNode.nodeName = "Email" Then
                                .Email = custChildNode.text
                            ElseIf custChildNode.nodeName = "TermsRef_ListID" Then
                                .TermsRef_ListID = custChildNode.text
                            ElseIf custChildNode.nodeName = "TermsRef_FullName" Then
                                .TermsRef_FullName = custChildNode.text
                            ElseIf custChildNode.nodeName = "Sublevel" Then
                                .Sublevel = custChildNode.text
                            End If

                        Next custChildNode ' end of one CustomerRet
                    End With

                    ' Add customer object to the collection with FullName as the key
                    customerCollection.Add(cust, cust.FullName)

                    ' Reset the customer object
                    cust = Nothing
                    cust = New Customer

                End If

            Next objChild
        Next

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


    Public Function GetRequestXML() As String

        GetRequestXML = requestXML

    End Function
    Public Function GetResponseXML() As String

        GetResponseXML = responseXML

    End Function
End Class