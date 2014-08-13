
'----------------------------------------------------------
' Form: CustomerAddForm
'
' Description:  This module demonstrates the simple use of QBFC in .NET,
'               by adding a new customer to QuickBooks.
'
' Created On: 09/10/2002
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'----------------------------------------------------------
Imports Interop.QBFC13


Public Class CustomerAddForm
    Inherits System.Windows.Forms.Form


#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents label3 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents CustomerName As System.Windows.Forms.TextBox
    Friend WithEvents Phone As System.Windows.Forms.TextBox
    Friend WithEvents AddCustomer As System.Windows.Forms.Button
    Friend WithEvents ExitForm As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.CustomerName = New System.Windows.Forms.TextBox()
        Me.Phone = New System.Windows.Forms.TextBox()
        Me.AddCustomer = New System.Windows.Forms.Button()
        Me.ExitForm = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'label3
        '
        Me.label3.Location = New System.Drawing.Point(16, 88)
        Me.label3.Name = "label3"
        Me.label3.Size = New System.Drawing.Size(200, 32)
        Me.label3.TabIndex = 13
        Me.label3.Text = "Note: You need to have QuickBooks with a company file opened."
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(48, 24)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "Name"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(48, 16)
        Me.Label2.TabIndex = 15
        Me.Label2.Text = "Phone"
        '
        'CustomerName
        '
        Me.CustomerName.Location = New System.Drawing.Point(88, 16)
        Me.CustomerName.Name = "CustomerName"
        Me.CustomerName.Size = New System.Drawing.Size(304, 20)
        Me.CustomerName.TabIndex = 16
        Me.CustomerName.Text = ""
        '
        'Phone
        '
        Me.Phone.Location = New System.Drawing.Point(88, 40)
        Me.Phone.Name = "Phone"
        Me.Phone.Size = New System.Drawing.Size(128, 20)
        Me.Phone.TabIndex = 17
        Me.Phone.Text = ""
        '
        'AddCustomer
        '
        Me.AddCustomer.Location = New System.Drawing.Point(280, 64)
        Me.AddCustomer.Name = "AddCustomer"
        Me.AddCustomer.Size = New System.Drawing.Size(112, 24)
        Me.AddCustomer.TabIndex = 18
        Me.AddCustomer.Text = "AddCustomer"
        '
        'ExitForm
        '
        Me.ExitForm.Location = New System.Drawing.Point(280, 96)
        Me.ExitForm.Name = "ExitForm"
        Me.ExitForm.Size = New System.Drawing.Size(112, 24)
        Me.ExitForm.TabIndex = 19
        Me.ExitForm.Text = "Exit"
        '
        'CustomerAddForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(408, 133)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.ExitForm, Me.AddCustomer, Me.Phone, Me.CustomerName, Me.Label2, Me.Label1, Me.label3})
        Me.Name = "CustomerAddForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "CustomerAdd"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub ExitForm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitForm.Click
        Me.Close()
    End Sub

    Private Sub AddCustomer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddCustomer.Click

        'step1: verify that Name is not empty
        Dim NewName As String = CustomerName.Text.Trim()
        If (NewName.Length = 0) Then
            MessageBox.Show("Please enter a value for Name.", "Input Validation")
            Return
        End If

        'step2: create QBFC session manager and prepare the request
        Dim sessManager As QBSessionManager
        Dim msgSetRs As IMsgSetResponse

        Try
            sessManager = New QBSessionManagerClass()

            Dim msgSetRq As IMsgSetRequest = sessManager.CreateMsgSetRequest("US", 2, 0)
            msgSetRq.Attributes.OnError = ENRqOnError.roeContinue
            Dim custAdd As ICustomerAdd = msgSetRq.AppendCustomerAddRq
            custAdd.Name.SetValue(NewName)
            If Phone.Text.Length > 0 Then
                custAdd.Phone.SetValue(Phone.Text)
            End If

            'step3: begin QB session and send the request
            sessManager.OpenConnection("App", "IDN AddCustomer QBFC2 VB.NET sample")
            sessManager.BeginSession("", ENOpenMode.omDontCare)
            msgSetRs = sessManager.DoRequests(msgSetRq)

        Catch
            MessageBox.Show(Err.Description, "COM Error")
            Return
        Finally
            If Not sessManager Is Nothing Then
                sessManager.EndSession()
                sessManager.CloseConnection()
            End If
        End Try

        'step4: retrieve returned data and popup a message
        Dim popupMessage As New System.Text.StringBuilder()

        If (msgSetRs.ResponseList.Count = 1) Then
            'we have one response for oyur single add request
            Dim rs As IResponse = msgSetRs.ResponseList.GetAt(0)

            'get the status Code, info and Severity
            Dim retStatusCode As String = rs.StatusCode
            Dim retStatusSeverity As String = rs.StatusSeverity
            Dim retStatusMessage As String = rs.StatusMessage
            popupMessage.AppendFormat("statusCode = {0}, statusSeverity = {1}, statusMessage = {2}", _
                retStatusCode, retStatusSeverity, retStatusMessage)

            'retrieve some CustomerRet values
            'Add and Mod Rq return a single Ret object in rs.Detail
            'Query return a RetList object in rs.Detail 
            Dim customerRet As ICustomerRet = rs.Detail
            If (Not customerRet Is Nothing) Then
                Dim listID As String = customerRet.ListID.GetValue()
                Dim fullName As String = customerRet.FullName.GetValue()

                popupMessage.AppendFormat( _
                    ControlChars.CrLf & "Customer ListID = {0} " & ControlChars.CrLf & "Customer FullName={1}", _
                    listID, fullName)

                'retrieve a field that may NOT be part of the QuickBook response    
                If Not (customerRet.Phone Is Nothing) Then
                    popupMessage.AppendFormat(ControlChars.CrLf & "Customer Phone = {0}", _
                      customerRet.Phone.GetValue)
                End If
            End If
            MessageBox.Show(popupMessage.ToString(), "QuickBooks response")
        End If
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
