Public Class SyncCustomerListForm
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
    Friend WithEvents CloseWindowButton As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents SyncCustomerListButton As System.Windows.Forms.Button
    Friend WithEvents CustomerListBox As System.Windows.Forms.ListBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents SyncTimeTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.CloseWindowButton = New System.Windows.Forms.Button()
        Me.SyncCustomerListButton = New System.Windows.Forms.Button()
        Me.CustomerListBox = New System.Windows.Forms.ListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.SyncTimeTextBox = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'CloseWindowButton
        '
        Me.CloseWindowButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.CloseWindowButton.Location = New System.Drawing.Point(258, 272)
        Me.CloseWindowButton.Name = "CloseWindowButton"
        Me.CloseWindowButton.Size = New System.Drawing.Size(148, 23)
        Me.CloseWindowButton.TabIndex = 0
        Me.CloseWindowButton.Text = "Close Window"
        '
        'SyncCustomerListButton
        '
        Me.SyncCustomerListButton.Location = New System.Drawing.Point(58, 272)
        Me.SyncCustomerListButton.Name = "SyncCustomerListButton"
        Me.SyncCustomerListButton.Size = New System.Drawing.Size(148, 23)
        Me.SyncCustomerListButton.TabIndex = 1
        Me.SyncCustomerListButton.Text = "Synchronize Customer List"
        '
        'CustomerListBox
        '
        Me.CustomerListBox.Location = New System.Drawing.Point(24, 40)
        Me.CustomerListBox.Name = "CustomerListBox"
        Me.CustomerListBox.SelectionMode = System.Windows.Forms.SelectionMode.None
        Me.CustomerListBox.Size = New System.Drawing.Size(424, 95)
        Me.CustomerListBox.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(112, 16)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Customer Names:"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(24, 152)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(208, 16)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = " Time of Last Customer Synchronization:"
        '
        'SyncTimeTextBox
        '
        Me.SyncTimeTextBox.Location = New System.Drawing.Point(240, 152)
        Me.SyncTimeTextBox.Name = "SyncTimeTextBox"
        Me.SyncTimeTextBox.Size = New System.Drawing.Size(208, 20)
        Me.SyncTimeTextBox.TabIndex = 6
        Me.SyncTimeTextBox.Text = ""
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(24, 184)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(424, 80)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "We have read the information for Customers from QuickBooks.  Now make changes to " & _
        "Customers using the QuickBooks UI, like restoring an old copy of the company fil" & _
        "e, add and modify some customer information, and delete some customers.  Then pr" & _
        "ess the ""Synchronize Customer List"" button.  We will issue a CompanyActivityQuer" & _
        "y, CustomerQuery by modified date, and ListDeletedQuery for these changes and up" & _
        "date the Customer List."
        '
        'SyncCustomerListForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.CloseWindowButton
        Me.ClientSize = New System.Drawing.Size(464, 317)
        Me.ControlBox = False
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label3, Me.SyncTimeTextBox, Me.Label2, Me.Label1, Me.CustomerListBox, Me.SyncCustomerListButton, Me.CloseWindowButton})
        Me.Name = "SyncCustomerListForm"
        Me.Text = "Synchronize Customer List"
        Me.ResumeLayout(False)

    End Sub

#End Region

    '----------------------------------------------------------
    ' Form: SyncCustomerListForm
    '
    ' Description: This form displays the names of the customers
    '           in the open company file.  Click on the 
    '           Synchronize Customer List button to issue the 
    '           CustomerActivityQuery, CustomerQuery by modified date
    '           and ListDeletedQuery.
    '
    ' Copyright © 2002-2013 Intuit Inc. All rights reserved.
    ' Use is subject to the terms specified at:
    '      http://developer.intuit.com/legal/devsite_tos.html
    '
    '----------------------------------------------------------
    ' Note:  We are doing connect and disconnect to QB around each set of requests
    ' so we can illustrate synching when the backup/restore changes to the company file.

    Private Sub SyncCustomerListForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        ' OpenConnection and BeginSession to QB
        OpenConnectionBeginSession()

        Dim timeStamp As Date
        timeStamp = New Date().Now()

        Dim bError As Boolean
        bError = False

        ' fill the list box with Customer FullNames
        GetCustomers(CustomerListBox, bError)

        ' only reset the sync time if we were really successful in resyncing the Customer list
        If (Not bError) Then
            ' remember the last time we sync'd the Customer list
            SetLastTimeSync(timeStamp)

            Dim timeString As String()
            timeString = timeStamp.GetDateTimeFormats("G"c)
            SyncTimeTextBox.Text = timeString(0)
        End If

        ' EndSession and CloseConnection with QB
        EndSessionCloseConnection()

    End Sub

    Private Sub CloseWindowButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CloseWindowButton.Click

        ' EndSession and CloseConnection with QB
        EndSessionCloseConnection()

        End
    End Sub

    Private Sub SyncCustomerListButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SyncCustomerListButton.Click

        Dim timeStamp As Date
        timeStamp = New Date().Now()

        Dim bError As Boolean
        bError = False

        ' OpenConnection and BeginSession to QB
        OpenConnectionBeginSession()

        ' Synchronize the Customer List of information
        SyncCustomerList(CustomerListBox, bError)

        ' EndSession and CloseConnection with QB
        EndSessionCloseConnection()

        ' only reset the sync time if we were really successful in resyncing the Customer list
        If (bError) Then
            MsgBox("Error synchronizing the Customer List")
        Else
            ' remember the last time we sync'd the Customer list
            SetLastTimeSync(timeStamp)

            Dim timeString As String()
            timeString = timeStamp.GetDateTimeFormats("G"c)
            SyncTimeTextBox.Text = timeString(0)

            MsgBox("Customer List successfully synchronized")
        End If
    End Sub
End Class
