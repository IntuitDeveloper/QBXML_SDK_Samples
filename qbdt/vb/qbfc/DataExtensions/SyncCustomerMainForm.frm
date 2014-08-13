VERSION 5.00
Begin VB.Form SyncCustomerListForm 
   Caption         =   "Synchronize Customer List"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5850
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   5850
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox SyncResultsTextBox 
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton CloseWindowButton 
      Cancel          =   -1  'True
      Caption         =   "Close Window"
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   4440
      Width           =   2535
   End
   Begin VB.CommandButton SyncCustomerListButton 
      Caption         =   "Synchronize Customer List"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   4440
      Width           =   2415
   End
   Begin VB.TextBox SyncTimeTextBox 
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   2520
      Width           =   2655
   End
   Begin VB.ListBox CustomerListBox 
      Height          =   1815
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   5295
   End
   Begin VB.Label Label3 
      Caption         =   $"SyncCustomerMainForm.frx":0000
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   5535
   End
   Begin VB.Label Label2 
      Caption         =   "Time of Last Customer Synchronization:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Customer Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "SyncCustomerListForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CloseWindowButton_Click()

    ' EndSession and CloseConnection with QB
    EndSessionCloseConnection

    End

End Sub

Private Sub Form_Load()
        ' OpenConnection and BeginSession to QB
        OpenConnectionBeginSession

        Dim timeStamp As Date
        timeStamp = Now

        Dim bError As Boolean
        bError = False

        ' fill the list box with Customer FullNames
        GetCustomers CustomerListBox, bError

        ' only reset the sync time if we were really successful in resyncing the Customer list
        If (Not bError) Then
            ' remember the last time we sync'd the Customer list
            SetLastTimeSync (timeStamp)

            Dim timeString As String
            timeString = Format(timeStamp, "General Date")
            SyncTimeTextBox.Text = timeString
        End If

        ' EndSession and CloseConnection with QB
        EndSessionCloseConnection

End Sub

Private Sub SyncCustomerListButton_Click()
        Dim timeStamp As Date
        timeStamp = Now

        Dim bError As Boolean
        bError = False

        ' OpenConnection and BeginSession to QB
        OpenConnectionBeginSession

        ' Synchronize the Customer List of information
        SyncCustomerList CustomerListBox, SyncResultsTextBox, bError

        ' EndSession and CloseConnection with QB
        EndSessionCloseConnection

        ' only reset the sync time if we were really successful in resyncing the Customer list
        If (bError) Then
            MsgBox ("Error synchronizing the Customer List")
        Else
            ' remember the last time we sync'd the Customer list
            SetLastTimeSync (timeStamp)

            Dim timeString As String
            timeString = Format(timeStamp, "General Date")
            SyncTimeTextBox.Text = timeString

            MsgBox ("Customer List successfully synchronized")
        End If
End Sub
