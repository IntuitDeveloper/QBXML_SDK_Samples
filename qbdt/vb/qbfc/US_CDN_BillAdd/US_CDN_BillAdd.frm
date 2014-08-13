VERSION 5.00
Begin VB.Form BillAdd 
   Caption         =   "BillAdd Sample"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   7185
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton EnterBill 
      Caption         =   "Enter Bill!"
      Height          =   495
      Left            =   1560
      TabIndex        =   8
      Top             =   4440
      Width           =   2775
   End
   Begin VB.TextBox BillAmount 
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Text            =   "0.00"
      Top             =   3480
      Width           =   3615
   End
   Begin VB.ComboBox AccountList 
      Height          =   315
      Left            =   3240
      TabIndex        =   4
      Top             =   3000
      Width           =   3615
   End
   Begin VB.ComboBox VendorList 
      Height          =   315
      Left            =   3240
      TabIndex        =   2
      Top             =   2520
      Width           =   3615
   End
   Begin VB.CommandButton CommPrefs 
      Caption         =   "Change Communication Preferences"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   1440
      Width           =   5655
   End
   Begin VB.Label Label4 
      Caption         =   "Specify the amount of the bill:"
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Select an item for the bill:"
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Select a vendor for the bill:"
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   7080
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   6960
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   6960
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7080
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label1 
      Caption         =   $"US_CDN_BillAdd.frx":0000
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "BillAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' This is the primary form for this application, it communicates with QuickBooks to
' populate list boxes with vendor and account lists and allows the user to select a
' vendor, an accounts payable account, and an amount for the bill and adds it to
' QuickBooks running either on the desktop or the online edition based on communication
' preferences set by the user (see the CommPrefsDlg).
'
' Updated to QBFC 12.0 and fixed setting max QBXML version: 09/2012

'
' We are both a desktop and an Online edition application, so we have some constants
' here that define us in either environment.
'
'Used for the registry and to Identify us to QuickBooks Desktop
'
Const appName = "IDN US-Canada Bill Add Sample"
'
' Used to identify us to QuickBooks Online Edition
'
Const appLogin = "billadd.developer.intuit.com"
Const appID = "60177302"
Const appVersion = "1.0"

'
' Private variables to hold our connection to QuickBooks
'
Private SessionManager As QBSessionManager

' User clicked the "Communication Preferences" button
'
Private Sub CommPrefs_Click()
    DoCommPrefs
    If (SessionManager Is Nothing) Then
        FillInForm
    End If
End Sub

Private Sub DoCommPrefs()
    '
    ' Show the Communication preferences dialog
    '
    CommPrefsDlg.Show vbModal
    
    '
    ' If the user clicked OK and we already had a session manger, get rid of
    ' the existing session manager and then fill the form back in with the
    ' data from QuickBooks
    '
    If (CommPrefsDlg.isOK And Not SessionManager Is Nothing) Then
        CommPrefsDlg.DisconnectFromQuickBooks SessionManager
        Set SessionManager = Nothing
    End If

End Sub

'
' This is where most of the work happens, we build the BillAddRq and send the bill to
' QuickBooks (note that we don't have to do anything special for Online or Desktop, though
' there are parts of the BillAddRq that are not supported by QuickBooks Online Edition (ItemLineAdd
' and ItemGroupLineAdd) we aren't using them here so we don't have to check if we are Online or
' on the desktop.
'
Private Sub EnterBill_Click()
    Dim vendor As String
    Dim account As String
    Dim amount As String
    '
    ' Do the validation of the bill information
    '
    vendor = VendorList.Text
    If ("" = vendor) Then
        MsgBox "Must specify a vendor for the bill!", vbExclamation, "Missing parameter"
        Exit Sub
    End If
    account = AccountList.Text
    If ("" = account) Then
        MsgBox "Must specify an account for this bill!", vbExclamation, "Missing Parameter"
    End If
    amount = BillAmount.Text
    If ("" = amount) Then
        MsgBox "Must specify an amount for the bill!", vbExclamation, "Missing parameter"
        Exit Sub
    End If
    '
    ' Make sure we have a session manager
    '
    If (SessionManager Is Nothing) Then
        Set SessionManager = CommPrefsDlg.ConnectToQuickBooks(appName)
    End If
    
    '
    ' Now build the BillAddRq
    '
    If (Not SessionManager Is Nothing) Then
        Dim BillAddSet As IMsgSetRequest
        Set BillAddSet = GetLatestMsgSetRequest(SessionManager)
        Dim BillAddRq As IBillAdd
        Set BillAddRq = BillAddSet.AppendBillAddRq
        
        BillAddRq.TxnDate.setValue Now
        BillAddRq.VendorRef.FullName.setValue vendor
        
        Dim Expense As IExpenseLineAdd
        Set Expense = BillAddRq.ExpenseLineAddList.Append
        Expense.amount.setValue amount
        Expense.AccountRef.FullName.setValue account
        Expense.Memo.setValue "Sample expense line"
        
        ' Perform the request and obtain a response from QuickBooks
        Dim ResponseSet As IMsgSetResponse
        Set ResponseSet = SessionManager.DoRequests(BillAddSet)
    
        ' Uncomment the following to see the request and response XML for debugging
        ' MsgBox BillAddSet.ToXMLString, vbOKOnly, "RequestXML"
        ' MsgBox ResponseSet.ToXMLString, vbOKOnly, "ResponseXML"
    
        ' Interpret the response
        Dim response As IResponse
    
        ' The response list contains only one response,
        ' which corresponds to our single BillAdd request
        Set response = ResponseSet.ResponseList.GetAt(0)
     
        msg = "Status: Code = " & CStr(response.StatusCode) & _
          ", Message = " & response.StatusMessage & _
          ", Severity = " & response.StatusSeverity & vbCrLf
        
        ' The Detail property of the IResponse object
        ' returns a Ret object for Add and Mod requests.
        ' In this case, the Ret object is IBillRet.
    
        ' For help finding out the Detail's type, uncomment the following line:
        ' MsgBox response.Detail.Type.GetAsString
    
        Dim billRet As IBillRet
        Set billRet = response.Detail
        '
        ' We have to make sure we really have a BillRet, since if something went wrong we
        ' won't.
        '
        If (Not (billRet Is Nothing)) Then
            '
            ' Retrieve fields that are guaranteed to be returned in the response
            ' We don't have to check that these fields exist in the response
            '
            msg = msg & vbCr & "BillRet: EditSequence = " & billRet.EditSequence.getValue
            msg = msg & ", Amount Due = " & billRet.AmountDue.GetAsString
        
            '
            ' Retrieve a field that is not guaranteed to be returned from QuickBooks
            ' Note that in this case we must first check whether the field exists in the
            ' response before trying to use it.
            '
            If (Not billRet.DueDate Is Nothing) Then
                msg = msg & ", Due Date = " & billRet.DueDate.getValue
            End If
        End If
        '
        ' Display a message showing some of the interesting fields of the response
        '
        MsgBox msg
    End If
End Sub

'
' Load the form, this is where we set up the information needed by the CommPrefsDlg
'
Private Sub Form_Load()
    CommPrefsDlg.SetAppLogin appName, appLogin
    CommPrefsDlg.SetAppID appName, appID
    CommPrefsDlg.SetAppVersion appName, appVersion
    DoCommPrefs
    FillInForm
End Sub

'
' Fill in the list boxes in the form with the vendors and accounts payable accounts
'
Private Sub FillInForm()
    '
    ' Make sure we are connected to QuickBooks
    '
    If (SessionManager Is Nothing) Then
        Set SessionManager = CommPrefsDlg.ConnectToQuickBooks(appName)
    End If
    '
    ' Now build the VendorQuery and AccountQuery requests to fill in our list boxes
    '
    If (Not SessionManager Is Nothing) Then
        '
        ' Set up the MsgSet to query the lists
        '
        Dim ListQuerySet As IMsgSetRequest
        Set ListQuerySet = GetLatestMsgSetRequest(SessionManager)
        ListQuerySet.Attributes.OnError = roeContinue
        
        '
        ' Add the VendorQuery
        '
        Dim VendorQuery As IVendorQuery
        Set VendorQuery = ListQuerySet.AppendVendorQueryRq
        
        '
        ' Add the AccountQuery to get CostOfGoodsSold and Expense accounts
        '
        Dim AccountQuery As IAccountQuery
        Set AccountQuery = ListQuerySet.AppendAccountQueryRq
        AccountQuery.ORAccountListQuery.AccountListFilter.AccountTypeList.Add atCostOfGoodsSold
        AccountQuery.ORAccountListQuery.AccountListFilter.AccountTypeList.Add atExpense
        Set AccountQuery = ListQuerySet.AppendAccountQueryRq
        
        '
        ' Now send the query and get QBFC to parse the response
        '
        Dim ListResponse As IMsgSetResponse
        Dim ListQueryXML As String
        ListQueryXML = ListQuerySet.ToXMLString

        Set ListResponse = SessionManager.DoRequests(ListQuerySet)
        Dim response As IResponse
        If (ListResponse.ResponseList Is Nothing) Then
            ' This could happen if there is no session ticket available to QBOE
            ' For example the session ticket we had expired
            Exit Sub
        End If

        '
        ' The response list contains only two responses,
        ' the first response corresponds to our vendor query, populate
        ' the VendorList listbox with the results of that query.
        '
        Set response = ListResponse.ResponseList.GetAt(0)
        Dim QBVendorList As IVendorRetList
        Set QBVendorList = response.Detail
        VendorList.Clear
        If (Not QBVendorList Is Nothing) Then
            For i = 0 To (QBVendorList.Count - 1)
                Dim vendor As IVendorRet
                Set vendor = QBVendorList.GetAt(i)
                VendorList.AddItem vendor.Name.getValue
            Next i
        End If
        
        '
        ' the second response corresponds to our account query, populate
        ' the AccountList listbox with the results of that query
        '
        Set response = ListResponse.ResponseList.GetAt(1)
        Dim QBAccountList As IAccountRetList
        Set QBAccountList = response.Detail
        AccountList.Clear
        If (Not QBAccountList Is Nothing) Then
            For i = 0 To (QBAccountList.Count - 1)
                Dim account As IAccountRet
                Set account = QBAccountList.GetAt(i)
                AccountList.AddItem account.FullName.getValue
            Next i
        End If
    End If
End Sub

Function QBFCLatestVersion(SessionManager As QBSessionManager) As String
    Dim strXMLVersions() As String
    'Should be able to use this, but there appears to be a bug that may cause 2.0 to be returned
    'when it should not.
    'strXMLVersions = SessionManager.QBXMLVersionsForSession
    
    Dim strVersions() As String
    
    Dim msgset As IMsgSetRequest
    On Error GoTo use2_0msgSetforOE
    'Use oldest version to ensure that we work with any QuickBooks (US)
    Set msgset = SessionManager.CreateMsgSetRequest("US", 1, 0)
    GoTo gotmsgset
use2_0msgSetforOE:
    'remove error handler to avoid a possible infinite loop
    On Error GoTo 0
    Set msgset = SessionManager.CreateMsgSetRequest("US", 2, 0)
    QBFCLatestVersion = "2.1"
    Exit Function
gotmsgset:
    msgset.AppendHostQueryRq
    Dim QueryResponse As IMsgSetResponse
    Set QueryResponse = SessionManager.DoRequests(msgset)
    Dim response As IResponse
    
    ' The response list contains only one response,
    ' which corresponds to our single HostQuery request
    Set response = QueryResponse.ResponseList.GetAt(0)
    Dim HostResponse As IHostRet
    Set HostResponse = response.Detail
    Dim supportedVersions As IBSTRList
    Set supportedVersions = HostResponse.SupportedQBXMLVersionList
    
    Dim i As Long
    Dim vers As Double
    Dim LastVers As Double
    LastVers = 0
    For i = 0 To supportedVersions.Count - 1
        vers = Val(supportedVersions.GetAt(i))
        If (vers > LastVers) Then
            LastVers = vers
            QBFCLatestVersion = supportedVersions.GetAt(i)
        End If
    Next i
End Function

Public Function GetLatestMsgSetRequest(SessionManager As QBSessionManager) As IMsgSetRequest
    Dim supportedVersion As String
    Dim strVersions() As String
       
  supportedVersion = Val(QBFCLatestVersion(SessionManager))
  ' If this is QB Canada, just use the 3.0 spec
  ' Ideally we should repeat the version logic just like for the US version
  strVersions = SessionManager.QBXMLVersionsForSession
    If InStr(1, strVersions(UBound(strVersions)), "CA") Then
     Set GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("CA", 2, 0)
    
    ElseIf (supportedVersion >= 5#) Then
        Set GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 5, 0)
    ElseIf (supportedVersion >= 4#) Then
        Set GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 4, 0)
    ElseIf (supportedVersion >= 3#) Then
        Set GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 3, 0)
    ElseIf (supportedVersion = 2#) Then
        Set GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 2, 0)
    ElseIf (supportedVersion = 1.1) Then
        Set GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 1, 1)
    Else
        MsgBox "You are apparently running QuickBooks 2002 Release 1, we strongly recommend that you use QuickBooks' online update feature to obtain the latest fixes and enhancements", vbExclamation
        Set GetLatestMsgSetRequest = SessionManager.CreateMsgSetRequest("US", 1, 0)
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    If Not SessionManager Is Nothing Then
        SessionManager.EndSession
        SessionManager.CloseConnection
    End If
    Unload CommPrefsDlg
    
End Sub
