Imports Interop.QBFC13
Imports System
Imports System.IO

Public Class ReceiveItems
    Inherits System.Windows.Forms.Form

    Private sessMgr As QBSessionManager

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
    Friend WithEvents VendorList As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents POList As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents LineItems As System.Windows.Forms.ListView
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents StoreIR As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.VendorList = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.POList = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.LineItems = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.StoreIR = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'VendorList
        '
        Me.VendorList.Location = New System.Drawing.Point(8, 24)
        Me.VendorList.Name = "VendorList"
        Me.VendorList.Size = New System.Drawing.Size(200, 21)
        Me.VendorList.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(120, 16)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "1. Select Vendor:"
        '
        'POList
        '
        Me.POList.Location = New System.Drawing.Point(8, 72)
        Me.POList.Name = "POList"
        Me.POList.Size = New System.Drawing.Size(568, 21)
        Me.POList.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 56)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 16)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "2. Select PO:"
        '
        'LineItems
        '
        Me.LineItems.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader4, Me.ColumnHeader2, Me.ColumnHeader3})
        Me.LineItems.LabelEdit = True
        Me.LineItems.Location = New System.Drawing.Point(8, 120)
        Me.LineItems.Name = "LineItems"
        Me.LineItems.Size = New System.Drawing.Size(568, 312)
        Me.LineItems.TabIndex = 4
        Me.LineItems.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "# Received"
        Me.ColumnHeader1.Width = 68
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "# Already Received"
        Me.ColumnHeader4.Width = 106
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "# Ordered"
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Item "
        Me.ColumnHeader3.Width = 436
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(8, 104)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(224, 16)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "3. Enter Number Received for Each Item:"
        '
        'StoreIR
        '
        Me.StoreIR.Location = New System.Drawing.Point(16, 456)
        Me.StoreIR.Name = "StoreIR"
        Me.StoreIR.Size = New System.Drawing.Size(192, 32)
        Me.StoreIR.TabIndex = 6
        Me.StoreIR.Text = "4. Store Item Receipt"
        '
        'ReceiveItems
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(592, 494)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.StoreIR, Me.Label3, Me.LineItems, Me.Label2, Me.POList, Me.Label1, Me.VendorList})
        Me.Name = "ReceiveItems"
        Me.Text = "Receive Items"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub VendorList_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VendorList.SelectedIndexChanged
        If (Not VendorList.SelectedItem Is Nothing) Then
            POList.Items.Clear()
            Dim selectedVendor As VendorItem
            selectedVendor = VendorList.SelectedItem
            Dim msgRq As IMsgSetRequest
            msgRq = sessMgr.CreateMsgSetRequest("US", 4, 0)
            Dim POquery As IPurchaseOrderQuery
            POquery = msgRq.AppendPurchaseOrderQueryRq
            POquery.ORTxnQuery.TxnFilter.EntityFilter.OREntityFilter.ListIDList.Add(selectedVendor.ListID)
            POquery.IncludeRetElementList.Add("TxnID")
            POquery.IncludeRetElementList.Add("RefNumber")
            POquery.IncludeRetElementList.Add("TxnNumber")
            POquery.IncludeRetElementList.Add("TxnDate")
            POquery.IncludeRetElementList.Add("IsManuallyClosed")
            POquery.IncludeRetElementList.Add("IsFullyReceived")
            Dim msgRs As IMsgSetResponse
            msgRs = sessMgr.DoRequests(msgRq)
            Dim resp As IResponse
            resp = msgRs.ResponseList.GetAt(0) ' Only one request, only one response

            '
            ' Check for errors
            '
            If (resp.StatusSeverity = "Error") Then
                MsgBox("Problem getting POs for vendor (" & resp.StatusCode & "): " & resp.StatusMessage)
                Exit Sub
            End If

            If (resp.Detail Is Nothing) Then
                Exit Sub
            End If

            Dim poRetList As IPurchaseOrderRetList
            poRetList = resp.Detail 'upcast to proper type
            Dim i As Integer
            For i = 0 To poRetList.Count - 1
                Dim poRet As IPurchaseOrderRet
                poRet = poRetList.GetAt(i)
                Dim ShowPO As Boolean
                ShowPO = True
                If (Not poRet.IsFullyReceived Is Nothing) Then
                    If (poRet.IsFullyReceived.GetValue()) Then
                        ShowPO = False
                    End If
                End If
                If (ShowPO And Not poRet.IsManuallyClosed Is Nothing) Then
                    If (poRet.IsManuallyClosed.GetValue()) Then
                        ShowPO = False
                    End If
                End If
                If (ShowPO) Then
                    Dim RefNumber As String
                    Dim TxnNumber As String
                    If (Not poRet.RefNumber Is Nothing) Then
                        RefNumber = poRet.RefNumber.GetValue
                    Else
                        RefNumber = ""
                    End If
                    If (Not poRet.TxnNumber Is Nothing) Then
                        TxnNumber = poRet.TxnNumber.GetValue
                    Else
                        TxnNumber = ""
                    End If
                    POList.Items.Add(New POItem(poRet.TxnID.GetValue(), poRet.TxnDate.GetValue(), RefNumber, TxnNumber))
                End If
            Next i
        End If
    End Sub

    Private Sub ReceiveItems_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        sessMgr = New QBSessionManager()
        sessMgr.OpenConnection("", "IDN Item Receipt Sample")
        sessMgr.BeginSession("", ENOpenMode.omDontCare)
        Dim maxVersion As Double
        maxVersion = Val(sessMgr.QBXMLVersionsForSession(UBound(sessMgr.QBXMLVersionsForSession)))
        If (maxVersion < 4.0) Then 'Careful -- Should handle CA4.0, etc. later
            MsgBox("This sample requires QuickBooks 2005 to support qbXML 4.0, current QuickBooks supports only " & maxVersion, MsgBoxStyle.Exclamation, "qbXML 4.0 not supported")
            sessMgr.EndSession()
            sessMgr.CloseConnection()
            Me.Close()
            Exit Sub
        End If
        Dim MsgRq As IMsgSetRequest
        MsgRq = sessMgr.CreateMsgSetRequest("US", 4, 0)
        Dim VQuery As IVendorQuery
        VQuery = MsgRq.AppendVendorQueryRq
        '
        ' Only get active vendors and only get the ListID and the name for efficiency
        '
        VQuery.ORVendorListQuery.VendorListFilter.ActiveStatus.SetValue(ENActiveStatus.asActiveOnly)
        VQuery.IncludeRetElementList.Add("ListID")
        VQuery.IncludeRetElementList.Add("Name")

        ' 
        ' Send request and get response
        Dim MsgRs As IMsgSetResponse
        MsgRs = sessMgr.DoRequests(MsgRq)
        Dim resp As IResponse
        resp = MsgRs.ResponseList.GetAt(0) ' Only one request, only one response

        '
        ' Check for errors
        '
        If (resp.StatusSeverity = "Error") Then
            MsgBox("Problem getting vendors (" & resp.StatusCode & "): " & resp.StatusMessage)
            Exit Sub
        End If

        If (resp.Detail Is Nothing) Then
            MsgBox("No detail to vendor query response!", MsgBoxStyle.Exclamation, "Unexpected Error")
            Exit Sub
        End If

        Dim vList As IVendorRetList
        vList = resp.Detail 'upcast to appropriate type for the query sent

        Dim i As Integer
        For i = 0 To vList.Count - 1
            Dim vRet As IVendorRet
            vRet = vList.GetAt(i)
            Dim vItem As VendorItem
            vItem = New VendorItem(vRet.ListID.GetValue(), vRet.Name.GetValue())
            VendorList.Items.Add(vItem)
        Next i
    End Sub

    Private Sub ReceiveItems_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        sessMgr.EndSession()
        sessMgr.CloseConnection()
        sessMgr = Nothing
    End Sub

    Private Sub POList_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles POList.SelectedIndexChanged
        Dim selPO As POItem
        LineItems.Items.Clear()
        selPO = POList.SelectedItem
        If (selPO Is Nothing) Then
            Exit Sub
        End If
        Dim msgRq As IMsgSetRequest
        msgRq = sessMgr.CreateMsgSetRequest("US", 4, 0)
        Dim poQ As IPurchaseOrderQuery
        poQ = msgRq.AppendPurchaseOrderQueryRq
        poQ.ORTxnQuery.TxnIDList.Add(selPO.TxnID)
        poQ.IncludeLineItems.SetValue(True)
        Dim msgRs As IMsgSetResponse
        msgRs = sessMgr.DoRequests(msgRq)
        Dim POresp As IResponse
        POresp = msgRs.ResponseList.GetAt(0) 'only one request
        '
        ' Check for errors
        '
        If (POresp.StatusSeverity = "Error") Then
            MsgBox("Problem getting selected PO (" & POresp.StatusCode & "): " & POresp.StatusMessage)
            Exit Sub
        End If

        If (POresp.Detail Is Nothing) Then
            Exit Sub
        End If

        Dim PORetList As IPurchaseOrderRetList
        PORetList = POresp.Detail 'upcast
        Dim PORet As IPurchaseOrderRet
        If (PORetList.Count < 1) Then
            MsgBox("Could not find the specified PO -- strange!")
            Exit Sub
        End If
        PORet = PORetList.GetAt(0)

        If (PORet.ORPurchaseOrderLineRetList Is Nothing) Then
            Exit Sub
        End If

        Dim lineList As IORPurchaseOrderLineRetList
        lineList = PORet.ORPurchaseOrderLineRetList
        AddLinesToPOLineList(lineList, selPO.TxnID)
    End Sub

    Private Sub AddLinesToPOLineList(ByVal lineList As IORPurchaseOrderLineRetList, ByVal TxnID As String)
        Dim i As Integer
        For i = 0 To lineList.Count - 1
            With lineList.GetAt(i)
                If (Not .PurchaseOrderLineGroupRet Is Nothing) Then
                    If (Not .PurchaseOrderLineGroupRet.PurchaseOrderLineRetList Is Nothing) Then
                        AddLinesToPOLineList(.PurchaseOrderLineGroupRet.PurchaseOrderLineRetList, TxnID)
                    End If
                End If
                If (Not .PurchaseOrderLineRet Is Nothing) Then
                    With .PurchaseOrderLineRet
                        Dim poItem As POLineItem
                        poItem = New POLineItem(TxnID, .TxnLineID.GetValue)
                        Dim lvItem As ListViewItem
                        lvItem = New ListViewItem("0")
                        lvItem.Tag = poItem
                        If (Not .ReceivedQuantity Is Nothing) Then
                            lvItem.SubItems.Add(.ReceivedQuantity.GetAsString())
                        Else
                            lvItem.SubItems.Add("0")
                        End If
                        If (Not .Quantity Is Nothing) Then
                            lvItem.SubItems.Add(.Quantity.GetAsString())
                        Else
                            lvItem.SubItems.Add("0")
                        End If
                        If (Not .Desc Is Nothing) Then
                            lvItem.SubItems.Add(.Desc.GetValue())
                        Else
                            If (Not .ItemRef Is Nothing) Then
                                lvItem.SubItems.Add(.ItemRef.FullName.GetValue())
                            End If
                        End If
                        LineItems.Items.Add(lvItem)
                    End With
                End If
            End With
        Next i
    End Sub

    Private Sub MyListView_AfterLabelEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.LabelEditEventArgs) Handles LineItems.AfterLabelEdit
        Dim lvItem As ListViewItem
        lvItem = LineItems.Items(e.Item)
        Dim poItem As POLineItem
        poItem = lvItem.Tag
        poItem.Received = e.Label
    End Sub

    Private Sub StoreIR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StoreIR.Click
        Dim msgRq As IMsgSetRequest
        msgRq = sessMgr.CreateMsgSetRequest("US", 4, 0)
        Dim IR As IItemReceiptAdd
        IR = msgRq.AppendItemReceiptAddRq
        Dim selVendor As VendorItem
        selVendor = VendorList.SelectedItem
        IR.VendorRef.ListID.SetValue(selVendor.ListID)
        Dim lvItem As ListViewItem
        Dim i As Integer
        For i = 0 To LineItems.Items.Count - 1
            lvItem = LineItems.Items.Item(i)
            Dim poLine As POLineItem
            poLine = lvItem.Tag
            If (poLine.IsDirty) Then
                Dim IRLine As IORItemLineAdd
                IRLine = IR.ORItemLineAddList.Append()
                With IRLine.ItemLineAdd
                    .LinkToTxn.TxnID.SetValue(poLine.TxnID)
                    .LinkToTxn.TxnLineID.SetValue(poLine.TxnLineID)
                    .Quantity.SetValue(poLine.Received)
                End With
            End If
        Next i
        Dim msgRs As IMsgSetResponse
        Dim sr As StreamWriter
        sr = New StreamWriter("ItemReceipt.xml")
        sr.Write(msgRq.ToXMLString())
        sr.Close()
        msgRs = sessMgr.DoRequests(msgRq)
        Dim resp As IResponse
        resp = msgRs.ResponseList.GetAt(0)
        If (resp.StatusSeverity <> "Info") Then
            MsgBox("Got message from storing ItemReceipt (" & resp.StatusCode & "): " & resp.StatusMessage)
        Else
            MsgBox("Stored ItemReceipt!")
        End If
    End Sub
End Class
