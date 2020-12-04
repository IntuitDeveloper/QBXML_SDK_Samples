Imports Interop.QBFC13

Public Class PriceLevelSample
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
    Friend WithEvents GetPriceLevels As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ItemList As System.Windows.Forms.ListView
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents ListName As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ListID As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents IsActive As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Fixed As System.Windows.Forms.RadioButton
    Friend WithEvents PerItem As System.Windows.Forms.RadioButton
    Friend WithEvents FixedPercent As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents NewPriceLevel As System.Windows.Forms.Button
    Friend WithEvents Submit As System.Windows.Forms.Button
    Friend WithEvents PriceLevels As System.Windows.Forms.ListBox
    Friend WithEvents AddItem As System.Windows.Forms.Button
    Friend WithEvents EditSequence As System.Windows.Forms.TextBox
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.GetPriceLevels = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ItemList = New System.Windows.Forms.ListView()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.EditSequence = New System.Windows.Forms.TextBox()
        Me.AddItem = New System.Windows.Forms.Button()
        Me.Submit = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.FixedPercent = New System.Windows.Forms.TextBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.PerItem = New System.Windows.Forms.RadioButton()
        Me.Fixed = New System.Windows.Forms.RadioButton()
        Me.IsActive = New System.Windows.Forms.CheckBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.ListID = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ListName = New System.Windows.Forms.TextBox()
        Me.NewPriceLevel = New System.Windows.Forms.Button()
        Me.PriceLevels = New System.Windows.Forms.ListBox()
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GetPriceLevels
        '
        Me.GetPriceLevels.Location = New System.Drawing.Point(8, 8)
        Me.GetPriceLevels.Name = "GetPriceLevels"
        Me.GetPriceLevels.Size = New System.Drawing.Size(144, 32)
        Me.GetPriceLevels.TabIndex = 1
        Me.GetPriceLevels.Text = "Get Current Price Levels"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 48)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(384, 16)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Current Price Levels:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'ItemList
        '
        Me.ItemList.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4})
        Me.ItemList.GridLines = True
        Me.ItemList.Location = New System.Drawing.Point(8, 136)
        Me.ItemList.Name = "ItemList"
        Me.ItemList.Size = New System.Drawing.Size(512, 152)
        Me.ItemList.TabIndex = 3
        Me.ItemList.View = System.Windows.Forms.View.Details
        Me.ItemList.Visible = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.EditSequence, Me.AddItem, Me.Submit, Me.Label4, Me.FixedPercent, Me.GroupBox2, Me.IsActive, Me.Label3, Me.ListID, Me.Label2, Me.ListName, Me.ItemList})
        Me.GroupBox1.Location = New System.Drawing.Point(8, 176)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(528, 344)
        Me.GroupBox1.TabIndex = 4
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Price Level Configuration"
        '
        'EditSequence
        '
        Me.EditSequence.Location = New System.Drawing.Point(328, 72)
        Me.EditSequence.Name = "EditSequence"
        Me.EditSequence.Size = New System.Drawing.Size(184, 20)
        Me.EditSequence.TabIndex = 14
        Me.EditSequence.Text = "TextBox1"
        Me.EditSequence.Visible = False
        '
        'AddItem
        '
        Me.AddItem.Location = New System.Drawing.Point(328, 104)
        Me.AddItem.Name = "AddItem"
        Me.AddItem.Size = New System.Drawing.Size(184, 24)
        Me.AddItem.TabIndex = 13
        Me.AddItem.Text = "Add Custom Item Price"
        '
        'Submit
        '
        Me.Submit.Location = New System.Drawing.Point(16, 304)
        Me.Submit.Name = "Submit"
        Me.Submit.Size = New System.Drawing.Size(152, 32)
        Me.Submit.TabIndex = 12
        Me.Submit.Text = "Submit Price Level Info"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(208, 88)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(100, 16)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "Fixed %:"
        '
        'FixedPercent
        '
        Me.FixedPercent.Location = New System.Drawing.Point(208, 104)
        Me.FixedPercent.Name = "FixedPercent"
        Me.FixedPercent.Size = New System.Drawing.Size(104, 20)
        Me.FixedPercent.TabIndex = 10
        Me.FixedPercent.Text = ""
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.AddRange(New System.Windows.Forms.Control() {Me.PerItem, Me.Fixed})
        Me.GroupBox2.Location = New System.Drawing.Point(16, 80)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(176, 48)
        Me.GroupBox2.TabIndex = 9
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Price Level Type"
        '
        'PerItem
        '
        Me.PerItem.Location = New System.Drawing.Point(88, 24)
        Me.PerItem.Name = "PerItem"
        Me.PerItem.Size = New System.Drawing.Size(72, 16)
        Me.PerItem.TabIndex = 1
        Me.PerItem.Text = "Per-Item"
        '
        'Fixed
        '
        Me.Fixed.Checked = True
        Me.Fixed.Location = New System.Drawing.Point(8, 24)
        Me.Fixed.Name = "Fixed"
        Me.Fixed.Size = New System.Drawing.Size(72, 16)
        Me.Fixed.TabIndex = 0
        Me.Fixed.TabStop = True
        Me.Fixed.Text = "Fixed %"
        '
        'IsActive
        '
        Me.IsActive.Location = New System.Drawing.Point(352, 40)
        Me.IsActive.Name = "IsActive"
        Me.IsActive.Size = New System.Drawing.Size(64, 16)
        Me.IsActive.TabIndex = 8
        Me.IsActive.Text = "Active"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(16, 24)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(128, 16)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "ListID"
        '
        'ListID
        '
        Me.ListID.Enabled = False
        Me.ListID.Location = New System.Drawing.Point(16, 40)
        Me.ListID.Name = "ListID"
        Me.ListID.Size = New System.Drawing.Size(184, 20)
        Me.ListID.TabIndex = 6
        Me.ListID.TabStop = False
        Me.ListID.Text = ""
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(208, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(104, 16)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Name:"
        '
        'ListName
        '
        Me.ListName.Location = New System.Drawing.Point(208, 40)
        Me.ListName.Name = "ListName"
        Me.ListName.Size = New System.Drawing.Size(136, 20)
        Me.ListName.TabIndex = 4
        Me.ListName.Text = ""
        '
        'NewPriceLevel
        '
        Me.NewPriceLevel.Location = New System.Drawing.Point(176, 8)
        Me.NewPriceLevel.Name = "NewPriceLevel"
        Me.NewPriceLevel.Size = New System.Drawing.Size(152, 32)
        Me.NewPriceLevel.TabIndex = 5
        Me.NewPriceLevel.Text = "Create New"
        '
        'PriceLevels
        '
        Me.PriceLevels.Location = New System.Drawing.Point(8, 64)
        Me.PriceLevels.Name = "PriceLevels"
        Me.PriceLevels.Size = New System.Drawing.Size(528, 95)
        Me.PriceLevels.TabIndex = 6
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "List ID"
        Me.ColumnHeader1.Width = 42
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Name"
        Me.ColumnHeader2.Width = 40
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Std. Price"
        Me.ColumnHeader3.Width = 58
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "New Price"
        Me.ColumnHeader4.Width = 368
        '
        'PriceLevelSample
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(552, 526)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.PriceLevels, Me.NewPriceLevel, Me.GroupBox1, Me.Label1, Me.GetPriceLevels})
        Me.Name = "PriceLevelSample"
        Me.Text = "Price Level Sample"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private sessMgr As QBSessionManager
    Private supports40 As Boolean


    Private Sub GetPriceLevels_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GetPriceLevels.Click
        Dim MsgRq As IMsgSetRequest
        MsgRq = sessMgr.CreateMsgSetRequest("US", 4, 0)
        Dim plQuery As IPriceLevelQuery
        plQuery = MsgRq.AppendPriceLevelQueryRq()
        plQuery.IncludeRetElementList.Add("Name")
        Dim MsgRs As IMsgSetResponse
        MsgRs = sessMgr.DoRequests(MsgRq)
        '
        ' Only one query, only one response (PriceLevelQueryRs)
        Dim plResp As IResponse
        plResp = MsgRs.ResponseList.GetAt(0)
        If (plResp.StatusSeverity = "Error") Then
            MsgBox("Error reading price levels: " & plResp.StatusCode & ":" & plResp.StatusMessage)
        Else
            If (Not plResp.Detail Is Nothing) Then
                Dim plList As IPriceLevelRetList
                plList = plResp.Detail
                Dim i As Integer
                For i = 0 To plList.Count - 1
                    Dim plRet As IPriceLevelRet
                    plRet = plList.GetAt(i)
                    PriceLevels.Items.Add(plRet.Name.GetValue)
                Next i
            End If
        End If
    End Sub

    Private Sub PriceLevelSample_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        supports40 = True
        sessMgr = New QBSessionManager()
        sessMgr.OpenConnection("", "IDN Price Level Sample")
        sessMgr.BeginSession("", ENOpenMode.omDontCare)
        Dim maxVersion As Double
        maxVersion = Val(sessMgr.QBXMLVersionsForSession(UBound(sessMgr.QBXMLVersionsForSession)))
        If (maxVersion < 4.0) Then 'Careful -- Should handle CA4.0, etc. later
            MsgBox("This sample requires QuickBooks 2005 to support qbXML 4.0, current QuickBooks supports only " & maxVersion, MsgBoxStyle.Exclamation, "qbXML 4.0 not supported")
            sessMgr.EndSession()
            sessMgr.CloseConnection()
            supports40 = False
            GetPriceLevels.Enabled = False
            NewPriceLevel.Enabled = False
            Submit.Enabled = False
        End If
    End Sub

    Private Sub PriceLevelSample_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        sessMgr.EndSession()
        sessMgr.CloseConnection()
        sessMgr = Nothing
    End Sub

    Private Sub FillInPLDetail(ByVal plRet As IPriceLevelRet)
        ListID.Text = plRet.ListID.GetValue
        ListName.Text = plRet.Name.GetValue
        EditSequence.Text = plRet.EditSequence.GetValue
        IsActive.Checked = plRet.IsActive.GetValue
        ItemList.Items.Clear()
        If (Not (plRet.ORPriceLevelRet.PriceLevelFixedPercentage Is Nothing)) Then
            Fixed.Checked = True
            FixedPercent.Text = plRet.ORPriceLevelRet.PriceLevelFixedPercentage.GetValue
        Else
            Fixed.Checked = False
            PerItem.Checked = True
            FixedPercent.Visible = False
            ItemList.Visible = True
            Dim i As Integer
            For i = 0 To plRet.ORPriceLevelRet.PriceLevelPerItemRetList.Count - 1
                Dim lvItem As ListViewItem
                Dim plItem As IPriceLevelPerItemRet
                plItem = plRet.ORPriceLevelRet.PriceLevelPerItemRetList.GetAt(i)
                lvItem = New ListViewItem(plItem.ItemRef.ListID.GetValue)
                lvItem.SubItems.Add(plItem.ItemRef.FullName.GetValue)
                lvItem.SubItems.Add("")
                lvItem.SubItems.Add(plItem.ORORCustomPrice.CustomPrice.GetAsString)
                ItemList.Items.Add(lvItem)
            Next i
        End If
    End Sub

    Private Sub PriceLevels_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PriceLevels.SelectedIndexChanged
        Dim MsgRq As IMsgSetRequest
        MsgRq = sessMgr.CreateMsgSetRequest("US", 4, 0)
        Dim plQuery As IPriceLevelQuery
        plQuery = MsgRq.AppendPriceLevelQueryRq()
        Dim selectedLevel As String
        selectedLevel = PriceLevels.Items.Item(PriceLevels.SelectedIndex)
        plQuery.ORPriceLevelQuery.FullNameList.Add(selectedLevel)
        Dim MsgRs As IMsgSetResponse
        MsgRs = sessMgr.DoRequests(MsgRq)
        '
        ' Only one query, only one response (PriceLevelQueryRs)
        Dim plResp As IResponse
        plResp = MsgRs.ResponseList.GetAt(0)
        If (plResp.StatusSeverity = "Error") Then
            MsgBox("Error reading selected price level: " & plResp.StatusCode & ":" & plResp.StatusMessage)
        Else
            If (Not plResp.Detail Is Nothing) Then
                Dim plList As IPriceLevelRetList
                plList = plResp.Detail
                Dim plRet As IPriceLevelRet
                plRet = plList.GetAt(0)
                FillInPLDetail(plRet)
            End If
        End If

    End Sub

    Private Sub NewPriceLevel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewPriceLevel.Click
        Fixed.Checked = True
        ListID.Text = ""
        ListName.Text = ""
        IsActive.Checked = True
        EditSequence.Text = ""
        ItemList.Items.Clear()
    End Sub

    Private Sub Fixed_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Fixed.CheckedChanged
        If (Fixed.Checked) Then
            FixedPercent.Visible = True
            AddItem.Visible = False
            AddItem.Enabled = False
            ItemList.Visible = False
        Else
            FixedPercent.Visible = False
            AddItem.Enabled = True
            AddItem.Visible = True
            ItemList.Visible = True
        End If
    End Sub

    Private Sub Submit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Submit.Click
        Dim MsgRq As IMsgSetRequest
        MsgRq = sessMgr.CreateMsgSetRequest("US", 4, 0)
        '
        ' Are we modifying or adding?  We can tell by whether the ListID control has a value
        '
        If (ListID.Text = "") Then
            ' Adding
            Dim plSubmit As IPriceLevelAdd
            plSubmit = MsgRq.AppendPriceLevelAddRq
            plSubmit.IsActive.SetValue(IsActive.Checked)
            plSubmit.Name.SetValue(ListName.Text)
            If (Fixed.Checked) Then
                plSubmit.ORPriceLevel.PriceLevelFixedPercentage.SetValue(FixedPercent.Text)
            Else
                Dim i As Integer
                For i = 0 To ItemList.Items.Count - 1
                    Dim lvItem As ListViewItem
                    Dim plItem As IPriceLevelPerItem
                    lvItem = ItemList.Items.Item(i)
                    'check if this item has a custom price set
                    If (Not lvItem.SubItems.Item(3).Text = "") Then
                        plItem = plSubmit.ORPriceLevel.PriceLevelPerItemList.Append()
                        plItem.ItemRef.ListID.SetValue(lvItem.Text)
                        plItem.ORPriceLevelPrice.ORCustomPrice.ORORCustomPrice.CustomPrice.SetValue(lvItem.SubItems.Item(3).Text)
                    End If
                Next i
            End If
        Else
            ' modifying
            Dim plSubmit As IPriceLevelMod
            plSubmit = MsgRq.AppendPriceLevelModRq
            plSubmit.ListID.SetValue(ListID.Text)
            plSubmit.IsActive.SetValue(IsActive.Checked)
            plSubmit.Name.SetValue(ListName.Text)
            plSubmit.EditSequence.SetValue(EditSequence.Text)
            If (Fixed.Checked) Then
                plSubmit.ORPriceLevel.PriceLevelFixedPercentage.SetValue(FixedPercent.Text)
            Else
                Dim i As Integer
                For i = 0 To ItemList.Items.Count - 1
                    Dim lvItem As ListViewItem
                    Dim plItem As IPriceLevelPerItem
                    lvItem = ItemList.Items.Item(i)
                    'check if this item has a custom price set
                    If (Not lvItem.SubItems.Item(3).Text = "") Then
                        plItem = plSubmit.ORPriceLevel.PriceLevelPerItemList.Append()
                        plItem.ItemRef.ListID.SetValue(lvItem.Text)
                        plItem.ORPriceLevelPrice.ORCustomPrice.ORORCustomPrice.CustomPrice.SetValue(lvItem.SubItems.Item(3).Text)
                    End If
                Next i
            End If
        End If

        '
        ' Send request and show response
        '
        Dim MsgRs As IMsgSetResponse
        MsgRs = sessMgr.DoRequests(MsgRq)
        Dim plResp As IResponse
        plResp = MsgRs.ResponseList.GetAt(0)
        If (plResp.StatusCode <> 0) Then
            MsgBox("Error from QuickBooks (" & plResp.StatusCode & "):" & plResp.StatusMessage, MsgBoxStyle.Exclamation, "Could not add/modify price level")
        Else
            Dim plRet As IPriceLevelRet
            plRet = plResp.Detail
            FillInPLDetail(plRet)
        End If
    End Sub

    Private Sub AddItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddItem.Click
        Dim AddItemDlg As AddItem
        AddItemDlg = New AddItem()
        AddItemDlg.sessMgr = sessMgr
        If (AddItemDlg.ShowDialog(Me) = DialogResult.OK) Then
            Dim MsgRq As IMsgSetRequest
            MsgRq = sessMgr.CreateMsgSetRequest("US", 4, 0)
            Dim itemQ As IItemQuery
            itemQ = MsgRq.AppendItemQueryRq
            Dim itemChecked As Object
            For Each itemChecked In AddItemDlg.ItemList.CheckedItems
                Dim itemName As String
                itemName = itemChecked.ToString
                itemQ.ORListQuery.FullNameList.Add(itemName)
            Next
            Dim MsgRs As IMsgSetResponse
            MsgRs = sessMgr.DoRequests(MsgRq)
            Dim Qresp As IResponse
            Qresp = MsgRs.ResponseList.GetAt(0)
            If (Qresp.StatusCode = 0) Then
                Dim ORItemList As IORItemRetList
                ORItemList = Qresp.Detail
                Dim i As Integer
                For i = 0 To ORItemList.Count - 1
                    Dim lvItem As ListViewItem
                    With ORItemList.GetAt(i)
                        If (Not .ItemInventoryRet Is Nothing) Then
                            With .ItemInventoryRet
                                lvItem = New ListViewItem(.ListID.GetValue)
                                lvItem.SubItems.Add(.Name.GetValue)
                                lvItem.SubItems.Add(.SalesPrice.GetValue)
                                lvItem.SubItems.Add("")
                            End With
                            ItemList.Items.Add(lvItem)
                        End If
                        If (Not .ItemNonInventoryRet Is Nothing) Then
                            With .ItemNonInventoryRet
                                lvItem = New ListViewItem(.ListID.GetValue)
                                lvItem.SubItems.Add(.Name.GetValue)
                                If (Not .ORSalesPurchase.SalesAndPurchase Is Nothing) Then
                                    lvItem.SubItems.Add(.ORSalesPurchase.SalesAndPurchase.SalesPrice.GetValue)
                                ElseIf (Not .ORSalesPurchase.SalesOrPurchase Is Nothing) Then
                                    If (Not .ORSalesPurchase.SalesOrPurchase.ORPrice.Price Is Nothing) Then
                                        lvItem.SubItems.Add(.ORSalesPurchase.SalesOrPurchase.ORPrice.Price.GetValue)
                                    Else
                                        lvItem.SubItems.Add("")
                                    End If
                                End If
                            End With
                            lvItem.SubItems.Add("")
                            ItemList.Items.Add(lvItem)
                        End If
                    End With
                Next i
            End If
        End If
    End Sub

    Private Sub ItemList_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ItemList.SelectedIndexChanged
        Dim selItem As ListViewItem
        If (ItemList.SelectedItems.Count > 0) Then
            selItem = ItemList.SelectedItems(0)
            Dim newPrice As String
            newPrice = InputBox("Enter special price for " & selItem.SubItems(0).Text & ":", "", selItem.SubItems(2).Text)
            selItem.SubItems(3).Text = newPrice
        End If
    End Sub
End Class
