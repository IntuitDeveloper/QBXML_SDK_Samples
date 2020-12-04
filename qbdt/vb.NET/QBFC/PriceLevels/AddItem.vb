Imports Interop.QBFC13

Public Class AddItem
    Inherits System.Windows.Forms.Form
    Public sessMgr As QBSessionManager
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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ItemList As System.Windows.Forms.CheckedListBox
    Friend WithEvents OK As System.Windows.Forms.Button
    Friend WithEvents Cancel As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ItemList = New System.Windows.Forms.CheckedListBox()
        Me.OK = New System.Windows.Forms.Button()
        Me.Cancel = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(264, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Select items to add custom prices in price level:"
        '
        'ItemList
        '
        Me.ItemList.CheckOnClick = True
        Me.ItemList.Location = New System.Drawing.Point(8, 32)
        Me.ItemList.Name = "ItemList"
        Me.ItemList.Size = New System.Drawing.Size(264, 184)
        Me.ItemList.TabIndex = 1
        '
        'OK
        '
        Me.OK.Location = New System.Drawing.Point(8, 224)
        Me.OK.Name = "OK"
        Me.OK.Size = New System.Drawing.Size(96, 32)
        Me.OK.TabIndex = 2
        Me.OK.Text = "OK"
        '
        'Cancel
        '
        Me.Cancel.Location = New System.Drawing.Point(120, 224)
        Me.Cancel.Name = "Cancel"
        Me.Cancel.Size = New System.Drawing.Size(96, 32)
        Me.Cancel.TabIndex = 3
        Me.Cancel.Text = "Cancel"
        '
        'AddItem
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 266)
        Me.ControlBox = False
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Cancel, Me.OK, Me.ItemList, Me.Label1})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "AddItem"
        Me.Text = "Select Item to Add"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub AddItem_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If (Not sessMgr Is Nothing) Then
            Dim MsgRq As IMsgSetRequest
            MsgRq = sessMgr.CreateMsgSetRequest("US", 4, 0)
            Dim itemQ As IItemQuery
            itemQ = MsgRq.AppendItemQueryRq
            itemQ.IncludeRetElementList.Add("ListID")
            itemQ.IncludeRetElementList.Add("FullName")
            itemQ.ORListQuery.ListFilter.ActiveStatus.SetValue(ENActiveStatus.asActiveOnly)
            Dim MsgRs As IMsgSetResponse
            MsgRs = sessMgr.DoRequests(MsgRq)
            Dim Qresp As IResponse
            Qresp = MsgRs.ResponseList.GetAt(0)
            If (Qresp.StatusCode = 0) Then
                Dim ORItemList As IORItemRetList
                ORItemList = Qresp.Detail
                Dim i As Integer
                For i = 0 To ORItemList.Count - 1
                    With ORItemList.GetAt(i)
                        '
                        ' We only care about Inventory and NonInventory items (I think)
                        '
                        If (Not .ItemInventoryRet Is Nothing) Then
                            With .ItemInventoryRet
                                ItemList.Items.Add(.FullName.GetValue, False)
                            End With
                        End If
                        If (Not .ItemNonInventoryRet Is Nothing) Then
                            With .ItemNonInventoryRet
                                ItemList.Items.Add(.FullName.GetValue, False)
                            End With
                        End If
                    End With
                Next i
            Else
                MsgBox("Error from QuickBooks (" & Qresp.StatusCode & "):" & Qresp.StatusMessage, MsgBoxStyle.Exclamation, "Problem querying items")
            End If
        End If
    End Sub

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        Me.DialogResult = DialogResult.OK
    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.DialogResult = DialogResult.Cancel
    End Sub
End Class
