<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class Mainfrm
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents Comm_View_Req As System.Windows.Forms.Button
	Public WithEvents Comm_View_Res As System.Windows.Forms.Button
	Public WithEvents Comm_Exit As System.Windows.Forms.Button
	Public WithEvents Comm_Submit_Add As System.Windows.Forms.Button
	Public WithEvents Text_CustomerName As System.Windows.Forms.TextBox
	Public WithEvents Comm_Submit_Remove As System.Windows.Forms.Button
	Public WithEvents Text_ListID As System.Windows.Forms.TextBox
	Public WithEvents Comm_Submit_Find As System.Windows.Forms.Button
    Public WithEvents Image_QBBANNER As System.Windows.Forms.PictureBox
    Public WithEvents Label7 As System.Windows.Forms.Label
	Public WithEvents Image_QBCUST As System.Windows.Forms.PictureBox
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Mainfrm))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Comm_View_Req = New System.Windows.Forms.Button()
        Me.Comm_View_Res = New System.Windows.Forms.Button()
        Me.Comm_Exit = New System.Windows.Forms.Button()
        Me.Comm_Submit_Add = New System.Windows.Forms.Button()
        Me.Text_CustomerName = New System.Windows.Forms.TextBox()
        Me.Comm_Submit_Remove = New System.Windows.Forms.Button()
        Me.Text_ListID = New System.Windows.Forms.TextBox()
        Me.Comm_Submit_Find = New System.Windows.Forms.Button()
        Me.Image_QBBANNER = New System.Windows.Forms.PictureBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Image_QBCUST = New System.Windows.Forms.PictureBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        CType(Me.Image_QBBANNER, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Image_QBCUST, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Comm_View_Req
        '
        Me.Comm_View_Req.BackColor = System.Drawing.SystemColors.Control
        Me.Comm_View_Req.Cursor = System.Windows.Forms.Cursors.Default
        Me.Comm_View_Req.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Comm_View_Req.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Comm_View_Req.Location = New System.Drawing.Point(512, 152)
        Me.Comm_View_Req.Name = "Comm_View_Req"
        Me.Comm_View_Req.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Comm_View_Req.Size = New System.Drawing.Size(137, 25)
        Me.Comm_View_Req.TabIndex = 12
        Me.Comm_View_Req.Text = "&View qbXML Request"
        Me.Comm_View_Req.UseVisualStyleBackColor = False
        '
        'Comm_View_Res
        '
        Me.Comm_View_Res.BackColor = System.Drawing.SystemColors.Control
        Me.Comm_View_Res.Cursor = System.Windows.Forms.Cursors.Default
        Me.Comm_View_Res.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Comm_View_Res.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Comm_View_Res.Location = New System.Drawing.Point(512, 200)
        Me.Comm_View_Res.Name = "Comm_View_Res"
        Me.Comm_View_Res.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Comm_View_Res.Size = New System.Drawing.Size(137, 25)
        Me.Comm_View_Res.TabIndex = 11
        Me.Comm_View_Res.Text = "Vi&ew qbXML Response"
        Me.Comm_View_Res.UseVisualStyleBackColor = False
        '
        'Comm_Exit
        '
        Me.Comm_Exit.BackColor = System.Drawing.SystemColors.Control
        Me.Comm_Exit.Cursor = System.Windows.Forms.Cursors.Default
        Me.Comm_Exit.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Comm_Exit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Comm_Exit.Location = New System.Drawing.Point(512, 248)
        Me.Comm_Exit.Name = "Comm_Exit"
        Me.Comm_Exit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Comm_Exit.Size = New System.Drawing.Size(137, 25)
        Me.Comm_Exit.TabIndex = 10
        Me.Comm_Exit.Text = "&Done"
        Me.Comm_Exit.UseVisualStyleBackColor = False
        '
        'Comm_Submit_Add
        '
        Me.Comm_Submit_Add.BackColor = System.Drawing.SystemColors.Control
        Me.Comm_Submit_Add.Cursor = System.Windows.Forms.Cursors.Default
        Me.Comm_Submit_Add.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Comm_Submit_Add.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Comm_Submit_Add.Location = New System.Drawing.Point(312, 128)
        Me.Comm_Submit_Add.Name = "Comm_Submit_Add"
        Me.Comm_Submit_Add.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Comm_Submit_Add.Size = New System.Drawing.Size(137, 25)
        Me.Comm_Submit_Add.TabIndex = 4
        Me.Comm_Submit_Add.Text = "&Add Customer"
        Me.Comm_Submit_Add.UseVisualStyleBackColor = False
        '
        'Text_CustomerName
        '
        Me.Text_CustomerName.AcceptsReturn = True
        Me.Text_CustomerName.BackColor = System.Drawing.SystemColors.Window
        Me.Text_CustomerName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text_CustomerName.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text_CustomerName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text_CustomerName.Location = New System.Drawing.Point(312, 104)
        Me.Text_CustomerName.MaxLength = 0
        Me.Text_CustomerName.Name = "Text_CustomerName"
        Me.Text_CustomerName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text_CustomerName.Size = New System.Drawing.Size(137, 20)
        Me.Text_CustomerName.TabIndex = 3
        '
        'Comm_Submit_Remove
        '
        Me.Comm_Submit_Remove.BackColor = System.Drawing.SystemColors.Control
        Me.Comm_Submit_Remove.Cursor = System.Windows.Forms.Cursors.Default
        Me.Comm_Submit_Remove.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Comm_Submit_Remove.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Comm_Submit_Remove.Location = New System.Drawing.Point(312, 368)
        Me.Comm_Submit_Remove.Name = "Comm_Submit_Remove"
        Me.Comm_Submit_Remove.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Comm_Submit_Remove.Size = New System.Drawing.Size(137, 25)
        Me.Comm_Submit_Remove.TabIndex = 2
        Me.Comm_Submit_Remove.Text = "Delete Customer"
        Me.Comm_Submit_Remove.UseVisualStyleBackColor = False
        '
        'Text_ListID
        '
        Me.Text_ListID.AcceptsReturn = True
        Me.Text_ListID.BackColor = System.Drawing.SystemColors.Window
        Me.Text_ListID.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text_ListID.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text_ListID.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text_ListID.Location = New System.Drawing.Point(40, 376)
        Me.Text_ListID.MaxLength = 0
        Me.Text_ListID.Name = "Text_ListID"
        Me.Text_ListID.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text_ListID.Size = New System.Drawing.Size(169, 20)
        Me.Text_ListID.TabIndex = 1
        '
        'Comm_Submit_Find
        '
        Me.Comm_Submit_Find.BackColor = System.Drawing.SystemColors.Control
        Me.Comm_Submit_Find.Cursor = System.Windows.Forms.Cursors.Default
        Me.Comm_Submit_Find.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Comm_Submit_Find.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Comm_Submit_Find.Location = New System.Drawing.Point(312, 216)
        Me.Comm_Submit_Find.Name = "Comm_Submit_Find"
        Me.Comm_Submit_Find.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Comm_Submit_Find.Size = New System.Drawing.Size(137, 30)
        Me.Comm_Submit_Find.TabIndex = 0
        Me.Comm_Submit_Find.Text = "Find An ""In Use"" Customer"
        Me.Comm_Submit_Find.UseVisualStyleBackColor = False
        '
        'Image_QBBANNER
        '
        Me.Image_QBBANNER.Cursor = System.Windows.Forms.Cursors.Default
        Me.Image_QBBANNER.Image = CType(resources.GetObject("Image_QBBANNER.Image"), System.Drawing.Image)
        Me.Image_QBBANNER.Location = New System.Drawing.Point(48, 448)
        Me.Image_QBBANNER.Name = "Image_QBBANNER"
        Me.Image_QBBANNER.Size = New System.Drawing.Size(430, 29)
        Me.Image_QBBANNER.TabIndex = 13
        Me.Image_QBBANNER.TabStop = False
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(316, 80)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(173, 17)
        Me.Label7.TabIndex = 9
        Me.Label7.Text = "&Customer (Required for Add):"
        '
        'Image_QBCUST
        '
        Me.Image_QBCUST.Cursor = System.Windows.Forms.Cursors.Default
        Me.Image_QBCUST.Image = CType(resources.GetObject("Image_QBCUST.Image"), System.Drawing.Image)
        Me.Image_QBCUST.Location = New System.Drawing.Point(16, 24)
        Me.Image_QBCUST.Name = "Image_QBCUST"
        Me.Image_QBCUST.Size = New System.Drawing.Size(30, 33)
        Me.Image_QBCUST.TabIndex = 14
        Me.Image_QBCUST.TabStop = False
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(40, 360)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(177, 17)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "List&ID (Required for Delete):"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(48, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(485, 29)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "This sample program illustrates how to delete a customer using the QuickBooks SDK" &
    "."
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(44, 88)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(237, 81)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = resources.GetString("Label2.Text")
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(40, 216)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(241, 113)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = resources.GetString("Label3.Text")
        '
        'Mainfrm
        '
        Me.AcceptButton = Me.Comm_Submit_Add
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(659, 526)
        Me.Controls.Add(Me.Comm_View_Req)
        Me.Controls.Add(Me.Comm_View_Res)
        Me.Controls.Add(Me.Comm_Exit)
        Me.Controls.Add(Me.Comm_Submit_Add)
        Me.Controls.Add(Me.Text_CustomerName)
        Me.Controls.Add(Me.Comm_Submit_Remove)
        Me.Controls.Add(Me.Text_ListID)
        Me.Controls.Add(Me.Comm_Submit_Find)
        Me.Controls.Add(Me.Image_QBBANNER)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Image_QBCUST)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label3)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(4, 23)
        Me.Name = "Mainfrm"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "qbXML Sample: Deleting a Customer"
        CType(Me.Image_QBBANNER, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Image_QBCUST, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
#End Region
End Class