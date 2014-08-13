Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmAddDataExtension
	Inherits System.Windows.Forms.Form
#Region "Windows Form Designer generated code "
	Public Sub New()
		MyBase.New()
		If m_vb6FormDefInstance Is Nothing Then
			If m_InitializingDefInstance Then
				m_vb6FormDefInstance = Me
			Else
				Try 
					'For the start-up form, the first instance created is the default instance.
					If System.Reflection.Assembly.GetExecutingAssembly.EntryPoint.DeclaringType Is Me.GetType Then
						m_vb6FormDefInstance = Me
					End If
				Catch
				End Try
			End If
		End If
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
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
	Public WithEvents Text1 As System.Windows.Forms.TextBox
	Public WithEvents lstUsedDataExts As System.Windows.Forms.ListBox
	Public WithEvents lstAvailableDataExts As System.Windows.Forms.ListBox
	Public WithEvents cmdCloseWindow As System.Windows.Forms.Button
	Public WithEvents cmdAddValue As System.Windows.Forms.Button
	Public WithEvents txtDataExtValue As System.Windows.Forms.TextBox
	Public WithEvents lstUnusedDataExtDefs As System.Windows.Forms.ListBox
	Public WithEvents lstCustomers As System.Windows.Forms.ListBox
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Text1 = New System.Windows.Forms.TextBox()
        Me.lstUsedDataExts = New System.Windows.Forms.ListBox()
        Me.lstAvailableDataExts = New System.Windows.Forms.ListBox()
        Me.cmdCloseWindow = New System.Windows.Forms.Button()
        Me.cmdAddValue = New System.Windows.Forms.Button()
        Me.txtDataExtValue = New System.Windows.Forms.TextBox()
        Me.lstUnusedDataExtDefs = New System.Windows.Forms.ListBox()
        Me.lstCustomers = New System.Windows.Forms.ListBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Text1
        '
        Me.Text1.AcceptsReturn = True
        Me.Text1.AutoSize = False
        Me.Text1.BackColor = System.Drawing.SystemColors.Menu
        Me.Text1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Text1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text1.Location = New System.Drawing.Point(8, 8)
        Me.Text1.MaxLength = 0
        Me.Text1.Multiline = True
        Me.Text1.Name = "Text1"
        Me.Text1.ReadOnly = True
        Me.Text1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text1.Size = New System.Drawing.Size(377, 73)
        Me.Text1.TabIndex = 10
        Me.Text1.Text = "1) Display unused Data Extension definitions by selecting a " & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10) & "customer" & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10) & "2) Select" & _
        " a Data Extension value to add by selecting one" & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10) & "3) Enter a value for the Data E" & _
        "xtension" & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10) & "4) Click on the Add Data Extension Value button" & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10)
        '
        'lstUsedDataExts
        '
        Me.lstUsedDataExts.BackColor = System.Drawing.SystemColors.Window
        Me.lstUsedDataExts.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstUsedDataExts.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstUsedDataExts.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstUsedDataExts.ItemHeight = 14
        Me.lstUsedDataExts.Items.AddRange(New Object() {"lstUsedDataExts"})
        Me.lstUsedDataExts.Location = New System.Drawing.Point(208, 416)
        Me.lstUsedDataExts.Name = "lstUsedDataExts"
        Me.lstUsedDataExts.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstUsedDataExts.Size = New System.Drawing.Size(24, 18)
        Me.lstUsedDataExts.Sorted = True
        Me.lstUsedDataExts.TabIndex = 9
        Me.lstUsedDataExts.Visible = False
        '
        'lstAvailableDataExts
        '
        Me.lstAvailableDataExts.BackColor = System.Drawing.SystemColors.Window
        Me.lstAvailableDataExts.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstAvailableDataExts.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstAvailableDataExts.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstAvailableDataExts.ItemHeight = 14
        Me.lstAvailableDataExts.Location = New System.Drawing.Point(176, 416)
        Me.lstAvailableDataExts.Name = "lstAvailableDataExts"
        Me.lstAvailableDataExts.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstAvailableDataExts.Size = New System.Drawing.Size(24, 18)
        Me.lstAvailableDataExts.Sorted = True
        Me.lstAvailableDataExts.TabIndex = 8
        Me.lstAvailableDataExts.Visible = False
        '
        'cmdCloseWindow
        '
        Me.cmdCloseWindow.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCloseWindow.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCloseWindow.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCloseWindow.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCloseWindow.Location = New System.Drawing.Point(264, 408)
        Me.cmdCloseWindow.Name = "cmdCloseWindow"
        Me.cmdCloseWindow.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCloseWindow.Size = New System.Drawing.Size(121, 41)
        Me.cmdCloseWindow.TabIndex = 7
        Me.cmdCloseWindow.Text = "Close Window"
        '
        'cmdAddValue
        '
        Me.cmdAddValue.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAddValue.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAddValue.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAddValue.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAddValue.Location = New System.Drawing.Point(8, 408)
        Me.cmdAddValue.Name = "cmdAddValue"
        Me.cmdAddValue.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAddValue.Size = New System.Drawing.Size(145, 41)
        Me.cmdAddValue.TabIndex = 6
        Me.cmdAddValue.Text = "Add Data Extension Value"
        '
        'txtDataExtValue
        '
        Me.txtDataExtValue.AcceptsReturn = True
        Me.txtDataExtValue.AutoSize = False
        Me.txtDataExtValue.BackColor = System.Drawing.SystemColors.Window
        Me.txtDataExtValue.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDataExtValue.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDataExtValue.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDataExtValue.Location = New System.Drawing.Point(136, 368)
        Me.txtDataExtValue.MaxLength = 0
        Me.txtDataExtValue.Name = "txtDataExtValue"
        Me.txtDataExtValue.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDataExtValue.Size = New System.Drawing.Size(177, 19)
        Me.txtDataExtValue.TabIndex = 5
        Me.txtDataExtValue.Text = ""
        '
        'lstUnusedDataExtDefs
        '
        Me.lstUnusedDataExtDefs.BackColor = System.Drawing.SystemColors.Window
        Me.lstUnusedDataExtDefs.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstUnusedDataExtDefs.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstUnusedDataExtDefs.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstUnusedDataExtDefs.ItemHeight = 14
        Me.lstUnusedDataExtDefs.Location = New System.Drawing.Point(8, 272)
        Me.lstUnusedDataExtDefs.Name = "lstUnusedDataExtDefs"
        Me.lstUnusedDataExtDefs.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstUnusedDataExtDefs.Size = New System.Drawing.Size(377, 74)
        Me.lstUnusedDataExtDefs.TabIndex = 2
        '
        'lstCustomers
        '
        Me.lstCustomers.BackColor = System.Drawing.SystemColors.Window
        Me.lstCustomers.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstCustomers.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstCustomers.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstCustomers.ItemHeight = 14
        Me.lstCustomers.Location = New System.Drawing.Point(8, 104)
        Me.lstCustomers.Name = "lstCustomers"
        Me.lstCustomers.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstCustomers.Size = New System.Drawing.Size(377, 130)
        Me.lstCustomers.TabIndex = 0
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(8, 368)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(112, 17)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Data Extension Value"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(8, 256)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(177, 17)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Unused Data Extension Definitions"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(8, 88)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(57, 17)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Customers"
        '
        'frmAddDataExtension
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(397, 460)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Text1, Me.lstUsedDataExts, Me.lstAvailableDataExts, Me.cmdCloseWindow, Me.cmdAddValue, Me.txtDataExtValue, Me.lstUnusedDataExtDefs, Me.lstCustomers, Me.Label3, Me.Label2, Me.Label1})
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(276, 202)
        Me.Name = "frmAddDataExtension"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Add Data Extension To Customer"
        Me.ResumeLayout(False)

    End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As frmAddDataExtension
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As frmAddDataExtension
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New frmAddDataExtension()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
    Public Shared ReadOnly Property IsInitializing() As Boolean
        Get
            IsInitializing = m_InitializingDefInstance
        End Get
    End Property
#End Region
    '----------------------------------------------------------
    ' Form: frmAddDataExtension
    '
    ' Description: Allows the user to select a select a customer causing
    '              the form to list the unused data extensions for that
    '              customer.  Then the user may highlight one of the
    '              unused data extensions, provide a value for it and
    '              add it to the customer record
    '
    ' Copyright © 2002-2013 Intuit Inc. All rights reserved.
    ' Use is subject to the terms specified at:
    '      http://developer.intuit.com/legal/devsite_tos.html
    '
    '----------------------------------------------------------
    Private Sub cmdAddValue_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAddValue.Click

        If lstCustomers.SelectedIndex < 0 Then
            MsgBox("You must select a customer first")
            Exit Sub
        End If

        If lstUnusedDataExtDefs.SelectedIndex < 0 Then
            MsgBox("You must select a Data Extension to add first")
            Exit Sub
        End If

        If txtDataExtValue.Text = "" Then
            MsgBox("You must supply a value for the Data Extension first")
            Exit Sub
        End If

        Dim strCustomer As String
        Dim strDataExtName As String
        Dim strDataExtValue As String
        strCustomer = VB6.GetItemString(lstCustomers, lstCustomers.SelectedIndex)
        strDataExtValue = txtDataExtValue.Text

        ' make sure there are data extensions available to add
        Dim noDataExtNameAvailable As String
        noDataExtNameAvailable = "All Data Extensions for " & strCustomer & " are already assigned values"
        strDataExtName = VB6.GetItemString(lstUnusedDataExtDefs, lstUnusedDataExtDefs.SelectedIndex)
        If (strDataExtName = noDataExtNameAvailable) Then
            MsgBox("There are no Data Extensions available to add for this Customer")
            Exit Sub
        End If
        strDataExtName = VB.Left(strDataExtName, InStr(strDataExtName, "  |") - 1)

        If AddDataExt(strCustomer, strDataExtName, strDataExtValue) Then
            lstUnusedDataExtDefs.Items.RemoveAt((lstUnusedDataExtDefs.SelectedIndex))
            txtDataExtValue.Text = ""
            txtDataExtValue.Refresh()
        End If

    End Sub

    Private Sub cmdCloseWindow_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCloseWindow.Click
        frmAddDataExtension.DefInstance.Close()
    End Sub

    Private Sub frmAddDataExtension_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        GetCustomers(lstCustomers)
    End Sub

    Private Sub lstCustomers_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstCustomers.SelectedIndexChanged
        If frmAddDataExtension.DefInstance.IsInitializing = True Then
            Exit Sub
        Else
            lstUnusedDataExtDefs.Items.Clear()
            FillUnusedDataExts(VB6.GetItemString(lstCustomers, lstCustomers.SelectedIndex), lstUnusedDataExtDefs)
        End If
    End Sub
End Class