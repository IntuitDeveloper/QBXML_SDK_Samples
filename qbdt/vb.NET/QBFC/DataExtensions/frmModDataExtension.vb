Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmModDataExtension
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
	Public WithEvents lstCustomers As System.Windows.Forms.ListBox
	Public WithEvents lstUsedDataExts As System.Windows.Forms.ListBox
	Public WithEvents txtDataExtValue As System.Windows.Forms.TextBox
	Public WithEvents cmdModValue As System.Windows.Forms.Button
	Public WithEvents cmdCloseWindow As System.Windows.Forms.Button
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Text1 = New System.Windows.Forms.TextBox()
        Me.lstCustomers = New System.Windows.Forms.ListBox()
        Me.lstUsedDataExts = New System.Windows.Forms.ListBox()
        Me.txtDataExtValue = New System.Windows.Forms.TextBox()
        Me.cmdModValue = New System.Windows.Forms.Button()
        Me.cmdCloseWindow = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
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
        Me.Text1.TabIndex = 8
        Me.Text1.TabStop = False
        Me.Text1.Text = "1) Display the data extension values added to a customer by " & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10) & "selecting that cust" & _
        "omer" & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10) & "2) Select a data extension value to modify" & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10) & "3) Enter a new data extension " & _
        "value" & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10) & "4) Click on the Modify Data Extension Value to complete the " & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10) & "modificatio" & _
        "n" & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10)
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
        Me.lstCustomers.TabIndex = 4
        Me.lstCustomers.TabStop = False
        '
        'lstUsedDataExts
        '
        Me.lstUsedDataExts.BackColor = System.Drawing.SystemColors.Window
        Me.lstUsedDataExts.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstUsedDataExts.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstUsedDataExts.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstUsedDataExts.ItemHeight = 14
        Me.lstUsedDataExts.Location = New System.Drawing.Point(8, 272)
        Me.lstUsedDataExts.Name = "lstUsedDataExts"
        Me.lstUsedDataExts.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstUsedDataExts.Size = New System.Drawing.Size(377, 74)
        Me.lstUsedDataExts.TabIndex = 3
        Me.lstUsedDataExts.TabStop = False
        '
        'txtDataExtValue
        '
        Me.txtDataExtValue.AcceptsReturn = True
        Me.txtDataExtValue.AutoSize = False
        Me.txtDataExtValue.BackColor = System.Drawing.SystemColors.Window
        Me.txtDataExtValue.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDataExtValue.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDataExtValue.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDataExtValue.Location = New System.Drawing.Point(160, 368)
        Me.txtDataExtValue.MaxLength = 0
        Me.txtDataExtValue.Name = "txtDataExtValue"
        Me.txtDataExtValue.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDataExtValue.Size = New System.Drawing.Size(177, 19)
        Me.txtDataExtValue.TabIndex = 2
        Me.txtDataExtValue.Text = ""
        '
        'cmdModValue
        '
        Me.cmdModValue.BackColor = System.Drawing.SystemColors.Control
        Me.cmdModValue.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdModValue.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdModValue.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdModValue.Location = New System.Drawing.Point(8, 408)
        Me.cmdModValue.Name = "cmdModValue"
        Me.cmdModValue.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdModValue.Size = New System.Drawing.Size(161, 41)
        Me.cmdModValue.TabIndex = 1
        Me.cmdModValue.Text = "Modify Data Extension Value"
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
        Me.cmdCloseWindow.TabIndex = 0
        Me.cmdCloseWindow.Text = "Close Window"
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
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Customers"
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
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Data Extension Values"
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
        Me.Label3.Size = New System.Drawing.Size(136, 17)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "New Data Extension Value"
        '
        'frmModDataExtension
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(398, 461)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Text1, Me.lstCustomers, Me.lstUsedDataExts, Me.txtDataExtValue, Me.cmdModValue, Me.cmdCloseWindow, Me.Label1, Me.Label2, Me.Label3})
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(289, 189)
        Me.Name = "frmModDataExtension"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Modify Data Extension Value for Customer"
        Me.ResumeLayout(False)

    End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As frmModDataExtension
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As frmModDataExtension
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New frmModDataExtension()
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
    ' Form: frmModDataExtension
    '
    ' Description: Allows the user to highlight a customer which brings
    '              up a list of used data extensions for that customer.
    '              The user may then highlight a data extension and modify
    '              the value given for that customer.
    '
    ' Copyright © 2002-2013 Intuit Inc. All rights reserved.
    ' Use is subject to the terms specified at:
    '      http://developer.intuit.com/legal/devsite_tos.html
    '
    '----------------------------------------------------------
    Private Sub cmdCloseWindow_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCloseWindow.Click
        frmModDataExtension.DefInstance.Close()
    End Sub

    Private Sub cmdModValue_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdModValue.Click
        If lstUsedDataExts.SelectedIndex < 0 Then
            MsgBox("You must select a data extension to modify")
            Exit Sub
        End If

        If txtDataExtValue.Text = "" Then
            MsgBox("You must select a data extension to modify")
            Exit Sub
        End If

        Dim strCustomer As String = VB6.GetItemString(lstCustomers, lstCustomers.SelectedIndex)
        Dim strDataExtName As String
        strCustomer = VB6.GetItemString(lstCustomers, lstCustomers.SelectedIndex)
        strDataExtName = VB6.GetItemString(lstUsedDataExts, lstUsedDataExts.SelectedIndex)
        If (strDataExtName = "This customer has no data extensions added to their record") Then
            MsgBox("There are no Data Extensions available to modify for this Customer")
            Exit Sub
        End If

        If ModDataExt(strCustomer, VB.Left(strDataExtName, InStr(strDataExtName, "  ") - 1), (txtDataExtValue.Text)) Then
            GetUsedCustomerDataExts(VB6.GetItemString(lstCustomers, lstCustomers.SelectedIndex), lstUsedDataExts, True)
        End If
    End Sub

    Private Sub frmModDataExtension_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        GetCustomers(lstCustomers)
    End Sub

    Private Sub lstCustomers_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstCustomers.SelectedIndexChanged
        If frmModDataExtension.DefInstance.IsInitializing = True Then
            Exit Sub
        Else
            lstUsedDataExts.Items.Clear()
            lstUsedDataExts.Refresh()

            GetUsedCustomerDataExts(VB6.GetItemString(lstCustomers, lstCustomers.SelectedIndex), lstUsedDataExts, True)

            If lstUsedDataExts.Items.Count = 0 Then
                lstUsedDataExts.Items.Add("This customer has no data extensions added to their record")
            End If
        End If
    End Sub
End Class