Option Strict Off
Option Explicit On
Friend Class frmDataExtSample
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
	Public WithEvents txtUtility As System.Windows.Forms.TextBox
	Public WithEvents txtDataExts As System.Windows.Forms.TextBox
	Public WithEvents cmdExit As System.Windows.Forms.Button
	Public WithEvents cmdModDataExt As System.Windows.Forms.Button
	Public WithEvents cmdAddDataExtension As System.Windows.Forms.Button
	Public WithEvents cmdDefineDataExt As System.Windows.Forms.Button
	Public WithEvents txtCustomFields As System.Windows.Forms.TextBox
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Text1 = New System.Windows.Forms.TextBox()
        Me.txtUtility = New System.Windows.Forms.TextBox()
        Me.txtDataExts = New System.Windows.Forms.TextBox()
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.cmdModDataExt = New System.Windows.Forms.Button()
        Me.cmdAddDataExtension = New System.Windows.Forms.Button()
        Me.cmdDefineDataExt = New System.Windows.Forms.Button()
        Me.txtCustomFields = New System.Windows.Forms.TextBox()
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
        Me.Text1.Location = New System.Drawing.Point(8, 32)
        Me.Text1.MaxLength = 0
        Me.Text1.Multiline = True
        Me.Text1.Name = "Text1"
        Me.Text1.ReadOnly = True
        Me.Text1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text1.Size = New System.Drawing.Size(385, 49)
        Me.Text1.TabIndex = 9
        Me.Text1.Text = "Select the buttons below to add data extension definitions to a company file, add" & _
        " data extension values to a customer record or modify the value of a data extens" & _
        "ion already added to a customer" & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10)
        '
        'txtUtility
        '
        Me.txtUtility.AcceptsReturn = True
        Me.txtUtility.AutoSize = False
        Me.txtUtility.BackColor = System.Drawing.SystemColors.Window
        Me.txtUtility.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtUtility.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtUtility.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtUtility.Location = New System.Drawing.Point(400, 264)
        Me.txtUtility.MaxLength = 0
        Me.txtUtility.Multiline = True
        Me.txtUtility.Name = "txtUtility"
        Me.txtUtility.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtUtility.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtUtility.Size = New System.Drawing.Size(10, 13)
        Me.txtUtility.TabIndex = 8
        Me.txtUtility.TabStop = False
        Me.txtUtility.Text = "T" & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10) & "e" & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10) & "x" & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10) & "t" & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10) & "1" & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10)
        Me.txtUtility.Visible = False
        Me.txtUtility.WordWrap = False
        '
        'txtDataExts
        '
        Me.txtDataExts.AcceptsReturn = True
        Me.txtDataExts.AutoSize = False
        Me.txtDataExts.BackColor = System.Drawing.SystemColors.Window
        Me.txtDataExts.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDataExts.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDataExts.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDataExts.Location = New System.Drawing.Point(8, 88)
        Me.txtDataExts.MaxLength = 0
        Me.txtDataExts.Multiline = True
        Me.txtDataExts.Name = "txtDataExts"
        Me.txtDataExts.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDataExts.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtDataExts.Size = New System.Drawing.Size(393, 161)
        Me.txtDataExts.TabIndex = 7
        Me.txtDataExts.Text = ""
        Me.txtDataExts.WordWrap = False
        '
        'cmdExit
        '
        Me.cmdExit.BackColor = System.Drawing.SystemColors.Control
        Me.cmdExit.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdExit.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdExit.Location = New System.Drawing.Point(328, 440)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdExit.Size = New System.Drawing.Size(73, 49)
        Me.cmdExit.TabIndex = 6
        Me.cmdExit.Text = "Exit"
        '
        'cmdModDataExt
        '
        Me.cmdModDataExt.BackColor = System.Drawing.SystemColors.Control
        Me.cmdModDataExt.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdModDataExt.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdModDataExt.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdModDataExt.Location = New System.Drawing.Point(216, 440)
        Me.cmdModDataExt.Name = "cmdModDataExt"
        Me.cmdModDataExt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdModDataExt.Size = New System.Drawing.Size(97, 49)
        Me.cmdModDataExt.TabIndex = 5
        Me.cmdModDataExt.Text = "Modify Customer Data Extension"
        '
        'cmdAddDataExtension
        '
        Me.cmdAddDataExtension.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAddDataExtension.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAddDataExtension.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAddDataExtension.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAddDataExtension.Location = New System.Drawing.Point(104, 440)
        Me.cmdAddDataExtension.Name = "cmdAddDataExtension"
        Me.cmdAddDataExtension.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAddDataExtension.Size = New System.Drawing.Size(97, 49)
        Me.cmdAddDataExtension.TabIndex = 4
        Me.cmdAddDataExtension.Text = "Add Customer Data Extension"
        '
        'cmdDefineDataExt
        '
        Me.cmdDefineDataExt.BackColor = System.Drawing.SystemColors.Control
        Me.cmdDefineDataExt.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdDefineDataExt.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDefineDataExt.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDefineDataExt.Location = New System.Drawing.Point(8, 440)
        Me.cmdDefineDataExt.Name = "cmdDefineDataExt"
        Me.cmdDefineDataExt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdDefineDataExt.Size = New System.Drawing.Size(81, 49)
        Me.cmdDefineDataExt.TabIndex = 3
        Me.cmdDefineDataExt.Text = "Define Data Extension"
        '
        'txtCustomFields
        '
        Me.txtCustomFields.AcceptsReturn = True
        Me.txtCustomFields.AutoSize = False
        Me.txtCustomFields.BackColor = System.Drawing.SystemColors.Window
        Me.txtCustomFields.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCustomFields.Enabled = False
        Me.txtCustomFields.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustomFields.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCustomFields.Location = New System.Drawing.Point(8, 288)
        Me.txtCustomFields.MaxLength = 0
        Me.txtCustomFields.Multiline = True
        Me.txtCustomFields.Name = "txtCustomFields"
        Me.txtCustomFields.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCustomFields.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtCustomFields.Size = New System.Drawing.Size(393, 145)
        Me.txtCustomFields.TabIndex = 2
        Me.txtCustomFields.Text = ""
        Me.txtCustomFields.WordWrap = False
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(8, 264)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(385, 17)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Defined Custom Fields (read only, to show query for Custom Field Defs)"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(400, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Data Extensions for OwnerID {E09C86CF-9D6E-4EF2-BCBE-4D66B6B0F754}"
        '
        'frmDataExtSample
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(424, 497)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Text1, Me.txtUtility, Me.txtDataExts, Me.cmdExit, Me.cmdModDataExt, Me.cmdAddDataExtension, Me.cmdDefineDataExt, Me.txtCustomFields, Me.Label2, Me.Label1})
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(276, 164)
        Me.Name = "frmDataExtSample"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Data Extension Sample Main Window"
        Me.ResumeLayout(False)

    End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As frmDataExtSample
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As frmDataExtSample
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New frmDataExtSample()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	'----------------------------------------------------------
	' Form: frmDataExtSample
	'
	' Description: This the main form and entry point for this sample
	'              program.  It displays the currently defined data
	'              extension definitions and custom fields for the open
	'              company file.  From here the user may choose to
	'              activate forms to add data extension definitions to the
	'              company file, define values for data extension
	'              for customers or modify data extension values for
	'              customers
	'
	'              The form calls OpenSessionBeginSession to make sure
	'              a company file is open.
	'
	' Copyright © 2002-2013 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	'----------------------------------------------------------
	Private Sub cmdAddDataExtension_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAddDataExtension.Click
		If CustomersHaveDataExts Then
            frmAddDataExtension.DefInstance.Show()
        Else
            MsgBox("You must add a data extension for Customers before adding a data extension value to a specified customer")
        End If
	End Sub
	
	Private Sub cmdDefineDataExt_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDefineDataExt.Click
        frmAddDataExtDef.DefInstance.Show()
    End Sub

    Private Sub cmdExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdExit.Click
        EndSessionCloseConnection()
        End
    End Sub

    Private Sub cmdModDataExt_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdModDataExt.Click
        If CustomersHaveDataExts() Then
            frmModDataExtension.DefInstance.Show()
        Else
            MsgBox("You must add a data extension for Customers before modifying a data extension value for a specific customer")
        End If
    End Sub

    Private Sub frmDataExtSample_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        OpenConnectionBeginSession()
        GetDataExts(txtDataExts)
        GetCustomFields(txtCustomFields)
    End Sub
End Class