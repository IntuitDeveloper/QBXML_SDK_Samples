Option Strict Off
Option Explicit On
Friend Class frmDepositAdd
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
	Public WithEvents cmdExit As System.Windows.Forms.Button
	Public WithEvents cmdDepositFunds As System.Windows.Forms.Button
	Public WithEvents lstFundsForDeposit As System.Windows.Forms.ListBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmDepositAdd))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me.cmdExit = New System.Windows.Forms.Button
		Me.cmdDepositFunds = New System.Windows.Forms.Button
		Me.lstFundsForDeposit = New System.Windows.Forms.ListBox
		Me.Text = "Deposit Add"
		Me.ClientSize = New System.Drawing.Size(501, 431)
		Me.Location = New System.Drawing.Point(4, 23)
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MaximizeBox = True
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmDepositAdd"
		Me.cmdExit.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdExit.Text = "Exit"
		Me.cmdExit.Size = New System.Drawing.Size(153, 57)
		Me.cmdExit.Location = New System.Drawing.Point(280, 360)
		Me.cmdExit.TabIndex = 2
		Me.cmdExit.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdExit.BackColor = System.Drawing.SystemColors.Control
		Me.cmdExit.CausesValidation = True
		Me.cmdExit.Enabled = True
		Me.cmdExit.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdExit.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdExit.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdExit.TabStop = True
		Me.cmdExit.Name = "cmdExit"
		Me.cmdDepositFunds.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdDepositFunds.Text = "Deposit Selected Funds"
		Me.cmdDepositFunds.Size = New System.Drawing.Size(153, 57)
		Me.cmdDepositFunds.Location = New System.Drawing.Point(64, 360)
		Me.cmdDepositFunds.TabIndex = 1
		Me.cmdDepositFunds.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdDepositFunds.BackColor = System.Drawing.SystemColors.Control
		Me.cmdDepositFunds.CausesValidation = True
		Me.cmdDepositFunds.Enabled = True
		Me.cmdDepositFunds.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdDepositFunds.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdDepositFunds.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdDepositFunds.TabStop = True
		Me.cmdDepositFunds.Name = "cmdDepositFunds"
		Me.lstFundsForDeposit.Size = New System.Drawing.Size(465, 332)
		Me.lstFundsForDeposit.Location = New System.Drawing.Point(16, 16)
		Me.lstFundsForDeposit.TabIndex = 0
		Me.lstFundsForDeposit.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lstFundsForDeposit.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lstFundsForDeposit.BackColor = System.Drawing.SystemColors.Window
		Me.lstFundsForDeposit.CausesValidation = True
		Me.lstFundsForDeposit.Enabled = True
		Me.lstFundsForDeposit.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lstFundsForDeposit.IntegralHeight = True
		Me.lstFundsForDeposit.Cursor = System.Windows.Forms.Cursors.Default
		Me.lstFundsForDeposit.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.lstFundsForDeposit.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lstFundsForDeposit.Sorted = False
		Me.lstFundsForDeposit.TabStop = True
		Me.lstFundsForDeposit.Visible = True
		Me.lstFundsForDeposit.MultiColumn = False
		Me.lstFundsForDeposit.Name = "lstFundsForDeposit"
		Me.Controls.Add(cmdExit)
		Me.Controls.Add(cmdDepositFunds)
		Me.Controls.Add(lstFundsForDeposit)
	End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As frmDepositAdd
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As frmDepositAdd
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New frmDepositAdd()
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
    '-----------------------------------------------------------
    ' Form Module: frmDepositAdd
    '
    ' Description: this form is to display payments available for deposit
    '              and allow the user to deposit them.
    '
    ' Copyright © 2002-2013 Intuit Inc. All rights reserved.
    ' Use is subject to the terms specified at:
    '      http://developer.intuit.com/legal/devsite_tos.html
    '
    '----------------------------------------------------------

    Dim booFundsSelected As Boolean

    Private Sub cmdDepositFunds_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDepositFunds.Click
        If Not booFundsSelected Then
            MsgBox("You must select funds to deposit before attempting to deposit them.")
            Exit Sub
        End If

        DepositFunds(VB6.GetItemString(lstFundsForDeposit, lstFundsForDeposit.SelectedIndex))
        GetFundsForDeposit(lstFundsForDeposit)
    End Sub

    Private Sub cmdExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdExit.Click
        EndSessionCloseConnection()
        End
    End Sub

    Private Sub frmDepositAdd_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        booFundsSelected = False
        Connect()
        GetFundsForDeposit(lstFundsForDeposit)
    End Sub

    Private Sub lstFundsForDeposit_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstFundsForDeposit.SelectedIndexChanged
        If frmDepositAdd.DefInstance.IsInitializing = True Then
            Exit Sub
        Else
            booFundsSelected = True
        End If
    End Sub
End Class