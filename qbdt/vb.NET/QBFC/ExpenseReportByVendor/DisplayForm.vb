Option Strict Off
Option Explicit On
Friend Class DisplayForm
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
	Public WithEvents closeButton As System.Windows.Forms.Button
	Public WithEvents displayText As System.Windows.Forms.TextBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.closeButton = New System.Windows.Forms.Button()
        Me.displayText = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'closeButton
        '
        Me.closeButton.BackColor = System.Drawing.SystemColors.Control
        Me.closeButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.closeButton.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.closeButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.closeButton.Location = New System.Drawing.Point(116, 272)
        Me.closeButton.Name = "closeButton"
        Me.closeButton.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.closeButton.Size = New System.Drawing.Size(157, 28)
        Me.closeButton.TabIndex = 1
        Me.closeButton.Text = "Close"
        Me.closeButton.UseVisualStyleBackColor = False
        '
        'displayText
        '
        Me.displayText.AcceptsReturn = True
        Me.displayText.BackColor = System.Drawing.SystemColors.Window
        Me.displayText.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.displayText.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.displayText.ForeColor = System.Drawing.SystemColors.WindowText
        Me.displayText.Location = New System.Drawing.Point(4, 4)
        Me.displayText.MaxLength = 0
        Me.displayText.Multiline = True
        Me.displayText.Name = "displayText"
        Me.displayText.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.displayText.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.displayText.Size = New System.Drawing.Size(381, 261)
        Me.displayText.TabIndex = 0
        '
        'DisplayForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(391, 304)
        Me.Controls.Add(Me.closeButton)
        Me.Controls.Add(Me.displayText)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(4, 23)
        Me.Name = "DisplayForm"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ResumeLayout(False)

    End Sub
#End Region
#Region "Upgrade Support "
    Private Shared m_vb6FormDefInstance As DisplayForm
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As DisplayForm
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New DisplayForm()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region
    ' DisplayForm.frm
    ' Created Sept, 2002
    '
    ' This form can be used to display the request and response
    ' qbxml to the user for illustrational purposes.
    '
    ' Copyright © 2002-2020 Intuit Inc. All rights reserved.
    ' Use is subject to the terms specified at:
    '      http://developer.intuit.com/legal/devsite_tos.html
    '
    '
    '


    Private xmlString As String
	
	
	'
	' Set the xml string to be displayed later.
	'
	Public Sub setXML(ByRef xmlToShow As String)
		xmlString = xmlToShow
	End Sub
	
	
	
	'
	' Display xml to the user.
	'
	Public Sub showXML(ByRef title As String)
		Me.Text = title
		displayText.Text = xmlString
		Me.Show()
	End Sub
    '
    ' Close the form.
    '
    Private Sub closeButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles closeButton.Click
        Me.Close()
    End Sub
End Class