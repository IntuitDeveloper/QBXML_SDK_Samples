<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class QBDataEventManagerDisplay
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Debug = New System.Windows.Forms.TextBox()
        Me.eventXML = New System.Windows.Forms.TextBox()
        Me.eventLabel = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Debug
        '
        Me.Debug.AcceptsReturn = True
        Me.Debug.BackColor = System.Drawing.SystemColors.Window
        Me.Debug.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Debug.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Debug.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Debug.Location = New System.Drawing.Point(132, 13)
        Me.Debug.MaxLength = 0
        Me.Debug.Name = "Debug"
        Me.Debug.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Debug.Size = New System.Drawing.Size(441, 20)
        Me.Debug.TabIndex = 5
        '
        'eventXML
        '
        Me.eventXML.AcceptsReturn = True
        Me.eventXML.BackColor = System.Drawing.SystemColors.Window
        Me.eventXML.CausesValidation = False
        Me.eventXML.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.eventXML.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.eventXML.ForeColor = System.Drawing.SystemColors.WindowText
        Me.eventXML.Location = New System.Drawing.Point(12, 37)
        Me.eventXML.MaxLength = 0
        Me.eventXML.Multiline = True
        Me.eventXML.Name = "eventXML"
        Me.eventXML.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.eventXML.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.eventXML.Size = New System.Drawing.Size(561, 441)
        Me.eventXML.TabIndex = 3
        Me.eventXML.Text = "eventXML appears here"
        Me.eventXML.WordWrap = False
        '
        'eventLabel
        '
        Me.eventLabel.BackColor = System.Drawing.SystemColors.Control
        Me.eventLabel.Cursor = System.Windows.Forms.Cursors.Default
        Me.eventLabel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.eventLabel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.eventLabel.Location = New System.Drawing.Point(12, 13)
        Me.eventLabel.Name = "eventLabel"
        Me.eventLabel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.eventLabel.Size = New System.Drawing.Size(114, 21)
        Me.eventLabel.TabIndex = 4
        Me.eventLabel.Text = "Received Event #"
        '
        'QBDataEventManagerDisplay
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(595, 492)
        Me.Controls.Add(Me.Debug)
        Me.Controls.Add(Me.eventXML)
        Me.Controls.Add(Me.eventLabel)
        Me.Name = "QBDataEventManagerDisplay"
        Me.Text = "QBDataEventManagerDisplay"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Public WithEvents Debug As Windows.Forms.TextBox
    Public WithEvents eventXML As Windows.Forms.TextBox
    Public WithEvents eventLabel As Windows.Forms.Label
End Class
