Option Strict Off
Option Explicit On
Friend Class frmDisplay
	Inherits System.Windows.Forms.Form
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' (c) 2003 Intuit Inc. All Rights Reserved                          '
	' Use is subject to IP Rights Notice and Restrictions available at: '
	' http://developer.quickbooks.com/legal/IPRNotice_021201.html       '
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click
		Me.Close()
	End Sub
	
	Private Sub frmDisplay_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim FileNum As Object
		Dim strInput As String
		Dim DisplayString As String

        FileNum = FreeFile()
        txtDisplay.Text = CStr(Nothing)
        FileOpen(FileNum, frmSDKTestPlus3.DisplayFile, OpenMode.Input)
        Do Until EOF(FileNum)
            strInput = LineInput(FileNum)
            txtDisplay.Text = txtDisplay.Text & strInput & vbCrLf
        Loop
        FileClose(FileNum)
        '  txtDisplay.Text = DisplayString
    End Sub
End Class