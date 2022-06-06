Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class UIForm
	Inherits System.Windows.Forms.Form
	'
	' Copyright © 2021-2022 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	'
	' $Id: UIForm.frm,v 1.2 2002/08/05 13:00:33 giza Exp $
	'
	'
	'
	'
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		' Ensure that there's an input xml file.
		Dim InputFilename As String
        InputFilename = inputFile.Text
        If InputFilename.Length Then
            Log.showURL((InputFilename))
        Else
            MessageBox.Show("Please select an input file.")
        End If
    End Sub
	
	'
	' This method is called by the VB runtime when the application is
	' first started.
	'
	Private Sub UIForm_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Log.clear()
		Dim args As Collection
		args = GetCommandLine
		If (args.Count() > 0) Then
            inputFile.Text = args.Item(1)
        End If
	End Sub

    Private Sub Form_Terminate_Renamed()
        Log.clearXML()
    End Sub

    '
    ' This method is invoked when the user hits the Go button.
    '
    Private Sub goBut_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles goBut.Click
		
		Log.clear()
		
		' Ensure that there's an input xml file.
		On Error GoTo problem
		Dim InputFilename As String
		InputFilename = inputFile.Text
		Dim fnum As Short
		Dim isOpen As Boolean
        fnum = FreeFile()
        If InputFilename.Length Then
            FileOpen(fnum, InputFilename, OpenMode.Input)
            isOpen = True
        Else
            MessageBox.Show("Please select an input file.")
            Exit Sub
        End If



        ' Load the file into a string
        Dim xmldoc As String
		xmldoc = InputString(fnum, LOF(fnum))
		FileClose(fnum)
		
		' Send the request to QuickBooks and display the response.
		Dim response As String
		response = qbooks.post(xmldoc)
		If (response = "") Then
			Exit Sub
		End If
		Log.showXMLStream((response))
		Exit Sub
problem: 
		If isOpen Then FileClose(fnum)
		If Err.Number Then Err.Raise(Err.Number,  , Err.Description)
	End Sub
	'
	' Vanilla Windows UI code for file browsing.
	'
	Private Sub browseButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles browseButton.Click
		On Error GoTo ErrorHandler

        ' Set CancelError to True

        ' Set filters
        browseDialogOpen.Filter = "All Files (*.*)|*.*|qbXML Request Files" & "(*.xml)|*.xml"

        ' Specify default filter
        browseDialogOpen.FilterIndex = 2

        ' Set flags
        browseDialogOpen.ShowReadOnly = False

        browseDialogOpen.InitialDirectory = My.Application.Info.DirectoryPath & "\..\..\..\..\xmlfiles"
		
		
		' Display the Open dialog box
		browseDialogOpen.ShowDialog()
		
		' Set up the selected file in textBox
		inputFile.Text = browseDialogOpen.FileName
		Exit Sub
		
ErrorHandler: 
		Exit Sub
	End Sub
	
	Private Function GetCommandLine() As Collection
		'Declare variables.
		Dim I, CmdLnLen, C, CmdLine, InArg, NumArgs As Object
		Dim InQuote As Boolean
		Dim ArgArray As New Collection
		Dim Arg As String
		Arg = ""
        'Get command line arguments and make sure we have a space at end.
        CmdLine = VB.Command() & " "
        CmdLnLen = Len(CmdLine)
        'Go thru command line one character
        'at a time.
        For I = 1 To CmdLnLen
            C = Mid(CmdLine, I, 1)
            '  Test for a quote
            If (C = Chr(34)) Then
                If (InQuote) Then
                    InQuote = False
                Else
                    InQuote = True
                End If
            Else
                'Test for space or tab.
                If (InQuote Or (C <> " " And C <> vbTab)) Then
                    'Neither space nor tab or in a quoted string
                    'Test if already in argument.
                    If Not InArg Then
                        'New argument begins.
                        'Test for too many arguments.
                        InArg = True
                    End If
                    'Concatenate character to current argument.
                    Arg = Arg & C
                Else
                    'Found a space or tab.
                    'Set InArg flag to False.
                    If (InArg) Then ArgArray.Add(Arg)
                    InArg = False
                    Arg = ""
                End If
            End If
        Next I
        GetCommandLine = ArgArray
	End Function
End Class