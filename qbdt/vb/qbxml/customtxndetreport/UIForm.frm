VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form UIForm 
   Caption         =   "QBSDK 2.0 -- Custom Transaction Detail Report"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   ScaleHeight     =   7665
   ScaleWidth      =   10080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClearFilters 
      Caption         =   "Clear Filters"
      Height          =   375
      Left            =   7200
      TabIndex        =   29
      Top             =   4920
      Width           =   2175
   End
   Begin VB.ListBox RowsList 
      Height          =   2010
      ItemData        =   "UIForm.frx":0000
      Left            =   3600
      List            =   "UIForm.frx":005B
      TabIndex        =   28
      Top             =   5340
      Width           =   3015
   End
   Begin VB.CommandButton showResButton 
      Caption         =   "Show Response qbXML"
      Height          =   375
      Left            =   5160
      TabIndex        =   26
      Top             =   1620
      Width           =   2175
   End
   Begin VB.CommandButton showReqButton 
      Caption         =   "Show Request qbXML"
      Height          =   375
      Left            =   2700
      TabIndex        =   25
      Top             =   1620
      Width           =   2175
   End
   Begin VB.CommandButton exitButton 
      Caption         =   "&Exit Program"
      Height          =   375
      Left            =   7620
      TabIndex        =   12
      Top             =   1620
      Width           =   2175
   End
   Begin VB.ListBox ColumnsList 
      Height          =   2010
      ItemData        =   "UIForm.frx":0193
      Left            =   240
      List            =   "UIForm.frx":0251
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   5340
      Width           =   3135
   End
   Begin VB.TextBox toDateText 
      Height          =   315
      Left            =   7800
      TabIndex        =   9
      Top             =   6120
      Width           =   1455
   End
   Begin VB.TextBox fromDateText 
      Height          =   315
      Left            =   7800
      TabIndex        =   8
      Top             =   5580
      Width           =   1455
   End
   Begin VB.ListBox TxnTypeFilterList 
      Height          =   2010
      ItemData        =   "UIForm.frx":0536
      Left            =   7575
      List            =   "UIForm.frx":057F
      TabIndex        =   6
      Top             =   2535
      Width           =   2055
   End
   Begin VB.ListBox ItemTypeFilterList 
      Height          =   2010
      ItemData        =   "UIForm.frx":06BE
      Left            =   5160
      List            =   "UIForm.frx":06E0
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   2535
      Width           =   2055
   End
   Begin VB.ListBox EntityFilterList 
      Height          =   1815
      ItemData        =   "UIForm.frx":0757
      Left            =   2700
      List            =   "UIForm.frx":0767
      TabIndex        =   4
      Top             =   2535
      Width           =   2055
   End
   Begin VB.ListBox AccountFilterList 
      Height          =   2010
      ItemData        =   "UIForm.frx":0792
      Left            =   240
      List            =   "UIForm.frx":0802
      TabIndex        =   3
      Top             =   2535
      Width           =   2055
   End
   Begin VB.CommandButton goButton 
      Caption         =   "&Generate Report"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   1620
      Width           =   2175
   End
   Begin VB.TextBox qbFile 
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
   Begin VB.CommandButton browseButton 
      Caption         =   "&Browse"
      Height          =   375
      Left            =   6360
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox htmlFile 
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   1020
      Width           =   4215
   End
   Begin MSComDlg.CommonDialog browseDialog 
      Left            =   -120
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Summarize Data By (required)"
      Height          =   495
      Left            =   3600
      TabIndex        =   27
      Top             =   4920
      Width           =   3015
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "(blank to use currently open file)"
      Height          =   195
      Left            =   180
      TabIndex        =   24
      Top             =   600
      Width           =   2475
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "(yyyy-mm-dd)"
      Height          =   255
      Left            =   6540
      TabIndex        =   23
      Top             =   6300
      Width           =   1200
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "(yyyy-mm-dd)"
      Height          =   255
      Left            =   6540
      TabIndex        =   22
      Top             =   5760
      Width           =   1200
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "[hold Ctrl to select more than one.]"
      Height          =   255
      Left            =   300
      TabIndex        =   21
      Top             =   5160
      Width           =   3015
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Columns (required)"
      Height          =   495
      Left            =   240
      TabIndex        =   20
      Top             =   4920
      Width           =   3135
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "To Date:"
      Height          =   255
      Index           =   0
      Left            =   6540
      TabIndex        =   19
      Top             =   6120
      Width           =   1200
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "From Date:"
      Height          =   255
      Left            =   6540
      TabIndex        =   18
      Top             =   5580
      Width           =   1200
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Transaction Type Filter"
      Height          =   255
      Left            =   7560
      TabIndex        =   17
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Item Type Filter"
      Height          =   255
      Left            =   5160
      TabIndex        =   16
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Entity Type Filter"
      Height          =   255
      Left            =   2700
      TabIndex        =   15
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Account Type Filter"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Select QuickBooks Company File:"
      Height          =   315
      Left            =   180
      TabIndex        =   13
      Top             =   420
      Width           =   2475
   End
   Begin VB.Label Label2 
      Caption         =   "Select Output Html File:"
      Height          =   315
      Left            =   900
      TabIndex        =   10
      Top             =   1080
      Width           =   2475
   End
End
Attribute VB_Name = "UIForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' UIForm.frm
' Created July, 2002
'
' This file contains the code for the User Interface form of
' the project.  The UI allows the user to first select a
' QuickBooks company file and a location to store the output
' .html file.  When the GenerateReport button is pressed,
' it kicks off all of the main action.
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'
'


' Variables needed for displaying qbXML request and response
' in DisplayForm forms.
Private requestDisplay As New DisplayForm
Private responseDisplay As New DisplayForm
Private displayInit As Boolean



Private Sub cmdClearFilters_Click()
  AccountFilterList.ListIndex = -1
  AccountFilterList.Refresh
  
  EntityFilterList.ListIndex = -1
  EntityFilterList.Refresh

  ItemTypeFilterList.ListIndex = -1
  ItemTypeFilterList.Refresh

  TxnTypeFilterList.ListIndex = -1
  TxnTypeFilterList.Refresh

  RowsList.ListIndex = -1
  RowsList.Refresh
End Sub

'
' Set the default output path when the form is first loaded.
'
Private Sub Form_Load()
    
    ' Default location for html file
    htmlFile.Text = App.Path & "\output.html"
    displayInit = False
    
End Sub



'
' Browse for a company file using the MS Common Dialog
' Control
'
Private Sub browseButton_Click()
  On Error GoTo ErrorHandler

  ' Set CancelError to True
  browseDialog.CancelError = True
  
  ' Set flags
  browseDialog.Flags = cdlOFNHideReadOnly
  
  ' Set filters
  browseDialog.Filter = "All Files (*.*)|*.*|QB Company Files" & _
    "(*.qbw)|*.qbw"
  
  ' Specify default filter
  browseDialog.FilterIndex = 2
  
  ' Display the Open dialog box
  browseDialog.ShowOpen
  
  ' set up the selected file in Text_Schema textBox
  qbFile.Text = browseDialog.fileName
  Exit Sub
  
ErrorHandler:
  ' error occurs, do nothing
  Exit Sub

End Sub






'
' Generate and display the report when the go button is
' clicked if a company file has already been selected.
'
Private Sub goButton_Click()
   ' On Error GoTo ErrorHandler

    ' Get company file name, use open file if blank
    Dim companyFile As String
    companyFile = qbFile.Text

    ' Make sure we have an output file path
    Dim url As String
    
    If ("" = htmlFile.Text) Then
        MsgBox "Please fill in a path for the html output file."
        Exit Sub
    Else
        url = htmlFile.Text
    End If
    
    ' Make sure at least one Column is selected
    If (ColumnsList.SelCount = 0) Then
        MsgBox "You must select at least one column to display."
        Exit Sub
    End If

    ' Make sure Summarize Rows By is selected
    If (RowsList.ListIndex < 0) Then
        MsgBox "You must select summary data."
        Exit Sub
    End If

    ' Make sure dates are in a format that can be understood
    ' by QuickBooks, i.e. yyyy-mm-dd
    If Not "" = fromDateText.Text Then
        If (Len(fromDateText.Text) <> 10) Or _
            Not Mid(fromDateText.Text, 5, 1) = "-" Or _
            Not Mid(fromDateText.Text, 8, 1) = "-" Then
            MsgBox "Dates must be in the format yyyy-mm-dd"
            Exit Sub
        End If
    End If
    
    If Not "" = toDateText.Text Then
        If (Len(toDateText.Text) <> 10) Or _
            Not Mid(toDateText.Text, 5, 1) = "-" Or _
            Not Mid(toDateText.Text, 8, 1) = "-" Then
            MsgBox "Dates must be in the format yyyy-mm-dd"
            Exit Sub
        End If
    End If

    UIForm.MousePointer = vbHourglass

    Dim NumColumns As Integer
    NumColumns = ColumnsList.SelCount
    
    Dim ColumnsArray() As String
    ReDim ColumnsArray(NumColumns) As String
    
    ' Loop through the columns list and add all of the
    ' selected columns to the array
    Dim counter As Integer
    Dim arrayIndex As Integer
    arrayIndex = 1
    For counter = 0 To ColumnsList.ListCount - 1
        If ColumnsList.Selected(counter) Then
            ColumnsArray(arrayIndex) = ColumnsList.List(counter)
            arrayIndex = arrayIndex + 1
        End If
    Next

    Dim AccountFilter As String
    Dim EntityFilter As String
    Dim ItemFilter As String
    Dim TxnFilter As String
    Dim FromDate As String
    Dim ToDate As String
    Dim SummarizeRowsBy As String
    
    If (AccountFilterList.SelCount = 1) Then
        AccountFilter = AccountFilterList.Text
    Else
        AccountFilter = ""
    End If
    
    If (EntityFilterList.SelCount = 1) Then
        EntityFilter = EntityFilterList.Text
    Else
        EntityFilter = ""
    End If
    
    If (ItemTypeFilterList.SelCount = 1) Then
        ItemFilter = ItemTypeFilterList.Text
    Else
        ItemFilter = ""
    End If
    
    If (TxnTypeFilterList.SelCount = 1) Then
        TxnFilter = TxnTypeFilterList.Text
    Else
        TxnFilter = ""
    End If
        
    FromDate = fromDateText.Text
    ToDate = toDateText.Text
    SummarizeRowsBy = RowsList.List(RowsList.ListIndex)
    
    Dim xmlRequest As String
    xmlRequest = generateXMLRequest(AccountFilter, EntityFilter, _
        ItemFilter, TxnFilter, FromDate, ToDate, ColumnsArray, SummarizeRowsBy)

    ' Exit sub if the xml was not generated correctly
    If (ErrorCode = xmlRequest) Then
        GoTo ErrorHandler2
    End If

    Dim xmlResponse As String
    xmlResponse = sendReqToQB(xmlRequest, companyFile)
    
    ' Exit sub if xml error from QuickBooks was already
    ' reported in sendReqToQB
    If (ErrorCode = xmlResponse) Then
        GoTo ErrorHandler2
    End If

    ' Set the two DisplayForm forms so that the user will be
    ' able to view the request and response by clicking on the
    ' Show Request and Show Response buttons.
    requestDisplay.setXML PrettyXMLString(xmlRequest)
    responseDisplay.setXML PrettyXMLString(xmlResponse)
    displayInit = True
    
    Dim newStatus As String
    newStatus = parseXMLResponse(xmlResponse, url)
    
    ' Exit sub if the xml was not parsed correctly
    If (ErrorCode = newStatus) Then
        GoTo ErrorHandler2
    End If
    
    BigDisplayForm.displayResults url
    
    UIForm.MousePointer = vbDefault
    Exit Sub

ErrorHandler:
    MsgBox Err.Description
ErrorHandler2:
    UIForm.MousePointer = vbDefault
    Exit Sub
End Sub



'
' Display XML Request if one has been generated already.
'
Private Sub showReqButton_Click()
    On Error GoTo showReqErrorHandler
    
    If displayInit Then
        requestDisplay.showXML "qbXML Request:"
    Else
        MsgBox "Please click Generate Report before trying to " _
            & "view the qbXML request."
    End If

    Exit Sub

showReqErrorHandler:
    MsgBox Err.Description
    Exit Sub
End Sub



'
' Display XML Response if one has been generated already.
'
Private Sub showResButton_Click()
    On Error GoTo showResErrorHandler
    
    If displayInit Then
        responseDisplay.showXML "qbXML Response:"
    Else
        MsgBox "Please click Generate Report before trying to " _
            & "view the qbXML response."
    End If
    
    Exit Sub
    
showResErrorHandler:
    MsgBox Err.Description
    Exit Sub
End Sub



'
' Quit the program.
'
Private Sub exitButton_Click()
    Unload Me
End Sub

'----------------------------------------------------------
'
' Routine: PrettyPrintXMLToFile
'
' Description
'
' Takes an XML message set string and a file name as input.
' Creates a new copy of the file and pretty prints the XML
' message set into the file.
'
'----------------------------------------------------------
Public Sub PrettyPrintXMLToFile(XMLString As String, _
                                 XMLFile As String)
                                       
  Dim SplitXMLString() As String
  Dim IndentString As String
  Dim XMLStringLength As Long
  Dim SplitIndex
  
  IndentString = Empty
  
  FileNum = FreeFile
  Open XMLFile For Output As FileNum
  
'Remove the linefeeds from the XML output string
  XMLString = Replace(XMLString, vbLf, vbNullString)
  
  SplitXMLString = Split(XMLString, "<")
  
'We're expecting the first character of the XML output to be "<"
'which result in an empty first array element, so skip it.
  SplitIndex = LBound(SplitXMLString) + 1
  
  Do
    If Left(SplitXMLString(SplitIndex), 1) = "/" Then
      IndentString = Left(IndentString, Len(IndentString) - 3)
      Print #FileNum, IndentString & "<" & _
                      SplitXMLString(SplitIndex)
      SplitIndex = SplitIndex + 1
    ElseIf Left(SplitXMLString(SplitIndex + 1), 1) = "/" Then
      If InStr(1, Left(SplitXMLString(SplitIndex), _
               InStr(1, SplitXMLString(SplitIndex), ">")), _
                " ") > 0 Then
        Print #FileNum, IndentString & "<" & _
                        SplitXMLString(SplitIndex)
        SplitIndex = SplitIndex + 1
      Else
        Print #FileNum, IndentString & "<" & _
                        SplitXMLString(SplitIndex) & "<" & _
                        SplitXMLString(SplitIndex + 1)
        SplitIndex = SplitIndex + 2
      End If
    Else
      Print #FileNum, IndentString & "<" & _
                      SplitXMLString(SplitIndex)
      If SplitXMLString(SplitIndex) <> "?xml version=""1.0"" ?>" And _
         SplitXMLString(SplitIndex) <> "!DOCTYPE QBXML PUBLIC '-//INTUIT//DTD QBXML QBD 1.0//EN' >" And _
         Right(SplitXMLString(SplitIndex), 2) <> "/>" Then
        IndentString = IndentString & "   "
      End If
      SplitIndex = SplitIndex + 1
    End If
  Loop Until SplitIndex >= UBound(SplitXMLString)
  
  If Left(SplitXMLString(UBound(SplitXMLString)), 1) = "/" Then
    IndentString = Left(IndentString, Len(IndentString) - 3)
  End If
  
  Print #FileNum, IndentString & "<" & _
                  SplitXMLString(UBound(SplitXMLString))
  
  Close FileNum
End Sub

