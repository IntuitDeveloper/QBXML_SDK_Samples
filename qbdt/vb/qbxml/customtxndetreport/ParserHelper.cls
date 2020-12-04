VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParserHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ParserHelper.cls
' Created July, 2002
'
' This class module is used to parse the response from QuickBooks.
' It creates the .html table from the parsed response.  As the
' CustomTxnDetail report is fairly general, this parser is general
' as well.  With a few modications, it could be used to easily create
' other variaties of reports.
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'


Option Explicit

' Column Description Type
Private Type ColDescType
    colID As Integer
    colTitle As String
    colType As String
End Type

' ReportRet Type
Private Type ReportRetType
    title As String
    reportBasis As String
    numRows As Integer
    numCols As Integer
    listColDesc() As ColDescType
    colDescNum As Integer
End Type


' Local Variables
Private output As String ' html output

Private currNode  As IXMLDOMNode ' current node being parsed
Private rowCount As Integer ' number of rows in report
Private currReportRet As ReportRetType ' current ReportRet

Private errorStatus  As Boolean ' true if there is an error
Private currStatusMsg As String

Private Const NUMOFCOLS = 1000



'
' Get output html String once it has been created.
'
Public Property Get OutputStr() As String
    OutputStr = output
End Property



'
' Start the xml parsing.  After this function has returned, output
' will contain the html body of the report we are generating.
'
Public Function ParseReportResponseXML(xmlStr As String) As Boolean
  
    Dim ret  As Boolean
 
    ' Put initial tags on html string
    output = "<html>" & vbCrLf
    output = output & "<body bgcolor=""#ffffff"">" & vbCrLf
    output = output & "<table cellpadding=""2"" cellspacing=""2"" border=""1"">" _
        & vbCrLf & vbCrLf
  
    rowCount = 0 ' initialize the number of rows in the report
    errorStatus = False ' initialize the error status
    
    ret = ParseResponseXML(xmlStr)
  
    ' ret will be true if ParseResponseXML has succeeded with no error
    If (Not ret) Or (errorStatus) Then
        ParseReportResponseXML = False
        Exit Function
    End If
  
    ' Make sure we have displayed the correct number of rows
    ' in the report.
    If rowCount <> currReportRet.numRows Then
        MsgBox "Error: Actual num of rows: " & CStr(rowCount) & _
            " doesn't match NumOfRows " & CStr(currReportRet.numRows)
    End If
  
    ' Put final tags on html string
    output = output & "</table>" & vbCrLf
    output = output & "</body>" & vbCrLf
    output = output & "</html>" & vbCrLf
  
    ParseReportResponseXML = True
  
End Function




'
' Create the MSXML DOMDocument object using the MSXML version 4
' library.  Call AccessNode to start parsing the object.
'
Private Function ParseResponseXML(xmlStr As String) As Boolean
    On Error GoTo LoadMSXML40ErrorHandler

    Dim ret As Boolean
  
    If xmlStr = Null Or xmlStr = "" Then
        ParseResponseXML = False
        Exit Function
    End If
  
    ' Create MSXML document object
    Dim xmlDoc As New MSXML2.DOMDocument40
  
    ' Load XML document from xmlStr
    xmlDoc.async = False
    ret = xmlDoc.loadXML(xmlStr)
      
    ' Set currNode to the root node of the document
    Set currNode = xmlDoc.documentElement
  
    ' Start parsing at the root level
    AccessNode
    
    ' Clean up
    Set xmlDoc = Nothing
    
    If errorStatus Then
        ' There has been a major error somewhere in the module.
        ParseResponseXML = False
    Else
        ' The parsing was successful
        ParseResponseXML = True
    End If
    
    Exit Function
  
LoadMSXML40ErrorHandler:
    MsgBox "Error:" + CStr(Hex(Err.Number)) + ", " + Err.Description, vbCritical, "Parse XML error: "
    errorStatus = True
    ParseResponseXML = False
    Exit Function

MSXML40ParseErrorHandler:
    MsgBox "Error:" + Err.Description, vbCritical, "Error in Parsing Response"
    errorStatus = True
    ParseResponseXML = False
    Exit Function

End Function




'
' Loop through nodes parsing out data and creating the table
' for display.
'
Private Sub AccessNode()
    On Error GoTo AccErrorhandler
  
    If errorStatus Then
        Exit Sub
    End If
  
    Dim currTagName As String ' curr Tag Name
    currTagName = currNode.nodeName
  
    ' If the node we are looking at is the CustomDetailReportQueryRs
    ' node, we need to check the error status.
    If currTagName = "CustomDetailReportQueryRs" Then
        If CheckRsStatus(currTagName) Then
            MsgBox currStatusMsg, vbCritical, "Error from QuickBooks."
            errorStatus = True
            Exit Sub
        End If
    End If
  
    ' If the node we are looking at is a ReportRet node, we need
    ' to initialize the current ReportRet variable (currReportRet).
    If currTagName = "ReportRet" Then
        InitCurrReportRet
    End If
  
    ' Loop through the children of the current node, if they exist.
    If currNode.hasChildNodes Then
        Dim length As Integer
        Dim it As Integer
      
        Dim tempNode As IXMLDOMNode
        Set tempNode = currNode
        length = tempNode.childNodes.length
    
        If length = 1 And _
            tempNode.childNodes.Item(0).nodeType = NODE_TEXT Then
      
            ' Get Report level info
            Dim valStr As String
            valStr = tempNode.childNodes.Item(0).Text
      
            Select Case currTagName
                Case "ReportTitle"
                    currReportRet.title = valStr
                Case "ReportBasis"
                    currReportRet.reportBasis = valStr
                Case "NumRows"
                    currReportRet.numRows = CInt(valStr)
                Case "NumColumns"
                    currReportRet.numCols = CInt(valStr)
            End Select
        
        Else ' Aggregate Level
    
            ' Loop through sub aggregate/elements
            For it = 0 To length - 1
                Set currNode = tempNode.childNodes.Item(it)
                currTagName = currNode.nodeName
        
                Select Case currTagName
                    Case "ColDesc"
                        ProcessColDesc
                    Case "ReportData"
                        PrintColDesc ' end of Head session, to print the header
                        ProcessReportData
                    Case Else
                        AccessNode ' go to next level
                End Select
            Next
        End If
    End If
    Exit Sub
  
AccErrorhandler:
    MsgBox Err.Description, vbCritical, "Error in Parsing Response"
    errorStatus = True
End Sub



'
' Parse the ColDesc elements.  This data will be used later when
' we are printing out the column headers.
'
Private Sub ProcessColDesc()
    On Error GoTo ProcErrorHandler:
  
    If currNode.hasChildNodes Then
    
        Dim length As Integer
        Dim it As Integer
        Dim m As Integer
        Dim currTagName As String
    
        ' Add to the reportRet.listColDesc list
        currReportRet.listColDesc(currReportRet.colDescNum).colID = GetColID
    
        Dim tempNode As IXMLDOMNode
        Set tempNode = currNode
     
        length = currNode.childNodes.length
        For it = 0 To length - 1
            Set currNode = tempNode.childNodes.Item(it)
            currTagName = currNode.nodeName
      
            Dim valStr As String
      
            'Set up each Column's ColType and ColTitle for use later
            If currTagName <> "#comment" Then
                Select Case currTagName
                    Case "ColTitle"
                        ' Get the Column Title from the attribute titleValue which
                        ' appears second after titleRow if it exists
                        If (currNode.Attributes.length = 2) Then
                            valStr = currNode.Attributes.Item(1).Text
                        Else
                            valStr = ""
                        End If
                        currReportRet.listColDesc(currReportRet.colDescNum).colTitle = valStr
                    Case "ColType"
                        valStr = currNode.childNodes.Item(0).Text
                        currReportRet.listColDesc(currReportRet.colDescNum).colType = valStr
                End Select
            End If
        Next
        currReportRet.colDescNum = currReportRet.colDescNum + 1
    End If
  
    Exit Sub
  
ProcErrorHandler:
    MsgBox Err.Description, vbCritical, "Error in ProcessColDesc"
    errorStatus = True
End Sub




'
' Parse everything contained in the ReportData element and store
' the information for later use.
'
Private Sub ProcessReportData()
    Dim currTagName As String
    Dim tempNode As IXMLDOMNode
    Set tempNode = currNode
    Dim length As Integer
    Dim it As Integer
        
    On Error GoTo ProcErrorHandler
    
    length = tempNode.childNodes.length
    For it = 0 To length - 1
        Set currNode = tempNode.childNodes.Item(it)
        
        currTagName = currNode.nodeName
        Dim valStr As String
        
        Select Case currTagName
          Case "DataRow"
            rowCount = rowCount + 1
            ProcessRow currTagName
          Case "TextRow"
            rowCount = rowCount + 1
          Case "SubtotalRow"
            rowCount = rowCount + 1
            ProcessRow currTagName
          Case "TotalRow"
            ' Do not display totals in this report
            rowCount = rowCount + 1
        End Select
            
    Next
    Exit Sub

ProcErrorHandler:
    MsgBox Err.Description, vbCritical, "Error in Process ReportData"
    errorStatus = True
End Sub



'
' Print a row of column headers, which have already been parsed
' from the response XML and stored, to the html report string.
'
Private Sub PrintColDesc()
    
    Dim it As Integer
    Dim colPos As Integer
    Dim startPos As Integer
    colPos = 0
    startPos = 0
    
    Dim colDesc As ColDescType
    
    output = output & "<tr>" & vbCrLf
    
    For it = 1 To currReportRet.colDescNum
        
        colDesc = GetColDescType(it - startPos, startPos)
        output = output & "<td align=""center""><b>" & colDesc.colTitle & "</b></td>" _
            & vbCrLf
      
        colPos = colPos + 1
      
        If colDesc.colID = currReportRet.numCols Then ' next desc line
            output = output & vbCrLf
            colPos = 0
            startPos = startPos + currReportRet.numCols
        End If
      
    Next
    
    output = output & "</tr>" & vbCrLf & vbCrLf

End Sub






'
' Parse the RowData and ColumnData elements and add the
' results to our output display table html string.
'
Private Sub ProcessRow(sTagName As String)
    On Error GoTo ProcErrorHandler
  
    Dim listCol() As String
    If currNode.hasChildNodes Then
    
    Dim length As Integer
    Dim it As Integer
    Dim m As Integer
    Dim rowType As String
    Dim rowValue As String
    Dim rowNum As String
    
    Dim currTagName As String
    
    ' Get rowNum from attribute list
    Dim attrib
    For m = 0 To currNode.Attributes.length - 1
        Set attrib = currNode.Attributes.Item(m)
        If (attrib.Name = "rowNumber") Then
                rowNum = attrib.nodeValue
        End If
    Next
  
    ' Validate rowNum
    If Val(rowNum) > currReportRet.numRows Or _
        Val(rowNum) < 0 Then
        MsgBox "Invalid rowNum: " & rowNum
        errorStatus = True
    End If
    
    Dim tempNode As IXMLDOMNode
    Set tempNode = currNode
    length = currNode.childNodes.length
    ReDim listCol(currReportRet.numCols + 1)
    
    For it = 0 To currReportRet.numCols
        listCol(it) = ""
    Next
    
    For it = 0 To length - 1
        Set currNode = tempNode.childNodes.Item(it)
        currTagName = currNode.nodeName
      
        Select Case currTagName
            Case "RowData"
                GetRowData rowType, rowValue

            Case "ColData" ' it could be a list
                Dim colID As Integer
                Dim colValue As String
          
                If colID > currReportRet.numCols Then
                    MsgBox "current colID value " & CStr(colID) & _
                    " is large than NumOfCol " & CStr(currReportRet.numCols), vbCritical, "Wrong Col ID"
                    errorStatus = True
                    Exit Sub
                End If
                GetColData colID, colValue
                listCol(colID) = colValue
        End Select
    Next
    
    ' Add the row to the html string
    output = output & "<tr>" & vbCrLf
    
    ' Loop through each column on this row
    For it = 1 To currReportRet.numCols
        If listCol(it) <> "" Then
            Dim formattedData As String
            formattedData = FormatData(listCol(it), GetColDescType(it, 0).colType)
            output = output & "<td align=""left"">" & _
                formattedData & "</td>" & vbCrLf
        Else
            output = output & "<td>&nbsp;</td>" & vbCrLf
        End If
    Next
    
    output = output & "</tr>" & vbCrLf & vbCrLf
    
    ReDim listCol(0)
    End If
    Exit Sub
  
ProcErrorHandler:
    MsgBox Err.Description, vbCritical, "Error in ProcessRow"
    errorStatus = True
End Sub




'
' Initialize the currReportRet variable when we get to a new ReportRet.
'
Private Sub InitCurrReportRet()
    ReDim currReportRet.listColDesc(NUMOFCOLS)
    currReportRet.numCols = 0
    currReportRet.numRows = 0
    currReportRet.reportBasis = ""
    currReportRet.title = ""
    currReportRet.colDescNum = 0
End Sub




'
' Check the statusCode attribute of the response to make sure
' there hasn't been an error from QuickBooks.
'
Private Function CheckRsStatus(tagName As String) As Boolean
    CheckRsStatus = False
    Dim m As Integer
    For m = 0 To currNode.Attributes.length - 1
        Dim attrib
        Set attrib = currNode.Attributes.Item(m)
        If attrib.Name = "statusCode" And attrib.nodeValue <> "0" Then
            CheckRsStatus = True
            errorStatus = True
        End If
    
        If attrib.Name = "statusMessage" Then
            currStatusMsg = attrib.nodeValue
        End If
    Next
  
End Function





Private Function GetColID() As Integer
    
    Dim m As Integer
    
    ' Get attribute
    For m = 0 To currNode.Attributes.length - 1
        Dim attrib
        Set attrib = currNode.Attributes.Item(m)
        If m = 0 Then
            GetColID = CInt(attrib.nodeValue)
            Exit Function
        End If
    Next
    
    GetColID = 0
End Function




Private Function GetColDescType(colID As Integer, startPos As Integer) As ColDescType
    Dim it As Integer
    For it = startPos To currReportRet.colDescNum - 1
        If currReportRet.listColDesc(it).colID = colID Then
            GetColDescType = currReportRet.listColDesc(it)
            Exit Function
        End If
  
    Next
  
    ' if can't find
    Dim tColDesc As ColDescType
    tColDesc.colID = -1
    GetColDescType = tColDesc
End Function




Private Sub GetRowData(ByRef rowType As String, ByRef rowValue As String)
    rowType = ""
    rowValue = ""
  
    If currNode.nodeName <> "RowData" Then
        Exit Sub
    End If
  
    Dim attrib
    Dim m As Integer
    For m = 0 To currNode.Attributes.length - 1
        Set attrib = currNode.Attributes.Item(m)
        Select Case attrib.Name
            Case "rowType"
                rowType = attrib.nodeValue
            Case "value"
                rowValue = attrib.nodeValue
        End Select
    Next
End Sub




Private Sub GetColData(ByRef colID As Integer, ByRef colValue As String)
    colID = -1
    colValue = ""
  
    If currNode.nodeName <> "ColData" Then
        Exit Sub
    End If
  
    Dim attrib
    Dim m As Integer
    For m = 0 To currNode.Attributes.length - 1
        Set attrib = currNode.Attributes.Item(m)
        Select Case attrib.Name
            Case "colID"
                colID = CInt(attrib.nodeValue)
            Case "value"
                colValue = attrib.nodeValue
        End Select
    Next
End Sub




'
' Takes the data from the qbxml and formats it for display as either
' a date, amount, or percentage depending on the type.
'
Private Function FormatData(sStr As String, sColType As String) As String
  Select Case sColType
    Case "Date" ' Date formate conversion
      FormatData = FormatDate(sStr)
      Exit Function
    Case "BilledDate" ' Date formate conversion
      FormatData = FormatDate(sStr)
      Exit Function
    Case "DueDate" ' Date formate conversion
      FormatData = FormatDate(sStr)
      Exit Function
    Case "ReceiveDate" ' Date formate conversion
      FormatData = FormatDate(sStr)
      Exit Function
    Case "ShipDate" ' Date formate conversion
      FormatData = FormatDate(sStr)
      Exit Function
    Case "ThruDate" ' Date formate conversion
      FormatData = FormatDate(sStr)
      Exit Function
    Case "Quantity"
      FormatData = sStr
      Exit Function
    Case "Total"
      FormatData = FormatAmt(sStr)
      Exit Function
    Case "AmountDifference"
      FormatData = FormatAmt(sStr)
      Exit Function
    Case "Amount"
      FormatData = FormatAmt(sStr)
      Exit Function
    Case "OriginalAmount"
      FormatData = FormatAmt(sStr)
      Exit Function
    Case "PaidAmount"
      FormatData = FormatAmt(sStr)
      Exit Function
    Case "RunningBalance"
      FormatData = FormatAmt(sStr)
      Exit Function
    Case "Debit"
      FormatData = FormatAmt(sStr)
      Exit Function
    Case "IncomeSubjectToTax"
      FormatData = FormatAmt(sStr)
      Exit Function
    Case "OpenBalance"
      FormatData = FormatAmt(sStr)
      Exit Function
    Case "PaidAmount"
      FormatData = FormatAmt(sStr)
      Exit Function
    Case "UnitPrice"
      FormatData = FormatAmt(sStr)
      Exit Function
    Case "WageBase"
      FormatData = FormatAmt(sStr)
      Exit Function
    Case "WageBaseTips"
      FormatData = FormatAmt(sStr)
      Exit Function
    Case "Percent"
      FormatData = FormatUncalculatedPercent(sStr)
      Exit Function
    Case "PercentChange"
      FormatData = FormatUncalculatedPercent(sStr)
      Exit Function
  End Select
  
  FormatData = sStr
End Function



'
' Format data representing numeric amounts.  Add "," to the integer part
' and a dollar sign in the correct spot.
'
Private Function FormatAmt(sStr As String) As String
  If IsNumeric(sStr) Then
    Dim intPart As String
    Dim desPart As String
    Dim pos As Integer
    
    Dim signStr As String
    
    signStr = ""
    
    If Left(sStr, 1) = "-" Then
      signStr = "-"
    End If
    
    pos = InStr(sStr, ".")
    If pos > 0 Then
      intPart = Left(sStr, pos - 1)
      desPart = Right(sStr, Len(sStr) - pos)
      
      If signStr = "-" Then
        intPart = Right(intPart, Len(intPart) - 1)
      End If
    Else
      intPart = sStr
      desPart = "00"
    End If
    
    If Len(desPart) = 0 Then
      desPart = desPart & "00"
    ElseIf Len(desPart) = 1 Then
      desPart = desPart & "0"
    ElseIf Len(desPart) > 2 Then
      'don't chop off
      'desPart = Left(desPart, 2)
    End If
      
    Dim tStr As String
    Dim sAddCm As Boolean
    sAddCm = False
    While (Len(intPart) > 3)
      If sAddCm Then
        tStr = Right(intPart, 3) & "," & tStr
      Else
        tStr = Right(intPart, 3) & tStr
        sAddCm = True
      End If
      intPart = Left(intPart, Len(intPart) - 3)
    
    Wend
    If Len(intPart) > 0 Then
      If sAddCm Then
        tStr = intPart & "," & tStr
      Else
        tStr = intPart & tStr
       
      End If
    End If
      
    FormatAmt = "$" & signStr & tStr & "." & desPart
    Exit Function
  
  Else
    FormatAmt = sStr
    MsgBox "Not numeric value: " & sStr
    Exit Function
  End If
End Function


'
' Format date data, converting to qbXML date yyyy-mm-dd to
' QB format mm/dd/yyyy.
'
Private Function FormatDate(sStr As String) As String
  If sStr = "" Then
    FormatDate = sStr
    Exit Function
  End If
  
  If Len(sStr) <> 10 Then
    MsgBox "Wrong date format: " & sStr
    FormatDate = sStr
    Exit Function
  End If
  
  ' Convert SDK Date format yyyy-mm-dd to the QB Date format
  
    FormatDate = Mid(sStr, 6, 2) & "/" & Mid(sStr, 9, 2) & "/" & Mid(sStr, 1, 4)
    Exit Function

  
End Function



'
' Format Percentages and add "," to the integer part.
'
Private Function FormatUncalculatedPercent(sStr As String) As String
  If Not IsNumeric(sStr) Then
    MsgBox "not numeric value: " & sStr
    FormatUncalculatedPercent = sStr
    Exit Function
  End If
  
  Dim intPart As String
  Dim desPart As String
  Dim pos As Integer
  
  Dim signStr As String
  
  signStr = ""
  
  If Left(sStr, 1) = "-" Then
    signStr = "-"
  End If
  
  pos = InStr(sStr, ".")
  If pos > 0 Then
    intPart = Left(sStr, pos - 1)
    desPart = Right(sStr, Len(sStr) - pos)
    
    If signStr = "-" Then
      intPart = Right(intPart, Len(intPart) - 1)
    End If
  Else
    intPart = sStr
    desPart = "00"
  End If
  
  If Len(desPart) = 0 Then
    desPart = desPart & "00"
  ElseIf Len(desPart) = 1 Then
    desPart = desPart & "0"
  ElseIf Len(desPart) > 2 Then
    'don't chop off
    'desPart = Left(desPart, 2)
  End If
  
  intPart = intPart & Left(desPart, 2)
  desPart = Right(desPart, Len(desPart) - 2) 'shift 2 positions
  If Len(desPart) = 0 Then
    desPart = "0" ' looks like just 1 position in the percentage format
  End If
  
  ' Format the Integer Part
  Dim tStr As String
  tStr = FormatInt(intPart)
  
  'trim of leading "0"
  While Len(tStr) > 0 And Left(tStr, 1) = "0"
    tStr = Right(tStr, Len(tStr) - 1)
  Wend
  
  If tStr = "" Then
    tStr = "0"
  End If
  FormatUncalculatedPercent = signStr & tStr & "." & desPart & "%"
  
End Function



'
' Format integer data  For example, convert 99999999 to 99,999,999.
'
Private Function FormatInt(intPart As String) As String
  ' Format the Integer Part
  Dim tStr As String
  Dim sAddCm As Boolean
  sAddCm = False
  While (Len(intPart) > 3)
    If sAddCm Then
      tStr = Right(intPart, 3) & "," & tStr
    Else
      tStr = Right(intPart, 3) & tStr
      sAddCm = True
    End If
    intPart = Left(intPart, Len(intPart) - 3)
  
  Wend
  If Len(intPart) > 0 Then
    If sAddCm Then
      tStr = intPart & "," & tStr
    Else
      tStr = intPart & tStr
    End If
  End If
  
  FormatInt = tStr
End Function

