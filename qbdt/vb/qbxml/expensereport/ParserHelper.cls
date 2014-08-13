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
' It creates the .html table from the parsed response, including
' the calculated percentage columns available through the QuickBooks UI.
' Although certain small pieces of this module are particular to the
' type of report being created for this project, most of the code is
' generic and could easily be adapted to any other QuickBooks report.
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
Private outPut As String ' html output

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
    OutputStr = outPut
End Property



'
' Start the xml parsing.  After this function has returned, output
' will contain the html body of the report we are generating.
'
Public Function ParseReportResponseXML(xmlStr As String) As Boolean
  
    Dim ret  As Boolean
 
    ' Put initial tags on html string
    outPut = "<html>" & vbCrLf
    outPut = outPut & "<body bgcolor=""#ffffff"">" & vbCrLf
    outPut = outPut & "<table cellpadding=""2"" cellspacing=""2"" border=""1"">" _
        & vbCrLf & vbCrLf
  
    rowCount = 0 ' initialize the number of rows in the report
    errorStatus = False ' initialize the error status
    
    ret = ParseResponseXML(xmlStr)
  
    ' ret will be true if ParseResponseXML has succeeded with no error
    If Not ret Then
        MsgBox Err.Description, vbCritical, "Parse Response XML Error:"
        ParseReportResponseXML = False
        Exit Function
    End If
  
    ' Make sure we have displayed the correct number of rows
    ' in the report.
    If rowCount <> currReportRet.numRows Then
        MsgBox "Error: Actual num of rows: " & CStr(rowCount) & _
            "doesn't match NumOfRows " & CStr(currReportRet.numRows)
    End If
  
    ' Put final tags on html string
    outPut = outPut & "</table>" & vbCrLf
    outPut = outPut & "</body>" & vbCrLf
    outPut = outPut & "</html>" & vbCrLf
  
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
    
    ParseResponseXML = True
    Exit Function
  
LoadMSXML40ErrorHandler:
    MsgBox "Error:" + CStr(Hex(Err.Number)) + ", " + Err.Description, vbCritical, "Parse XML error: "
    ParseResponseXML = False
    Exit Function

MSXML40ParseErrorHandler:
    MsgBox "Error:" + Err.Description, vbCritical, "Error in Parsing Response"
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
  
    ' If the node we are looking at is the GeneralSummaryReportQueryRs
    ' node, we need to check the error status.
    If currTagName = "GeneralSummaryReportQueryRs" Then
        If CheckRsStatus(currTagName) Then
            MsgBox currStatusMsg, vbCritical, "Error in Response"
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
        
        rowCount = rowCount + 1
        ProcessRow currTagName
            
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
    
    outPut = outPut & "<tr>" & vbCrLf
    
    ' Handle the first column separately, this will be the list of vendors
    outPut = outPut & "<td align=""center""><b>Vendor</b></td>" & vbCrLf
    
    For it = 2 To currReportRet.colDescNum
        
        colDesc = GetColDescType(it - startPos, startPos)
        outPut = outPut & "<td align=""center""><b>" & colDesc.colTitle & "</b></td>" _
            & vbCrLf
            
        ' Need to add the calculated column headers after each with the
        ' exception of the first (vendor name) and last (total) columns
        If (it < currReportRet.colDescNum) Then
            outPut = outPut & "<td align=""center""><b>Percent of Row</b></td>" _
                & vbCrLf
        End If
      
        colPos = colPos + 1
      
        If colDesc.colID = currReportRet.numCols Then ' next desc line
            outPut = outPut & vbCrLf
            colPos = 0
            startPos = startPos + currReportRet.numCols
        End If
      
    Next
    
    outPut = outPut & "</tr>" & vbCrLf
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
                    Exit Sub
                End If
                GetColData colID, colValue
                listCol(colID) = colValue
        End Select
    Next
    
    ' Add the row to the html string
    outPut = outPut & "<tr>" & vbCrLf
    
    Dim calculation As Double ' calculated percentage values
    calculation = 0
    
    Dim rowTotal As Double
    rowTotal = CDbl(listCol(currReportRet.numCols))
    
    ' The first column will contain the vendor name
    outPut = outPut & "<td align=""left"">" & listCol(1) & "</td>" & vbCrLf
    
    ' The rest of the columns will contain numeric amounts
    For it = 2 To currReportRet.numCols - 1
        If listCol(it) <> "" Then
        
            outPut = outPut & "<td align=""right"">" & _
                FormatAmt(listCol(it)) & "</td>" & vbCrLf

            ' Calculate the special % column and display it in
            ' percentage format.  Be sure to avoid divide by 0.
            If 0 = rowTotal Then
                calculation = 0
            Else
                calculation = 100 * CDbl(listCol(it)) / rowTotal
            End If
                
            outPut = outPut & "<td align=""right"">" & _
                FormatCalculatedPercent(CStr(calculation)) & "</td>"
                
        End If
    Next
    
    ' Handle the last row separately since the total should not have
    ' a calculated percentage column.
    outPut = outPut & "<td align=""right"">" & _
        FormatAmt(listCol(currReportRet.numCols)) & "</td>" & vbCrLf
    outPut = outPut & "</tr>" & vbCrLf & vbCrLf
    
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



'
' Get the colID value of current ColDesc being parsed.
'
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




'
' Get the type (e.g. Amount) of the ColDesc currently being parsed.
'
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



'
' Get the data for a particular RowData element.
'
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
            Case "rowValue"
                rowValue = attrib.nodeValue
        End Select
    Next
End Sub




'
' Get data for a particular ColData element.
'
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
' Format Amount, add "," to the integer part and "$" in front
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
' Format Percentages and add "," to the integer part.  Note that
' the percentage should already be calculated, fractional values
' will be taken as a fraction of one percent.
'
Private Function FormatCalculatedPercent(sStr As String) As String

    If Not IsNumeric(sStr) Then
        Log "not numeric value: " & sStr
        FormatCalculatedPercent = sStr
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
    
    ' Want to make the decimal part two digits
    If Len(desPart) = 0 Then
        desPart = desPart & "00"
    ElseIf Len(desPart) = 1 Then
        desPart = desPart & "0"
    ElseIf Len(desPart) > 2 Then
        ' chop off if there are more than 2 decimal places
        desPart = Left(desPart, 2)
    End If
    
    ' Format the Integer Part
    Dim tStr As String
    tStr = FormatInt(intPart)
    
    'trim off leading "0"
    While Len(tStr) > 0 And Left(tStr, 1) = "0"
        tStr = Right(tStr, Len(tStr) - 1)
    Wend
    
    If tStr = "" Then
        tStr = "0"
    End If
    FormatCalculatedPercent = signStr & tStr & "." & desPart & "%"
    Exit Function

End Function


'
' Format Integers, for example convert 99999999 to 99,999,999.
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
