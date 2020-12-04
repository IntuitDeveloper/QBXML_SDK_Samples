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
    ' Created Sept, 2002
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
         listColDesc() As ParserHelper.ColDescType
         colDescNum As Integer
    End Type
    
    ' Local Variables
    Private outPut As String ' html output
    
    Private rowCount As Integer ' number of rows in report
    Private currReportRet As ParserHelper.ReportRetType ' current ReportRet

    Private Const NUMOFCOLS = 1000

    '
    ' Get output html String once it has been created.
    '
    Public Property Get OutputStr() As String
        OutputStr = outPut
    End Property

    '
    ' Start the response parsing.  After this function has returned, output
    ' will contain the html body of the report we are generating.
    '
    Public Function ParseReportResponseSet(ByRef responseSet As IMsgSetResponse) As Boolean

        Dim bSuccess As Boolean

        ' Put initial tags on html string
        outPut = "<html>" & vbCrLf
        outPut = outPut & "<body bgcolor=""#ffffff"">" & vbCrLf
        outPut = outPut & "<table cellpadding=""2"" cellspacing=""2"" border=""1"">" & vbCrLf & vbCrLf

        rowCount = 0 ' initialize the number of rows in the report

        bSuccess = ParseResponseSet(responseSet)

        ' ret will be true if ParseResponseSet has succeeded with no error
        If Not bSuccess Then
            MsgBox Err.Description, vbOKOnly, "Parse Response Set Error:"
            ParseReportResponseSet = False
            Exit Function
        End If

        ' Make sure we have displayed the correct number of rows
        ' in the report.
        If rowCount <> currReportRet.numRows Then
            MsgBox ("Error: Actual num of rows: " & CStr(rowCount) & "doesn't match NumOfRows " & CStr(currReportRet.numRows))
        End If

        ' Put final tags on html string
        outPut = outPut & "</table>" & vbCrLf
        outPut = outPut & "</body>" & vbCrLf
        outPut = outPut & "</html>" & vbCrLf

        ParseReportResponseSet = True

    End Function




    '
    ' Call AccessNode to start parsing the object.
    '
    Private Function ParseResponseSet(ByRef responseSet As IMsgSetResponse) _
        As Boolean
        On Error GoTo ErrorHandler

        ' check to make sure we have objects to access first
        ' and that there are responses in the list
        If (responseSet Is Nothing) Then
            ParseResponseSet = False
            Exit Function
        End If
        If (responseSet.responseList Is Nothing) Then
            ParseResponseSet = False
            Exit Function
        End If
        If (responseSet.responseList.count <= 0) Then
            ParseResponseSet = False
            Exit Function
        End If

        ' we need to check the error status.
        ' there will not be a status message in the QBXMLMsgsRs
        ' but the following code would illustrate if we did
        If CheckRsStatus(responseSet) Then
            ParseResponseSet = False
            Exit Function
        End If

        ' Start parsing the response list
        Dim responseList As IResponseList
        Set responseList = responseSet.responseList

        Dim count As Integer
        count = responseList.count
        Dim index As Integer
        ' go thru each response and process the response.
        ' this example will only have one response in the list:  GeneralSummaryReportQueryRs
        ' but this loop will show how to get all responses in the case that more than one
        ' request was sent.
        For index = 0 To (count - 1)
            Dim response As IResponse
            Set response = responseList.GetAt(index)
            If (Not response Is Nothing) Then
                If response.StatusSeverity <> "Error" Then
                    ParseGeneralSummaryReportQueryRs response
                Else
                    MsgBox "Status Code: " & CStr(response.StatusCode) & ", " & response.StatusMessage, vbOKOnly, "ParseGeneralSummaryReportQueryRs error: "
                End If
            End If
        Next

        ParseResponseSet = True
        Exit Function

ErrorHandler:
        MsgBox "Error:" & CStr(Hex(Err.Number)) & ", " & Err.Description, vbOKOnly, "Parse error: "
        ParseResponseSet = False
        Exit Function

    End Function

    '
    ' go thru response gathering data and creating the table
    ' for display.
    '
    Private Sub ParseGeneralSummaryReportQueryRs(ByRef response As IResponse)
        On Error GoTo AccErrorhandler

        ' first make sure we have a response object to handle
        If (response Is Nothing) Then
            Exit Sub
        End If
        If (response.Type Is Nothing) Or _
            (response.Detail Is Nothing) Then
            Exit Sub
        End If
        If (response.Detail.Type Is Nothing) Then
            Exit Sub
        End If

        ' make sure we are processing the GeneralSummaryReportQueryRs and
        ' the ReportRet responses in this response list
        Dim reportRet As IReportRet
        Dim responseType As ENResponseType
        Dim responseDetailType As ENObjectType
        responseType = response.Type.GetValue()
        responseDetailType = response.Detail.Type.GetValue()
        If (responseType = ENResponseType.rtGeneralSummaryReportQueryRs) And _
            (responseDetailType = ENObjectType.otReportRet) Then
            ' save the response detail in the appropriate object type
            ' since we have first verified the type of the response object
            Set reportRet = response.Detail
        Else
            ' bail, we do not have the responses we were expecting
            Exit Sub
        End If

        ' we need to initialize the current ReportRet variable (currReportRet).
        InitCurrReportRet

        ' Get Report level info
        If (Not reportRet.ReportTitle Is Nothing) Then
            currReportRet.title = reportRet.ReportTitle.GetValue()
        End If
        If (Not reportRet.reportBasis Is Nothing) Then
            currReportRet.reportBasis = reportRet.reportBasis.GetValue()
        End If
        If (Not reportRet.numRows Is Nothing) Then
            currReportRet.numRows = reportRet.numRows.GetValue()
        End If
        If (Not reportRet.NumColumns Is Nothing) Then
            currReportRet.numCols = reportRet.NumColumns.GetValue()
        End If

        Dim count As Integer
        Dim index As Integer

        'Get Column Descriptions
        Dim colDesc As IColDesc
        If (Not reportRet.ColDescList Is Nothing) Then
            count = reportRet.ColDescList.count
            For index = 0 To count - 1
                Set colDesc = reportRet.ColDescList.GetAt(index)
                If (Not colDesc Is Nothing) Then
                    ProcessColDesc colDesc
                End If
            Next
        End If

        ' end of Head session, to print the header
        PrintColDesc

        ' Get Report Data
        Dim orReportData As IORReportData
        If (Not reportRet.ReportData Is Nothing) Then
            If (Not reportRet.ReportData.ORReportDataList Is Nothing) Then
                count = reportRet.ReportData.ORReportDataList.count
                For index = 0 To count - 1
                    Set orReportData = reportRet.ReportData.ORReportDataList.GetAt(index)
                    If (Not orReportData Is Nothing) Then
                        rowCount = rowCount + 1
                        ProcessRow orReportData
                    End If
                Next
            End If
        End If
        Exit Sub

AccErrorhandler:
        MsgBox "Error:" & CStr(Hex(Err.Number)) & ", " & Err.Description, vbOKOnly, "Error in Parsing Response"
    End Sub

    '
    ' Parse the ColDesc elements.  This data will be used later when
    ' we are printing out the column headers.
    '
    Private Sub ProcessColDesc(ByRef colDesc As IColDesc)
        On Error GoTo ProcErrorHandler

        If (Not colDesc.colID Is Nothing) Then
            currReportRet.listColDesc(currReportRet.colDescNum).colID = colDesc.colID.GetValue()
        End If
        If (Not colDesc.ColTitleList Is Nothing) Then
            If (colDesc.ColTitleList.count >= 1) Then
                Dim colTitle As IColTitle
                Set colTitle = colDesc.ColTitleList.GetAt(0)
                If (Not colTitle Is Nothing) Then
                    If (Not colTitle.titleRow Is Nothing) And _
                        (Not colTitle.Value Is Nothing) Then
                        currReportRet.listColDesc(currReportRet.colDescNum).colTitle = colTitle.Value.GetValue()
                    End If
                End If
            End If
        End If

        If (Not colDesc.colType Is Nothing) Then
            currReportRet.listColDesc(currReportRet.colDescNum).colType = colDesc.colType.GetAsString()
        End If
        currReportRet.colDescNum = currReportRet.colDescNum + 1

        Exit Sub

ProcErrorHandler:
        MsgBox "Error:" & CStr(Hex(Err.Number)) & ", " & Err.Description, vbOKOnly, "Error in ProcessColDesc"
    End Sub

    '
    ' Print a row of column headers, which have already been parsed
    ' from the response and stored, to the html report string.
    '
    Private Sub PrintColDesc()

        Dim it As Integer
        Dim colPos As Integer
        Dim startPos As Integer
        colPos = 0
        startPos = 0

        Dim colDesc As ParserHelper.ColDescType

        outPut = outPut & "<tr>" & vbCrLf

        ' Handle the first column separately, this will be the list of vendors
        outPut = outPut & "<td align=""center""><b>Vendor</b></td>" & vbCrLf

        For it = 2 To currReportRet.colDescNum

            colDesc = GetColDescType(it - startPos, startPos)
            outPut = outPut & "<td align=""center""><b>" & colDesc.colTitle & "</b></td>" & vbCrLf

            ' Need to add the calculated column headers after each with the
            ' exception of the first (vendor name) and last (total) columns
            If (it < currReportRet.colDescNum) Then
                outPut = outPut & "<td align=""center""><b>Percent of Row</b></td>" & vbCrLf
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
    Private Sub ProcessRow(ByRef orReportData As IORReportData)
        On Error GoTo ProcErrorHandler

        Dim rowNum As String
        Dim count As Integer
        Dim index As Integer
        Dim colDataList As IColDataList
        Dim colData As IColData
        Dim colID As Integer
        Dim colValue As String
        Dim listCol() As String
        Dim it As Integer

        ' first make sure we have an object to process
        If (orReportData Is Nothing) Then
            Exit Sub
        End If

        ' we will process two types of the orReportData, SubTotalRow and DataRow
        Select Case orReportData.ortype
            Case ENORReportData.orrdSubtotalRow
                ' what to do for the subtotal row
                If (orReportData.SubtotalRow Is Nothing) Then
                    Exit Sub
                End If
                If (orReportData.SubtotalRow.rowNumber Is Nothing) Then
                    Exit Sub
                End If
                rowNum = orReportData.SubtotalRow.rowNumber.GetValue()
                ' Validate rowNum
                If Val(rowNum) > currReportRet.numRows Or Val(rowNum) < 0 Then
                    MsgBox ("Invalid rowNum: " & rowNum)
                End If

                ReDim listCol(currReportRet.numCols + 1)

                For it = 0 To currReportRet.numCols
                    listCol(it) = ""
                Next

                If (orReportData.SubtotalRow.colDataList Is Nothing) Then
                    Exit Sub
                End If
                Set colDataList = orReportData.SubtotalRow.colDataList
                count = colDataList.count
                For index = 0 To count - 1
                    Set colData = colDataList.GetAt(index)
                    If (Not colData Is Nothing) Then
                        If (Not colData.colID Is Nothing) And _
                            (Not colData.colValue Is Nothing) Then
                            colID = colData.colID.GetValue()
                            colValue = colData.colValue.GetValue()
                            If colID > currReportRet.numCols Then
                                MsgBox "current colID value " & CStr(colID) & " is large than NumOfCol " & CStr(currReportRet.numCols), vbOKOnly, "Wrong Col ID"
                                Exit Sub
                            End If
                            listCol(colID) = colValue
                        End If
                    End If
                Next
            Case ENORReportData.orrdDataRow
                If (orReportData.DataRow Is Nothing) Then
                    Exit Sub
                End If
                If (orReportData.DataRow.rowNumber Is Nothing) Then
                    Exit Sub
                End If

                'save the row number
                rowNum = orReportData.DataRow.rowNumber.GetValue()
                ' Validate rowNum
                If Val(rowNum) > currReportRet.numRows Or Val(rowNum) < 0 Then
                    MsgBox ("Invalid rowNum: " & rowNum)
                End If

                ReDim listCol(currReportRet.numCols + 1)

                For it = 0 To currReportRet.numCols
                    listCol(it) = ""
                Next

                If (orReportData.DataRow.RowData Is Nothing) Then
                    Exit Sub
                End If
                If (orReportData.DataRow.colDataList Is Nothing) Then
                    Exit Sub
                End If
                Set colDataList = orReportData.DataRow.colDataList
                count = colDataList.count
                For index = 0 To count - 1
                    Set colData = colDataList.GetAt(index)
                    If (Not colData Is Nothing) Then
                        If (Not colData.colID Is Nothing) And _
                            (Not colData.Value Is Nothing) Then
                            ' save the ColID and ColValue
                            colID = colData.colID.GetValue()
                            colValue = colData.Value.GetValue()
                            If colID > currReportRet.numCols Then
                                MsgBox "current colID value " & CStr(colID) & " is large than NumOfCol " & CStr(currReportRet.numCols), vbOKOnly, "Wrong Col ID"
                                Exit Sub
                            End If
                            listCol(colID) = colValue
                        End If
                    End If
                Next
            Case ENORReportData.orrdTotalRow
                If (orReportData.TotalRow Is Nothing) Then
                    Exit Sub
                End If
                If (orReportData.TotalRow.rowNumber Is Nothing) Then
                    Exit Sub
                End If
                rowNum = orReportData.TotalRow.rowNumber.GetValue()
                ' Validate rowNum
                If Val(rowNum) > currReportRet.numRows Or Val(rowNum) < 0 Then
                    MsgBox ("Invalid rowNum: " & rowNum)
                End If

                ReDim listCol(currReportRet.numCols + 1)

                For it = 0 To currReportRet.numCols
                    listCol(it) = ""
                Next

                If (orReportData.TotalRow.colDataList Is Nothing) Then
                    Exit Sub
                End If
                Set colDataList = orReportData.TotalRow.colDataList
                count = colDataList.count
                For index = 0 To count - 1
                    Set colData = colDataList.GetAt(index)
                    If (Not colData Is Nothing) Then
                        If (Not colData.colID Is Nothing) And _
                            (Not colData.Value Is Nothing) Then
                            colID = colData.colID.GetValue()
                            colValue = colData.Value.GetValue()
                            If colID > currReportRet.numCols Then
                                MsgBox "current colID value " & CStr(colID) & " is large than NumOfCol " & CStr(currReportRet.numCols), vbOKOnly, "Wrong Col ID"
                                Exit Sub
                            End If
                            listCol(colID) = colValue
                        End If
                    End If
                Next
        End Select

        Dim calculation As Double
        Dim rowTotal As Double ' calculated percentage values

        ' Add the row to the html string
        outPut = outPut & "<tr>" & vbCrLf

        calculation = 0

        rowTotal = CDbl(listCol(currReportRet.numCols))

        ' The first column will contain the vendor name
        outPut = outPut & "<td align=""left"">" & listCol(1) & "</td>" & vbCrLf

        ' The rest of the columns will contain numeric amounts
        For it = 2 To currReportRet.numCols - 1
            If listCol(it) <> "" Then

                outPut = outPut & "<td align=""right"">" & FormatAmt(listCol(it)) & "</td>" & vbCrLf

                ' Calculate the special % column and display it in
                ' percentage format.  Be sure to avoid divide by 0.
                If 0 = rowTotal Then
                    calculation = 0
                Else
                    calculation = 100 * CDbl(listCol(it)) / rowTotal
                End If

                outPut = outPut & "<td align=""right"">" & FormatCalculatedPercent(CStr(calculation)) & "</td>"

            End If
        Next

        ' Handle the last row separately since the total should not have
        ' a calculated percentage column.
        outPut = outPut & "<td align=""right"">" & FormatAmt(listCol(currReportRet.numCols)) & "</td>" & vbCrLf
        outPut = outPut & "</tr>" & vbCrLf & vbCrLf

        ReDim listCol(0)
        Exit Sub

ProcErrorHandler:
        MsgBox "Error:" & CStr(Hex(Err.Number)) & ", " & Err.Description, vbOKOnly, "Error in ProcessRow"
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
    Private Function CheckRsStatus(ByRef responseSet As IMsgSetResponse) As Boolean
        On Error GoTo ErrorHandler
        CheckRsStatus = False

        Dim currStatusMsg As String
        Dim attrib As IAttributesRsSet
        Set attrib = responseSet.Attributes()
        If (Not attrib Is Nothing) Then
            currStatusMsg = attrib.MessageSetStatusCode()
            If currStatusMsg <> "0" Then
                MsgBox currStatusMsg, vbOKOnly, "Message Set Status Code Error in Response"
                CheckRsStatus = True
            End If
        End If
        Exit Function

ErrorHandler:
        MsgBox "Error:" & CStr(Hex(Err.Number)) & ", " & Err.Description, vbOKOnly, "Error in CheckRsStatus"
    End Function

    '
    ' Get the type (e.g. Amount) of the ColDesc currently being parsed.
    '
    Private Function GetColDescType(ByRef colID As Integer, ByRef startPos As Integer) As ParserHelper.ColDescType
        Dim it As Integer
        For it = startPos To currReportRet.colDescNum - 1
            If currReportRet.listColDesc(it).colID = colID Then
                GetColDescType = currReportRet.listColDesc(it)
                Exit Function
            End If

        Next

        ' if can't find
        Dim tColDesc As ParserHelper.ColDescType
        tColDesc.colID = -1
        GetColDescType = tColDesc
    End Function

    '
    ' Format Amount, add "," to the integer part and "$" in front
    '
    Private Function FormatAmt(ByRef sStr As String) As String

        Dim intPart As String
        Dim desPart As String
        Dim pos As Integer
        Dim signStr As String
        Dim tStr As String
        Dim sAddCm As Boolean
        If IsNumeric(sStr) Then

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
            MsgBox ("Not numeric value: " & sStr)
            Exit Function
        End If
    End Function

    '
    ' Format Percentages and add "," to the integer part.  Note that
    ' the percentage should already be calculated, fractional values
    ' will be taken as a fraction of one percent.
    '
    Private Function FormatCalculatedPercent(ByRef sStr As String) As String

        If Not IsNumeric(sStr) Then
            System.Math.Log (CDbl("not numeric value: " & sStr))
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
    Private Function FormatInt(ByRef intPart As String) As String
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


