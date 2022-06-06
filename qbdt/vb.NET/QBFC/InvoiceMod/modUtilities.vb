Option Strict Off
Option Explicit On
Imports MSXML2

Module modUtilities
	'----------------------------------------------------------
	' Copyright © 2003-2013 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/rdmgr/?ID=100
	'
	'----------------------------------------------------------
	
	
	Dim stroldMessageSetID As String
	
	Public Function DateValid(ByRef strDate As String, ByRef strMsgText As String) As Boolean
		
		'Accept empty values as valid
		If Trim(strDate) = "" Then
			DateValid = True
			Exit Function
		End If
		
		'Check to see if the date string is a valid date string
		If Not IsDate(strDate) Then
			MsgBox("The " & strMsgText & " date you've supplied is illegal")
			DateValid = False
			Exit Function
		End If
		
		'Check to see if the date is in the valid date range
		If CDate(strDate) < CDate("1/1/1970") Or CDate(strDate) > CDate("1/18/2038") Then
			MsgBox("The " & strMsgText & " date you supply must be between Jan 1, 1970 and Jan 18, 2038")
			DateValid = False
			Exit Function
		End If
		
		'Check to see if the date is of the proper format
		Dim strSplits() As String
		strSplits = Split(Trim(strDate), "-")
		
		If UBound(strSplits) = 2 Then
			If Len(strSplits(0)) = 4 And Len(strSplits(1)) = 2 And Len(strSplits(2)) = 2 Then
				DateValid = True
				Exit Function
			End If
		End If
		
		MsgBox("The " & strMsgText & " date must be of the form yyyy-mm-dd")
		DateValid = False
		
	End Function
	
	
	Public Function TimeValid(ByRef strTime As String, ByRef strMsgText As String) As Boolean

        'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If String.IsNullOrEmpty(Trim(strTime)) Then
            TimeValid = True
            Exit Function
        End If

        Dim strSplits() As String
		strSplits = Split(Trim(strTime), ":")
		
		If UBound(strSplits) = 1 Then
			If IsNumeric(strSplits(0)) And IsNumeric(strSplits(1)) Then
				If Len(strSplits(1)) = 2 Then
					If CShort(strSplits(0)) >= 1 And CShort(strSplits(0)) <= 12 And CShort(strSplits(0)) >= 0 And CShort(strSplits(0)) <= 59 Then
						TimeValid = True
						Exit Function
					End If
				End If
			End If
		End If
		
		MsgBox("Invalid " & strMsgText & " time value")
		TimeValid = False
	End Function
	
	
	Public Function DateTimeString(ByRef strDate As String, ByRef strTime As String, ByRef booAM As Boolean, ByRef booSupportsDateTime As Boolean) As String
		
		Dim strSplits() As String

        'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If String.IsNullOrEmpty(Trim(strDate)) Then
            DateTimeString = ""
            Exit Function
        End If

        'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If String.IsNullOrEmpty(Trim(strTime)) Then
            DateTimeString = Trim(strDate)
            Exit Function
        End If

        strSplits = Split(Trim(strTime), ":")
		
		If booAM Then
			If strSplits(0) = "12" Then strSplits(0) = "00"
			If Len(strSplits(0)) = 1 Then strSplits(0) = "0" & strSplits(0)
		Else
			If CShort(strSplits(0)) < 12 Then strSplits(0) = Str(CShort(strSplits(0)) + 12)
		End If
		
		DateTimeString = Trim(strDate) & "T" & Trim(strSplits(0)) & ":" & strSplits(1)
	End Function


    Public Sub AddMSXMLNode(ByRef strNodeName As String, ByRef objDOMDocument As MSXML2.DOMDocument60, ByRef objParentNode As MSXML2.IXMLDOMNode, ByRef objNode As MSXML2.IXMLDOMNode)

        objNode = objDOMDocument.createElement(strNodeName)
        objParentNode.appendChild(objNode)
    End Sub


    Public Sub AddMSXMLAttribute(ByRef strAttributeName As String, ByRef strAttributeValue As String, ByRef objDOMDocument As DOMDocument60, ByRef objParentNode As IXMLDOMNode, ByRef objAttribute As IXMLDOMAttribute)

        objAttribute = objDOMDocument.createAttribute(strAttributeName)
        objAttribute.text = strAttributeValue
        'UPGRADE_WARNING: Couldn't resolve default property of object objAttribute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        objParentNode.attributes.setNamedItem(objAttribute)
    End Sub

    Public Sub AddMSXMLElement(ByRef strElementName As String, ByRef strElementValue As String, ByRef objDOMDocument As DOMDocument60, ByRef objParentNode As IXMLDOMNode, ByRef objElement As IXMLDOMElement)


        objElement = objDOMDocument.createElement(strElementName)
        objElement.text = strElementValue
        'UPGRADE_WARNING: Couldn't resolve default property of object objElement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        objParentNode.appendChild(objElement)
    End Sub


    Public Sub CreateStandardRequestNode(ByRef boonewMessageSetID As Boolean, ByRef stronError As String, ByRef objDOMDocument As DOMDocument60, ByRef objRootNode As MSXML2.IXMLDOMNode, ByRef objRequestNode As MSXML2.IXMLDOMNode, ByRef objAttribute As IXMLDOMAttribute)

        objRootNode = objDOMDocument.createElement("QBXML")
        objDOMDocument.appendChild(objRootNode)

        objRequestNode = objDOMDocument.createElement("QBXMLMsgsRq")
        objRootNode.appendChild(objRequestNode)

        If boonewMessageSetID Then
            stroldMessageSetID = NewMessageSetID()
            AddMSXMLAttribute("newMessageSetID", stroldMessageSetID, objDOMDocument, objRequestNode, objAttribute)
            stroldMessageSetID = CStr(Nothing)
        End If

        AddMSXMLAttribute("onError", stronError, objDOMDocument, objRequestNode, objAttribute)
    End Sub


    Public Sub GetTags(ByRef strInString As String, ByRef strStartTag As String, ByRef strEndTag As String, ByRef intTagLength As Short)
		
		'We expect the string passed in to be of the form:
		'
		' "<tag>...</tag>..."
		'
		'If the tag delimiters aren't here, there aren't any tags or if the first
		'part of the string isn't the start of a tag exit without returning tags
		If InStr(1, strInString, "<") = 0 Or InStr(1, strInString, ">") = 0 Or InStr(1, strInString, "</") = 0 Or Left(strInString, 1) <> "<" Then
			strStartTag = CStr(Nothing)
			strEndTag = CStr(Nothing)
			intTagLength = 0
			Exit Sub
		End If
		
		strStartTag = Left(strInString, InStr(1, strInString, ">"))
		intTagLength = Len(strStartTag)
		strEndTag = Replace(strStartTag, "<", "</")
		
		If InStr(1, strInString, strEndTag) = 0 Then
			strStartTag = CStr(Nothing)
			strEndTag = CStr(Nothing)
			intTagLength = 0
		End If
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
	Public Sub PrettyPrintXMLToFile(ByRef XMLString As String, ByRef XMLFile As String)
		
		Dim SplitXMLString() As String
		Dim IndentString As String
		Dim XMLStringLength As Integer
		Dim SplitIndex As Object
		Dim FileNum As Short

        'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If String.IsNullOrEmpty(Trim(XMLString)) Then Exit Sub

        IndentString = CStr(Nothing)

        FileNum = FreeFile()
        FileOpen(FileNum, XMLFile, OpenMode.Output)
		
		'Remove the linefeeds from the XML output string
		XMLString = Replace(XMLString, vbLf, vbNullString)
		
		SplitXMLString = Split(XMLString, "<")
		
		'We're expecting the first character of the XML output to be "<"
		'which result in an empty first array element, so skip it.
		'UPGRADE_WARNING: Couldn't resolve default property of object SplitIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		SplitIndex = LBound(SplitXMLString) + 1
		
		Do 
			'UPGRADE_WARNING: Couldn't resolve default property of object SplitIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Left(SplitXMLString(SplitIndex), 1) = "/" And Len(IndentString) > 2 Then
				IndentString = Left(IndentString, Len(IndentString) - 3)
				'UPGRADE_WARNING: Couldn't resolve default property of object SplitIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PrintLine(FileNum, IndentString & "<" & SplitXMLString(SplitIndex))
				'UPGRADE_WARNING: Couldn't resolve default property of object SplitIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				SplitIndex = SplitIndex + 1
				'UPGRADE_WARNING: Couldn't resolve default property of object SplitIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf Left(SplitXMLString(SplitIndex + 1), 1) = "/" Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object SplitIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If InStr(1, Left(SplitXMLString(SplitIndex), InStr(1, SplitXMLString(SplitIndex), ">")), " ") > 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object SplitIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					PrintLine(FileNum, IndentString & "<" & SplitXMLString(SplitIndex))
					'UPGRADE_WARNING: Couldn't resolve default property of object SplitIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					SplitIndex = SplitIndex + 1
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object SplitIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					PrintLine(FileNum, IndentString & "<" & SplitXMLString(SplitIndex) & "<" & SplitXMLString(SplitIndex + 1))
					'UPGRADE_WARNING: Couldn't resolve default property of object SplitIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					SplitIndex = SplitIndex + 2
				End If
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object SplitIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				PrintLine(FileNum, IndentString & "<" & SplitXMLString(SplitIndex))
				'UPGRADE_WARNING: Couldn't resolve default property of object SplitIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If SplitXMLString(SplitIndex) <> "?xml version=""1.0"" ?>" And SplitXMLString(SplitIndex) <> "!DOCTYPE QBXML PUBLIC '-//INTUIT//DTD QBXML QBD 1.0//EN' >" And InStr(1, SplitXMLString(SplitIndex), "qbxml version") = 0 Then
					IndentString = IndentString & "   "
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object SplitIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				SplitIndex = SplitIndex + 1
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object SplitIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Loop Until SplitIndex >= UBound(SplitXMLString)
		
		If Left(SplitXMLString(UBound(SplitXMLString)), 1) = "/" Then
			If Len(IndentString) >= 3 Then
				IndentString = Left(IndentString, Len(IndentString) - 3)
			End If
		End If
		
		PrintLine(FileNum, IndentString & "<" & SplitXMLString(UBound(SplitXMLString)))
		
		FileClose(FileNum)
	End Sub
	
	
	'----------------------------------------------------------
	'
	' Routine: PrettyPrintXMLToString
	'
	' Description
	'
	' Takes an XML message set string  as input.
	' Creates a new copy of the string and pretty prints the XML
	' message set into the string.
	'
	'----------------------------------------------------------
	Public Function PrettyPrintXMLToString(ByRef strInXMLString As String) As String
		
		Dim SplitXMLString() As String
		Dim IndentString As String
		Dim XMLStringLength As Integer
		Dim SplitIndex As Object
		Dim FileNum As Short
		
		Dim XMLString As String
		Dim strOut As String
		
		XMLString = strInXMLString
		strOut = CStr(Nothing)

        'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If String.IsNullOrEmpty(Trim(XMLString)) Then
            PrettyPrintXMLToString = CStr(Nothing)
            Exit Function
        End If

        IndentString = CStr(Nothing)
		
		'Remove the linefeeds from the XML output string
		XMLString = Replace(XMLString, vbLf, vbNullString)
		
		SplitXMLString = Split(XMLString, "<")
		
		'We're expecting the first character of the XML output to be "<"
		'which result in an empty first array element, so skip it.
		'UPGRADE_WARNING: Couldn't resolve default property of object SplitIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		SplitIndex = LBound(SplitXMLString) + 1
		
		Do 
			'UPGRADE_WARNING: Couldn't resolve default property of object SplitIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Left(SplitXMLString(SplitIndex), 1) = "/" And Len(IndentString) > 2 Then
				IndentString = Left(IndentString, Len(IndentString) - 3)
				'UPGRADE_WARNING: Couldn't resolve default property of object SplitIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				strOut = strOut & IndentString & "<" & SplitXMLString(SplitIndex) & vbCrLf
				'UPGRADE_WARNING: Couldn't resolve default property of object SplitIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				SplitIndex = SplitIndex + 1
				'UPGRADE_WARNING: Couldn't resolve default property of object SplitIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf Left(SplitXMLString(SplitIndex + 1), 1) = "/" Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object SplitIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If InStr(1, Left(SplitXMLString(SplitIndex), InStr(1, SplitXMLString(SplitIndex), ">")), " ") > 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object SplitIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					strOut = strOut & IndentString & "<" & SplitXMLString(SplitIndex) & vbCrLf
					'UPGRADE_WARNING: Couldn't resolve default property of object SplitIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					SplitIndex = SplitIndex + 1
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object SplitIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					strOut = strOut & IndentString & "<" & SplitXMLString(SplitIndex) & "<" & SplitXMLString(SplitIndex + 1) & vbCrLf
					'UPGRADE_WARNING: Couldn't resolve default property of object SplitIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					SplitIndex = SplitIndex + 2
				End If
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object SplitIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				strOut = strOut & IndentString & "<" & SplitXMLString(SplitIndex) & vbCrLf
				'UPGRADE_WARNING: Couldn't resolve default property of object SplitIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If SplitXMLString(SplitIndex) <> "?xml version=""1.0"" ?>" And SplitXMLString(SplitIndex) <> "!DOCTYPE QBXML PUBLIC '-//INTUIT//DTD QBXML QBD 1.0//EN' >" And InStr(1, SplitXMLString(SplitIndex), "qbxml version") = 0 Then
					IndentString = IndentString & "   "
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object SplitIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				SplitIndex = SplitIndex + 1
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object SplitIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Loop Until SplitIndex >= UBound(SplitXMLString)
		
		If Left(SplitXMLString(UBound(SplitXMLString)), 1) = "/" Then
			If Len(IndentString) >= 3 Then
				IndentString = Left(IndentString, Len(IndentString) - 3)
			End If
		End If
		
		strOut = strOut & IndentString & "<" & SplitXMLString(UBound(SplitXMLString))
		
		PrettyPrintXMLToString = strOut
	End Function
	
	
	Private Function NewMessageSetID() As String
		'Because this is a sample program that isn't expected to run on multiple
		'machines and isn't expected to produce more than one modify request per
		'second, we can simply use the date time string for our message set ID.
		'If we were expecting multiple machines we'd add some sort of system
		'identifier.  If we were expecting more than one request message set per
		'second we'd use a smaller time sequence.
		NewMessageSetID = VB6.Format(Now, "XyyyymmddThhmmss")
	End Function
End Module